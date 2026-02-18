#include "FormulaEngine.h"
#include "Spreadsheet.h"
#include <cmath>
#include <algorithm>
#include <cctype>
#include <limits>
#include <QDateTime>
#include <QRegularExpression>

FormulaEngine::FormulaEngine(Spreadsheet* spreadsheet)
    : m_spreadsheet(spreadsheet) {
}

void FormulaEngine::setSpreadsheet(Spreadsheet* spreadsheet) {
    m_spreadsheet = spreadsheet;
}

QVariant FormulaEngine::evaluate(const QString& formula) {
    m_lastError.clear();
    m_lastDependencies.clear();

    if (formula.isEmpty()) return QVariant();

    QString expr = formula.startsWith('=') ? formula.mid(1) : formula;

    try {
        return parseExpression(expr);
    } catch (const std::exception& e) {
        m_lastError = QString::fromStdString(e.what());
        return QVariant("#ERROR!");
    }
}

QString FormulaEngine::getLastError() const { return m_lastError; }
bool FormulaEngine::hasError() const { return !m_lastError.isEmpty(); }
void FormulaEngine::clearCache() { m_cache.clear(); }
void FormulaEngine::invalidateCell(const CellAddress& addr) { m_cache.erase(addr.toString().toStdString()); }

void FormulaEngine::skipWhitespace(const QString& expr, int& pos) {
    while (pos < expr.length() && expr[pos].isSpace()) pos++;
}

std::vector<QVariant> FormulaEngine::flattenArgs(const std::vector<QVariant>& args) {
    std::vector<QVariant> flat;
    for (const auto& arg : args) {
        if (arg.canConvert<std::vector<QVariant>>()) {
            auto nested = arg.value<std::vector<QVariant>>();
            for (const auto& v : nested) flat.push_back(v);
        } else {
            flat.push_back(arg);
        }
    }
    return flat;
}

QVariant FormulaEngine::parseExpression(const QString& expr) {
    int pos = 0;
    QVariant result = evaluateComparison(expr, pos);
    skipWhitespace(expr, pos);
    return result;
}

QVariant FormulaEngine::evaluateComparison(const QString& expr, int& pos) {
    QVariant left = evaluateTerm(expr, pos);
    skipWhitespace(expr, pos);
    while (pos < expr.length()) {
        if (pos + 1 < expr.length() && expr[pos] == '<' && expr[pos + 1] == '>') {
            pos += 2; QVariant right = evaluateTerm(expr, pos);
            left = QVariant(toNumber(left) != toNumber(right));
        } else if (pos + 1 < expr.length() && expr[pos] == '<' && expr[pos + 1] == '=') {
            pos += 2; QVariant right = evaluateTerm(expr, pos);
            left = QVariant(toNumber(left) <= toNumber(right));
        } else if (pos + 1 < expr.length() && expr[pos] == '>' && expr[pos + 1] == '=') {
            pos += 2; QVariant right = evaluateTerm(expr, pos);
            left = QVariant(toNumber(left) >= toNumber(right));
        } else if (expr[pos] == '<') {
            pos++; QVariant right = evaluateTerm(expr, pos);
            left = QVariant(toNumber(left) < toNumber(right));
        } else if (expr[pos] == '>') {
            pos++; QVariant right = evaluateTerm(expr, pos);
            left = QVariant(toNumber(left) > toNumber(right));
        } else if (expr[pos] == '=') {
            pos++; QVariant right = evaluateTerm(expr, pos);
            left = QVariant(toNumber(left) == toNumber(right));
        } else break;
        skipWhitespace(expr, pos);
    }
    return left;
}

QVariant FormulaEngine::evaluateTerm(const QString& expr, int& pos) {
    QVariant result = evaluateMultiplicative(expr, pos);
    skipWhitespace(expr, pos);
    while (pos < expr.length()) {
        if (expr[pos] == '+') { pos++; result = toNumber(result) + toNumber(evaluateMultiplicative(expr, pos)); }
        else if (expr[pos] == '-') { pos++; result = toNumber(result) - toNumber(evaluateMultiplicative(expr, pos)); }
        else break;
        skipWhitespace(expr, pos);
    }
    return result;
}

QVariant FormulaEngine::evaluateMultiplicative(const QString& expr, int& pos) {
    QVariant result = evaluateUnary(expr, pos);
    skipWhitespace(expr, pos);
    while (pos < expr.length()) {
        if (expr[pos] == '*') { pos++; result = toNumber(result) * toNumber(evaluateUnary(expr, pos)); }
        else if (expr[pos] == '/') {
            pos++; double d = toNumber(evaluateUnary(expr, pos));
            if (d == 0.0) { m_lastError = "Division by zero"; return QVariant("#DIV/0!"); }
            result = toNumber(result) / d;
        } else break;
        skipWhitespace(expr, pos);
    }
    return result;
}

QVariant FormulaEngine::evaluateUnary(const QString& expr, int& pos) {
    skipWhitespace(expr, pos);
    if (pos < expr.length() && expr[pos] == '-') { pos++; return -toNumber(evaluateUnary(expr, pos)); }
    return evaluatePower(expr, pos);
}

QVariant FormulaEngine::evaluatePower(const QString& expr, int& pos) {
    QVariant base = evaluateFactor(expr, pos);
    skipWhitespace(expr, pos);
    if (pos < expr.length() && expr[pos] == '^') {
        pos++; return std::pow(toNumber(base), toNumber(evaluateUnary(expr, pos)));
    }
    return base;
}

QVariant FormulaEngine::evaluateFactor(const QString& expr, int& pos) {
    skipWhitespace(expr, pos);
    if (pos >= expr.length()) return QVariant();

    // Numbers
    if (expr[pos].isDigit() || (expr[pos] == '.' && pos + 1 < expr.length() && expr[pos + 1].isDigit())) {
        int start = pos;
        while (pos < expr.length() && (expr[pos].isDigit() || expr[pos] == '.')) pos++;
        return expr.mid(start, pos - start).toDouble();
    }

    // Strings
    if (expr[pos] == '"') {
        pos++; int start = pos;
        while (pos < expr.length() && expr[pos] != '"') pos++;
        QString result = expr.mid(start, pos - start);
        if (pos < expr.length()) pos++;
        return result;
    }

    // Letter tokens: functions, cell refs, ranges
    if (pos < expr.length() && (expr[pos].isLetter() || expr[pos] == '$')) {
        int start = pos;
        while (pos < expr.length() && (expr[pos].isLetterOrNumber() || expr[pos] == ':' || expr[pos] == '$' || expr[pos] == '_')) pos++;
        QString token = expr.mid(start, pos - start);
        skipWhitespace(expr, pos);

        // Function call
        if (pos < expr.length() && expr[pos] == '(') {
            pos++;
            std::vector<QVariant> args;
            skipWhitespace(expr, pos);
            while (pos < expr.length() && expr[pos] != ')') {
                args.push_back(evaluateComparison(expr, pos));
                skipWhitespace(expr, pos);
                if (pos < expr.length() && expr[pos] == ',') pos++;
                skipWhitespace(expr, pos);
            }
            if (pos < expr.length() && expr[pos] == ')') pos++;
            return evaluateFunction(token.toUpper(), args);
        }

        // Range
        if (token.contains(':')) {
            CellRange range(token);
            auto cells = range.getCells();
            for (const auto& c : cells) m_lastDependencies.push_back(c);
            std::vector<QVariant> values = getRangeValues(range);
            return QVariant::fromValue(values);
        }

        // Cell ref
        CellAddress addr = CellAddress::fromString(token);
        m_lastDependencies.push_back(addr);
        return getCellValue(addr);
    }

    // Parentheses
    if (pos < expr.length() && expr[pos] == '(') {
        pos++;
        QVariant result = evaluateComparison(expr, pos);
        skipWhitespace(expr, pos);
        if (pos < expr.length() && expr[pos] == ')') pos++;
        return result;
    }

    return QVariant();
}

QVariant FormulaEngine::evaluateFunction(const QString& fn, const std::vector<QVariant>& args) {
    // Aggregate
    if (fn == "SUM") return funcSUM(args);
    if (fn == "AVERAGE") return funcAVERAGE(args);
    if (fn == "COUNT") return funcCOUNT(args);
    if (fn == "COUNTA") return funcCOUNTA(args);
    if (fn == "MIN") return funcMIN(args);
    if (fn == "MAX") return funcMAX(args);
    if (fn == "IF") return funcIF(args);
    if (fn == "CONCAT" || fn == "CONCATENATE") return funcCONCAT(args);
    if (fn == "LEN") return funcLEN(args);
    if (fn == "UPPER") return funcUPPER(args);
    if (fn == "LOWER") return funcLOWER(args);
    if (fn == "TRIM") return funcTRIM(args);
    // Math
    if (fn == "ROUND") return funcROUND(args);
    if (fn == "ABS") return funcABS(args);
    if (fn == "SQRT") return funcSQRT(args);
    if (fn == "POWER") return funcPOWER(args);
    if (fn == "MOD") return funcMOD(args);
    if (fn == "INT") return funcINT(args);
    if (fn == "CEILING") return funcCEILING(args);
    if (fn == "FLOOR") return funcFLOOR(args);
    // Logical
    if (fn == "AND") return funcAND(args);
    if (fn == "OR") return funcOR(args);
    if (fn == "NOT") return funcNOT(args);
    if (fn == "IFERROR") return funcIFERROR(args);
    // Text
    if (fn == "LEFT") return funcLEFT(args);
    if (fn == "RIGHT") return funcRIGHT(args);
    if (fn == "MID") return funcMID(args);
    if (fn == "FIND") return funcFIND(args);
    if (fn == "SUBSTITUTE") return funcSUBSTITUTE(args);
    if (fn == "TEXT") return funcTEXT(args);
    // Statistical
    if (fn == "COUNTIF") return funcCOUNTIF(args);
    if (fn == "SUMIF") return funcSUMIF(args);
    // Date
    if (fn == "NOW") return funcNOW(args);
    if (fn == "TODAY") return funcTODAY(args);
    if (fn == "YEAR") return funcYEAR(args);
    if (fn == "MONTH") return funcMONTH(args);
    if (fn == "DAY") return funcDAY(args);

    m_lastError = "Unknown function: " + fn;
    return QVariant("#NAME?");
}

// ---- Aggregate functions ----

QVariant FormulaEngine::funcSUM(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args); double sum = 0;
    for (const auto& a : flat) if (!a.isNull() && a.isValid()) sum += toNumber(a);
    return sum;
}

QVariant FormulaEngine::funcAVERAGE(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    if (flat.empty()) return QVariant("#DIV/0!");
    double sum = 0; int count = 0;
    for (const auto& a : flat) if (!a.isNull() && a.isValid()) { sum += toNumber(a); count++; }
    return count == 0 ? QVariant("#DIV/0!") : QVariant(sum / count);
}

QVariant FormulaEngine::funcCOUNT(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args); int count = 0;
    for (const auto& a : flat) {
        if (!a.isNull() && a.isValid()) {
            bool ok = false; a.toDouble(&ok);
            if (ok || a.typeId() == QMetaType::Int || a.typeId() == QMetaType::Double) count++;
        }
    }
    return count;
}

QVariant FormulaEngine::funcCOUNTA(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args); int count = 0;
    for (const auto& a : flat) if (!a.isNull() && a.isValid() && !a.toString().isEmpty()) count++;
    return count;
}

QVariant FormulaEngine::funcMIN(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args); double min = std::numeric_limits<double>::max(); bool found = false;
    for (const auto& a : flat) if (!a.isNull() && a.isValid()) { double v = toNumber(a); if (!found || v < min) { min = v; found = true; } }
    return found ? QVariant(min) : QVariant();
}

QVariant FormulaEngine::funcMAX(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args); double max = std::numeric_limits<double>::lowest(); bool found = false;
    for (const auto& a : flat) if (!a.isNull() && a.isValid()) { double v = toNumber(a); if (!found || v > max) { max = v; found = true; } }
    return found ? QVariant(max) : QVariant();
}

QVariant FormulaEngine::funcIF(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    bool cond = toBoolean(args[0]);
    if (cond) return args[1];
    return args.size() >= 3 ? args[2] : QVariant(false);
}

QVariant FormulaEngine::funcCONCAT(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args); QString result;
    for (const auto& a : flat) result += toString(a);
    return result;
}

QVariant FormulaEngine::funcLEN(const std::vector<QVariant>& args) {
    if (args.empty()) return 0;
    return static_cast<int>(toString(args[0]).length());
}

QVariant FormulaEngine::funcUPPER(const std::vector<QVariant>& args) { return args.empty() ? QVariant("") : QVariant(toString(args[0]).toUpper()); }
QVariant FormulaEngine::funcLOWER(const std::vector<QVariant>& args) { return args.empty() ? QVariant("") : QVariant(toString(args[0]).toLower()); }
QVariant FormulaEngine::funcTRIM(const std::vector<QVariant>& args) { return args.empty() ? QVariant("") : QVariant(toString(args[0]).trimmed()); }

// ---- Math functions ----

QVariant FormulaEngine::funcROUND(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double val = toNumber(args[0]);
    int decimals = args.size() >= 2 ? static_cast<int>(toNumber(args[1])) : 0;
    double factor = std::pow(10.0, decimals);
    return std::round(val * factor) / factor;
}

QVariant FormulaEngine::funcABS(const std::vector<QVariant>& args) {
    return args.empty() ? QVariant("#VALUE!") : QVariant(std::abs(toNumber(args[0])));
}

QVariant FormulaEngine::funcSQRT(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double val = toNumber(args[0]);
    return val < 0 ? QVariant("#NUM!") : QVariant(std::sqrt(val));
}

QVariant FormulaEngine::funcPOWER(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    return std::pow(toNumber(args[0]), toNumber(args[1]));
}

QVariant FormulaEngine::funcMOD(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double d = toNumber(args[1]);
    return d == 0.0 ? QVariant("#DIV/0!") : QVariant(std::fmod(toNumber(args[0]), d));
}

QVariant FormulaEngine::funcINT(const std::vector<QVariant>& args) {
    return args.empty() ? QVariant("#VALUE!") : QVariant(std::floor(toNumber(args[0])));
}

QVariant FormulaEngine::funcCEILING(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double val = toNumber(args[0]);
    double sig = args.size() >= 2 ? toNumber(args[1]) : 1.0;
    if (sig == 0.0) return QVariant(0.0);
    return std::ceil(val / sig) * sig;
}

QVariant FormulaEngine::funcFLOOR(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    double val = toNumber(args[0]);
    double sig = args.size() >= 2 ? toNumber(args[1]) : 1.0;
    if (sig == 0.0) return QVariant(0.0);
    return std::floor(val / sig) * sig;
}

// ---- Logical functions ----

QVariant FormulaEngine::funcAND(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    for (const auto& a : flat) if (!toBoolean(a)) return QVariant(false);
    return QVariant(true);
}

QVariant FormulaEngine::funcOR(const std::vector<QVariant>& args) {
    auto flat = flattenArgs(args);
    for (const auto& a : flat) if (toBoolean(a)) return QVariant(true);
    return QVariant(false);
}

QVariant FormulaEngine::funcNOT(const std::vector<QVariant>& args) {
    return args.empty() ? QVariant("#VALUE!") : QVariant(!toBoolean(args[0]));
}

QVariant FormulaEngine::funcIFERROR(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    QString str = toString(args[0]);
    if (str.startsWith('#')) return args[1];
    return args[0];
}

// ---- Text functions ----

QVariant FormulaEngine::funcLEFT(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    int count = args.size() >= 2 ? static_cast<int>(toNumber(args[1])) : 1;
    return toString(args[0]).left(count);
}

QVariant FormulaEngine::funcRIGHT(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    int count = args.size() >= 2 ? static_cast<int>(toNumber(args[1])) : 1;
    return toString(args[0]).right(count);
}

QVariant FormulaEngine::funcMID(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    QString str = toString(args[0]);
    int start = static_cast<int>(toNumber(args[1])) - 1; // 1-based to 0-based
    int count = static_cast<int>(toNumber(args[2]));
    return str.mid(start, count);
}

QVariant FormulaEngine::funcFIND(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    QString search = toString(args[0]);
    QString text = toString(args[1]);
    int startPos = args.size() >= 3 ? static_cast<int>(toNumber(args[2])) - 1 : 0;
    int idx = text.indexOf(search, startPos);
    return idx >= 0 ? QVariant(idx + 1) : QVariant("#VALUE!"); // 1-based
}

QVariant FormulaEngine::funcSUBSTITUTE(const std::vector<QVariant>& args) {
    if (args.size() < 3) return QVariant("#VALUE!");
    QString text = toString(args[0]);
    QString oldText = toString(args[1]);
    QString newText = toString(args[2]);
    return text.replace(oldText, newText);
}

QVariant FormulaEngine::funcTEXT(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    double val = toNumber(args[0]);
    QString fmt = toString(args[1]);
    if (fmt.contains('#') || fmt.contains('0')) {
        // Simple number formatting
        int decimals = 0;
        int dotPos = fmt.indexOf('.');
        if (dotPos >= 0) decimals = fmt.length() - dotPos - 1;
        return QString::number(val, 'f', decimals);
    }
    return QString::number(val);
}

// ---- Statistical functions ----

bool FormulaEngine::matchesCriteria(const QVariant& value, const QString& criteria) {
    if (criteria.startsWith(">=")) return toNumber(value) >= criteria.mid(2).toDouble();
    if (criteria.startsWith("<=")) return toNumber(value) <= criteria.mid(2).toDouble();
    if (criteria.startsWith("<>")) return value.toString() != criteria.mid(2);
    if (criteria.startsWith(">")) return toNumber(value) > criteria.mid(1).toDouble();
    if (criteria.startsWith("<")) return toNumber(value) < criteria.mid(1).toDouble();
    if (criteria.startsWith("=")) return value.toString() == criteria.mid(1);

    // Direct comparison
    bool ok = false;
    double critNum = criteria.toDouble(&ok);
    if (ok) return toNumber(value) == critNum;
    return value.toString() == criteria;
}

QVariant FormulaEngine::funcCOUNTIF(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto flat = flattenArgs({args[0]});
    QString criteria = toString(args[1]);
    int count = 0;
    for (const auto& v : flat) if (matchesCriteria(v, criteria)) count++;
    return count;
}

QVariant FormulaEngine::funcSUMIF(const std::vector<QVariant>& args) {
    if (args.size() < 2) return QVariant("#VALUE!");
    auto range = flattenArgs({args[0]});
    QString criteria = toString(args[1]);
    auto sumRange = args.size() >= 3 ? flattenArgs({args[2]}) : range;
    double sum = 0;
    for (size_t i = 0; i < range.size() && i < sumRange.size(); ++i) {
        if (matchesCriteria(range[i], criteria)) sum += toNumber(sumRange[i]);
    }
    return sum;
}

// ---- Date functions ----

QVariant FormulaEngine::funcNOW(const std::vector<QVariant>&) {
    return QDateTime::currentDateTime().toString("yyyy-MM-dd hh:mm:ss");
}

QVariant FormulaEngine::funcTODAY(const std::vector<QVariant>&) {
    return QDate::currentDate().toString("yyyy-MM-dd");
}

QVariant FormulaEngine::funcYEAR(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QDate date = QDate::fromString(toString(args[0]), "yyyy-MM-dd");
    return date.isValid() ? QVariant(date.year()) : QVariant("#VALUE!");
}

QVariant FormulaEngine::funcMONTH(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QDate date = QDate::fromString(toString(args[0]), "yyyy-MM-dd");
    return date.isValid() ? QVariant(date.month()) : QVariant("#VALUE!");
}

QVariant FormulaEngine::funcDAY(const std::vector<QVariant>& args) {
    if (args.empty()) return QVariant("#VALUE!");
    QDate date = QDate::fromString(toString(args[0]), "yyyy-MM-dd");
    return date.isValid() ? QVariant(date.day()) : QVariant("#VALUE!");
}

// ---- Helper functions ----

double FormulaEngine::toNumber(const QVariant& value) {
    if (value.typeId() == QMetaType::Double || value.typeId() == QMetaType::Int) return value.toDouble();
    if (value.typeId() == QMetaType::Bool) return value.toBool() ? 1.0 : 0.0;
    bool ok = false; double r = value.toString().toDouble(&ok);
    return ok ? r : 0.0;
}

QString FormulaEngine::toString(const QVariant& value) { return value.toString(); }

bool FormulaEngine::toBoolean(const QVariant& value) {
    if (value.typeId() == QMetaType::Bool) return value.toBool();
    return toNumber(value) != 0;
}

QVariant FormulaEngine::getCellValue(const CellAddress& addr) {
    if (!m_spreadsheet) return QVariant();
    return m_spreadsheet->getCellValue(addr);
}

std::vector<QVariant> FormulaEngine::getRangeValues(const CellRange& range) {
    std::vector<QVariant> values;
    if (!m_spreadsheet) return values;
    for (const auto& addr : range.getCells()) values.push_back(m_spreadsheet->getCellValue(addr));
    return values;
}
