#include "NumberFormat.h"
#include <QLocale>
#include <QDate>
#include <QTime>
#include <QDateTime>
#include <QRegularExpression>
#include <cmath>

static const QList<CurrencyDef> s_currencies = {
    {"USD", "$",       "US Dollar ($)"},
    {"EUR", "\u20AC",  "Euro (\u20AC)"},
    {"GBP", "\u00A3",  "British Pound (\u00A3)"},
    {"JPY", "\u00A5",  "Japanese Yen (\u00A5)"},
    {"INR", "\u20B9",  "Indian Rupee (\u20B9)"},
    {"CNY", "\u00A5",  "Chinese Yuan (\u00A5)"},
    {"KRW", "\u20A9",  "Korean Won (\u20A9)"},
    {"CAD", "CA$",     "Canadian Dollar (CA$)"},
    {"AUD", "A$",      "Australian Dollar (A$)"},
    {"CHF", "CHF",     "Swiss Franc (CHF)"},
    {"BRL", "R$",      "Brazilian Real (R$)"},
    {"MXN", "MX$",     "Mexican Peso (MX$)"},
};

const QList<CurrencyDef>& NumberFormat::currencies() {
    return s_currencies;
}

QString NumberFormat::getCurrencySymbol(const QString& code) {
    for (const auto& c : s_currencies) {
        if (c.code == code) return c.symbol;
    }
    return "$";
}

NumberFormatType NumberFormat::typeFromString(const QString& str) {
    QString lower = str.toLower();
    if (lower == "number") return NumberFormatType::Number;
    if (lower == "currency") return NumberFormatType::Currency;
    if (lower == "accounting") return NumberFormatType::Accounting;
    if (lower == "percentage") return NumberFormatType::Percentage;
    if (lower == "date") return NumberFormatType::Date;
    if (lower == "time") return NumberFormatType::Time;
    if (lower == "text") return NumberFormatType::Text;
    if (lower == "custom") return NumberFormatType::Custom;
    return NumberFormatType::General;
}

QString NumberFormat::typeToString(NumberFormatType type) {
    switch (type) {
        case NumberFormatType::Number: return "Number";
        case NumberFormatType::Currency: return "Currency";
        case NumberFormatType::Accounting: return "Accounting";
        case NumberFormatType::Percentage: return "Percentage";
        case NumberFormatType::Date: return "Date";
        case NumberFormatType::Time: return "Time";
        case NumberFormatType::Text: return "Text";
        case NumberFormatType::Custom: return "Custom";
        default: return "General";
    }
}

static QString formatNumber(double num, int decimals, bool useThousands) {
    QLocale locale(QLocale::English, QLocale::UnitedStates);
    if (useThousands) {
        return locale.toString(num, 'f', decimals);
    } else {
        // No grouping separator
        QString result = QString::number(num, 'f', decimals);
        return result;
    }
}

static QDate parseDate(const QString& value) {
    // Try ISO format
    QDate d = QDate::fromString(value, Qt::ISODate);
    if (d.isValid()) return d;

    // Try common formats
    d = QDate::fromString(value, "MM/dd/yyyy");
    if (d.isValid()) return d;
    d = QDate::fromString(value, "dd/MM/yyyy");
    if (d.isValid()) return d;
    d = QDate::fromString(value, "yyyy-MM-dd");
    if (d.isValid()) return d;

    // Try Excel serial date
    bool ok;
    double serial = value.toDouble(&ok);
    if (ok && serial > 0 && serial < 200000) {
        QDate epoch(1899, 12, 30);
        return epoch.addDays(static_cast<int>(serial));
    }

    return QDate();
}

QString NumberFormat::format(const QString& value, const NumberFormatOptions& options) {
    if (value.isEmpty() || options.type == NumberFormatType::General ||
        options.type == NumberFormatType::Text) {
        return value;
    }

    bool ok;
    double num = value.toDouble(&ok);

    switch (options.type) {
        case NumberFormatType::Number: {
            if (!ok) return value;
            return formatNumber(num, options.decimalPlaces, options.useThousandsSeparator);
        }

        case NumberFormatType::Currency: {
            if (!ok) return value;
            QString symbol = getCurrencySymbol(options.currencyCode);
            QString formatted = formatNumber(std::abs(num), options.decimalPlaces, true);
            if (num < 0) return "-" + symbol + formatted;
            return symbol + formatted;
        }

        case NumberFormatType::Accounting: {
            if (!ok) return value;
            QString symbol = getCurrencySymbol(options.currencyCode);
            QString formatted = formatNumber(std::abs(num), options.decimalPlaces, true);
            if (num < 0) return "(" + symbol + formatted + ")";
            return symbol + formatted;
        }

        case NumberFormatType::Percentage: {
            if (!ok) return value;
            double pct = num * 100.0;
            return formatNumber(pct, options.decimalPlaces, false) + "%";
        }

        case NumberFormatType::Date: {
            QDate date = parseDate(value);
            if (!date.isValid()) return value;

            if (options.dateFormatId == "yyyy-mm-dd" || options.dateFormatId == "yyyy-MM-dd")
                return date.toString("yyyy-MM-dd");
            if (options.dateFormatId == "dd/mm/yyyy" || options.dateFormatId == "dd/MM/yyyy")
                return date.toString("dd/MM/yyyy");
            if (options.dateFormatId == "mmm d, yyyy")
                return date.toString("MMM d, yyyy");
            if (options.dateFormatId == "mmmm d, yyyy")
                return date.toString("MMMM d, yyyy");
            if (options.dateFormatId == "d-mmm-yy")
                return date.toString("d-MMM-yy");
            if (options.dateFormatId == "mm/dd")
                return date.toString("MM/dd");
            // Default: mm/dd/yyyy
            return date.toString("MM/dd/yyyy");
        }

        case NumberFormatType::Time: {
            QDate date = parseDate(value);
            if (date.isValid()) {
                // For serial dates, extract time portion
                double serial = value.toDouble(&ok);
                if (ok) {
                    double fraction = serial - std::floor(serial);
                    int totalSecs = static_cast<int>(fraction * 86400);
                    QTime time(totalSecs / 3600, (totalSecs % 3600) / 60, totalSecs % 60);
                    return time.toString("hh:mm:ss AP");
                }
            }
            QTime time = QTime::fromString(value);
            if (time.isValid()) return time.toString("hh:mm:ss AP");
            return value;
        }

        case NumberFormatType::Custom: {
            return applyCustomFormat(value, options.customFormat);
        }

        default:
            return value;
    }
}

QString NumberFormat::applyCustomFormat(const QString& value, const QString& formatStr) {
    if (formatStr.isEmpty()) return value;

    bool ok;
    double num = value.toDouble(&ok);
    if (!ok) return value;

    // Split format for positive;negative;zero
    QStringList parts = formatStr.split(';');
    QString fmt = parts[0];
    if (num < 0 && parts.size() > 1) fmt = parts[1];
    else if (num == 0.0 && parts.size() > 2) fmt = parts[2];

    // Strip color codes like [Red]
    fmt.replace(QRegularExpression("\\[[A-Za-z]+\\]"), "");

    bool isPercent = fmt.contains('%');
    double val = isPercent ? num * 100.0 : num;
    double absVal = std::abs(val);

    // Count decimal places
    int dotIdx = fmt.indexOf('.');
    int decimals = 0;
    if (dotIdx != -1) {
        decimals = fmt.length() - dotIdx - 1;
    }

    bool useComma = fmt.contains(',');

    QString formatted = formatNumber(absVal, decimals, useComma);

    // Extract prefix (currency symbol etc)
    QString prefix;
    QRegularExpression prefixRe("^([^#0,.]+)");
    auto prefixMatch = prefixRe.match(fmt);
    if (prefixMatch.hasMatch()) {
        prefix = prefixMatch.captured(1).trimmed();
    }

    QString suffix;
    if (isPercent) suffix = "%";

    QString sign = (val < 0) ? "-" : "";
    return sign + prefix + formatted + suffix;
}
