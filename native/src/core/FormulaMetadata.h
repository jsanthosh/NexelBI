#ifndef FORMULAMETADATA_H
#define FORMULAMETADATA_H

#include <QString>
#include <QVector>
#include <QMap>

struct FormulaParamInfo {
    QString name;
    bool optional = false;
};

struct FormulaFuncInfo {
    QString name;
    QString description;
    QString syntax;
    QVector<FormulaParamInfo> params;
};

inline const QMap<QString, FormulaFuncInfo>& formulaRegistry() {
    static const QMap<QString, FormulaFuncInfo> reg = {
        // === Aggregate ===
        {"SUM", {"SUM", "Adds all the numbers in a range of cells",
            "SUM(number1, [number2], ...)",
            {{"number1", false}, {"number2", true}}}},
        {"AVERAGE", {"AVERAGE", "Returns the average of its arguments",
            "AVERAGE(number1, [number2], ...)",
            {{"number1", false}, {"number2", true}}}},
        {"COUNT", {"COUNT", "Counts the number of cells that contain numbers",
            "COUNT(value1, [value2], ...)",
            {{"value1", false}, {"value2", true}}}},
        {"COUNTA", {"COUNTA", "Counts the number of non-empty cells",
            "COUNTA(value1, [value2], ...)",
            {{"value1", false}, {"value2", true}}}},
        {"MIN", {"MIN", "Returns the smallest value in a set of values",
            "MIN(number1, [number2], ...)",
            {{"number1", false}, {"number2", true}}}},
        {"MAX", {"MAX", "Returns the largest value in a set of values",
            "MAX(number1, [number2], ...)",
            {{"number1", false}, {"number2", true}}}},

        // === Conditional ===
        {"IF", {"IF", "Checks a condition and returns one value if TRUE, another if FALSE",
            "IF(logical_test, value_if_true, [value_if_false])",
            {{"logical_test", false}, {"value_if_true", false}, {"value_if_false", true}}}},
        {"IFERROR", {"IFERROR", "Returns a value if no error; otherwise returns an alternate value",
            "IFERROR(value, value_if_error)",
            {{"value", false}, {"value_if_error", false}}}},
        {"AND", {"AND", "Returns TRUE if all arguments are TRUE",
            "AND(logical1, [logical2], ...)",
            {{"logical1", false}, {"logical2", true}}}},
        {"OR", {"OR", "Returns TRUE if any argument is TRUE",
            "OR(logical1, [logical2], ...)",
            {{"logical1", false}, {"logical2", true}}}},
        {"NOT", {"NOT", "Reverses the logic of its argument",
            "NOT(logical)",
            {{"logical", false}}}},

        // === String ===
        {"CONCAT", {"CONCAT", "Joins several text strings into one",
            "CONCAT(text1, [text2], ...)",
            {{"text1", false}, {"text2", true}}}},
        {"CONCATENATE", {"CONCATENATE", "Joins several text strings into one",
            "CONCATENATE(text1, [text2], ...)",
            {{"text1", false}, {"text2", true}}}},
        {"LEN", {"LEN", "Returns the number of characters in a text string",
            "LEN(text)",
            {{"text", false}}}},
        {"UPPER", {"UPPER", "Converts text to uppercase",
            "UPPER(text)",
            {{"text", false}}}},
        {"LOWER", {"LOWER", "Converts text to lowercase",
            "LOWER(text)",
            {{"text", false}}}},
        {"TRIM", {"TRIM", "Removes extra spaces from text",
            "TRIM(text)",
            {{"text", false}}}},
        {"LEFT", {"LEFT", "Returns the leftmost characters from a text string",
            "LEFT(text, [num_chars])",
            {{"text", false}, {"num_chars", true}}}},
        {"RIGHT", {"RIGHT", "Returns the rightmost characters from a text string",
            "RIGHT(text, [num_chars])",
            {{"text", false}, {"num_chars", true}}}},
        {"MID", {"MID", "Returns characters from the middle of a text string",
            "MID(text, start_num, num_chars)",
            {{"text", false}, {"start_num", false}, {"num_chars", false}}}},
        {"FIND", {"FIND", "Finds one text string within another (case-sensitive)",
            "FIND(find_text, within_text, [start_num])",
            {{"find_text", false}, {"within_text", false}, {"start_num", true}}}},
        {"SEARCH", {"SEARCH", "Finds one text string within another (case-insensitive)",
            "SEARCH(find_text, within_text, [start_num])",
            {{"find_text", false}, {"within_text", false}, {"start_num", true}}}},
        {"SUBSTITUTE", {"SUBSTITUTE", "Substitutes new text for old text in a string",
            "SUBSTITUTE(text, old_text, new_text, [instance_num])",
            {{"text", false}, {"old_text", false}, {"new_text", false}, {"instance_num", true}}}},
        {"TEXT", {"TEXT", "Formats a number as text with a given format",
            "TEXT(value, format_text)",
            {{"value", false}, {"format_text", false}}}},
        {"PROPER", {"PROPER", "Capitalizes the first letter of each word",
            "PROPER(text)",
            {{"text", false}}}},
        {"REPT", {"REPT", "Repeats text a given number of times",
            "REPT(text, number_times)",
            {{"text", false}, {"number_times", false}}}},
        {"EXACT", {"EXACT", "Checks whether two text strings are exactly the same",
            "EXACT(text1, text2)",
            {{"text1", false}, {"text2", false}}}},
        {"VALUE", {"VALUE", "Converts a text string that represents a number to a number",
            "VALUE(text)",
            {{"text", false}}}},

        // === Math ===
        {"ROUND", {"ROUND", "Rounds a number to a specified number of digits",
            "ROUND(number, num_digits)",
            {{"number", false}, {"num_digits", false}}}},
        {"ROUNDUP", {"ROUNDUP", "Rounds a number up, away from zero",
            "ROUNDUP(number, num_digits)",
            {{"number", false}, {"num_digits", false}}}},
        {"ROUNDDOWN", {"ROUNDDOWN", "Rounds a number down, toward zero",
            "ROUNDDOWN(number, num_digits)",
            {{"number", false}, {"num_digits", false}}}},
        {"ABS", {"ABS", "Returns the absolute value of a number",
            "ABS(number)",
            {{"number", false}}}},
        {"SQRT", {"SQRT", "Returns the square root of a number",
            "SQRT(number)",
            {{"number", false}}}},
        {"POWER", {"POWER", "Returns the result of a number raised to a power",
            "POWER(number, power)",
            {{"number", false}, {"power", false}}}},
        {"MOD", {"MOD", "Returns the remainder after division",
            "MOD(number, divisor)",
            {{"number", false}, {"divisor", false}}}},
        {"INT", {"INT", "Rounds a number down to the nearest integer",
            "INT(number)",
            {{"number", false}}}},
        {"CEILING", {"CEILING", "Rounds a number up to the nearest multiple",
            "CEILING(number, significance)",
            {{"number", false}, {"significance", false}}}},
        {"FLOOR", {"FLOOR", "Rounds a number down to the nearest multiple",
            "FLOOR(number, significance)",
            {{"number", false}, {"significance", false}}}},
        {"LOG", {"LOG", "Returns the logarithm of a number to a specified base",
            "LOG(number, [base])",
            {{"number", false}, {"base", true}}}},
        {"LN", {"LN", "Returns the natural logarithm of a number",
            "LN(number)",
            {{"number", false}}}},
        {"EXP", {"EXP", "Returns e raised to the power of a number",
            "EXP(number)",
            {{"number", false}}}},
        {"RAND", {"RAND", "Returns a random number between 0 and 1",
            "RAND()", {}}},
        {"RANDBETWEEN", {"RANDBETWEEN", "Returns a random integer between two values",
            "RANDBETWEEN(bottom, top)",
            {{"bottom", false}, {"top", false}}}},

        // === Statistical ===
        {"COUNTIF", {"COUNTIF", "Counts cells that meet a condition",
            "COUNTIF(range, criteria)",
            {{"range", false}, {"criteria", false}}}},
        {"SUMIF", {"SUMIF", "Adds cells that meet a condition",
            "SUMIF(range, criteria, [sum_range])",
            {{"range", false}, {"criteria", false}, {"sum_range", true}}}},
        {"AVERAGEIF", {"AVERAGEIF", "Returns the average of cells that meet a condition",
            "AVERAGEIF(range, criteria, [average_range])",
            {{"range", false}, {"criteria", false}, {"average_range", true}}}},
        {"COUNTBLANK", {"COUNTBLANK", "Counts the number of empty cells in a range",
            "COUNTBLANK(range)",
            {{"range", false}}}},
        {"SUMPRODUCT", {"SUMPRODUCT", "Returns the sum of products of corresponding arrays",
            "SUMPRODUCT(array1, [array2], ...)",
            {{"array1", false}, {"array2", true}}}},
        {"MEDIAN", {"MEDIAN", "Returns the median of the given numbers",
            "MEDIAN(number1, [number2], ...)",
            {{"number1", false}, {"number2", true}}}},
        {"MODE", {"MODE", "Returns the most frequently occurring value",
            "MODE(number1, [number2], ...)",
            {{"number1", false}, {"number2", true}}}},
        {"STDEV", {"STDEV", "Estimates standard deviation based on a sample",
            "STDEV(number1, [number2], ...)",
            {{"number1", false}, {"number2", true}}}},
        {"VAR", {"VAR", "Estimates variance based on a sample",
            "VAR(number1, [number2], ...)",
            {{"number1", false}, {"number2", true}}}},
        {"LARGE", {"LARGE", "Returns the k-th largest value in a data set",
            "LARGE(array, k)",
            {{"array", false}, {"k", false}}}},
        {"SMALL", {"SMALL", "Returns the k-th smallest value in a data set",
            "SMALL(array, k)",
            {{"array", false}, {"k", false}}}},
        {"RANK", {"RANK", "Returns the rank of a number in a list of numbers",
            "RANK(number, ref, [order])",
            {{"number", false}, {"ref", false}, {"order", true}}}},
        {"PERCENTILE", {"PERCENTILE", "Returns the k-th percentile of values in a range",
            "PERCENTILE(array, k)",
            {{"array", false}, {"k", false}}}},

        // === Date/Time ===
        {"NOW", {"NOW", "Returns the current date and time",
            "NOW()", {}}},
        {"TODAY", {"TODAY", "Returns the current date",
            "TODAY()", {}}},
        {"YEAR", {"YEAR", "Returns the year of a date",
            "YEAR(serial_number)",
            {{"serial_number", false}}}},
        {"MONTH", {"MONTH", "Returns the month of a date",
            "MONTH(serial_number)",
            {{"serial_number", false}}}},
        {"DAY", {"DAY", "Returns the day of a date",
            "DAY(serial_number)",
            {{"serial_number", false}}}},
        {"DATE", {"DATE", "Returns a date from year, month, and day",
            "DATE(year, month, day)",
            {{"year", false}, {"month", false}, {"day", false}}}},
        {"HOUR", {"HOUR", "Returns the hour of a time value",
            "HOUR(serial_number)",
            {{"serial_number", false}}}},
        {"MINUTE", {"MINUTE", "Returns the minutes of a time value",
            "MINUTE(serial_number)",
            {{"serial_number", false}}}},
        {"SECOND", {"SECOND", "Returns the seconds of a time value",
            "SECOND(serial_number)",
            {{"serial_number", false}}}},
        {"DATEDIF", {"DATEDIF", "Calculates the number of days, months, or years between two dates",
            "DATEDIF(start_date, end_date, unit)",
            {{"start_date", false}, {"end_date", false}, {"unit", false}}}},
        {"NETWORKDAYS", {"NETWORKDAYS", "Returns the number of whole working days between two dates",
            "NETWORKDAYS(start_date, end_date, [holidays])",
            {{"start_date", false}, {"end_date", false}, {"holidays", true}}}},
        {"WEEKDAY", {"WEEKDAY", "Returns the day of the week for a date",
            "WEEKDAY(serial_number, [return_type])",
            {{"serial_number", false}, {"return_type", true}}}},
        {"EDATE", {"EDATE", "Returns a date that is a given number of months before or after a date",
            "EDATE(start_date, months)",
            {{"start_date", false}, {"months", false}}}},
        {"EOMONTH", {"EOMONTH", "Returns the last day of the month before or after a given number of months",
            "EOMONTH(start_date, months)",
            {{"start_date", false}, {"months", false}}}},
        {"DATEVALUE", {"DATEVALUE", "Converts a date in text format to a serial number",
            "DATEVALUE(date_text)",
            {{"date_text", false}}}},

        // === Lookup ===
        {"VLOOKUP", {"VLOOKUP", "Looks for a value in the leftmost column and returns a value in the same row",
            "VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])",
            {{"lookup_value", false}, {"table_array", false}, {"col_index_num", false}, {"range_lookup", true}}}},
        {"HLOOKUP", {"HLOOKUP", "Looks for a value in the top row and returns a value in the same column",
            "HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])",
            {{"lookup_value", false}, {"table_array", false}, {"row_index_num", false}, {"range_lookup", true}}}},
        {"XLOOKUP", {"XLOOKUP", "Searches a range and returns a corresponding item",
            "XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode])",
            {{"lookup_value", false}, {"lookup_array", false}, {"return_array", false}, {"if_not_found", true}, {"match_mode", true}}}},
        {"INDEX", {"INDEX", "Returns the value of a cell in a given range by row and column number",
            "INDEX(array, row_num, [column_num])",
            {{"array", false}, {"row_num", false}, {"column_num", true}}}},
        {"MATCH", {"MATCH", "Returns the relative position of a value in a range",
            "MATCH(lookup_value, lookup_array, [match_type])",
            {{"lookup_value", false}, {"lookup_array", false}, {"match_type", true}}}},

        // === Info ===
        {"ISBLANK", {"ISBLANK", "Returns TRUE if a cell is empty",
            "ISBLANK(value)",
            {{"value", false}}}},
        {"ISERROR", {"ISERROR", "Returns TRUE if a value is an error",
            "ISERROR(value)",
            {{"value", false}}}},
        {"ISNUMBER", {"ISNUMBER", "Returns TRUE if a value is a number",
            "ISNUMBER(value)",
            {{"value", false}}}},
        {"ISTEXT", {"ISTEXT", "Returns TRUE if a value is text",
            "ISTEXT(value)",
            {{"value", false}}}},
        {"CHOOSE", {"CHOOSE", "Returns a value from a list based on an index number",
            "CHOOSE(index_num, value1, [value2], ...)",
            {{"index_num", false}, {"value1", false}, {"value2", true}}}},
        {"SWITCH", {"SWITCH", "Evaluates an expression against a list of values and returns the matching result",
            "SWITCH(expression, value1, result1, [value2, result2], ..., [default])",
            {{"expression", false}, {"value1", false}, {"result1", false}, {"default", true}}}},
    };
    return reg;
}

// Helper: find which function the cursor is inside, and which parameter index
struct FormulaContext {
    QString funcName;
    int paramIndex = -1;  // 0-based, -1 if not inside a function
};

inline FormulaContext findFormulaContext(const QString& text, int cursorPos) {
    FormulaContext ctx;
    if (!text.startsWith("=")) return ctx;

    // Walk backwards from cursor to find the innermost open paren
    int depth = 0;
    int parenPos = -1;
    for (int i = cursorPos - 1; i >= 1; --i) {
        QChar ch = text[i];
        if (ch == ')') depth++;
        else if (ch == '(') {
            if (depth == 0) { parenPos = i; break; }
            depth--;
        }
    }
    if (parenPos < 0) return ctx;

    // Extract function name before the paren
    int nameEnd = parenPos;
    int nameStart = nameEnd - 1;
    while (nameStart >= 1 && (text[nameStart].isLetterOrNumber() || text[nameStart] == '_'))
        nameStart--;
    nameStart++;
    if (nameStart >= nameEnd) return ctx;

    ctx.funcName = text.mid(nameStart, nameEnd - nameStart).toUpper();

    // Count commas between parenPos+1 and cursorPos to determine param index
    int commas = 0;
    int d = 0;
    for (int i = parenPos + 1; i < cursorPos && i < text.length(); ++i) {
        QChar ch = text[i];
        if (ch == '(') d++;
        else if (ch == ')') d--;
        else if (ch == ',' && d == 0) commas++;
    }
    ctx.paramIndex = commas;
    return ctx;
}

// Build rich-text syntax string with current param highlighted
inline QString buildParamHintHtml(const FormulaFuncInfo& info, int paramIndex) {
    QString html = "<span style='color:#333; font-size:12px;'>";
    html += "<b>" + info.name + "</b>(";
    for (int i = 0; i < info.params.size(); ++i) {
        if (i > 0) html += ", ";
        const auto& p = info.params[i];
        QString pName = p.optional ? "[" + p.name + "]" : p.name;
        if (i == paramIndex) {
            html += "<b style='color:#0078D4;'>" + pName + "</b>";
        } else {
            html += "<span style='color:#888;'>" + pName + "</span>";
        }
    }
    // Add "..." for variadic functions (those ending with optional params)
    if (!info.params.isEmpty() && info.params.last().optional) {
        html += ", ...";
    }
    html += ")</span>";
    return html;
}

#endif // FORMULAMETADATA_H
