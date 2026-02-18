#ifndef NUMBERFORMAT_H
#define NUMBERFORMAT_H

#include <QString>

enum class NumberFormatType {
    General,
    Number,
    Currency,
    Accounting,
    Percentage,
    Date,
    Time,
    Text,
    Custom
};

struct NumberFormatOptions {
    NumberFormatType type = NumberFormatType::General;
    int decimalPlaces = 2;
    bool useThousandsSeparator = false;
    QString currencyCode = "USD";
    QString dateFormatId = "mm/dd/yyyy";
    QString customFormat = "";
};

struct CurrencyDef {
    QString code;
    QString symbol;
    QString label;
};

class NumberFormat {
public:
    static QString format(const QString& value, const NumberFormatOptions& options);
    static QString applyCustomFormat(const QString& value, const QString& formatStr);
    static NumberFormatType typeFromString(const QString& str);
    static QString typeToString(NumberFormatType type);
    static QString getCurrencySymbol(const QString& code);

    static const QList<CurrencyDef>& currencies();
};

#endif // NUMBERFORMAT_H
