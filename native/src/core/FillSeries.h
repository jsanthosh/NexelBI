#ifndef FILLSERIES_H
#define FILLSERIES_H

#include <QString>
#include <QStringList>

class FillSeries {
public:
    static QStringList generateSeries(const QStringList& seeds, int count);

private:
    static const QStringList MONTHS_FULL;
    static const QStringList MONTHS_SHORT;
    static const QStringList DAYS_FULL;
    static const QStringList DAYS_SHORT;
    static QString matchCase(const QString& templateStr, const QString& value);
};

#endif // FILLSERIES_H
