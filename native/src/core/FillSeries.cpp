#include "FillSeries.h"
#include <QRegularExpression>
#include <cmath>

const QStringList FillSeries::MONTHS_FULL = {
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
};

const QStringList FillSeries::MONTHS_SHORT = {
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
};

const QStringList FillSeries::DAYS_FULL = {
    "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"
};

const QStringList FillSeries::DAYS_SHORT = {
    "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"
};

QString FillSeries::matchCase(const QString& templateStr, const QString& value) {
    if (templateStr == templateStr.toUpper()) return value.toUpper();
    if (!templateStr.isEmpty() && templateStr[0].isUpper())
        return value[0].toUpper() + value.mid(1).toLower();
    return value.toLower();
}

QStringList FillSeries::generateSeries(const QStringList& seeds, int count) {
    if (seeds.isEmpty() || count == 0) return {};

    QString s0 = seeds[0].trimmed();
    QString s1 = seeds.size() > 1 ? seeds[1].trimmed() : QString();

    // --- Numeric ---
    bool ok0, ok1;
    double n0 = s0.toDouble(&ok0);
    if (ok0 && !s0.isEmpty()) {
        double n1 = s1.toDouble(&ok1);
        double step = (ok1 && seeds.size() > 1) ? (n1 - n0) : 1.0;
        bool isInt = (n0 == std::floor(n0)) && (step == std::floor(step));

        QStringList result;
        result.reserve(count);
        for (int i = 0; i < count; ++i) {
            double val = n0 + step * i;
            if (isInt) {
                result.append(QString::number(static_cast<long long>(std::round(val))));
            } else {
                result.append(QString::number(val, 'g', 10));
            }
        }
        return result;
    }

    // --- List-based (months / days) ---
    const QList<const QStringList*> lists = {&MONTHS_FULL, &MONTHS_SHORT, &DAYS_FULL, &DAYS_SHORT};
    for (const auto* list : lists) {
        int idx0 = -1;
        for (int i = 0; i < list->size(); ++i) {
            if ((*list)[i].compare(s0, Qt::CaseInsensitive) == 0) {
                idx0 = i;
                break;
            }
        }
        if (idx0 == -1) continue;

        int step = 1;
        if (seeds.size() > 1) {
            int idx1 = -1;
            for (int i = 0; i < list->size(); ++i) {
                if ((*list)[i].compare(s1, Qt::CaseInsensitive) == 0) {
                    idx1 = i;
                    break;
                }
            }
            if (idx1 != -1) {
                step = ((idx1 - idx0) % list->size() + list->size()) % list->size();
                if (step == 0) step = list->size();
            }
        }

        QStringList result;
        result.reserve(count);
        for (int i = 0; i < count; ++i) {
            int idx = (idx0 + step * i) % list->size();
            result.append(matchCase(s0, (*list)[idx]));
        }
        return result;
    }

    // --- Text + trailing number: "Item 1", "Q1", "Week 1" ---
    static QRegularExpression re("^(.*?)(\\d+)$");
    auto m0 = re.match(s0);
    if (m0.hasMatch()) {
        QString prefix = m0.captured(1);
        int start = m0.captured(2).toInt();
        int step = 1;

        if (seeds.size() > 1) {
            auto m1 = re.match(s1);
            if (m1.hasMatch() && m1.captured(1) == prefix) {
                step = m1.captured(2).toInt() - start;
            }
        }

        QStringList result;
        result.reserve(count);
        for (int i = 0; i < count; ++i) {
            result.append(prefix + QString::number(start + step * i));
        }
        return result;
    }

    // --- Fallback: repeat first seed ---
    QStringList result;
    result.reserve(count);
    for (int i = 0; i < count; ++i) {
        result.append(s0);
    }
    return result;
}
