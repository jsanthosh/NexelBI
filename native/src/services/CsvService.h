#ifndef CSVSERVICE_H
#define CSVSERVICE_H

#include <QString>
#include <QStringList>
#include <memory>
#include "../core/Spreadsheet.h"

class CsvService {
public:
    static std::shared_ptr<Spreadsheet> importFromFile(const QString& filePath);
    static bool exportToFile(const Spreadsheet& spreadsheet, const QString& filePath);

private:
    static QStringList parseCsvLine(const QString& line);
};

#endif // CSVSERVICE_H
