#include "CsvService.h"
#include <QFile>
#include <QTextStream>

std::shared_ptr<Spreadsheet> CsvService::importFromFile(const QString& filePath) {
    QFile file(filePath);
    if (!file.open(QIODevice::ReadOnly | QIODevice::Text)) {
        return nullptr;
    }

    auto spreadsheet = std::make_shared<Spreadsheet>();
    spreadsheet->setAutoRecalculate(false);

    QTextStream stream(&file);
    int row = 0;

    while (!stream.atEnd()) {
        QString line = stream.readLine();
        QStringList fields = parseCsvLine(line);

        for (int col = 0; col < fields.size(); ++col) {
            QString value = fields[col].trimmed();
            if (value.isEmpty()) continue;

            CellAddress addr(row, col);

            // Try to detect numeric values
            bool ok = false;
            double numValue = value.toDouble(&ok);
            if (ok) {
                spreadsheet->setCellValue(addr, numValue);
            } else if (value.startsWith('=')) {
                spreadsheet->setCellFormula(addr, value);
            } else {
                spreadsheet->setCellValue(addr, value);
            }
        }
        row++;
    }

    file.close();

    // Auto-expand row/column count to fit imported data
    int maxRow = spreadsheet->getMaxRow();
    int maxCol = spreadsheet->getMaxColumn();
    spreadsheet->setRowCount(std::max(spreadsheet->getRowCount(), maxRow + 100));
    spreadsheet->setColumnCount(std::max(spreadsheet->getColumnCount(), maxCol + 10));

    spreadsheet->setAutoRecalculate(true);
    return spreadsheet;
}

bool CsvService::exportToFile(const Spreadsheet& spreadsheet, const QString& filePath) {
    QFile file(filePath);
    if (!file.open(QIODevice::WriteOnly | QIODevice::Text)) {
        return false;
    }

    QTextStream stream(&file);
    int maxRow = spreadsheet.getMaxRow();
    int maxCol = spreadsheet.getMaxColumn();

    for (int r = 0; r <= maxRow; ++r) {
        QStringList rowData;
        for (int c = 0; c <= maxCol; ++c) {
            CellAddress addr(r, c);
            QVariant value = const_cast<Spreadsheet&>(spreadsheet).getCellValue(addr);
            QString str = value.toString();

            // Quote fields containing commas, quotes, or newlines
            if (str.contains(',') || str.contains('"') || str.contains('\n')) {
                str.replace('"', "\"\"");
                str = '"' + str + '"';
            }
            rowData.append(str);
        }

        // Remove trailing empty fields
        while (!rowData.isEmpty() && rowData.last().isEmpty()) {
            rowData.removeLast();
        }

        stream << rowData.join(',') << "\n";
    }

    file.close();
    return true;
}

// RFC 4180 compliant CSV line parser
QStringList CsvService::parseCsvLine(const QString& line) {
    QStringList fields;
    QString field;
    bool inQuotes = false;

    for (int i = 0; i < line.length(); ++i) {
        QChar ch = line[i];

        if (inQuotes) {
            if (ch == '"') {
                if (i + 1 < line.length() && line[i + 1] == '"') {
                    field += '"';
                    i++; // Skip escaped quote
                } else {
                    inQuotes = false;
                }
            } else {
                field += ch;
            }
        } else {
            if (ch == '"') {
                inQuotes = true;
            } else if (ch == ',') {
                fields.append(field);
                field.clear();
            } else {
                field += ch;
            }
        }
    }
    fields.append(field);

    return fields;
}
