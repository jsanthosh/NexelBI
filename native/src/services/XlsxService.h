#ifndef XLSXSERVICE_H
#define XLSXSERVICE_H

#include <QString>
#include <QStringList>
#include <QColor>
#include <memory>
#include <vector>
#include "../core/Spreadsheet.h"
#include "../core/Cell.h"

class XlsxService {
public:
    // Returns a vector of sheets (one per worksheet in the xlsx file)
    static std::vector<std::shared_ptr<Spreadsheet>> importFromFile(const QString& filePath);

    // Export sheets to XLSX with all formatting
    static bool exportToFile(const std::vector<std::shared_ptr<Spreadsheet>>& sheets, const QString& filePath);

private:
    // Export helpers
    static QString columnIndexToLetter(int col);
    static QByteArray generateContentTypes(int sheetCount);
    static QByteArray generateRels();
    static QByteArray generateWorkbook(const std::vector<std::shared_ptr<Spreadsheet>>& sheets);
    static QByteArray generateWorkbookRels(int sheetCount);
    static QByteArray generateStyles(const std::vector<std::shared_ptr<Spreadsheet>>& sheets,
                                      std::map<QString, int>& styleIndexMap);
    static QByteArray generateSheet(Spreadsheet* sheet, const std::map<QString, int>& styleIndexMap,
                                     QStringList& sharedStrings);
    static QByteArray generateSharedStrings(const QStringList& sharedStrings);
    static QString cellStyleKey(const CellStyle& style);

    struct SheetInfo {
        QString name;
        QString rId;
        QString filePath;  // e.g. "worksheets/sheet1.xml"
    };

    struct XlsxFont {
        QString name = "Arial";
        int size = 11;
        bool bold = false;
        bool italic = false;
        bool underline = false;
        bool strikethrough = false;
        QColor color = QColor("#000000");
    };

    struct XlsxFill {
        QColor fgColor = QColor("#FFFFFF");
        bool hasFg = false;
    };

    struct XlsxCellXf {
        int fontId = 0;
        int fillId = 0;
        int numFmtId = 0;
        HorizontalAlignment hAlign = HorizontalAlignment::General;
        VerticalAlignment vAlign = VerticalAlignment::Bottom;
        bool applyFont = false;
        bool applyFill = false;
        bool applyAlignment = false;
        bool applyNumberFormat = false;
    };

    static QStringList parseSharedStrings(const QByteArray& xmlData);
    static std::vector<SheetInfo> parseWorkbook(const QByteArray& workbookXml, const QByteArray& relsXml);
    static std::vector<XlsxFont> parseFonts(const QByteArray& stylesXml);
    static std::vector<XlsxFill> parseFills(const QByteArray& stylesXml);
    static std::vector<XlsxCellXf> parseCellXfs(const QByteArray& stylesXml);
    static std::map<int, QString> parseNumFmts(const QByteArray& stylesXml);
    static CellStyle buildCellStyle(const XlsxCellXf& xf,
                                     const std::vector<XlsxFont>& fonts,
                                     const std::vector<XlsxFill>& fills,
                                     int numFmtId,
                                     const std::map<int, QString>& customNumFmts);
    static void parseSheet(const QByteArray& xmlData, const QStringList& sharedStrings,
                           const std::vector<CellStyle>& styles, Spreadsheet* sheet);
    static int columnLetterToIndex(const QString& letters);
    static QString mapNumFmtId(int id, const std::map<int, QString>& customNumFmts);
    static bool isDateFormatCode(const QString& formatCode);
};

#endif // XLSXSERVICE_H
