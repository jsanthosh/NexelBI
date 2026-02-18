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

private:
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
    static CellStyle buildCellStyle(const XlsxCellXf& xf,
                                     const std::vector<XlsxFont>& fonts,
                                     const std::vector<XlsxFill>& fills,
                                     int numFmtId);
    static void parseSheet(const QByteArray& xmlData, const QStringList& sharedStrings,
                           const std::vector<CellStyle>& styles, Spreadsheet* sheet);
    static int columnLetterToIndex(const QString& letters);
    static QString mapNumFmtId(int id);
};

#endif // XLSXSERVICE_H
