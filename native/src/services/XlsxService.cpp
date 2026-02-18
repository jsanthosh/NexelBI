#include "XlsxService.h"
#include <QFile>
#include <QXmlStreamReader>
#include <QRegularExpression>
#include <QDate>
#include <QtCore/private/qzipreader_p.h>
#include <algorithm>

std::vector<std::shared_ptr<Spreadsheet>> XlsxService::importFromFile(const QString& filePath) {
    std::vector<std::shared_ptr<Spreadsheet>> sheets;

    QZipReader zip(filePath);
    if (!zip.isReadable()) {
        return sheets;
    }

    // Read shared strings
    QStringList sharedStrings;
    QByteArray ssData = zip.fileData("xl/sharedStrings.xml");
    if (!ssData.isEmpty()) {
        sharedStrings = parseSharedStrings(ssData);
    }

    // Parse styles
    std::vector<CellStyle> styles;
    QByteArray stylesData = zip.fileData("xl/styles.xml");
    if (!stylesData.isEmpty()) {
        auto fonts = parseFonts(stylesData);
        auto fills = parseFills(stylesData);
        auto cellXfs = parseCellXfs(stylesData);
        for (const auto& xf : cellXfs) {
            styles.push_back(buildCellStyle(xf, fonts, fills, xf.numFmtId));
        }
    }

    // Parse workbook for sheet names and paths
    QByteArray workbookData = zip.fileData("xl/workbook.xml");
    QByteArray relsData = zip.fileData("xl/_rels/workbook.xml.rels");
    auto sheetInfos = parseWorkbook(workbookData, relsData);

    if (sheetInfos.empty()) {
        // Fallback: try sheet1.xml directly
        SheetInfo si;
        si.name = "Sheet1";
        si.filePath = "worksheets/sheet1.xml";
        sheetInfos.push_back(si);
    }

    // Parse each sheet
    for (const auto& info : sheetInfos) {
        QString path = "xl/" + info.filePath;
        QByteArray sheetData = zip.fileData(path);
        if (sheetData.isEmpty()) continue;

        auto spreadsheet = std::make_shared<Spreadsheet>();
        spreadsheet->setAutoRecalculate(false);
        spreadsheet->setSheetName(info.name);

        parseSheet(sheetData, sharedStrings, styles, spreadsheet.get());

        // Auto-expand row/col count
        int maxRow = spreadsheet->getMaxRow();
        int maxCol = spreadsheet->getMaxColumn();
        spreadsheet->setRowCount(std::max(1000, maxRow + 100));
        spreadsheet->setColumnCount(std::max(256, maxCol + 10));

        spreadsheet->setAutoRecalculate(true);
        sheets.push_back(spreadsheet);
    }

    zip.close();
    return sheets;
}

std::vector<XlsxService::SheetInfo> XlsxService::parseWorkbook(const QByteArray& workbookXml,
                                                                  const QByteArray& relsXml) {
    std::vector<SheetInfo> sheets;

    // Parse relationships: rId -> target path
    std::map<QString, QString> relMap;
    if (!relsXml.isEmpty()) {
        QXmlStreamReader relsReader(relsXml);
        while (!relsReader.atEnd()) {
            relsReader.readNext();
            if (relsReader.isStartElement() && relsReader.name() == u"Relationship") {
                QString id = relsReader.attributes().value("Id").toString();
                QString target = relsReader.attributes().value("Target").toString();
                relMap[id] = target;
            }
        }
    }

    // Parse workbook: sheet names and rIds
    if (!workbookXml.isEmpty()) {
        QXmlStreamReader xml(workbookXml);
        while (!xml.atEnd()) {
            xml.readNext();
            if (xml.isStartElement() && xml.name() == u"sheet") {
                SheetInfo info;
                info.name = xml.attributes().value("name").toString();
                info.rId = xml.attributes().value(u"r:id").toString();
                // Try r:id first, then Id
                if (info.rId.isEmpty()) {
                    // Some xlsx files use different namespace
                    for (const auto& attr : xml.attributes()) {
                        if (attr.name() == u"id" || attr.qualifiedName().toString().contains("id", Qt::CaseInsensitive)) {
                            info.rId = attr.value().toString();
                            break;
                        }
                    }
                }
                auto it = relMap.find(info.rId);
                if (it != relMap.end()) {
                    info.filePath = it->second;
                }
                if (!info.filePath.isEmpty()) {
                    sheets.push_back(info);
                }
            }
        }
    }

    return sheets;
}

QStringList XlsxService::parseSharedStrings(const QByteArray& xmlData) {
    QStringList strings;
    QXmlStreamReader xml(xmlData);

    QString currentString;
    bool inSi = false;

    while (!xml.atEnd()) {
        xml.readNext();
        if (xml.isStartElement()) {
            if (xml.name() == u"si") {
                inSi = true;
                currentString.clear();
            }
        } else if (xml.isCharacters() && inSi) {
            currentString += xml.text().toString();
        } else if (xml.isEndElement()) {
            if (xml.name() == u"si") {
                strings.append(currentString);
                inSi = false;
            }
        }
    }

    return strings;
}

std::vector<XlsxService::XlsxFont> XlsxService::parseFonts(const QByteArray& stylesXml) {
    std::vector<XlsxFont> fonts;
    QXmlStreamReader xml(stylesXml);

    bool inFonts = false;
    bool inFont = false;
    XlsxFont currentFont;

    while (!xml.atEnd()) {
        xml.readNext();
        if (xml.isStartElement()) {
            if (xml.name() == u"fonts") {
                inFonts = true;
            } else if (inFonts && xml.name() == u"font") {
                inFont = true;
                currentFont = XlsxFont();
            } else if (inFont) {
                if (xml.name() == u"name") {
                    currentFont.name = xml.attributes().value("val").toString();
                } else if (xml.name() == u"sz") {
                    currentFont.size = xml.attributes().value("val").toInt();
                } else if (xml.name() == u"b") {
                    QString val = xml.attributes().value("val").toString();
                    currentFont.bold = (val.isEmpty() || val != "0");
                } else if (xml.name() == u"i") {
                    QString val = xml.attributes().value("val").toString();
                    currentFont.italic = (val.isEmpty() || val != "0");
                } else if (xml.name() == u"u") {
                    currentFont.underline = true;
                } else if (xml.name() == u"strike") {
                    currentFont.strikethrough = true;
                } else if (xml.name() == u"color") {
                    QString rgb = xml.attributes().value("rgb").toString();
                    if (!rgb.isEmpty()) {
                        // ARGB format: strip alpha if 8 chars
                        if (rgb.length() == 8) rgb = rgb.mid(2);
                        currentFont.color = QColor("#" + rgb);
                    }
                }
            }
        } else if (xml.isEndElement()) {
            if (xml.name() == u"font" && inFonts) {
                fonts.push_back(currentFont);
                inFont = false;
            } else if (xml.name() == u"fonts") {
                inFonts = false;
                break;  // done with fonts section
            }
        }
    }

    return fonts;
}

std::vector<XlsxService::XlsxFill> XlsxService::parseFills(const QByteArray& stylesXml) {
    std::vector<XlsxFill> fills;
    QXmlStreamReader xml(stylesXml);

    bool inFills = false;
    bool inFill = false;
    XlsxFill currentFill;

    while (!xml.atEnd()) {
        xml.readNext();
        if (xml.isStartElement()) {
            if (xml.name() == u"fills") {
                inFills = true;
            } else if (inFills && xml.name() == u"fill") {
                inFill = true;
                currentFill = XlsxFill();
            } else if (inFill && xml.name() == u"fgColor") {
                QString rgb = xml.attributes().value("rgb").toString();
                if (!rgb.isEmpty()) {
                    if (rgb.length() == 8) rgb = rgb.mid(2);
                    currentFill.fgColor = QColor("#" + rgb);
                    currentFill.hasFg = true;
                }
            }
        } else if (xml.isEndElement()) {
            if (xml.name() == u"fill" && inFills) {
                fills.push_back(currentFill);
                inFill = false;
            } else if (xml.name() == u"fills") {
                inFills = false;
                break;
            }
        }
    }

    return fills;
}

std::vector<XlsxService::XlsxCellXf> XlsxService::parseCellXfs(const QByteArray& stylesXml) {
    std::vector<XlsxCellXf> xfs;
    QXmlStreamReader xml(stylesXml);

    bool inCellXfs = false;
    bool inXf = false;
    XlsxCellXf currentXf;

    while (!xml.atEnd()) {
        xml.readNext();
        if (xml.isStartElement()) {
            if (xml.name() == u"cellXfs") {
                inCellXfs = true;
            } else if (inCellXfs && xml.name() == u"xf") {
                inXf = true;
                currentXf = XlsxCellXf();
                currentXf.fontId = xml.attributes().value("fontId").toInt();
                currentXf.fillId = xml.attributes().value("fillId").toInt();
                currentXf.numFmtId = xml.attributes().value("numFmtId").toInt();
                currentXf.applyFont = (xml.attributes().value("applyFont") == u"1");
                currentXf.applyFill = (xml.attributes().value("applyFill") == u"1");
                currentXf.applyAlignment = (xml.attributes().value("applyAlignment") == u"1");
                currentXf.applyNumberFormat = (xml.attributes().value("applyNumberFormat") == u"1");
            } else if (inXf && xml.name() == u"alignment") {
                QString h = xml.attributes().value("horizontal").toString();
                QString v = xml.attributes().value("vertical").toString();
                if (h == "left") currentXf.hAlign = HorizontalAlignment::Left;
                else if (h == "center") currentXf.hAlign = HorizontalAlignment::Center;
                else if (h == "right") currentXf.hAlign = HorizontalAlignment::Right;
                else currentXf.hAlign = HorizontalAlignment::General;

                if (v == "top") currentXf.vAlign = VerticalAlignment::Top;
                else if (v == "center") currentXf.vAlign = VerticalAlignment::Middle;
                else currentXf.vAlign = VerticalAlignment::Bottom;
                currentXf.applyAlignment = true;
            }
        } else if (xml.isEndElement()) {
            if (xml.name() == u"xf" && inCellXfs) {
                xfs.push_back(currentXf);
                inXf = false;
            } else if (xml.name() == u"cellXfs") {
                inCellXfs = false;
                break;
            }
        }
    }

    return xfs;
}

CellStyle XlsxService::buildCellStyle(const XlsxCellXf& xf,
                                        const std::vector<XlsxFont>& fonts,
                                        const std::vector<XlsxFill>& fills,
                                        int numFmtId) {
    CellStyle style;

    // Apply font
    if (xf.fontId >= 0 && xf.fontId < static_cast<int>(fonts.size())) {
        const auto& f = fonts[xf.fontId];
        style.fontName = f.name;
        style.fontSize = f.size;
        style.bold = f.bold;
        style.italic = f.italic;
        style.underline = f.underline;
        style.strikethrough = f.strikethrough;
        style.foregroundColor = f.color.name();
    }

    // Apply fill
    if (xf.fillId >= 0 && xf.fillId < static_cast<int>(fills.size())) {
        const auto& fl = fills[xf.fillId];
        if (fl.hasFg) {
            style.backgroundColor = fl.fgColor.name();
        }
    }

    // Apply alignment
    style.hAlign = xf.hAlign;
    style.vAlign = xf.vAlign;

    // Apply number format
    style.numberFormat = mapNumFmtId(numFmtId);

    return style;
}

QString XlsxService::mapNumFmtId(int id) {
    // XLSX built-in number format IDs
    if (id == 0) return "General";
    if (id >= 1 && id <= 4) return "Number";
    if (id >= 5 && id <= 8) return "Currency";
    if (id >= 9 && id <= 10) return "Percentage";
    if (id >= 14 && id <= 22) return "Date";
    if (id >= 45 && id <= 47) return "Time";
    if (id == 49) return "Text";
    // Custom formats (id >= 164) default to General
    return "General";
}

void XlsxService::parseSheet(const QByteArray& xmlData, const QStringList& sharedStrings,
                              const std::vector<CellStyle>& styles, Spreadsheet* sheet) {
    QXmlStreamReader xml(xmlData);

    static QRegularExpression cellRefRe("^([A-Z]+)(\\d+)$");

    while (!xml.atEnd()) {
        xml.readNext();

        if (xml.isStartElement() && xml.name() == u"c") {
            QString ref = xml.attributes().value("r").toString();
            QString type = xml.attributes().value("t").toString();
            int styleIdx = xml.attributes().value("s").toInt();

            auto match = cellRefRe.match(ref);
            if (!match.hasMatch()) continue;

            int col = columnLetterToIndex(match.captured(1));
            int row = match.captured(2).toInt() - 1; // XLSX is 1-based

            // Read child elements: <f> (formula), <v> (value), <is> (inline string)
            QString value;
            QString formula;
            QString inlineStr;
            bool hasInlineStr = false;
            int depth = 1; // track nesting depth within <c>

            while (!xml.atEnd() && depth > 0) {
                xml.readNext();
                if (xml.isStartElement()) {
                    if (xml.name() == u"v") {
                        value = xml.readElementText();
                    } else if (xml.name() == u"f") {
                        formula = xml.readElementText();
                    } else if (xml.name() == u"is") {
                        // Inline string - read <t> children
                        hasInlineStr = true;
                    } else if (hasInlineStr && xml.name() == u"t") {
                        inlineStr += xml.readElementText();
                    } else {
                        depth++;
                    }
                } else if (xml.isEndElement()) {
                    if (xml.name() == u"c") {
                        depth = 0; // exit
                    } else if (xml.name() == u"is") {
                        hasInlineStr = false;
                    }
                }
            }

            CellAddress addr(row, col);
            bool cellSet = false;

            // Handle inline strings first (type="inlineStr")
            if (type == "inlineStr" && !inlineStr.isEmpty()) {
                sheet->setCellValue(addr, inlineStr);
                cellSet = true;
            }
            // Handle shared string reference
            else if (type == "s" && !value.isEmpty()) {
                int ssIdx = value.toInt();
                if (ssIdx >= 0 && ssIdx < sharedStrings.size()) {
                    sheet->setCellValue(addr, sharedStrings[ssIdx]);
                    cellSet = true;
                }
            }
            // Handle boolean
            else if (type == "b" && !value.isEmpty()) {
                sheet->setCellValue(addr, value == "1" ? "TRUE" : "FALSE");
                cellSet = true;
            }
            // Handle string formula result (type="str")
            else if (type == "str" && !value.isEmpty()) {
                sheet->setCellValue(addr, value);
                cellSet = true;
            }
            // Handle numeric / date values
            else if (!value.isEmpty()) {
                bool ok;
                double num = value.toDouble(&ok);
                if (ok) {
                    // Check if this cell has a Date/Time format - convert serial to date string
                    bool isDateFmt = false;
                    if (styleIdx > 0 && styleIdx < static_cast<int>(styles.size())) {
                        const auto& fmt = styles[styleIdx].numberFormat;
                        isDateFmt = (fmt == "Date" || fmt == "Time");
                    }
                    if (isDateFmt && num > 0 && num < 2958466) {
                        // Valid Excel serial date (1 to Dec 31, 9999)
                        // Convert to date string so it displays correctly
                        QDate epoch(1899, 12, 30);
                        QDate date = epoch.addDays(static_cast<qint64>(num));
                        if (date.isValid()) {
                            sheet->setCellValue(addr, date.toString("MM/dd/yyyy"));
                        } else {
                            sheet->setCellValue(addr, num);
                        }
                    } else {
                        sheet->setCellValue(addr, num);
                        // If style says Date but value is not a valid serial, override to General
                        if (isDateFmt) {
                            if (styleIdx > 0 && styleIdx < static_cast<int>(styles.size())) {
                                // We'll apply a modified style below
                            }
                        }
                    }
                } else {
                    sheet->setCellValue(addr, value);
                }
                cellSet = true;
            }

            // Apply style (even for cells without values, e.g. styled empty cells)
            if ((cellSet || styleIdx > 0) && styleIdx < static_cast<int>(styles.size())) {
                auto cell = sheet->getCell(addr);
                if (cell) {
                    CellStyle cellStyle = styles[styleIdx];
                    // If date format was applied to a non-date value, override to General
                    if ((cellStyle.numberFormat == "Date" || cellStyle.numberFormat == "Time")) {
                        bool ok;
                        double num = value.toDouble(&ok);
                        if (ok && (num < 0 || num >= 2958466)) {
                            cellStyle.numberFormat = "General";
                        }
                    }
                    cell->setStyle(cellStyle);
                }
            }
        }
    }
}

int XlsxService::columnLetterToIndex(const QString& letters) {
    int result = 0;
    for (int i = 0; i < letters.length(); ++i) {
        result = result * 26 + (letters[i].unicode() - 'A' + 1);
    }
    return result - 1;
}
