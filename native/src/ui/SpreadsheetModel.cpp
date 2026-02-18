#include "SpreadsheetModel.h"
#include "../core/Spreadsheet.h"
#include "../core/NumberFormat.h"
#include "../core/UndoManager.h"
#include "../core/TableStyle.h"
#include "../core/ConditionalFormatting.h"
#include <QFont>
#include <QMessageBox>
#include <QApplication>
#include <QTimer>
#include <QColor>

SpreadsheetModel::SpreadsheetModel(std::shared_ptr<Spreadsheet> spreadsheet, QObject* parent)
    : QAbstractTableModel(parent), m_spreadsheet(spreadsheet) {
}

int SpreadsheetModel::rowCount(const QModelIndex& parent) const {
    if (parent.isValid()) return 0;
    return m_spreadsheet ? m_spreadsheet->getRowCount() : 100;
}

int SpreadsheetModel::columnCount(const QModelIndex& parent) const {
    if (parent.isValid()) return 0;
    return m_spreadsheet ? m_spreadsheet->getColumnCount() : 26;
}

QVariant SpreadsheetModel::data(const QModelIndex& index, int role) const {
    if (!index.isValid() || !m_spreadsheet) {
        return QVariant();
    }

    auto cell = m_spreadsheet->getCell(CellAddress(index.row(), index.column()));
    if (!cell) {
        return QVariant();
    }

    switch (role) {
        case Qt::DisplayRole: {
            auto value = m_spreadsheet->getCellValue(CellAddress(index.row(), index.column()));
            QString text = value.toString();
            // Apply number formatting for display
            const auto& style = cell->getStyle();
            if (style.numberFormat != "General" && !text.isEmpty()) {
                NumberFormatOptions opts;
                opts.type = NumberFormat::typeFromString(style.numberFormat);
                opts.decimalPlaces = style.decimalPlaces;
                opts.useThousandsSeparator = style.useThousandsSeparator;
                opts.currencyCode = style.currencyCode;
                opts.dateFormatId = style.dateFormatId;
                return NumberFormat::format(text, opts);
            }
            return value;
        }
        case Qt::EditRole: {
            // Return raw value for editing
            return m_spreadsheet->getCellValue(CellAddress(index.row(), index.column()));
        }
        case Qt::FontRole: {
            const auto& baseStyle = cell->getStyle();
            // Apply conditional formatting
            CellAddress addr(index.row(), index.column());
            auto cellValue = m_spreadsheet->getCellValue(addr);
            CellStyle style = m_spreadsheet->getConditionalFormatting().getEffectiveStyle(addr, cellValue, baseStyle);

            QFont font(style.fontName);
            font.setPointSize(style.fontSize);
            font.setBold(style.bold);
            font.setItalic(style.italic);
            font.setUnderline(style.underline);
            font.setStrikeOut(style.strikethrough);
            // Table header row: force bold
            auto* table = m_spreadsheet->getTableAt(index.row(), index.column());
            if (table && table->hasHeaderRow && index.row() == table->range.getStart().row) {
                font.setBold(true);
            }
            return font;
        }
        case Qt::ForegroundRole: {
            const auto& baseStyle = cell->getStyle();
            CellAddress addr(index.row(), index.column());
            auto cellValue = m_spreadsheet->getCellValue(addr);
            CellStyle style = m_spreadsheet->getConditionalFormatting().getEffectiveStyle(addr, cellValue, baseStyle);
            // Table header row: use header foreground
            auto* table = m_spreadsheet->getTableAt(index.row(), index.column());
            if (table && table->hasHeaderRow && index.row() == table->range.getStart().row) {
                return table->theme.headerFg;
            }
            return QColor(style.foregroundColor);
        }
        case Qt::BackgroundRole: {
            const auto& baseStyle = cell->getStyle();
            CellAddress addr(index.row(), index.column());
            auto cellValue = m_spreadsheet->getCellValue(addr);
            CellStyle style = m_spreadsheet->getConditionalFormatting().getEffectiveStyle(addr, cellValue, baseStyle);
            // Check if cell is in a table
            auto* table = m_spreadsheet->getTableAt(index.row(), index.column());
            if (table) {
                int tableStartRow = table->range.getStart().row;
                if (table->hasHeaderRow && index.row() == tableStartRow) {
                    return table->theme.headerBg;
                }
                int dataRow = index.row() - tableStartRow - (table->hasHeaderRow ? 1 : 0);
                if (table->bandedRows) {
                    return (dataRow % 2 == 0) ? table->theme.bandedRow1 : table->theme.bandedRow2;
                }
                return table->theme.bandedRow1;
            }
            // Highlight invalid cells with red tint
            if (m_highlightInvalid) {
                QString cellText = m_spreadsheet->getCellValue(addr).toString();
                if (!cellText.isEmpty() && !m_spreadsheet->validateCell(index.row(), index.column(), cellText)) {
                    return QColor(255, 200, 200); // Light red
                }
            }
            return QColor(style.backgroundColor);
        }
        case Qt::TextAlignmentRole: {
            const auto& style = cell->getStyle();

            // Vertical alignment
            int alignment = 0;
            if (style.vAlign == VerticalAlignment::Top) {
                alignment |= Qt::AlignTop;
            } else if (style.vAlign == VerticalAlignment::Bottom) {
                alignment |= Qt::AlignBottom;
            } else {
                alignment |= Qt::AlignVCenter;
            }

            // Horizontal alignment
            if (style.hAlign == HorizontalAlignment::Left) {
                alignment |= Qt::AlignLeft;
            } else if (style.hAlign == HorizontalAlignment::Right) {
                alignment |= Qt::AlignRight;
            } else if (style.hAlign == HorizontalAlignment::Center) {
                alignment |= Qt::AlignHCenter;
            } else {
                // General: right-align numbers, left-align text
                auto value = m_spreadsheet->getCellValue(CellAddress(index.row(), index.column()));
                bool isNumber = false;
                if (value.typeId() == QMetaType::Double || value.typeId() == QMetaType::Int ||
                    value.typeId() == QMetaType::LongLong || value.typeId() == QMetaType::Float) {
                    isNumber = true;
                } else {
                    value.toString().toDouble(&isNumber);
                }
                alignment |= isNumber ? Qt::AlignRight : Qt::AlignLeft;
            }
            return alignment;
        }
        default:
            return QVariant();
    }
}

QVariant SpreadsheetModel::headerData(int section, Qt::Orientation orientation, int role) const {
    if (role == Qt::DisplayRole) {
        if (orientation == Qt::Horizontal) {
            return columnIndexToLetter(section);
        } else {
            return section + 1;
        }
    }

    if (role == Qt::TextAlignmentRole) {
        return Qt::AlignCenter;
    }

    return QVariant();
}

Qt::ItemFlags SpreadsheetModel::flags(const QModelIndex& index) const {
    if (!index.isValid()) {
        return Qt::NoItemFlags;
    }
    return Qt::ItemIsEnabled | Qt::ItemIsSelectable | Qt::ItemIsEditable;
}

bool SpreadsheetModel::setData(const QModelIndex& index, const QVariant& value, int role) {
    if (!index.isValid() || role != Qt::EditRole || !m_spreadsheet) {
        return false;
    }

    CellAddress addr(index.row(), index.column());
    QString strValue = value.toString();

    // Data validation check (skip for formulas)
    if (!strValue.startsWith("=") && !m_suppressUndo) {
        if (!m_spreadsheet->validateCell(index.row(), index.column(), strValue)) {
            const auto* rule = m_spreadsheet->getValidationAt(index.row(), index.column());
            if (rule && rule->showErrorAlert) {
                QString errorMsg = rule->errorMessage.isEmpty()
                    ? "The value you entered is not valid.\nA user has restricted values that can be entered into this cell."
                    : rule->errorMessage;
                QString errorTitle = rule->errorTitle.isEmpty()
                    ? "Invalid Input" : rule->errorTitle;
                auto errorStyle = rule->errorStyle;

                // Defer the dialog to avoid re-entrant event loop crash
                // (showing a modal dialog inside setData() destroys the editor mid-call)
                QTimer::singleShot(0, qApp, [errorTitle, errorMsg, errorStyle]() {
                    QWidget* parent = QApplication::activeWindow();
                    if (errorStyle == Spreadsheet::DataValidationRule::Stop) {
                        QMessageBox::critical(parent, errorTitle, errorMsg);
                    } else if (errorStyle == Spreadsheet::DataValidationRule::Warning) {
                        QMessageBox::warning(parent, errorTitle, errorMsg);
                    } else {
                        QMessageBox::information(parent, errorTitle, errorMsg);
                    }
                });

                // For Stop style, reject the value; for Warning/Info, allow it
                if (rule->errorStyle == Spreadsheet::DataValidationRule::Stop) {
                    return false;
                }
            }
        }
    }

    if (!m_suppressUndo) {
        // Single-cell edit: capture before/after for undo
        CellSnapshot before = m_spreadsheet->takeCellSnapshot(addr);

        if (strValue.startsWith("=")) {
            m_spreadsheet->setCellFormula(addr, strValue);
        } else {
            m_spreadsheet->setCellValue(addr, value);
        }

        CellSnapshot after = m_spreadsheet->takeCellSnapshot(addr);
        m_spreadsheet->getUndoManager().pushCommand(
            std::make_unique<CellEditCommand>(before, after));
    } else {
        // Bulk operation: caller handles undo tracking
        if (strValue.startsWith("=")) {
            m_spreadsheet->setCellFormula(addr, strValue);
        } else {
            m_spreadsheet->setCellValue(addr, value);
        }
    }

    emit dataChanged(index, index, {Qt::DisplayRole, Qt::EditRole});
    return true;
}

QString SpreadsheetModel::columnIndexToLetter(int column) const {
    QString result;
    while (column >= 0) {
        result = QChar('A' + (column % 26)) + result;
        column = column / 26 - 1;
    }
    return result;
}
