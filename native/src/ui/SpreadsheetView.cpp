#include "SpreadsheetView.h"
#include "SpreadsheetModel.h"
#include "CellDelegate.h"
#include "../core/Spreadsheet.h"
#include "../core/UndoManager.h"
#include "../core/FillSeries.h"
#include "../core/TableStyle.h"
#include "../services/DocumentService.h"
#include <QHeaderView>
#include <QKeyEvent>
#include <QClipboard>
#include <QApplication>
#include <QPainter>
#include <QMenu>
#include <QFontMetrics>
#include <QSet>
#include <QLineEdit>
#include <QDate>
#include <algorithm>

SpreadsheetView::SpreadsheetView(QWidget* parent)
    : QTableView(parent), m_zoomLevel(100) {

    m_spreadsheet = DocumentService::instance().getCurrentSpreadsheet();
    if (!m_spreadsheet) {
        m_spreadsheet = std::make_shared<Spreadsheet>();
    }

    initializeView();
    setupConnections();
}

void SpreadsheetView::setSpreadsheet(std::shared_ptr<Spreadsheet> spreadsheet) {
    m_spreadsheet = spreadsheet;

    if (m_model) {
        delete m_model;
    }
    m_model = new SpreadsheetModel(m_spreadsheet, this);
    setModel(m_model);
}

std::shared_ptr<Spreadsheet> SpreadsheetView::getSpreadsheet() const {
    return m_spreadsheet;
}

void SpreadsheetView::initializeView() {
    m_model = new SpreadsheetModel(m_spreadsheet, this);
    setModel(m_model);

    m_delegate = new CellDelegate(this);
    setItemDelegate(m_delegate);

    // Column/row sizing
    horizontalHeader()->setDefaultSectionSize(80);
    verticalHeader()->setDefaultSectionSize(22);
    horizontalHeader()->setStretchLastSection(false);
    verticalHeader()->setStretchLastSection(false);
    horizontalHeader()->setMinimumSectionSize(30);
    verticalHeader()->setMinimumSectionSize(14);
    horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);
    verticalHeader()->setSectionResizeMode(QHeaderView::Interactive);

    // Delegate handles all painting — disable QTableView gridlines
    setShowGrid(false);
    setSelectionBehavior(QAbstractItemView::SelectItems);
    setSelectionMode(QAbstractItemView::ExtendedSelection);

    setFont(QFont("Arial", 11));

    // Clean, modern stylesheet
    setStyleSheet(
        "QTableView {"
        "   background-color: #ffffff;"
        "   border: none;"
        "   outline: none;"
        "}"
        "QTableView::item {"
        "   padding: 0px;"
        "   border: none;"
        "   background-color: transparent;"
        "}"
        "QTableView::item:selected {"
        "   background-color: transparent;"
        "}"
        "QTableView::item:focus {"
        "   border: none;"
        "   outline: none;"
        "}"
        "QHeaderView::section {"
        "   background-color: #F3F3F3;"
        "   padding: 2px 4px;"
        "   border: none;"
        "   border-right: 1px solid #DADCE0;"
        "   border-bottom: 1px solid #DADCE0;"
        "   font-size: 11px;"
        "   color: #333333;"
        "}"
        "QHeaderView {"
        "   background-color: #F3F3F3;"
        "}"
        "QTableCornerButton::section {"
        "   background-color: #F3F3F3;"
        "   border: none;"
        "   border-right: 1px solid #DADCE0;"
        "   border-bottom: 1px solid #DADCE0;"
        "}"
    );

    // Enable mouse tracking for fill handle cursor changes
    viewport()->setMouseTracking(true);

    // Cell context menu (right-click)
    setContextMenuPolicy(Qt::CustomContextMenu);
    connect(this, &QTableView::customContextMenuRequested,
            this, &SpreadsheetView::showCellContextMenu);

    // Setup header context menus
    setupHeaderContextMenus();
}

void SpreadsheetView::setupConnections() {
    connect(this, &QTableView::clicked, this, &SpreadsheetView::onCellClicked);
    connect(this, &QTableView::doubleClicked, this, &SpreadsheetView::onCellDoubleClicked);
    if (m_model) {
        connect(m_model, &QAbstractTableModel::dataChanged, this, &SpreadsheetView::onDataChanged);
    }

    // Formula edit mode from cell editor
    if (m_delegate) {
        connect(m_delegate, &CellDelegate::formulaEditModeChanged,
                this, &SpreadsheetView::setFormulaEditMode);
    }

    // Multi-select resize
    connect(horizontalHeader(), &QHeaderView::sectionResized,
            this, &SpreadsheetView::onHorizontalSectionResized);
    connect(verticalHeader(), &QHeaderView::sectionResized,
            this, &SpreadsheetView::onVerticalSectionResized);
}

void SpreadsheetView::setupHeaderContextMenus() {
    horizontalHeader()->setContextMenuPolicy(Qt::CustomContextMenu);
    connect(horizontalHeader(), &QHeaderView::customContextMenuRequested,
            this, [this](const QPoint& pos) {
        QMenu menu;
        menu.addAction("Autofit Column Width", this, &SpreadsheetView::autofitSelectedColumns);
        menu.addSeparator();
        int col = horizontalHeader()->logicalIndexAt(pos);
        menu.addAction("Insert Column", [this, col]() {
            if (m_spreadsheet) {
                m_spreadsheet->insertColumn(col);
                refreshView();
            }
        });
        menu.addAction("Delete Column", [this, col]() {
            if (m_spreadsheet) {
                m_spreadsheet->deleteColumn(col);
                refreshView();
            }
        });
        menu.exec(horizontalHeader()->mapToGlobal(pos));
    });

    verticalHeader()->setContextMenuPolicy(Qt::CustomContextMenu);
    connect(verticalHeader(), &QHeaderView::customContextMenuRequested,
            this, [this](const QPoint& pos) {
        QMenu menu;
        menu.addAction("Autofit Row Height", this, &SpreadsheetView::autofitSelectedRows);
        menu.addSeparator();
        int row = verticalHeader()->logicalIndexAt(pos);
        menu.addAction("Insert Row", [this, row]() {
            if (m_spreadsheet) {
                m_spreadsheet->insertRow(row);
                refreshView();
            }
        });
        menu.addAction("Delete Row", [this, row]() {
            if (m_spreadsheet) {
                m_spreadsheet->deleteRow(row);
                refreshView();
            }
        });
        menu.exec(verticalHeader()->mapToGlobal(pos));
    });
}

void SpreadsheetView::emitCellSelected(const QModelIndex& index) {
    if (!index.isValid() || !m_spreadsheet) return;

    CellAddress addr(index.row(), index.column());
    auto cell = m_spreadsheet->getCell(addr);

    QString content;
    if (cell->getType() == CellType::Formula) {
        content = cell->getFormula();
    } else {
        content = cell->getValue().toString();
    }

    emit cellSelected(index.row(), index.column(), content, addr.toString());
}

void SpreadsheetView::currentChanged(const QModelIndex& current, const QModelIndex& previous) {
    QTableView::currentChanged(current, previous);

    // Force repaint of previous cell to clear its focus border and fill handle
    if (previous.isValid()) {
        QRect prevRect = visualRect(previous);
        // Expand to cover 2px focus border + fill handle (7px square at corner)
        viewport()->update(prevRect.adjusted(-2, -2, 6, 6));
    }
    // Also invalidate the old fill handle rect
    if (!m_fillHandleRect.isNull()) {
        viewport()->update(m_fillHandleRect.adjusted(-2, -2, 2, 2));
    }

    emitCellSelected(current);
}

// ============== Clipboard operations ==============

void SpreadsheetView::cut() {
    copy();
    deleteSelection();
}

void SpreadsheetView::copy() {
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    std::sort(selected.begin(), selected.end(), [](const QModelIndex& a, const QModelIndex& b) {
        if (a.row() != b.row()) return a.row() < b.row();
        return a.column() < b.column();
    });

    QString data;
    int lastRow = selected.first().row();
    bool firstInRow = true;
    for (const auto& index : selected) {
        if (index.row() != lastRow) {
            data += "\n";
            lastRow = index.row();
            firstInRow = true;
        }
        if (!firstInRow) {
            data += "\t";
        }
        data += index.data().toString();
        firstInRow = false;
    }

    QApplication::clipboard()->setText(data);
}

void SpreadsheetView::paste() {
    QString data = QApplication::clipboard()->text();
    QModelIndex current = currentIndex();
    if (!current.isValid() || !m_spreadsheet) return;

    int startRow = current.row();
    int startCol = current.column();

    std::vector<CellSnapshot> before, after;

    m_model->setSuppressUndo(true);
    QStringList rows = data.split("\n");
    for (int r = 0; r < rows.size(); ++r) {
        QStringList cols = rows[r].split("\t");
        for (int c = 0; c < cols.size(); ++c) {
            CellAddress addr(startRow + r, startCol + c);
            before.push_back(m_spreadsheet->takeCellSnapshot(addr));

            QModelIndex index = m_model->index(startRow + r, startCol + c);
            m_model->setData(index, cols[c]);

            after.push_back(m_spreadsheet->takeCellSnapshot(addr));
        }
    }
    m_model->setSuppressUndo(false);

    m_spreadsheet->getUndoManager().pushCommand(
        std::make_unique<MultiCellEditCommand>(before, after, "Paste"));
}

void SpreadsheetView::deleteSelection() {
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty() || !m_spreadsheet) return;

    std::vector<CellSnapshot> before, after;

    m_model->setSuppressUndo(true);
    for (const auto& index : selected) {
        CellAddress addr(index.row(), index.column());
        before.push_back(m_spreadsheet->takeCellSnapshot(addr));
        m_model->setData(index, "");
        after.push_back(m_spreadsheet->takeCellSnapshot(addr));
    }
    m_model->setSuppressUndo(false);

    m_spreadsheet->getUndoManager().pushCommand(
        std::make_unique<MultiCellEditCommand>(before, after, "Delete"));
}

void SpreadsheetView::selectAll() {
    QTableView::selectAll();
}

// ============== Style operations ==============

// Efficient style application: for large selections (select all), only iterate occupied cells
void SpreadsheetView::applyStyleChange(std::function<void(CellStyle&)> modifier, const QList<int>& roles) {
    if (!m_spreadsheet) return;

    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    // For large selections (>5000 cells), only apply to occupied cells
    static constexpr int LARGE_SELECTION_THRESHOLD = 5000;
    bool isLargeSelection = selected.size() > LARGE_SELECTION_THRESHOLD;

    std::vector<CellSnapshot> before, after;

    if (isLargeSelection) {
        // Build a bounding box from selection, then iterate only occupied cells
        int minRow = INT_MAX, maxRow = 0, minCol = INT_MAX, maxCol = 0;
        for (const auto& idx : selected) {
            minRow = qMin(minRow, idx.row());
            maxRow = qMax(maxRow, idx.row());
            minCol = qMin(minCol, idx.column());
            maxCol = qMax(maxCol, idx.column());
        }

        // Use forEachCell to iterate only occupied cells within the selection bounds
        m_spreadsheet->forEachCell([&](int row, int col, const Cell&) {
            if (row < minRow || row > maxRow || col < minCol || col > maxCol) return;

            CellAddress addr(row, col);
            before.push_back(m_spreadsheet->takeCellSnapshot(addr));

            auto cell = m_spreadsheet->getCell(addr);
            CellStyle style = cell->getStyle();
            modifier(style);
            cell->setStyle(style);

            after.push_back(m_spreadsheet->takeCellSnapshot(addr));
        });
    } else {
        for (const auto& index : selected) {
            CellAddress addr(index.row(), index.column());
            before.push_back(m_spreadsheet->takeCellSnapshot(addr));

            auto cell = m_spreadsheet->getCell(addr);
            CellStyle style = cell->getStyle();
            modifier(style);
            cell->setStyle(style);

            after.push_back(m_spreadsheet->takeCellSnapshot(addr));
        }
    }

    if (!before.empty()) {
        m_spreadsheet->getUndoManager().execute(
            std::make_unique<StyleChangeCommand>(before, after), m_spreadsheet.get());
    }

    if (m_model) {
        if (isLargeSelection) {
            m_model->resetModel();
        } else {
            emit m_model->dataChanged(selected.first(), selected.last(), {roles.begin(), roles.end()});
        }
    }
}

void SpreadsheetView::applyBold() {
    applyStyleChange([](CellStyle& s) { s.bold = !s.bold; }, {Qt::FontRole});
}

void SpreadsheetView::applyItalic() {
    applyStyleChange([](CellStyle& s) { s.italic = !s.italic; }, {Qt::FontRole});
}

void SpreadsheetView::applyUnderline() {
    applyStyleChange([](CellStyle& s) { s.underline = !s.underline; }, {Qt::FontRole});
}

void SpreadsheetView::applyStrikethrough() {
    applyStyleChange([](CellStyle& s) { s.strikethrough = !s.strikethrough; }, {Qt::FontRole});
}

void SpreadsheetView::applyFontFamily(const QString& family) {
    applyStyleChange([&family](CellStyle& s) { s.fontName = family; }, {Qt::FontRole});
}

void SpreadsheetView::applyFontSize(int size) {
    applyStyleChange([size](CellStyle& s) { s.fontSize = size; }, {Qt::FontRole});
}

void SpreadsheetView::applyForegroundColor(const QColor& color) {
    applyStyleChange([&color](CellStyle& s) { s.foregroundColor = color.name(); }, {Qt::ForegroundRole});
}

void SpreadsheetView::applyBackgroundColor(const QColor& color) {
    applyStyleChange([&color](CellStyle& s) { s.backgroundColor = color.name(); }, {Qt::BackgroundRole});
}

void SpreadsheetView::applyThousandSeparator() {
    applyStyleChange([](CellStyle& s) {
        s.useThousandsSeparator = !s.useThousandsSeparator;
        if (s.numberFormat == "General") s.numberFormat = "Number";
    }, {Qt::DisplayRole});
}

void SpreadsheetView::applyNumberFormat(const QString& format) {
    applyStyleChange([&format](CellStyle& s) {
        s.numberFormat = format;
    }, {Qt::DisplayRole});
}

// ============== Alignment ==============

void SpreadsheetView::applyHAlign(HorizontalAlignment align) {
    applyStyleChange([align](CellStyle& s) { s.hAlign = align; }, {Qt::TextAlignmentRole});
}

void SpreadsheetView::applyVAlign(VerticalAlignment align) {
    applyStyleChange([align](CellStyle& s) { s.vAlign = align; }, {Qt::TextAlignmentRole});
}

// ============== Format Painter ==============

void SpreadsheetView::activateFormatPainter() {
    QModelIndex current = currentIndex();
    if (!current.isValid() || !m_spreadsheet) return;

    CellAddress addr(current.row(), current.column());
    auto cell = m_spreadsheet->getCell(addr);
    m_copiedStyle = cell->getStyle();
    m_formatPainterActive = true;
    viewport()->setCursor(Qt::CrossCursor);
}

// ============== Sorting ==============

void SpreadsheetView::sortAscending() {
    QModelIndex current = currentIndex();
    if (!current.isValid() || !m_spreadsheet) return;

    int col = current.column();
    int maxRow = m_spreadsheet->getMaxRow();
    int maxCol = m_spreadsheet->getMaxColumn();
    if (maxRow < 1 && maxCol < 1) return;
    if (maxRow < 1) maxRow = 1;

    CellRange range(CellAddress(0, 0), CellAddress(maxRow, qMax(maxCol, col)));
    m_spreadsheet->sortRange(range, col, true);

    // Full model reset to ensure view refreshes completely
    if (m_model) {
        m_model->resetModel();
    }
}

void SpreadsheetView::sortDescending() {
    QModelIndex current = currentIndex();
    if (!current.isValid() || !m_spreadsheet) return;

    int col = current.column();
    int maxRow = m_spreadsheet->getMaxRow();
    int maxCol = m_spreadsheet->getMaxColumn();
    if (maxRow < 1 && maxCol < 1) return;
    if (maxRow < 1) maxRow = 1;

    CellRange range(CellAddress(0, 0), CellAddress(maxRow, qMax(maxCol, col)));
    m_spreadsheet->sortRange(range, col, false);

    if (m_model) {
        m_model->resetModel();
    }
}

// ============== Table Style ==============

CellRange SpreadsheetView::detectDataRegion(int startRow, int startCol) const {
    if (!m_spreadsheet) return CellRange(CellAddress(startRow, startCol), CellAddress(startRow, startCol));

    // Expand outward from the starting cell to find the contiguous data region
    // Similar to Excel's Ctrl+Shift+* or Ctrl+T auto-detection

    int maxRow = m_spreadsheet->getMaxRow();
    int maxCol = m_spreadsheet->getMaxColumn();

    // Find the left boundary
    int left = startCol;
    while (left > 0) {
        bool colHasData = false;
        for (int r = 0; r <= maxRow; ++r) {
            auto val = m_spreadsheet->getCellValue(CellAddress(r, left - 1));
            if (val.isValid() && !val.toString().isEmpty()) {
                colHasData = true;
                break;
            }
        }
        if (!colHasData) break;
        left--;
    }

    // Find the right boundary
    int right = startCol;
    while (right < maxCol) {
        bool colHasData = false;
        for (int r = 0; r <= maxRow; ++r) {
            auto val = m_spreadsheet->getCellValue(CellAddress(r, right + 1));
            if (val.isValid() && !val.toString().isEmpty()) {
                colHasData = true;
                break;
            }
        }
        if (!colHasData) break;
        right++;
    }

    // Find the top boundary
    int top = startRow;
    while (top > 0) {
        bool rowHasData = false;
        for (int c = left; c <= right; ++c) {
            auto val = m_spreadsheet->getCellValue(CellAddress(top - 1, c));
            if (val.isValid() && !val.toString().isEmpty()) {
                rowHasData = true;
                break;
            }
        }
        if (!rowHasData) break;
        top--;
    }

    // Find the bottom boundary
    int bottom = startRow;
    while (bottom < maxRow) {
        bool rowHasData = false;
        for (int c = left; c <= right; ++c) {
            auto val = m_spreadsheet->getCellValue(CellAddress(bottom + 1, c));
            if (val.isValid() && !val.toString().isEmpty()) {
                rowHasData = true;
                break;
            }
        }
        if (!rowHasData) break;
        bottom++;
    }

    return CellRange(CellAddress(top, left), CellAddress(bottom, right));
}

void SpreadsheetView::applyTableStyle(int themeIndex) {
    if (!m_spreadsheet) return;

    auto themes = getBuiltinTableThemes();
    if (themeIndex < 0 || themeIndex >= static_cast<int>(themes.size())) return;

    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    int minRow, maxRow, minCol, maxCol;

    // Auto-detect: if only a single cell is selected, detect the contiguous data region
    if (selected.size() == 1) {
        CellRange region = detectDataRegion(selected.first().row(), selected.first().column());
        minRow = region.getStart().row;
        maxRow = region.getEnd().row;
        minCol = region.getStart().col;
        maxCol = region.getEnd().col;
    } else {
        // Find selection bounding box
        minRow = INT_MAX; maxRow = 0; minCol = INT_MAX; maxCol = 0;
        for (const auto& idx : selected) {
            minRow = qMin(minRow, idx.row());
            maxRow = qMax(maxRow, idx.row());
            minCol = qMin(minCol, idx.column());
            maxCol = qMax(maxCol, idx.column());
        }
    }

    SpreadsheetTable table;
    table.range = CellRange(CellAddress(minRow, minCol), CellAddress(maxRow, maxCol));
    table.theme = themes[themeIndex];
    table.hasHeaderRow = true;
    table.bandedRows = true;

    // Auto-name
    int tableNum = static_cast<int>(m_spreadsheet->getTables().size()) + 1;
    table.name = QString("Table%1").arg(tableNum);

    // Extract column names from header row
    for (int c = minCol; c <= maxCol; ++c) {
        auto val = m_spreadsheet->getCellValue(CellAddress(minRow, c));
        QString name = val.toString();
        if (name.isEmpty()) name = QString("Column%1").arg(c - minCol + 1);
        table.columnNames.append(name);
    }

    m_spreadsheet->addTable(table);
    refreshView();
}

// ============== Context Menu ==============

void SpreadsheetView::showCellContextMenu(const QPoint& pos) {
    QMenu menu(this);

    menu.addAction("Cut", this, &SpreadsheetView::cut, QKeySequence::Cut);
    menu.addAction("Copy", this, &SpreadsheetView::copy, QKeySequence::Copy);
    menu.addAction("Paste", this, &SpreadsheetView::paste, QKeySequence::Paste);

    menu.addSeparator();

    // Insert submenu
    QMenu* insertMenu = menu.addMenu("Insert...");
    insertMenu->addAction("Shift cells right", this, &SpreadsheetView::insertCellsShiftRight);
    insertMenu->addAction("Shift cells down", this, &SpreadsheetView::insertCellsShiftDown);
    insertMenu->addSeparator();
    insertMenu->addAction("Entire row", this, &SpreadsheetView::insertEntireRow);
    insertMenu->addAction("Entire column", this, &SpreadsheetView::insertEntireColumn);

    // Delete submenu
    QMenu* deleteMenu = menu.addMenu("Delete...");
    deleteMenu->addAction("Shift cells left", this, &SpreadsheetView::deleteCellsShiftLeft);
    deleteMenu->addAction("Shift cells up", this, &SpreadsheetView::deleteCellsShiftUp);
    deleteMenu->addSeparator();
    deleteMenu->addAction("Entire row", this, &SpreadsheetView::deleteEntireRow);
    deleteMenu->addAction("Entire column", this, &SpreadsheetView::deleteEntireColumn);

    menu.addSeparator();

    menu.addAction("Format Cells...", this, [this]() {
        emit formatCellsRequested();
    }, QKeySequence(Qt::CTRL | Qt::Key_1));

    menu.addSeparator();

    menu.addAction("Sort Ascending", this, &SpreadsheetView::sortAscending);
    menu.addAction("Sort Descending", this, &SpreadsheetView::sortDescending);

    menu.exec(viewport()->mapToGlobal(pos));
}

// ============== Insert/Delete with shift ==============

void SpreadsheetView::insertCellsShiftRight() {
    if (!m_spreadsheet) return;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    int minRow = selected.first().row(), maxRow = minRow;
    int minCol = selected.first().column(), maxCol = minCol;
    for (const auto& idx : selected) {
        minRow = qMin(minRow, idx.row());
        maxRow = qMax(maxRow, idx.row());
        minCol = qMin(minCol, idx.column());
        maxCol = qMax(maxCol, idx.column());
    }

    CellRange range(CellAddress(minRow, minCol), CellAddress(maxRow, maxCol));
    m_spreadsheet->insertCellsShiftRight(range);
    refreshView();
}

void SpreadsheetView::insertCellsShiftDown() {
    if (!m_spreadsheet) return;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    int minRow = selected.first().row(), maxRow = minRow;
    int minCol = selected.first().column(), maxCol = minCol;
    for (const auto& idx : selected) {
        minRow = qMin(minRow, idx.row());
        maxRow = qMax(maxRow, idx.row());
        minCol = qMin(minCol, idx.column());
        maxCol = qMax(maxCol, idx.column());
    }

    CellRange range(CellAddress(minRow, minCol), CellAddress(maxRow, maxCol));
    m_spreadsheet->insertCellsShiftDown(range);
    refreshView();
}

void SpreadsheetView::insertEntireRow() {
    if (!m_spreadsheet) return;
    QModelIndex current = currentIndex();
    if (!current.isValid()) return;

    // Get all selected rows
    QSet<int> rows;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) {
        rows.insert(current.row());
    } else {
        for (const auto& idx : selected) {
            rows.insert(idx.row());
        }
    }

    // Insert from bottom to top to preserve row indices
    QList<int> sortedRows(rows.begin(), rows.end());
    std::sort(sortedRows.rbegin(), sortedRows.rend());
    for (int row : sortedRows) {
        m_spreadsheet->insertRow(row);
    }
    refreshView();
}

void SpreadsheetView::insertEntireColumn() {
    if (!m_spreadsheet) return;
    QModelIndex current = currentIndex();
    if (!current.isValid()) return;

    QSet<int> cols;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) {
        cols.insert(current.column());
    } else {
        for (const auto& idx : selected) {
            cols.insert(idx.column());
        }
    }

    QList<int> sortedCols(cols.begin(), cols.end());
    std::sort(sortedCols.rbegin(), sortedCols.rend());
    for (int col : sortedCols) {
        m_spreadsheet->insertColumn(col);
    }
    refreshView();
}

void SpreadsheetView::deleteCellsShiftLeft() {
    if (!m_spreadsheet) return;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    int minRow = selected.first().row(), maxRow = minRow;
    int minCol = selected.first().column(), maxCol = minCol;
    for (const auto& idx : selected) {
        minRow = qMin(minRow, idx.row());
        maxRow = qMax(maxRow, idx.row());
        minCol = qMin(minCol, idx.column());
        maxCol = qMax(maxCol, idx.column());
    }

    CellRange range(CellAddress(minRow, minCol), CellAddress(maxRow, maxCol));
    m_spreadsheet->deleteCellsShiftLeft(range);
    refreshView();
}

void SpreadsheetView::deleteCellsShiftUp() {
    if (!m_spreadsheet) return;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    int minRow = selected.first().row(), maxRow = minRow;
    int minCol = selected.first().column(), maxCol = minCol;
    for (const auto& idx : selected) {
        minRow = qMin(minRow, idx.row());
        maxRow = qMax(maxRow, idx.row());
        minCol = qMin(minCol, idx.column());
        maxCol = qMax(maxCol, idx.column());
    }

    CellRange range(CellAddress(minRow, minCol), CellAddress(maxRow, maxCol));
    m_spreadsheet->deleteCellsShiftUp(range);
    refreshView();
}

void SpreadsheetView::deleteEntireRow() {
    if (!m_spreadsheet) return;
    QModelIndex current = currentIndex();
    if (!current.isValid()) return;

    QSet<int> rows;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) {
        rows.insert(current.row());
    } else {
        for (const auto& idx : selected) {
            rows.insert(idx.row());
        }
    }

    // Delete from bottom to top to preserve row indices
    QList<int> sortedRows(rows.begin(), rows.end());
    std::sort(sortedRows.rbegin(), sortedRows.rend());
    for (int row : sortedRows) {
        m_spreadsheet->deleteRow(row);
    }
    refreshView();
}

void SpreadsheetView::deleteEntireColumn() {
    if (!m_spreadsheet) return;
    QModelIndex current = currentIndex();
    if (!current.isValid()) return;

    QSet<int> cols;
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) {
        cols.insert(current.column());
    } else {
        for (const auto& idx : selected) {
            cols.insert(idx.column());
        }
    }

    QList<int> sortedCols(cols.begin(), cols.end());
    std::sort(sortedCols.rbegin(), sortedCols.rend());
    for (int col : sortedCols) {
        m_spreadsheet->deleteColumn(col);
    }
    refreshView();
}

// ============== Autofit ==============

void SpreadsheetView::autofitColumn(int column) {
    if (!m_model) return;
    int maxWidth = 40;
    for (int row = 0; row < m_model->rowCount(); ++row) {
        QModelIndex idx = m_model->index(row, column);
        QString text = idx.data(Qt::DisplayRole).toString();
        if (text.isEmpty()) continue;

        QFont cellFont = font();
        QVariant fontData = idx.data(Qt::FontRole);
        if (fontData.isValid()) {
            cellFont = fontData.value<QFont>();
        }
        QFontMetrics fm(cellFont);
        int width = fm.horizontalAdvance(text) + 16;
        maxWidth = qMax(maxWidth, width);
    }
    horizontalHeader()->resizeSection(column, maxWidth);
}

void SpreadsheetView::autofitRow(int row) {
    if (!m_model) return;
    int maxHeight = 18;
    for (int col = 0; col < m_model->columnCount(); ++col) {
        QModelIndex idx = m_model->index(row, col);
        QString text = idx.data(Qt::DisplayRole).toString();
        if (text.isEmpty()) continue;

        QFont cellFont = font();
        QVariant fontData = idx.data(Qt::FontRole);
        if (fontData.isValid()) {
            cellFont = fontData.value<QFont>();
        }
        QFontMetrics fm(cellFont);
        int height = fm.height() + 6;
        maxHeight = qMax(maxHeight, height);
    }
    verticalHeader()->resizeSection(row, maxHeight);
}

void SpreadsheetView::autofitSelectedColumns() {
    QModelIndexList selected = selectionModel()->selectedColumns();
    if (selected.isEmpty()) {
        autofitColumn(currentIndex().column());
    } else {
        for (const auto& idx : selected) {
            autofitColumn(idx.column());
        }
    }
}

void SpreadsheetView::autofitSelectedRows() {
    QModelIndexList selected = selectionModel()->selectedRows();
    if (selected.isEmpty()) {
        autofitRow(currentIndex().row());
    } else {
        for (const auto& idx : selected) {
            autofitRow(idx.row());
        }
    }
}

// ============== UI Operations ==============

void SpreadsheetView::refreshView() {
    if (m_model) {
        m_model->layoutChanged();
    }
}

void SpreadsheetView::setFrozenRow(int row) {
    if (row < 0) {
        // Unfreeze rows
        for (int r = 0; r < model()->rowCount(); ++r) {
            if (verticalHeader()->sectionPosition(r) >= 0) {
                // No built-in freeze in QTableView; approximate with header resize
            }
        }
    }
    // QTableView doesn't natively support freeze panes.
    // We use a simpler approach: keep headers fixed position (already default behavior)
    // For a true freeze, the user sees the rows above 'row' always fixed.
    // This is a best-effort implementation using QHeaderView.
}

void SpreadsheetView::setFrozenColumn(int col) {
    Q_UNUSED(col);
    // Similar to setFrozenRow - QTableView doesn't natively freeze columns.
    // Headers are always visible, which partially mimics this behavior.
}

void SpreadsheetView::zoomIn() {
    m_zoomLevel += 10;
    if (m_zoomLevel > 200) m_zoomLevel = 200;

    QFont f = this->font();
    f.setPointSize(f.pointSize() * m_zoomLevel / 100);
    setFont(f);
}

void SpreadsheetView::zoomOut() {
    m_zoomLevel -= 10;
    if (m_zoomLevel < 50) m_zoomLevel = 50;

    QFont f = this->font();
    f.setPointSize(f.pointSize() * 100 / m_zoomLevel);
    setFont(f);
}

void SpreadsheetView::resetZoom() {
    m_zoomLevel = 100;
    setFont(QFont("Arial", 11));
}

// ============== Event handlers ==============

void SpreadsheetView::keyPressEvent(QKeyEvent* event) {
    bool ctrl = event->modifiers() & Qt::ControlModifier;
    bool shift = event->modifiers() & Qt::ShiftModifier;

    // Delete / Backspace: clear selection (on Mac, "Delete" key = Backspace)
    if (event->key() == Qt::Key_Delete || event->key() == Qt::Key_Backspace) {
        if (state() != QAbstractItemView::EditingState) {
            deleteSelection();
            event->accept();
            return;
        }
        // If editing, let the editor handle the key
        QTableView::keyPressEvent(event);
        return;
    }

    // Enter/Return: commit and move down
    if (event->key() == Qt::Key_Return || event->key() == Qt::Key_Enter) {
        m_formulaEditMode = false;
        if (state() == QAbstractItemView::EditingState) {
            QWidget* editor = indexWidget(currentIndex());
            if (editor) {
                commitData(editor);
                closeEditor(editor, QAbstractItemDelegate::NoHint);
            }
        }
        int newRow = currentIndex().row() + (shift ? -1 : 1);
        newRow = qBound(0, newRow, model()->rowCount() - 1);
        QModelIndex next = model()->index(newRow, currentIndex().column());
        if (next.isValid()) {
            setCurrentIndex(next);
        }
        event->accept();
        return;
    }

    // Tab: commit and move right; Shift+Tab: move left
    if (event->key() == Qt::Key_Tab || event->key() == Qt::Key_Backtab) {
        if (state() == QAbstractItemView::EditingState) {
            QWidget* editor = indexWidget(currentIndex());
            if (editor) {
                commitData(editor);
                closeEditor(editor, QAbstractItemDelegate::NoHint);
            }
        }
        int newCol = currentIndex().column() + (event->key() == Qt::Key_Backtab ? -1 : 1);
        newCol = qBound(0, newCol, model()->columnCount() - 1);
        QModelIndex next = model()->index(currentIndex().row(), newCol);
        if (next.isValid()) {
            setCurrentIndex(next);
        }
        event->accept();
        return;
    }

    // F2: Edit current cell (like Excel)
    if (event->key() == Qt::Key_F2) {
        QModelIndex current = currentIndex();
        if (current.isValid() && state() != QAbstractItemView::EditingState) {
            edit(current);
        }
        event->accept();
        return;
    }

    // Escape: cancel editing / format painter / formula edit mode
    if (event->key() == Qt::Key_Escape) {
        m_formulaEditMode = false;
        if (m_formatPainterActive) {
            m_formatPainterActive = false;
            viewport()->setCursor(Qt::ArrowCursor);
            event->accept();
            return;
        }
    }

    // ===== Ctrl/Cmd+D: Fill Down (copy value from cell above into selection) =====
    if (ctrl && event->key() == Qt::Key_D) {
        if (!m_spreadsheet) { QTableView::keyPressEvent(event); return; }

        QModelIndexList selected = selectionModel()->selectedIndexes();
        if (selected.size() <= 1) {
            // Single cell: copy from cell above
            QModelIndex cur = currentIndex();
            if (cur.isValid() && cur.row() > 0) {
                auto valAbove = m_spreadsheet->getCellValue(CellAddress(cur.row() - 1, cur.column()));
                auto cellAbove = m_spreadsheet->getCell(CellAddress(cur.row() - 1, cur.column()));

                std::vector<CellSnapshot> before, after;
                CellAddress addr(cur.row(), cur.column());
                before.push_back(m_spreadsheet->takeCellSnapshot(addr));

                if (cellAbove->getType() == CellType::Formula) {
                    m_model->setData(cur, cellAbove->getFormula());
                } else {
                    m_model->setData(cur, valAbove);
                }

                after.push_back(m_spreadsheet->takeCellSnapshot(addr));
            }
        } else {
            // Multi-cell selection: for each column, copy the topmost selected cell value down
            // Find bounds
            int minRow = INT_MAX;
            std::map<int, int> colMinRow; // col -> min row in selection
            for (const auto& idx : selected) {
                if (colMinRow.find(idx.column()) == colMinRow.end() || idx.row() < colMinRow[idx.column()]) {
                    colMinRow[idx.column()] = idx.row();
                }
                minRow = qMin(minRow, idx.row());
            }

            std::vector<CellSnapshot> before, after;
            m_model->setSuppressUndo(true);

            for (const auto& idx : selected) {
                int sourceRow = colMinRow[idx.column()];
                if (idx.row() == sourceRow) continue; // Skip source cells

                CellAddress addr(idx.row(), idx.column());
                before.push_back(m_spreadsheet->takeCellSnapshot(addr));

                auto srcCell = m_spreadsheet->getCell(CellAddress(sourceRow, idx.column()));
                if (srcCell->getType() == CellType::Formula) {
                    m_model->setData(idx, srcCell->getFormula());
                } else {
                    m_model->setData(idx, srcCell->getValue());
                }

                after.push_back(m_spreadsheet->takeCellSnapshot(addr));
            }

            m_model->setSuppressUndo(false);
            if (!before.empty()) {
                m_spreadsheet->getUndoManager().pushCommand(
                    std::make_unique<MultiCellEditCommand>(before, after, "Fill Down"));
            }
        }
        event->accept();
        return;
    }

    // ===== Ctrl/Cmd+Arrow: Jump to edge of data region =====
    if (ctrl && !shift && (event->key() == Qt::Key_Up || event->key() == Qt::Key_Down ||
                           event->key() == Qt::Key_Left || event->key() == Qt::Key_Right)) {
        QModelIndex cur = currentIndex();
        if (!cur.isValid() || !m_spreadsheet) { QTableView::keyPressEvent(event); return; }

        int row = cur.row();
        int col = cur.column();
        int maxRow = m_spreadsheet->getRowCount() - 1;
        int maxCol = m_spreadsheet->getColumnCount() - 1;

        auto hasData = [this](int r, int c) -> bool {
            auto val = m_spreadsheet->getCellValue(CellAddress(r, c));
            return val.isValid() && !val.toString().isEmpty();
        };

        if (event->key() == Qt::Key_Up) {
            if (row == 0) { /* already at top */ }
            else if (hasData(row - 1, col)) {
                // Move up through data until hitting empty cell or top
                while (row > 0 && hasData(row - 1, col)) row--;
            } else {
                // Move up through empty cells until hitting data or top
                while (row > 0 && !hasData(row - 1, col)) row--;
            }
        } else if (event->key() == Qt::Key_Down) {
            if (row >= maxRow) { /* already at bottom */ }
            else if (hasData(row + 1, col)) {
                while (row < maxRow && hasData(row + 1, col)) row++;
            } else {
                while (row < maxRow && !hasData(row + 1, col)) row++;
            }
        } else if (event->key() == Qt::Key_Left) {
            if (col == 0) { /* already at left */ }
            else if (hasData(row, col - 1)) {
                while (col > 0 && hasData(row, col - 1)) col--;
            } else {
                while (col > 0 && !hasData(row, col - 1)) col--;
            }
        } else if (event->key() == Qt::Key_Right) {
            if (col >= maxCol) { /* already at right */ }
            else if (hasData(row, col + 1)) {
                while (col < maxCol && hasData(row, col + 1)) col++;
            } else {
                while (col < maxCol && !hasData(row, col + 1)) col++;
            }
        }

        QModelIndex target = model()->index(row, col);
        if (target.isValid()) {
            setCurrentIndex(target);
            scrollTo(target);
        }
        event->accept();
        return;
    }

    // ===== Ctrl/Cmd+Home: Go to cell A1 =====
    if (ctrl && event->key() == Qt::Key_Home) {
        QModelIndex first = model()->index(0, 0);
        setCurrentIndex(first);
        scrollTo(first);
        event->accept();
        return;
    }

    // ===== Ctrl/Cmd+End: Go to last used cell =====
    if (ctrl && event->key() == Qt::Key_End) {
        if (m_spreadsheet) {
            int maxRow = m_spreadsheet->getMaxRow();
            int maxCol = m_spreadsheet->getMaxColumn();
            if (maxRow >= 0 && maxCol >= 0) {
                QModelIndex last = model()->index(maxRow, maxCol);
                setCurrentIndex(last);
                scrollTo(last);
            }
        }
        event->accept();
        return;
    }

    // ===== Ctrl+Shift+Arrow: Extend selection to data edge =====
    if (ctrl && shift && (event->key() == Qt::Key_Up || event->key() == Qt::Key_Down ||
                           event->key() == Qt::Key_Left || event->key() == Qt::Key_Right)) {
        QModelIndex cur = currentIndex();
        if (!cur.isValid() || !m_spreadsheet) { QTableView::keyPressEvent(event); return; }

        int row = cur.row();
        int col = cur.column();
        int maxRowIdx = m_spreadsheet->getRowCount() - 1;
        int maxColIdx = m_spreadsheet->getColumnCount() - 1;

        auto hasData = [this](int r, int c) -> bool {
            auto val = m_spreadsheet->getCellValue(CellAddress(r, c));
            return val.isValid() && !val.toString().isEmpty();
        };

        if (event->key() == Qt::Key_Up) {
            if (row > 0 && hasData(row - 1, col)) {
                while (row > 0 && hasData(row - 1, col)) row--;
            } else {
                while (row > 0 && !hasData(row - 1, col)) row--;
            }
        } else if (event->key() == Qt::Key_Down) {
            if (row < maxRowIdx && hasData(row + 1, col)) {
                while (row < maxRowIdx && hasData(row + 1, col)) row++;
            } else {
                while (row < maxRowIdx && !hasData(row + 1, col)) row++;
            }
        } else if (event->key() == Qt::Key_Left) {
            if (col > 0 && hasData(row, col - 1)) {
                while (col > 0 && hasData(row, col - 1)) col--;
            } else {
                while (col > 0 && !hasData(row, col - 1)) col--;
            }
        } else if (event->key() == Qt::Key_Right) {
            if (col < maxColIdx && hasData(row, col + 1)) {
                while (col < maxColIdx && hasData(row, col + 1)) col++;
            } else {
                while (col < maxColIdx && !hasData(row, col + 1)) col++;
            }
        }

        QModelIndex target = model()->index(row, col);
        if (target.isValid()) {
            // Extend selection from current to target
            QItemSelection sel(cur, target);
            selectionModel()->select(sel, QItemSelectionModel::ClearAndSelect);
            setCurrentIndex(target);
            scrollTo(target);
        }
        event->accept();
        return;
    }

    // ===== Ctrl+R: Fill Right (copy from left cell) =====
    if (ctrl && event->key() == Qt::Key_R) {
        if (!m_spreadsheet) { QTableView::keyPressEvent(event); return; }

        QModelIndexList selected = selectionModel()->selectedIndexes();
        if (selected.size() <= 1) {
            QModelIndex cur = currentIndex();
            if (cur.isValid() && cur.column() > 0) {
                auto valLeft = m_spreadsheet->getCellValue(CellAddress(cur.row(), cur.column() - 1));
                auto cellLeft = m_spreadsheet->getCell(CellAddress(cur.row(), cur.column() - 1));

                CellAddress addr(cur.row(), cur.column());
                std::vector<CellSnapshot> before, after;
                before.push_back(m_spreadsheet->takeCellSnapshot(addr));

                if (cellLeft->getType() == CellType::Formula) {
                    m_model->setData(cur, cellLeft->getFormula());
                } else {
                    m_model->setData(cur, valLeft);
                }

                after.push_back(m_spreadsheet->takeCellSnapshot(addr));
            }
        }
        event->accept();
        return;
    }

    // ===== Ctrl+; : Insert current date =====
    if (ctrl && event->key() == Qt::Key_Semicolon) {
        QModelIndex cur = currentIndex();
        if (cur.isValid()) {
            QString today = QDate::currentDate().toString("MM/dd/yyyy");
            m_model->setData(cur, today);
        }
        event->accept();
        return;
    }

    QTableView::keyPressEvent(event);
}

void SpreadsheetView::setFormulaEditMode(bool active) {
    m_formulaEditMode = active;
    if (active) {
        m_formulaEditCell = currentIndex();
    }
}

void SpreadsheetView::insertCellReference(const QString& ref) {
    // Insert into the active cell editor if editing inline
    if (state() == QAbstractItemView::EditingState) {
        QWidget* editor = indexWidget(currentIndex());
        QLineEdit* lineEdit = qobject_cast<QLineEdit*>(editor);
        if (lineEdit) {
            lineEdit->insert(ref);
            return;
        }
    }
    // Otherwise signal for the formula bar to handle
    emit cellReferenceInserted(ref);
}

void SpreadsheetView::mousePressEvent(QMouseEvent* event) {
    // Formula edit mode: clicking a cell inserts its reference
    if (m_formulaEditMode && event->button() == Qt::LeftButton) {
        QModelIndex clickedIdx = indexAt(event->pos());
        if (clickedIdx.isValid() && clickedIdx != m_formulaEditCell) {
            CellAddress addr(clickedIdx.row(), clickedIdx.column());
            insertCellReference(addr.toString());
            event->accept();
            return;
        }
    }

    // Format painter: apply copied style
    if (m_formatPainterActive && event->button() == Qt::LeftButton) {
        QModelIndex idx = indexAt(event->pos());
        if (idx.isValid() && m_spreadsheet) {
            std::vector<CellSnapshot> before, after;
            CellAddress addr(idx.row(), idx.column());
            before.push_back(m_spreadsheet->takeCellSnapshot(addr));

            auto cell = m_spreadsheet->getCell(addr);
            cell->setStyle(m_copiedStyle);

            after.push_back(m_spreadsheet->takeCellSnapshot(addr));
            m_spreadsheet->getUndoManager().execute(
                std::make_unique<StyleChangeCommand>(before, after), m_spreadsheet.get());

            if (m_model) {
                emit m_model->dataChanged(idx, idx);
            }
        }
        m_formatPainterActive = false;
        viewport()->setCursor(Qt::ArrowCursor);
        return;
    }

    // Fill handle drag start
    if (event->button() == Qt::LeftButton && isOverFillHandle(event->pos())) {
        m_fillDragging = true;
        m_fillDragStart = currentIndex();
        m_fillDragCurrent = event->pos();
        event->accept();
        return;
    }

    QTableView::mousePressEvent(event);
}

void SpreadsheetView::mouseMoveEvent(QMouseEvent* event) {
    if (m_fillDragging) {
        m_fillDragCurrent = event->pos();
        viewport()->update();
        return;
    }

    // Change cursor when hovering over fill handle
    if (isOverFillHandle(event->pos())) {
        viewport()->setCursor(Qt::CrossCursor);
    } else if (!m_formatPainterActive) {
        viewport()->setCursor(Qt::ArrowCursor);
    }

    QTableView::mouseMoveEvent(event);
}

void SpreadsheetView::mouseReleaseEvent(QMouseEvent* event) {
    if (m_fillDragging) {
        m_fillDragging = false;
        performFillSeries();
        viewport()->update();
        return;
    }

    QTableView::mouseReleaseEvent(event);
}

void SpreadsheetView::paintEvent(QPaintEvent* event) {
    QTableView::paintEvent(event);

    // Draw fill handle on current selection
    QModelIndex current = currentIndex();
    if (current.isValid() && !m_fillDragging) {
        QRect selRect = getSelectionBoundingRect();
        if (!selRect.isNull()) {
            int handleSize = 7;
            m_fillHandleRect = QRect(
                selRect.right() - handleSize / 2,
                selRect.bottom() - handleSize / 2,
                handleSize, handleSize);

            QPainter painter(viewport());
            painter.setRenderHint(QPainter::Antialiasing, false);
            painter.fillRect(m_fillHandleRect, QColor(16, 124, 16));
            painter.setPen(QPen(Qt::white, 1));
            painter.drawRect(m_fillHandleRect);
        }
    }

    // Draw fill drag preview
    if (m_fillDragging && m_fillDragStart.isValid()) {
        QModelIndex dragTarget = indexAt(m_fillDragCurrent);
        if (dragTarget.isValid()) {
            QRect selRect = getSelectionBoundingRect();
            QRect targetRect = visualRect(dragTarget);

            QPainter painter(viewport());
            painter.setRenderHint(QPainter::Antialiasing, false);

            // Determine fill direction and draw dashed border
            QRect fillRect;
            if (dragTarget.row() > m_fillDragStart.row()) {
                fillRect = QRect(selRect.left(), selRect.bottom() + 1,
                                 selRect.width(), targetRect.bottom() - selRect.bottom());
            } else if (dragTarget.column() > m_fillDragStart.column()) {
                fillRect = QRect(selRect.right() + 1, selRect.top(),
                                 targetRect.right() - selRect.right(), selRect.height());
            }

            if (!fillRect.isNull()) {
                QPen dashPen(QColor(16, 124, 16), 1, Qt::DashLine);
                painter.setPen(dashPen);
                painter.setBrush(QColor(198, 217, 240, 40));
                painter.drawRect(fillRect);
            }
        }
    }
}

// ============== Fill series helpers ==============

QRect SpreadsheetView::getSelectionBoundingRect() const {
    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) {
        QModelIndex current = currentIndex();
        if (current.isValid()) return visualRect(current);
        return QRect();
    }

    QRect result;
    for (const auto& idx : selected) {
        QRect r = visualRect(idx);
        if (result.isNull()) result = r;
        else result = result.united(r);
    }
    return result;
}

bool SpreadsheetView::isOverFillHandle(const QPoint& pos) const {
    if (m_fillHandleRect.isNull()) return false;
    QRect hitRect = m_fillHandleRect.adjusted(-2, -2, 2, 2);
    return hitRect.contains(pos);
}

void SpreadsheetView::performFillSeries() {
    if (!m_spreadsheet || !m_fillDragStart.isValid()) return;

    QModelIndex dragTarget = indexAt(m_fillDragCurrent);
    if (!dragTarget.isValid()) return;

    QModelIndexList selected = selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    // Find selection bounds
    int selMinRow = selected.first().row(), selMaxRow = selMinRow;
    int selMinCol = selected.first().column(), selMaxCol = selMinCol;
    for (const auto& idx : selected) {
        selMinRow = qMin(selMinRow, idx.row());
        selMaxRow = qMax(selMaxRow, idx.row());
        selMinCol = qMin(selMinCol, idx.column());
        selMaxCol = qMax(selMaxCol, idx.column());
    }

    std::vector<CellSnapshot> before, after;

    m_model->setSuppressUndo(true);

    // Fill down
    if (dragTarget.row() > selMaxRow) {
        int fillCount = dragTarget.row() - selMaxRow;
        for (int col = selMinCol; col <= selMaxCol; ++col) {
            QStringList seeds;
            for (int row = selMinRow; row <= selMaxRow; ++row) {
                auto cell = m_spreadsheet->getCell(CellAddress(row, col));
                seeds.append(cell->getValue().toString());
            }

            int totalCount = seeds.size() + fillCount;
            QStringList series = FillSeries::generateSeries(seeds, totalCount);

            for (int i = 0; i < fillCount; ++i) {
                int targetRow = selMaxRow + 1 + i;
                CellAddress addr(targetRow, col);
                before.push_back(m_spreadsheet->takeCellSnapshot(addr));

                QModelIndex idx = m_model->index(targetRow, col);
                m_model->setData(idx, series[seeds.size() + i]);

                after.push_back(m_spreadsheet->takeCellSnapshot(addr));
            }
        }
    }
    // Fill right
    else if (dragTarget.column() > selMaxCol) {
        int fillCount = dragTarget.column() - selMaxCol;
        for (int row = selMinRow; row <= selMaxRow; ++row) {
            QStringList seeds;
            for (int col = selMinCol; col <= selMaxCol; ++col) {
                auto cell = m_spreadsheet->getCell(CellAddress(row, col));
                seeds.append(cell->getValue().toString());
            }

            int totalCount = seeds.size() + fillCount;
            QStringList series = FillSeries::generateSeries(seeds, totalCount);

            for (int i = 0; i < fillCount; ++i) {
                int targetCol = selMaxCol + 1 + i;
                CellAddress addr(row, targetCol);
                before.push_back(m_spreadsheet->takeCellSnapshot(addr));

                QModelIndex idx = m_model->index(row, targetCol);
                m_model->setData(idx, series[seeds.size() + i]);

                after.push_back(m_spreadsheet->takeCellSnapshot(addr));
            }
        }
    }

    m_model->setSuppressUndo(false);

    if (!before.empty()) {
        m_spreadsheet->getUndoManager().pushCommand(
            std::make_unique<MultiCellEditCommand>(before, after, "Fill Series"));
    }
}

// ============== Multi-select resize ==============

void SpreadsheetView::onHorizontalSectionResized(int logicalIndex, int /*oldSize*/, int newSize) {
    if (m_resizingMultiple) return;
    m_resizingMultiple = true;

    QModelIndexList selected = selectionModel()->selectedColumns();
    if (selected.size() > 1) {
        for (const auto& idx : selected) {
            if (idx.column() != logicalIndex) {
                horizontalHeader()->resizeSection(idx.column(), newSize);
            }
        }
    }

    m_resizingMultiple = false;
}

void SpreadsheetView::onVerticalSectionResized(int logicalIndex, int /*oldSize*/, int newSize) {
    if (m_resizingMultiple) return;
    m_resizingMultiple = true;

    QModelIndexList selected = selectionModel()->selectedRows();
    if (selected.size() > 1) {
        for (const auto& idx : selected) {
            if (idx.row() != logicalIndex) {
                verticalHeader()->resizeSection(idx.row(), newSize);
            }
        }
    }

    m_resizingMultiple = false;
}

// ============== Slots ==============

void SpreadsheetView::onCellClicked(const QModelIndex& index) {
    emitCellSelected(index);
}

void SpreadsheetView::onCellDoubleClicked(const QModelIndex& index) {
    edit(index);
}

void SpreadsheetView::onDataChanged(const QModelIndex& topLeft, const QModelIndex& bottomRight) {
    update(topLeft);
    update(bottomRight);
}
