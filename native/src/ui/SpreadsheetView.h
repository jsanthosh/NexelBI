#ifndef SPREADSHEETVIEW_H
#define SPREADSHEETVIEW_H

#include <QTableView>
#include <QAbstractTableModel>
#include <QColor>
#include <QMenu>
#include <memory>
#include <functional>
#include "../core/Cell.h"
#include "../core/CellRange.h"

class Spreadsheet;
class SpreadsheetModel;
class CellDelegate;

class SpreadsheetView : public QTableView {
    Q_OBJECT

public:
    explicit SpreadsheetView(QWidget* parent = nullptr);
    ~SpreadsheetView() = default;

    void setSpreadsheet(std::shared_ptr<Spreadsheet> spreadsheet);
    std::shared_ptr<Spreadsheet> getSpreadsheet() const;
    SpreadsheetModel* getModel() const { return m_model; }

    // Editing operations
    void cut();
    void copy();
    void paste();
    void deleteSelection();
    void selectAll() override;

    // Style operations
    void applyBold();
    void applyItalic();
    void applyUnderline();
    void applyStrikethrough();
    void applyFontFamily(const QString& family);
    void applyFontSize(int size);
    void applyForegroundColor(const QColor& color);
    void applyBackgroundColor(const QColor& color);
    void applyThousandSeparator();
    void applyNumberFormat(const QString& format);

    // Alignment
    void applyHAlign(HorizontalAlignment align);
    void applyVAlign(VerticalAlignment align);

    // Format painter
    void activateFormatPainter();

    // Sorting
    void sortAscending();
    void sortDescending();

    // Insert/Delete with shift
    void insertCellsShiftRight();
    void insertCellsShiftDown();
    void insertEntireRow();
    void insertEntireColumn();
    void deleteCellsShiftLeft();
    void deleteCellsShiftUp();
    void deleteEntireRow();
    void deleteEntireColumn();

    // Tables
    void applyTableStyle(int themeIndex);

    // Autofit
    void autofitColumn(int column);
    void autofitRow(int row);
    void autofitSelectedColumns();
    void autofitSelectedRows();

    // Formula editing: insert cell reference on click
    void setFormulaEditMode(bool active);
    bool isFormulaEditMode() const { return m_formulaEditMode; }
    void insertCellReference(const QString& ref);

    // Freeze Panes
    void setFrozenRow(int row);
    void setFrozenColumn(int col);

    // UI Operations
    void refreshView();
    void zoomIn();
    void zoomOut();
    void resetZoom();

signals:
    void cellSelected(int row, int col, const QString& content, const QString& address);
    void formatCellsRequested();
    void cellReferenceInserted(const QString& ref);

protected:
    void keyPressEvent(QKeyEvent* event) override;
    void mousePressEvent(QMouseEvent* event) override;
    void mouseMoveEvent(QMouseEvent* event) override;
    void mouseReleaseEvent(QMouseEvent* event) override;
    void paintEvent(QPaintEvent* event) override;
    void currentChanged(const QModelIndex& current, const QModelIndex& previous) override;

private slots:
    void onCellClicked(const QModelIndex& index);
    void onCellDoubleClicked(const QModelIndex& index);
    void onDataChanged(const QModelIndex& topLeft, const QModelIndex& bottomRight);
    void onHorizontalSectionResized(int logicalIndex, int oldSize, int newSize);
    void onVerticalSectionResized(int logicalIndex, int oldSize, int newSize);

private:
    std::shared_ptr<Spreadsheet> m_spreadsheet;
    SpreadsheetModel* m_model = nullptr;
    CellDelegate* m_delegate = nullptr;
    int m_zoomLevel;

    // Format painter state
    bool m_formatPainterActive = false;
    CellStyle m_copiedStyle;

    // Fill series state
    bool m_fillDragging = false;
    QModelIndex m_fillDragStart;
    QPoint m_fillDragCurrent;
    QRect m_fillHandleRect;

    // Multi-resize guard
    bool m_resizingMultiple = false;

    // Formula edit mode: when active, clicking cells inserts references
    bool m_formulaEditMode = false;
    QModelIndex m_formulaEditCell;  // The cell being edited with the formula

    void initializeView();
    void setupConnections();
    void emitCellSelected(const QModelIndex& index);
    void setupHeaderContextMenus();
    void showCellContextMenu(const QPoint& pos);

    // Fill series helpers
    bool isOverFillHandle(const QPoint& pos) const;
    void performFillSeries();
    QRect getSelectionBoundingRect() const;

    // Auto-detect contiguous data region from a cell
    CellRange detectDataRegion(int startRow, int startCol) const;

    // Efficient style application: only process occupied cells for large selections
    void applyStyleChange(std::function<void(CellStyle&)> modifier, const QList<int>& roles);
};

#endif // SPREADSHEETVIEW_H
