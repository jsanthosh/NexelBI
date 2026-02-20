#ifndef SPREADSHEET_H
#define SPREADSHEET_H

#include <QString>
#include <QVariant>
#include <unordered_map>
#include <map>
#include <memory>
#include <vector>
#include <functional>
#include "Cell.h"
#include "CellRange.h"
#include "TableStyle.h"
#include "FormulaEngine.h"
#include "UndoManager.h"
#include "DependencyGraph.h"
#include "ConditionalFormatting.h"

class Spreadsheet {
public:
    Spreadsheet();
    ~Spreadsheet() = default;

    // Cell access and modification
    std::shared_ptr<Cell> getCell(const CellAddress& addr);
    std::shared_ptr<Cell> getCell(int row, int col);
    QVariant getCellValue(const CellAddress& addr);
    void setCellValue(const CellAddress& addr, const QVariant& value);
    void setCellFormula(const CellAddress& addr, const QString& formula);

    // Range operations
    void fillRange(const CellRange& range, const QVariant& value);
    void clearRange(const CellRange& range);
    std::vector<std::shared_ptr<Cell>> getRange(const CellRange& range);

    // Row/Column operations
    void insertRow(int row, int count = 1);
    void insertColumn(int column, int count = 1);
    void deleteRow(int row, int count = 1);
    void deleteColumn(int column, int count = 1);

    // Sheet properties
    QString getSheetName() const;
    void setSheetName(const QString& name);
    int getMaxRow() const;
    int getMaxColumn() const;
    int getRowCount() const { return m_rowCount; }
    int getColumnCount() const { return m_columnCount; }
    void setRowCount(int count) { m_rowCount = count; }
    void setColumnCount(int count) { m_columnCount = count; }

    // Dirty tracking
    std::vector<CellAddress> getDirtyCells() const;
    void clearDirtyFlag();

    // Undo/Redo
    void startTransaction();
    void commitTransaction();
    void rollbackTransaction();

    // Formula engine access
    FormulaEngine& getFormulaEngine();

    // Cell iteration (for serialization)
    void forEachCell(std::function<void(int row, int col, const Cell&)> callback) const;

    // Undo/Redo
    UndoManager& getUndoManager() { return m_undoManager; }
    CellSnapshot takeCellSnapshot(const CellAddress& addr);

    // Sorting
    void sortRange(const CellRange& range, int sortColumn, bool ascending);

    // Cell shift insert/delete
    void insertCellsShiftRight(const CellRange& range);
    void insertCellsShiftDown(const CellRange& range);
    void deleteCellsShiftLeft(const CellRange& range);
    void deleteCellsShiftUp(const CellRange& range);

    // Table support
    void addTable(const SpreadsheetTable& table);
    void removeTable(const QString& name);
    const SpreadsheetTable* getTableAt(int row, int col) const;
    const std::vector<SpreadsheetTable>& getTables() const { return m_tables; }

    // Conditional formatting
    ConditionalFormatting& getConditionalFormatting() { return m_conditionalFormatting; }
    const ConditionalFormatting& getConditionalFormatting() const { return m_conditionalFormatting; }

    // Merge cells
    struct MergedRegion {
        CellRange range;
    };
    void mergeCells(const CellRange& range);
    void unmergeCells(const CellRange& range);
    const MergedRegion* getMergedRegionAt(int row, int col) const;
    std::vector<MergedRegion>& getMergedRegions() { return m_mergedRegions; }
    const std::vector<MergedRegion>& getMergedRegions() const { return m_mergedRegions; }

    // Data validation
    struct DataValidationRule {
        CellRange range;
        enum Type { WholeNumber, Decimal, List, TextLength, Date, Custom } type = WholeNumber;
        enum Operator { Between, NotBetween, EqualTo, NotEqualTo, GreaterThan, LessThan,
                        GreaterThanOrEqual, LessThanOrEqual } op = Between;
        QString value1;
        QString value2;
        QStringList listItems;
        QString customFormula;
        QString inputTitle;
        QString inputMessage;
        QString errorTitle;
        QString errorMessage;
        enum ErrorStyle { Stop, Warning, Information } errorStyle = Stop;
        bool showInputMessage = true;
        bool showErrorAlert = true;
    };

    void addValidationRule(const DataValidationRule& rule);
    void removeValidationRule(int index);
    const DataValidationRule* getValidationAt(int row, int col) const;
    std::vector<DataValidationRule>& getValidationRules() { return m_validationRules; }
    const std::vector<DataValidationRule>& getValidationRules() const { return m_validationRules; }
    bool validateCell(int row, int col, const QString& value) const;

    // Performance settings
    void setAutoRecalculate(bool enabled);
    bool getAutoRecalculate() const;

private:
    struct CellKey {
        int row, col;
        bool operator<(const CellKey& other) const {
            if (row != other.row) return row < other.row;
            return col < other.col;
        }
    };

    std::map<CellKey, std::shared_ptr<Cell>> m_cells;
    std::unique_ptr<FormulaEngine> m_formulaEngine;
    DependencyGraph m_depGraph;
    UndoManager m_undoManager;
    QString m_sheetName;
    int m_rowCount;
    int m_columnCount;
    bool m_autoRecalculate;
    bool m_inTransaction;
    std::vector<SpreadsheetTable> m_tables;
    ConditionalFormatting m_conditionalFormatting;
    std::vector<DataValidationRule> m_validationRules;
    std::vector<MergedRegion> m_mergedRegions;

    void recalculate(const CellAddress& addr);
    void recalculateAll();
    void updateDependencies(const CellAddress& addr);
    void recalculateDependents(const CellAddress& addr);
};

#endif // SPREADSHEET_H
