#include "Spreadsheet.h"
#include <algorithm>

Spreadsheet::Spreadsheet()
    : m_sheetName("Sheet1"), m_rowCount(1000), m_columnCount(256),
      m_autoRecalculate(true), m_inTransaction(false) {
    m_formulaEngine = std::make_unique<FormulaEngine>(this);
}

std::shared_ptr<Cell> Spreadsheet::getCell(const CellAddress& addr) {
    return getCell(addr.row, addr.col);
}

std::shared_ptr<Cell> Spreadsheet::getCell(int row, int col) {
    CellKey key{row, col};
    auto it = m_cells.find(key);
    
    if (it != m_cells.end()) {
        return it->second;
    }
    
    // Create cell on demand
    auto cell = std::make_shared<Cell>();
    m_cells[key] = cell;
    return cell;
}

QVariant Spreadsheet::getCellValue(const CellAddress& addr) {
    auto cell = getCell(addr);
    
    if (cell->getType() == CellType::Formula) {
        return cell->getComputedValue();
    }
    
    return cell->getValue();
}

void Spreadsheet::setCellValue(const CellAddress& addr, const QVariant& value) {
    auto cell = getCell(addr);
    cell->setValue(value);
    m_depGraph.removeDependencies(addr);

    if (m_autoRecalculate && !m_inTransaction) {
        recalculateDependents(addr);
    }
}

void Spreadsheet::setCellFormula(const CellAddress& addr, const QString& formula) {
    auto cell = getCell(addr);
    cell->setFormula(formula);

    updateDependencies(addr);

    if (m_depGraph.hasCircularDependency(addr)) {
        cell->setComputedValue(QVariant("#CIRCULAR!"));
        return;
    }

    if (m_autoRecalculate && !m_inTransaction) {
        recalculate(addr);
        recalculateDependents(addr);
    }
}

void Spreadsheet::fillRange(const CellRange& range, const QVariant& value) {
    for (const auto& addr : range.getCells()) {
        setCellValue(addr, value);
    }
}

void Spreadsheet::clearRange(const CellRange& range) {
    CellKey key;
    for (int r = range.getStart().row; r <= range.getEnd().row; ++r) {
        for (int c = range.getStart().col; c <= range.getEnd().col; ++c) {
            key = {r, c};
            auto it = m_cells.find(key);
            if (it != m_cells.end()) {
                it->second->clear();
            }
        }
    }
}

std::vector<std::shared_ptr<Cell>> Spreadsheet::getRange(const CellRange& range) {
    std::vector<std::shared_ptr<Cell>> result;
    for (const auto& addr : range.getCells()) {
        result.push_back(getCell(addr));
    }
    return result;
}

void Spreadsheet::insertRow(int row, int count) {
    // Shift all cells below this row down
    std::vector<CellKey> keysToUpdate;
    for (const auto& pair : m_cells) {
        if (pair.first.row >= row) {
            keysToUpdate.push_back(pair.first);
        }
    }
    
    // Sort in reverse to avoid conflicts
    std::sort(keysToUpdate.rbegin(), keysToUpdate.rend());
    
    for (const auto& key : keysToUpdate) {
        auto cell = m_cells[key];
        m_cells.erase(key);
        CellKey newKey{key.row + count, key.col};
        m_cells[newKey] = cell;
    }
    
    m_rowCount += count;
}

void Spreadsheet::insertColumn(int column, int count) {
    std::vector<CellKey> keysToUpdate;
    for (const auto& pair : m_cells) {
        if (pair.first.col >= column) {
            keysToUpdate.push_back(pair.first);
        }
    }
    
    std::sort(keysToUpdate.rbegin(), keysToUpdate.rend());
    
    for (const auto& key : keysToUpdate) {
        auto cell = m_cells[key];
        m_cells.erase(key);
        CellKey newKey{key.row, key.col + count};
        m_cells[newKey] = cell;
    }
    
    m_columnCount += count;
}

void Spreadsheet::deleteRow(int row, int count) {
    // Remove cells in the deleted rows
    std::vector<CellKey> keysToRemove;
    std::vector<CellKey> keysToShift;
    
    for (const auto& pair : m_cells) {
        if (pair.first.row >= row && pair.first.row < row + count) {
            keysToRemove.push_back(pair.first);
        } else if (pair.first.row >= row + count) {
            keysToShift.push_back(pair.first);
        }
    }
    
    for (const auto& key : keysToRemove) {
        m_cells.erase(key);
    }
    
    std::sort(keysToShift.rbegin(), keysToShift.rend());
    
    for (const auto& key : keysToShift) {
        auto cell = m_cells[key];
        m_cells.erase(key);
        CellKey newKey{key.row - count, key.col};
        m_cells[newKey] = cell;
    }
    
    m_rowCount -= count;
}

void Spreadsheet::deleteColumn(int column, int count) {
    std::vector<CellKey> keysToRemove;
    std::vector<CellKey> keysToShift;
    
    for (const auto& pair : m_cells) {
        if (pair.first.col >= column && pair.first.col < column + count) {
            keysToRemove.push_back(pair.first);
        } else if (pair.first.col >= column + count) {
            keysToShift.push_back(pair.first);
        }
    }
    
    for (const auto& key : keysToRemove) {
        m_cells.erase(key);
    }
    
    std::sort(keysToShift.rbegin(), keysToShift.rend());
    
    for (const auto& key : keysToShift) {
        auto cell = m_cells[key];
        m_cells.erase(key);
        CellKey newKey{key.row, key.col - count};
        m_cells[newKey] = cell;
    }
    
    m_columnCount -= count;
}

QString Spreadsheet::getSheetName() const {
    return m_sheetName;
}

void Spreadsheet::setSheetName(const QString& name) {
    m_sheetName = name;
}

int Spreadsheet::getMaxRow() const {
    int max = 0;
    for (const auto& pair : m_cells) {
        max = std::max(max, pair.first.row);
    }
    return max;
}

int Spreadsheet::getMaxColumn() const {
    int max = 0;
    for (const auto& pair : m_cells) {
        max = std::max(max, pair.first.col);
    }
    return max;
}

std::vector<CellAddress> Spreadsheet::getDirtyCells() const {
    std::vector<CellAddress> dirty;
    for (const auto& pair : m_cells) {
        if (pair.second->isDirty()) {
            dirty.push_back(CellAddress(pair.first.row, pair.first.col));
        }
    }
    return dirty;
}

void Spreadsheet::clearDirtyFlag() {
    for (auto& pair : m_cells) {
        pair.second->setDirty(false);
    }
}

void Spreadsheet::startTransaction() {
    m_inTransaction = true;
}

void Spreadsheet::commitTransaction() {
    m_inTransaction = false;
    if (m_autoRecalculate) {
        recalculateAll();
    }
}

void Spreadsheet::rollbackTransaction() {
    m_inTransaction = false;
}

FormulaEngine& Spreadsheet::getFormulaEngine() {
    return *m_formulaEngine;
}

void Spreadsheet::setAutoRecalculate(bool enabled) {
    m_autoRecalculate = enabled;
}

bool Spreadsheet::getAutoRecalculate() const {
    return m_autoRecalculate;
}

CellSnapshot Spreadsheet::takeCellSnapshot(const CellAddress& addr) {
    auto cell = getCell(addr);
    CellSnapshot snap;
    snap.addr = addr;
    snap.value = cell->getValue();
    snap.formula = cell->getFormula();
    snap.style = cell->getStyle();
    snap.type = cell->getType();
    return snap;
}

void Spreadsheet::forEachCell(std::function<void(int row, int col, const Cell&)> callback) const {
    for (const auto& pair : m_cells) {
        if (pair.second && pair.second->getType() != CellType::Empty) {
            callback(pair.first.row, pair.first.col, *pair.second);
        }
    }
}

void Spreadsheet::recalculate(const CellAddress& addr) {
    auto cell = getCell(addr);
    if (cell->getType() == CellType::Formula) {
        QVariant result = m_formulaEngine->evaluate(cell->getFormula());
        cell->setComputedValue(result);
    }
}

void Spreadsheet::recalculateAll() {
    for (auto& pair : m_cells) {
        if (pair.second->getType() == CellType::Formula) {
            CellAddress addr(pair.first.row, pair.first.col);
            QVariant result = m_formulaEngine->evaluate(pair.second->getFormula());
            pair.second->setComputedValue(result);
            updateDependencies(addr);
        }
    }
}

void Spreadsheet::updateDependencies(const CellAddress& addr) {
    m_depGraph.removeDependencies(addr);
    auto cell = getCell(addr);
    if (cell->getType() == CellType::Formula) {
        // Evaluate to discover dependencies
        m_formulaEngine->evaluate(cell->getFormula());
        for (const auto& dep : m_formulaEngine->getLastDependencies()) {
            m_depGraph.addDependency(addr, dep);
        }
    }
}

void Spreadsheet::recalculateDependents(const CellAddress& addr) {
    auto order = m_depGraph.getRecalcOrder(addr);
    for (const auto& depAddr : order) {
        auto cell = getCell(depAddr);
        if (cell->getType() == CellType::Formula) {
            QVariant result = m_formulaEngine->evaluate(cell->getFormula());
            cell->setComputedValue(result);
        }
    }
}

void Spreadsheet::sortRange(const CellRange& range, int sortColumn, bool ascending) {
    int startRow = range.getStart().row;
    int endRow = range.getEnd().row;
    int startCol = range.getStart().col;
    int endCol = range.getEnd().col;

    if (startRow >= endRow) return;

    // Collect rows as vectors of cell data
    struct RowData {
        int originalRow;
        QVariant sortValue;
        std::vector<std::pair<int, std::shared_ptr<Cell>>> cells;
    };

    std::vector<RowData> rows;
    for (int r = startRow; r <= endRow; ++r) {
        RowData rd;
        rd.originalRow = r;
        auto sortCell = getCell(CellAddress(r, sortColumn));
        rd.sortValue = getCellValue(CellAddress(r, sortColumn));
        for (int c = startCol; c <= endCol; ++c) {
            CellKey key{r, c};
            auto it = m_cells.find(key);
            if (it != m_cells.end()) {
                rd.cells.push_back({c, it->second});
            }
        }
        rows.push_back(std::move(rd));
    }

    // Sort rows
    std::sort(rows.begin(), rows.end(), [ascending](const RowData& a, const RowData& b) {
        QVariant va = a.sortValue;
        QVariant vb = b.sortValue;

        // Empty cells go to the bottom
        bool aEmpty = !va.isValid() || va.toString().isEmpty();
        bool bEmpty = !vb.isValid() || vb.toString().isEmpty();
        if (aEmpty && bEmpty) return false;
        if (aEmpty) return false;
        if (bEmpty) return true;

        // Try numeric comparison
        bool aOk, bOk;
        double aNum = va.toString().toDouble(&aOk);
        double bNum = vb.toString().toDouble(&bOk);
        if (aOk && bOk) {
            return ascending ? (aNum < bNum) : (aNum > bNum);
        }

        // String comparison
        int cmp = va.toString().compare(vb.toString(), Qt::CaseInsensitive);
        return ascending ? (cmp < 0) : (cmp > 0);
    });

    // Remove old cells in the range
    for (int r = startRow; r <= endRow; ++r) {
        for (int c = startCol; c <= endCol; ++c) {
            m_cells.erase(CellKey{r, c});
        }
    }

    // Write sorted rows back
    for (int i = 0; i < static_cast<int>(rows.size()); ++i) {
        int targetRow = startRow + i;
        for (auto& [col, cell] : rows[i].cells) {
            m_cells[CellKey{targetRow, col}] = cell;
        }
    }
}

void Spreadsheet::insertCellsShiftRight(const CellRange& range) {
    int startRow = range.getStart().row;
    int endRow = range.getEnd().row;
    int startCol = range.getStart().col;
    int colCount = range.getEnd().col - startCol + 1;

    // For each affected row, shift cells from startCol rightward
    for (int r = startRow; r <= endRow; ++r) {
        // Collect keys to shift (in this row, at or after startCol)
        std::vector<CellKey> keysToShift;
        for (const auto& pair : m_cells) {
            if (pair.first.row == r && pair.first.col >= startCol) {
                keysToShift.push_back(pair.first);
            }
        }
        // Sort descending by column to avoid conflicts
        std::sort(keysToShift.begin(), keysToShift.end(),
                  [](const CellKey& a, const CellKey& b) { return a.col > b.col; });

        for (const auto& key : keysToShift) {
            auto cell = m_cells[key];
            m_cells.erase(key);
            m_cells[CellKey{r, key.col + colCount}] = cell;
        }
    }
}

void Spreadsheet::insertCellsShiftDown(const CellRange& range) {
    int startRow = range.getStart().row;
    int startCol = range.getStart().col;
    int endCol = range.getEnd().col;
    int rowCount = range.getEnd().row - startRow + 1;

    // For each affected column, shift cells from startRow downward
    for (int c = startCol; c <= endCol; ++c) {
        std::vector<CellKey> keysToShift;
        for (const auto& pair : m_cells) {
            if (pair.first.col == c && pair.first.row >= startRow) {
                keysToShift.push_back(pair.first);
            }
        }
        std::sort(keysToShift.begin(), keysToShift.end(),
                  [](const CellKey& a, const CellKey& b) { return a.row > b.row; });

        for (const auto& key : keysToShift) {
            auto cell = m_cells[key];
            m_cells.erase(key);
            m_cells[CellKey{key.row + rowCount, c}] = cell;
        }
    }
}

void Spreadsheet::deleteCellsShiftLeft(const CellRange& range) {
    int startRow = range.getStart().row;
    int endRow = range.getEnd().row;
    int startCol = range.getStart().col;
    int endCol = range.getEnd().col;
    int colCount = endCol - startCol + 1;

    for (int r = startRow; r <= endRow; ++r) {
        // Remove cells in the deleted range
        for (int c = startCol; c <= endCol; ++c) {
            m_cells.erase(CellKey{r, c});
        }

        // Shift cells after the range leftward
        std::vector<CellKey> keysToShift;
        for (const auto& pair : m_cells) {
            if (pair.first.row == r && pair.first.col > endCol) {
                keysToShift.push_back(pair.first);
            }
        }
        std::sort(keysToShift.begin(), keysToShift.end(),
                  [](const CellKey& a, const CellKey& b) { return a.col < b.col; });

        for (const auto& key : keysToShift) {
            auto cell = m_cells[key];
            m_cells.erase(key);
            m_cells[CellKey{r, key.col - colCount}] = cell;
        }
    }
}

// ============== Table Support ==============

void Spreadsheet::addTable(const SpreadsheetTable& table) {
    m_tables.push_back(table);
}

void Spreadsheet::removeTable(const QString& name) {
    m_tables.erase(
        std::remove_if(m_tables.begin(), m_tables.end(),
                        [&name](const SpreadsheetTable& t) { return t.name == name; }),
        m_tables.end());
}

const SpreadsheetTable* Spreadsheet::getTableAt(int row, int col) const {
    for (const auto& table : m_tables) {
        int startRow = table.range.getStart().row;
        int endRow = table.range.getEnd().row;
        int startCol = table.range.getStart().col;
        int endCol = table.range.getEnd().col;
        if (row >= startRow && row <= endRow && col >= startCol && col <= endCol) {
            return &table;
        }
    }
    return nullptr;
}

void Spreadsheet::deleteCellsShiftUp(const CellRange& range) {
    int startRow = range.getStart().row;
    int endRow = range.getEnd().row;
    int startCol = range.getStart().col;
    int endCol = range.getEnd().col;
    int rowCount = endRow - startRow + 1;

    for (int c = startCol; c <= endCol; ++c) {
        // Remove cells in the deleted range
        for (int r = startRow; r <= endRow; ++r) {
            m_cells.erase(CellKey{r, c});
        }

        // Shift cells below the range upward
        std::vector<CellKey> keysToShift;
        for (const auto& pair : m_cells) {
            if (pair.first.col == c && pair.first.row > endRow) {
                keysToShift.push_back(pair.first);
            }
        }
        std::sort(keysToShift.begin(), keysToShift.end(),
                  [](const CellKey& a, const CellKey& b) { return a.row < b.row; });

        for (const auto& key : keysToShift) {
            auto cell = m_cells[key];
            m_cells.erase(key);
            m_cells[CellKey{key.row - rowCount, c}] = cell;
        }
    }
}

// ============== Merge Cells ==============

void Spreadsheet::mergeCells(const CellRange& range) {
    // Check if any cell in the range is already part of a merged region
    for (const auto& mr : m_mergedRegions) {
        if (mr.range.intersects(range)) {
            return; // Can't merge overlapping regions
        }
    }
    m_mergedRegions.push_back({range});
}

void Spreadsheet::unmergeCells(const CellRange& range) {
    m_mergedRegions.erase(
        std::remove_if(m_mergedRegions.begin(), m_mergedRegions.end(),
                        [&range](const MergedRegion& mr) { return mr.range.intersects(range); }),
        m_mergedRegions.end());
}

const Spreadsheet::MergedRegion* Spreadsheet::getMergedRegionAt(int row, int col) const {
    for (const auto& mr : m_mergedRegions) {
        if (mr.range.contains(row, col)) {
            return &mr;
        }
    }
    return nullptr;
}

// ============== Data Validation ==============

void Spreadsheet::addValidationRule(const DataValidationRule& rule) {
    m_validationRules.push_back(rule);
}

void Spreadsheet::removeValidationRule(int index) {
    if (index >= 0 && index < static_cast<int>(m_validationRules.size())) {
        m_validationRules.erase(m_validationRules.begin() + index);
    }
}

const Spreadsheet::DataValidationRule* Spreadsheet::getValidationAt(int row, int col) const {
    for (const auto& rule : m_validationRules) {
        if (rule.range.contains(row, col)) {
            return &rule;
        }
    }
    return nullptr;
}

bool Spreadsheet::validateCell(int row, int col, const QString& value) const {
    const auto* rule = getValidationAt(row, col);
    if (!rule) return true; // No validation rule = always valid

    if (value.isEmpty()) return true; // Allow empty cells

    switch (rule->type) {
        case DataValidationRule::WholeNumber: {
            bool ok;
            int num = value.toInt(&ok);
            if (!ok) return false;
            int v1 = rule->value1.toInt();
            int v2 = rule->value2.toInt();
            switch (rule->op) {
                case DataValidationRule::Between: return num >= v1 && num <= v2;
                case DataValidationRule::NotBetween: return num < v1 || num > v2;
                case DataValidationRule::EqualTo: return num == v1;
                case DataValidationRule::NotEqualTo: return num != v1;
                case DataValidationRule::GreaterThan: return num > v1;
                case DataValidationRule::LessThan: return num < v1;
                case DataValidationRule::GreaterThanOrEqual: return num >= v1;
                case DataValidationRule::LessThanOrEqual: return num <= v1;
            }
            break;
        }
        case DataValidationRule::Decimal: {
            bool ok;
            double num = value.toDouble(&ok);
            if (!ok) return false;
            double v1 = rule->value1.toDouble();
            double v2 = rule->value2.toDouble();
            switch (rule->op) {
                case DataValidationRule::Between: return num >= v1 && num <= v2;
                case DataValidationRule::NotBetween: return num < v1 || num > v2;
                case DataValidationRule::EqualTo: return qFuzzyCompare(num, v1);
                case DataValidationRule::NotEqualTo: return !qFuzzyCompare(num, v1);
                case DataValidationRule::GreaterThan: return num > v1;
                case DataValidationRule::LessThan: return num < v1;
                case DataValidationRule::GreaterThanOrEqual: return num >= v1;
                case DataValidationRule::LessThanOrEqual: return num <= v1;
            }
            break;
        }
        case DataValidationRule::List:
            return rule->listItems.contains(value, Qt::CaseInsensitive);
        case DataValidationRule::TextLength: {
            int len = value.length();
            int v1 = rule->value1.toInt();
            int v2 = rule->value2.toInt();
            switch (rule->op) {
                case DataValidationRule::Between: return len >= v1 && len <= v2;
                case DataValidationRule::NotBetween: return len < v1 || len > v2;
                case DataValidationRule::EqualTo: return len == v1;
                case DataValidationRule::NotEqualTo: return len != v1;
                case DataValidationRule::GreaterThan: return len > v1;
                case DataValidationRule::LessThan: return len < v1;
                case DataValidationRule::GreaterThanOrEqual: return len >= v1;
                case DataValidationRule::LessThanOrEqual: return len <= v1;
            }
            break;
        }
        case DataValidationRule::Custom:
            // Custom formula validation - would need formula engine
            return true;
        default:
            return true;
    }
    return true;
}
