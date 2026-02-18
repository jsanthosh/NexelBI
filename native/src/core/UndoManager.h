#ifndef UNDOMANAGER_H
#define UNDOMANAGER_H

#include <QString>
#include <QVariant>
#include <vector>
#include <memory>
#include "Cell.h"
#include "CellRange.h"

class Spreadsheet;

struct CellSnapshot {
    CellAddress addr;
    QVariant value;
    QString formula;
    CellStyle style;
    CellType type;
};

class UndoCommand {
public:
    virtual ~UndoCommand() = default;
    virtual void undo(Spreadsheet* sheet) = 0;
    virtual void redo(Spreadsheet* sheet) = 0;
    virtual QString description() const = 0;
    virtual CellAddress targetCell() const { return CellAddress(0, 0); }
};

class CellEditCommand : public UndoCommand {
public:
    CellEditCommand(const CellSnapshot& before, const CellSnapshot& after);
    void undo(Spreadsheet* sheet) override;
    void redo(Spreadsheet* sheet) override;
    QString description() const override { return "Edit Cell"; }
    CellAddress targetCell() const override { return m_before.addr; }
private:
    CellSnapshot m_before;
    CellSnapshot m_after;
};

class MultiCellEditCommand : public UndoCommand {
public:
    MultiCellEditCommand(const std::vector<CellSnapshot>& before, const std::vector<CellSnapshot>& after, const QString& desc);
    void undo(Spreadsheet* sheet) override;
    void redo(Spreadsheet* sheet) override;
    QString description() const override { return m_description; }
    CellAddress targetCell() const override { return m_before.empty() ? CellAddress(0,0) : m_before.front().addr; }
private:
    std::vector<CellSnapshot> m_before;
    std::vector<CellSnapshot> m_after;
    QString m_description;
};

class StyleChangeCommand : public UndoCommand {
public:
    StyleChangeCommand(const std::vector<CellSnapshot>& before, const std::vector<CellSnapshot>& after);
    void undo(Spreadsheet* sheet) override;
    void redo(Spreadsheet* sheet) override;
    QString description() const override { return "Change Style"; }
private:
    std::vector<CellSnapshot> m_before;
    std::vector<CellSnapshot> m_after;
};

class UndoManager {
public:
    UndoManager() = default;

    void execute(std::unique_ptr<UndoCommand> cmd, Spreadsheet* sheet);
    void pushCommand(std::unique_ptr<UndoCommand> cmd);
    void undo(Spreadsheet* sheet);
    void redo(Spreadsheet* sheet);

    bool canUndo() const;
    bool canRedo() const;
    QString undoText() const;
    QString redoText() const;
    CellAddress lastUndoTarget() const;
    CellAddress lastRedoTarget() const;
    void clear();

private:
    std::vector<std::unique_ptr<UndoCommand>> m_undoStack;
    std::vector<std::unique_ptr<UndoCommand>> m_redoStack;
    static constexpr size_t MAX_UNDO = 100;
};

#endif // UNDOMANAGER_H
