#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QDragEnterEvent>
#include <QDropEvent>
#include <QCloseEvent>
#include <QTabBar>
#include <QToolButton>
#include <QJsonArray>
#include <memory>
#include <vector>

class Spreadsheet;
class SpreadsheetView;
class FormulaBar;
class Toolbar;
class FormatCellsDialog;
class FindReplaceDialog;
class ChatPanel;

class MainWindow : public QMainWindow {
    Q_OBJECT

public:
    MainWindow(QWidget* parent = nullptr);
    ~MainWindow() = default;

protected:
    void closeEvent(QCloseEvent* event) override;
    void dragEnterEvent(QDragEnterEvent* event) override;
    void dropEvent(QDropEvent* event) override;

private slots:
    void onNewDocument();
    void onOpenDocument();
    void onSaveDocument();
    void onSaveAs();
    void onImportCsv();
    void onExportCsv();
    void onUndo();
    void onRedo();
    void onCut();
    void onCopy();
    void onPaste();
    void onDelete();
    void onSelectAll();
    void onFormatCells();
    void onFindReplace();
    void onGoTo();
    void onConditionalFormat();
    void onDataValidation();
    void onFreezePane();

    // Find/Replace operations
    void onFindNext();
    void onFindPrevious();
    void onReplaceOne();
    void onReplaceAll();

    // Sheet tab management
    void onSheetTabChanged(int index);
    void onSheetTabDoubleClicked(int index);
    void onAddSheet();
    void onDeleteSheet();
    void onDuplicateSheet();
    void showSheetContextMenu(const QPoint& pos);
    void onHighlightInvalidCells();
    void onChatActions(const QJsonArray& actions);

private:
    void createMenuBar();
    void createToolBar();
    void createStatusBar();
    void createSheetTabBar();
    void connectSignals();
    bool saveCurrentDocument();
    void openFile(const QString& fileName);
    void setSheets(const std::vector<std::shared_ptr<Spreadsheet>>& sheets);
    void switchToSheet(int index);
    int nextSheetNumber() const;
    bool cellMatchesSearch(int row, int col, const QString& searchText, bool matchCase, bool wholeCell) const;
    void updateStatusBarSummary();

    SpreadsheetView* m_spreadsheetView;
    FormulaBar* m_formulaBar;
    Toolbar* m_toolbar;
    QWidget* m_bottomBar;
    QTabBar* m_sheetTabBar;
    QToolButton* m_addSheetBtn;
    FindReplaceDialog* m_findReplaceDialog = nullptr;
    ChatPanel* m_chatPanel = nullptr;
    QDockWidget* m_chatDock = nullptr;
    QString m_currentFilePath;  // Track the file path for Ctrl+S

    // Multi-sheet storage
    std::vector<std::shared_ptr<Spreadsheet>> m_sheets;
    int m_activeSheetIndex = 0;
    bool m_frozenPanes = false;
};

#endif // MAINWINDOW_H
