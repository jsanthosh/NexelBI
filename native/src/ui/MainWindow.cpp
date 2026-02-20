#include "MainWindow.h"
#include "SpreadsheetView.h"
#include "SpreadsheetModel.h"
#include "FormulaBar.h"
#include "Toolbar.h"
#include "FormatCellsDialog.h"
#include "FindReplaceDialog.h"
#include "GoToDialog.h"
#include "ConditionalFormatDialog.h"
#include "DataValidationDialog.h"
#include "ChatPanel.h"
#include "../core/Spreadsheet.h"
#include "../core/UndoManager.h"
#include "../core/CellRange.h"
#include "../services/DocumentService.h"
#include "../services/CsvService.h"
#include "../services/XlsxService.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QMenuBar>
#include <QMenu>
#include <QFileDialog>
#include <QMessageBox>
#include <QStatusBar>
#include <QMimeData>
#include <QUrl>
#include <QFileInfo>
#include <QInputDialog>
#include <QTabBar>
#include <QToolButton>
#include <QScrollBar>
#include <QDockWidget>
#include <QJsonObject>
#include <QJsonArray>

MainWindow::MainWindow(QWidget* parent)
    : QMainWindow(parent) {
    setWindowTitle("NativeSpreadsheet");
    setGeometry(100, 100, 1280, 800);

    // Initialize with one default sheet
    auto defaultSheet = std::make_shared<Spreadsheet>();
    defaultSheet->setSheetName("Sheet1");
    m_sheets.push_back(defaultSheet);
    m_activeSheetIndex = 0;

    QWidget* centralWidget = new QWidget(this);
    QVBoxLayout* layout = new QVBoxLayout(centralWidget);
    layout->setContentsMargins(0, 0, 0, 0);
    layout->setSpacing(0);

    m_toolbar = new Toolbar(this);
    addToolBar(m_toolbar);
    addToolBarBreak(Qt::TopToolBarArea);
    QToolBar* toolbar2 = m_toolbar->createSecondaryToolbar(this);
    addToolBar(toolbar2);

    m_formulaBar = new FormulaBar(this);
    layout->addWidget(m_formulaBar);

    m_spreadsheetView = new SpreadsheetView(this);
    m_spreadsheetView->setSpreadsheet(m_sheets[0]);
    layout->addWidget(m_spreadsheetView);

    // Sheet tab bar at bottom
    createSheetTabBar();
    layout->addWidget(m_bottomBar);

    setCentralWidget(centralWidget);

    // Chat assistant panel (dock widget on the right) — created before menu bar so menu can connect
    m_chatPanel = new ChatPanel(this);
    m_chatPanel->setSpreadsheet(m_sheets[0]);
    m_chatDock = new QDockWidget("Claude Assistant", this);
    m_chatDock->setWidget(m_chatPanel);
    m_chatDock->setFeatures(QDockWidget::DockWidgetClosable | QDockWidget::DockWidgetMovable);
    m_chatDock->setMinimumWidth(300);
    m_chatDock->setStyleSheet(
        "QDockWidget { border: none; }"
        "QDockWidget::title { background: #1B5E3B; color: white; padding: 6px; font-weight: bold; text-align: center; }"
        "QDockWidget::close-button { background: transparent; }"
    );
    addDockWidget(Qt::RightDockWidgetArea, m_chatDock);
    m_chatDock->hide(); // Hidden by default, toggled from View menu

    createMenuBar();
    createStatusBar();
    connectSignals();

    setAcceptDrops(true);

    setStyleSheet(
        "QMainWindow { background-color: #F0F2F5; }"
        "QMenuBar { background-color: #1B5E3B; color: white; border: none; padding: 2px; font-size: 12px; }"
        "QMenuBar::item { padding: 4px 12px; border-radius: 3px; }"
        "QMenuBar::item:selected { background-color: #155030; }"
        "QMenu { background-color: #FFFFFF; border: 1px solid #D0D5DD; border-radius: 6px; padding: 4px; }"
        "QMenu::item { padding: 6px 30px 6px 20px; border-radius: 4px; }"
        "QMenu::item:selected { background-color: #E8F0FE; }"
        "QMenu::separator { height: 1px; background: #E0E3E8; margin: 4px 8px; }"
    );
}

void MainWindow::createSheetTabBar() {
    m_bottomBar = new QWidget(this);
    m_bottomBar->setFixedHeight(28);
    m_bottomBar->setStyleSheet(
        "QWidget { background-color: #F3F3F3; border-top: 1px solid #D0D0D0; }");

    QHBoxLayout* bottomLayout = new QHBoxLayout(m_bottomBar);
    bottomLayout->setContentsMargins(4, 0, 0, 0);
    bottomLayout->setSpacing(2);

    // Add sheet button
    m_addSheetBtn = new QToolButton(m_bottomBar);
    m_addSheetBtn->setText("+");
    m_addSheetBtn->setFixedSize(24, 22);
    m_addSheetBtn->setToolTip("Add New Sheet");
    m_addSheetBtn->setStyleSheet(
        "QToolButton { background: transparent; border: 1px solid transparent; "
        "border-radius: 3px; font-size: 16px; font-weight: bold; color: #555; }"
        "QToolButton:hover { background-color: #E0E0E0; border-color: #C0C0C0; }");
    bottomLayout->addWidget(m_addSheetBtn);
    connect(m_addSheetBtn, &QToolButton::clicked, this, &MainWindow::onAddSheet);

    // Tab bar
    m_sheetTabBar = new QTabBar(m_bottomBar);
    m_sheetTabBar->setExpanding(false);
    m_sheetTabBar->setMovable(true);
    m_sheetTabBar->setTabsClosable(false);
    m_sheetTabBar->setDocumentMode(true);
    m_sheetTabBar->setStyleSheet(
        "QTabBar { background: transparent; border: none; }"
        "QTabBar::tab {"
        "   background-color: #E8E8E8;"
        "   border: 1px solid #C8C8C8;"
        "   border-bottom: none;"
        "   padding: 3px 16px;"
        "   margin-right: 2px;"
        "   font-size: 11px;"
        "   min-width: 60px;"
        "   border-top-left-radius: 3px;"
        "   border-top-right-radius: 3px;"
        "}"
        "QTabBar::tab:selected {"
        "   background-color: white;"
        "   border-bottom: 2px solid #217346;"
        "   font-weight: bold;"
        "}"
        "QTabBar::tab:hover:!selected {"
        "   background-color: #D8D8D8;"
        "}"
    );

    // Populate tabs from m_sheets
    for (const auto& sheet : m_sheets) {
        m_sheetTabBar->addTab(sheet->getSheetName());
    }

    bottomLayout->addWidget(m_sheetTabBar);
    bottomLayout->addStretch();

    // Connect signals
    connect(m_sheetTabBar, &QTabBar::currentChanged, this, &MainWindow::onSheetTabChanged);
    connect(m_sheetTabBar, &QTabBar::tabBarDoubleClicked, this, &MainWindow::onSheetTabDoubleClicked);

    // Right-click context menu on tab bar
    m_sheetTabBar->setContextMenuPolicy(Qt::CustomContextMenu);
    connect(m_sheetTabBar, &QTabBar::customContextMenuRequested, this, &MainWindow::showSheetContextMenu);
}

void MainWindow::onSheetTabChanged(int index) {
    if (index < 0 || index >= static_cast<int>(m_sheets.size())) return;
    switchToSheet(index);
}

void MainWindow::switchToSheet(int index) {
    if (index < 0 || index >= static_cast<int>(m_sheets.size())) return;
    m_activeSheetIndex = index;
    m_spreadsheetView->setSpreadsheet(m_sheets[index]);
    m_spreadsheetView->refreshView();
    if (m_chatPanel) m_chatPanel->setSpreadsheet(m_sheets[index]);

    // Reset scroll position and focus to A1
    QModelIndex first = m_spreadsheetView->getModel()->index(0, 0);
    m_spreadsheetView->setCurrentIndex(first);
    m_spreadsheetView->scrollTo(first, QAbstractItemView::PositionAtTop);
    // Also reset horizontal scroll
    m_spreadsheetView->horizontalScrollBar()->setValue(0);
    m_spreadsheetView->verticalScrollBar()->setValue(0);
}

void MainWindow::onSheetTabDoubleClicked(int index) {
    if (index < 0 || index >= static_cast<int>(m_sheets.size())) return;
    QString currentName = m_sheetTabBar->tabText(index);
    bool ok;
    QString newName = QInputDialog::getText(this, "Rename Sheet",
        "Sheet name:", QLineEdit::Normal, currentName, &ok);
    if (ok && !newName.isEmpty()) {
        m_sheetTabBar->setTabText(index, newName);
        m_sheets[index]->setSheetName(newName);
    }
}

void MainWindow::onAddSheet() {
    int num = nextSheetNumber();
    QString name = QString("Sheet%1").arg(num);
    auto sheet = std::make_shared<Spreadsheet>();
    sheet->setSheetName(name);
    m_sheets.push_back(sheet);
    m_sheetTabBar->addTab(name);
    m_sheetTabBar->setCurrentIndex(m_sheetTabBar->count() - 1);
    statusBar()->showMessage("Added: " + name);
}

void MainWindow::onDeleteSheet() {
    if (m_sheets.size() <= 1) {
        QMessageBox::information(this, "Delete Sheet", "Cannot delete the only sheet.");
        return;
    }

    int idx = m_sheetTabBar->currentIndex();
    if (idx < 0) return;

    QString name = m_sheetTabBar->tabText(idx);
    if (QMessageBox::question(this, "Delete Sheet",
            QString("Delete sheet \"%1\"?").arg(name)) != QMessageBox::Yes) {
        return;
    }

    // Block signals during removal to avoid triggering onSheetTabChanged prematurely
    m_sheetTabBar->blockSignals(true);
    m_sheetTabBar->removeTab(idx);
    m_sheets.erase(m_sheets.begin() + idx);
    m_sheetTabBar->blockSignals(false);

    int newIdx = qMin(idx, static_cast<int>(m_sheets.size()) - 1);
    m_sheetTabBar->setCurrentIndex(newIdx);
    switchToSheet(newIdx);
    statusBar()->showMessage("Deleted: " + name);
}

void MainWindow::onDuplicateSheet() {
    int idx = m_sheetTabBar->currentIndex();
    if (idx < 0 || idx >= static_cast<int>(m_sheets.size())) return;

    auto source = m_sheets[idx];
    auto copy = std::make_shared<Spreadsheet>();
    copy->setSheetName(source->getSheetName() + " (Copy)");
    copy->setAutoRecalculate(false);

    // Copy all cells
    source->forEachCell([&](int row, int col, const Cell& cell) {
        CellAddress addr(row, col);
        auto val = source->getCellValue(addr);
        if (val.isValid() && !val.toString().isEmpty()) {
            copy->setCellValue(addr, val);
        }
        auto srcCell = source->getCell(addr);
        auto dstCell = copy->getCell(addr);
        dstCell->setStyle(srcCell->getStyle());
    });

    copy->setRowCount(source->getRowCount());
    copy->setColumnCount(source->getColumnCount());
    copy->setAutoRecalculate(true);

    m_sheets.insert(m_sheets.begin() + idx + 1, copy);
    m_sheetTabBar->insertTab(idx + 1, copy->getSheetName());
    m_sheetTabBar->setCurrentIndex(idx + 1);
    statusBar()->showMessage("Duplicated sheet");
}

void MainWindow::showSheetContextMenu(const QPoint& pos) {
    int tabIdx = m_sheetTabBar->tabAt(pos);
    if (tabIdx < 0) tabIdx = m_sheetTabBar->currentIndex();

    QMenu menu(this);
    menu.addAction("Insert New Sheet", this, &MainWindow::onAddSheet);
    menu.addAction("Duplicate Sheet", this, &MainWindow::onDuplicateSheet);
    menu.addSeparator();
    menu.addAction("Rename Sheet", this, [this, tabIdx]() {
        onSheetTabDoubleClicked(tabIdx);
    });
    menu.addAction("Delete Sheet", this, &MainWindow::onDeleteSheet);
    menu.exec(m_sheetTabBar->mapToGlobal(pos));
}

int MainWindow::nextSheetNumber() const {
    int maxNum = 0;
    for (const auto& sheet : m_sheets) {
        QString name = sheet->getSheetName();
        if (name.startsWith("Sheet")) {
            bool ok;
            int num = name.mid(5).toInt(&ok);
            if (ok && num > maxNum) maxNum = num;
        }
    }
    return maxNum + 1;
}

void MainWindow::setSheets(const std::vector<std::shared_ptr<Spreadsheet>>& sheets) {
    m_sheets = sheets;
    m_activeSheetIndex = 0;

    // Rebuild tab bar
    m_sheetTabBar->blockSignals(true);
    while (m_sheetTabBar->count() > 0) {
        m_sheetTabBar->removeTab(0);
    }
    for (const auto& sheet : m_sheets) {
        m_sheetTabBar->addTab(sheet->getSheetName());
    }
    m_sheetTabBar->setCurrentIndex(0);
    m_sheetTabBar->blockSignals(false);

    switchToSheet(0);
}

void MainWindow::createMenuBar() {
    QMenuBar* menuBar = new QMenuBar(this);
    setMenuBar(menuBar);

    QMenu* fileMenu = menuBar->addMenu("&File");
    fileMenu->addAction("&New", this, &MainWindow::onNewDocument, QKeySequence::New);
    fileMenu->addAction("&Open", this, &MainWindow::onOpenDocument, QKeySequence::Open);
    fileMenu->addAction("&Save", this, &MainWindow::onSaveDocument, QKeySequence::Save);
    fileMenu->addAction("Save &As", this, &MainWindow::onSaveAs, QKeySequence::SaveAs);
    fileMenu->addSeparator();
    fileMenu->addAction("&Import CSV...", this, &MainWindow::onImportCsv);
    fileMenu->addAction("&Export CSV...", this, &MainWindow::onExportCsv);
    fileMenu->addSeparator();
    fileMenu->addAction("E&xit", this, &QWidget::close, QKeySequence::Quit);

    QMenu* editMenu = menuBar->addMenu("&Edit");
    editMenu->addAction("&Undo", this, &MainWindow::onUndo, QKeySequence::Undo);
    editMenu->addAction("&Redo", this, &MainWindow::onRedo, QKeySequence::Redo);
    editMenu->addSeparator();
    editMenu->addAction("Cu&t", this, &MainWindow::onCut, QKeySequence::Cut);
    editMenu->addAction("&Copy", this, &MainWindow::onCopy, QKeySequence::Copy);
    editMenu->addAction("&Paste", this, &MainWindow::onPaste, QKeySequence::Paste);
    editMenu->addAction("&Delete", this, &MainWindow::onDelete, QKeySequence::Delete);
    editMenu->addSeparator();
    editMenu->addAction("Select &All", this, &MainWindow::onSelectAll, QKeySequence::SelectAll);
    editMenu->addSeparator();
    editMenu->addAction("&Find and Replace...", this, &MainWindow::onFindReplace, QKeySequence::Find);
    editMenu->addAction("&Go To...", this, &MainWindow::onGoTo,
                         QKeySequence(Qt::CTRL | Qt::Key_G));

    QMenu* formatMenu = menuBar->addMenu("F&ormat");
    formatMenu->addAction("Format &Cells...", this, &MainWindow::onFormatCells,
                          QKeySequence(Qt::CTRL | Qt::Key_1));
    formatMenu->addSeparator();
    formatMenu->addAction("&Conditional Formatting...", this, &MainWindow::onConditionalFormat);
    formatMenu->addSeparator();
    formatMenu->addAction("Autofit Column Width", m_spreadsheetView,
                          &SpreadsheetView::autofitSelectedColumns);
    formatMenu->addAction("Autofit Row Height", m_spreadsheetView,
                          &SpreadsheetView::autofitSelectedRows);

    QMenu* dataMenu = menuBar->addMenu("&Data");
    dataMenu->addAction("Sort &Ascending", m_spreadsheetView, &SpreadsheetView::sortAscending);
    dataMenu->addAction("Sort &Descending", m_spreadsheetView, &SpreadsheetView::sortDescending);
    dataMenu->addSeparator();
    dataMenu->addAction("&Data Validation...", this, &MainWindow::onDataValidation);
    dataMenu->addSeparator();
    QAction* highlightAction = dataMenu->addAction("&Circle Invalid Data", this, &MainWindow::onHighlightInvalidCells);
    highlightAction->setCheckable(true);

    QMenu* viewMenu = menuBar->addMenu("&View");
    viewMenu->addAction("&Freeze Panes", QKeySequence(Qt::CTRL | Qt::SHIFT | Qt::Key_F),
                         this, &MainWindow::onFreezePane);
    viewMenu->addSeparator();
    QAction* chatAction = viewMenu->addAction("&Claude Assistant");
    chatAction->setCheckable(true);
    chatAction->setChecked(false);
    connect(chatAction, &QAction::toggled, this, [this](bool checked) {
        if (checked) {
            m_chatDock->show();
        } else {
            m_chatDock->hide();
        }
    });
    connect(m_chatDock, &QDockWidget::visibilityChanged, chatAction, &QAction::setChecked);
}

void MainWindow::createToolBar() {}

void MainWindow::createStatusBar() {
    statusBar()->showMessage("Ready");
    statusBar()->setStyleSheet(
        "QStatusBar { background-color: #217346; color: white; border: none; font-size: 11px; padding: 2px 8px; }"
    );
}

void MainWindow::connectSignals() {
    connect(m_toolbar, &Toolbar::newDocument, this, &MainWindow::onNewDocument);
    connect(m_toolbar, &Toolbar::saveDocument, this, &MainWindow::onSaveDocument);
    connect(m_toolbar, &Toolbar::undo, this, &MainWindow::onUndo);
    connect(m_toolbar, &Toolbar::redo, this, &MainWindow::onRedo);

    connect(m_toolbar, &Toolbar::bold, m_spreadsheetView, &SpreadsheetView::applyBold);
    connect(m_toolbar, &Toolbar::italic, m_spreadsheetView, &SpreadsheetView::applyItalic);
    connect(m_toolbar, &Toolbar::underline, m_spreadsheetView, &SpreadsheetView::applyUnderline);
    connect(m_toolbar, &Toolbar::strikethrough, m_spreadsheetView, &SpreadsheetView::applyStrikethrough);
    connect(m_toolbar, &Toolbar::fontFamilyChanged, m_spreadsheetView, &SpreadsheetView::applyFontFamily);
    connect(m_toolbar, &Toolbar::fontSizeChanged, m_spreadsheetView, &SpreadsheetView::applyFontSize);

    connect(m_toolbar, &Toolbar::foregroundColorChanged, m_spreadsheetView, &SpreadsheetView::applyForegroundColor);
    connect(m_toolbar, &Toolbar::backgroundColorChanged, m_spreadsheetView, &SpreadsheetView::applyBackgroundColor);

    connect(m_toolbar, &Toolbar::hAlignChanged, m_spreadsheetView, &SpreadsheetView::applyHAlign);
    connect(m_toolbar, &Toolbar::vAlignChanged, m_spreadsheetView, &SpreadsheetView::applyVAlign);

    connect(m_toolbar, &Toolbar::thousandSeparatorToggled, m_spreadsheetView, &SpreadsheetView::applyThousandSeparator);
    connect(m_toolbar, &Toolbar::numberFormatChanged, m_spreadsheetView, &SpreadsheetView::applyNumberFormat);
    connect(m_toolbar, &Toolbar::formatCellsRequested, this, &MainWindow::onFormatCells);

    connect(m_toolbar, &Toolbar::formatPainterToggled, m_spreadsheetView, &SpreadsheetView::activateFormatPainter);

    connect(m_toolbar, &Toolbar::sortAscending, m_spreadsheetView, &SpreadsheetView::sortAscending);
    connect(m_toolbar, &Toolbar::sortDescending, m_spreadsheetView, &SpreadsheetView::sortDescending);
    connect(m_toolbar, &Toolbar::filterToggled, m_spreadsheetView, &SpreadsheetView::toggleAutoFilter);

    connect(m_toolbar, &Toolbar::tableStyleSelected, m_spreadsheetView, &SpreadsheetView::applyTableStyle);

    connect(m_toolbar, &Toolbar::borderStyleSelected, m_spreadsheetView, &SpreadsheetView::applyBorderStyle);
    connect(m_toolbar, &Toolbar::mergeCellsRequested, m_spreadsheetView, &SpreadsheetView::mergeCells);
    connect(m_toolbar, &Toolbar::unmergeCellsRequested, m_spreadsheetView, &SpreadsheetView::unmergeCells);
    connect(m_toolbar, &Toolbar::increaseIndent, m_spreadsheetView, &SpreadsheetView::increaseIndent);
    connect(m_toolbar, &Toolbar::decreaseIndent, m_spreadsheetView, &SpreadsheetView::decreaseIndent);

    connect(m_toolbar, &Toolbar::conditionalFormatRequested, this, &MainWindow::onConditionalFormat);
    connect(m_toolbar, &Toolbar::dataValidationRequested, this, &MainWindow::onDataValidation);

    // Chat assistant toggle
    connect(m_toolbar, &Toolbar::chatToggleRequested, this, [this]() {
        if (m_chatDock->isVisible()) {
            m_chatDock->hide();
        } else {
            m_chatDock->show();
            m_chatDock->raise();
        }
    });

    // Chat NLP actions
    connect(m_chatPanel, &ChatPanel::executeActions, this, &MainWindow::onChatActions);

    connect(m_spreadsheetView, &SpreadsheetView::formatCellsRequested, this, &MainWindow::onFormatCells);

    connect(m_spreadsheetView, &SpreadsheetView::cellSelected,
            this, [this](int, int, const QString& content, const QString& address) {
        // Don't update formula bar if we're in formula editing mode (would overwrite the formula)
        if (m_spreadsheetView->isFormulaEditMode() || m_formulaBar->isFormulaEditing()) {
            return;
        }
        m_formulaBar->setCellAddress(address);
        m_formulaBar->setCellContent(content);

        // Update status bar with selection summary (SUM, AVERAGE, COUNT like Excel)
        updateStatusBarSummary();
    });

    // Selection change also updates status bar summary
    connect(m_spreadsheetView->selectionModel(), &QItemSelectionModel::selectionChanged,
            this, [this]() { updateStatusBarSummary(); });

    // Formula bar -> cell reference insertion
    connect(m_formulaBar, &FormulaBar::formulaEditModeChanged,
            m_spreadsheetView, &SpreadsheetView::setFormulaEditMode);

    // When SpreadsheetView inserts a cell reference via click, insert it into formula bar
    connect(m_spreadsheetView, &SpreadsheetView::cellReferenceInserted,
            m_formulaBar, &FormulaBar::insertText);

    connect(m_formulaBar, &FormulaBar::contentEdited,
            this, [this](const QString& content) {
        auto index = m_spreadsheetView->currentIndex();
        if (index.isValid()) {
            auto model = m_spreadsheetView->getModel();
            if (model) {
                model->setData(index, content);
            }
        }
    });
}

void MainWindow::onFormatCells() {
    if (!m_spreadsheetView) return;
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (!sheet) return;

    QModelIndex current = m_spreadsheetView->currentIndex();
    if (!current.isValid()) return;

    CellAddress addr(current.row(), current.column());
    auto cell = sheet->getCell(addr);
    CellStyle currentStyle = cell->getStyle();

    FormatCellsDialog dialog(currentStyle, this);
    if (dialog.exec() == QDialog::Accepted) {
        CellStyle newStyle = dialog.getStyle();

        QModelIndexList selected = m_spreadsheetView->selectionModel()->selectedIndexes();
        if (selected.isEmpty()) selected.append(current);

        std::vector<CellSnapshot> before, after;
        for (const auto& idx : selected) {
            CellAddress a(idx.row(), idx.column());
            before.push_back(sheet->takeCellSnapshot(a));
            auto c = sheet->getCell(a);
            c->setStyle(newStyle);
            after.push_back(sheet->takeCellSnapshot(a));
        }

        sheet->getUndoManager().execute(
            std::make_unique<StyleChangeCommand>(before, after), sheet.get());

        m_spreadsheetView->refreshView();
        statusBar()->showMessage("Format applied");
    }
}

void MainWindow::openFile(const QString& fileName) {
    if (fileName.isEmpty()) return;

    m_currentFilePath = fileName;
    QString ext = QFileInfo(fileName).suffix().toLower();

    if (ext == "xlsx" || ext == "xls") {
        auto sheets = XlsxService::importFromFile(fileName);
        if (!sheets.empty()) {
            setSheets(sheets);
            setWindowTitle("NativeSpreadsheet - " + QFileInfo(fileName).fileName());
            statusBar()->showMessage("Opened: " + fileName);
        } else {
            QMessageBox::warning(this, "Open Failed", "Could not open file: " + fileName);
        }
    } else {
        auto spreadsheet = CsvService::importFromFile(fileName);
        if (spreadsheet) {
            spreadsheet->setSheetName(QFileInfo(fileName).baseName());
            std::vector<std::shared_ptr<Spreadsheet>> sheets = { spreadsheet };
            setSheets(sheets);
            setWindowTitle("NativeSpreadsheet - " + QFileInfo(fileName).fileName());
            statusBar()->showMessage("Opened: " + fileName);
        } else {
            QMessageBox::warning(this, "Open Failed", "Could not open file: " + fileName);
        }
    }
}

void MainWindow::onNewDocument() {
    if (DocumentService::instance().getCurrentDocument()) {
        if (QMessageBox::question(this, "New Document", "Save current document before creating new one?")
            == QMessageBox::Yes) {
            onSaveDocument();
        }
    }

    auto sheet = std::make_shared<Spreadsheet>();
    sheet->setSheetName("Sheet1");
    std::vector<std::shared_ptr<Spreadsheet>> sheets = { sheet };
    setSheets(sheets);

    DocumentService::instance().createNewDocument("Untitled");
    setWindowTitle("NativeSpreadsheet");
    statusBar()->showMessage("New document created");
}

void MainWindow::onOpenDocument() {
    QString fileName = QFileDialog::getOpenFileName(this, "Open Document", "",
        "All Spreadsheet Files (*.xlsx *.csv *.txt);;Excel Files (*.xlsx);;CSV Files (*.csv);;All Files (*)");
    openFile(fileName);
}

void MainWindow::onSaveDocument() {
    if (m_currentFilePath.isEmpty()) {
        onSaveAs();
        return;
    }

    QString ext = QFileInfo(m_currentFilePath).suffix().toLower();
    bool success = false;

    if (ext == "xlsx" || ext == "xls") {
        success = XlsxService::exportToFile(m_sheets, m_currentFilePath);
    } else {
        auto spreadsheet = m_spreadsheetView->getSpreadsheet();
        if (spreadsheet) success = CsvService::exportToFile(*spreadsheet, m_currentFilePath);
    }

    if (success) {
        statusBar()->showMessage("Saved: " + m_currentFilePath);
    } else {
        QMessageBox::warning(this, "Save Failed", "Could not save file.");
    }
}

void MainWindow::onSaveAs() {
    QString fileName = QFileDialog::getSaveFileName(this, "Save Document As", "",
        "Excel Workbook (*.xlsx);;CSV Files (*.csv);;All Files (*)");
    if (fileName.isEmpty()) return;

    QString ext = QFileInfo(fileName).suffix().toLower();
    bool success = false;

    if (ext == "xlsx") {
        success = XlsxService::exportToFile(m_sheets, fileName);
    } else {
        auto spreadsheet = m_spreadsheetView->getSpreadsheet();
        if (spreadsheet) success = CsvService::exportToFile(*spreadsheet, fileName);
    }

    if (success) {
        m_currentFilePath = fileName;
        setWindowTitle("NativeSpreadsheet - " + QFileInfo(fileName).fileName());
        statusBar()->showMessage("Saved: " + fileName);
    } else {
        QMessageBox::warning(this, "Save Failed", "Could not save file.");
    }
}

void MainWindow::onUndo() {
    if (!m_spreadsheetView) return;
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (sheet && sheet->getUndoManager().canUndo()) {
        sheet->getUndoManager().undo(sheet.get());
        // Navigate to affected cell
        CellAddress target = sheet->getUndoManager().lastUndoTarget();
        auto model = m_spreadsheetView->getModel();
        if (model) {
            model->resetModel();
            QModelIndex idx = model->index(target.row, target.col);
            m_spreadsheetView->setCurrentIndex(idx);
            m_spreadsheetView->scrollTo(idx);
        }
        statusBar()->showMessage("Undo");
    }
}

void MainWindow::onRedo() {
    if (!m_spreadsheetView) return;
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (sheet && sheet->getUndoManager().canRedo()) {
        sheet->getUndoManager().redo(sheet.get());
        // Navigate to affected cell
        CellAddress target = sheet->getUndoManager().lastRedoTarget();
        auto model = m_spreadsheetView->getModel();
        if (model) {
            model->resetModel();
            QModelIndex idx = model->index(target.row, target.col);
            m_spreadsheetView->setCurrentIndex(idx);
            m_spreadsheetView->scrollTo(idx);
        }
        statusBar()->showMessage("Redo");
    }
}

void MainWindow::onCut() { if (m_spreadsheetView) m_spreadsheetView->cut(); }
void MainWindow::onCopy() { if (m_spreadsheetView) m_spreadsheetView->copy(); }
void MainWindow::onPaste() { if (m_spreadsheetView) m_spreadsheetView->paste(); }
void MainWindow::onDelete() { if (m_spreadsheetView) m_spreadsheetView->deleteSelection(); }
void MainWindow::onSelectAll() { if (m_spreadsheetView) m_spreadsheetView->selectAll(); }

void MainWindow::onImportCsv() {
    QString fileName = QFileDialog::getOpenFileName(this, "Import CSV", "",
        "CSV Files (*.csv);;Text Files (*.txt);;All Files (*)");
    if (!fileName.isEmpty()) openFile(fileName);
}

void MainWindow::onExportCsv() {
    QString fileName = QFileDialog::getSaveFileName(this, "Export CSV", "",
        "CSV Files (*.csv);;All Files (*)");
    if (fileName.isEmpty()) return;

    auto spreadsheet = m_spreadsheetView->getSpreadsheet();
    if (spreadsheet && CsvService::exportToFile(*spreadsheet, fileName)) {
        statusBar()->showMessage("Exported: " + fileName);
    } else {
        QMessageBox::warning(this, "Export Failed", "Could not export CSV file.");
    }
}

void MainWindow::closeEvent(QCloseEvent* event) {
    if (saveCurrentDocument()) event->accept();
    else event->ignore();
}

void MainWindow::dragEnterEvent(QDragEnterEvent* event) {
    if (event->mimeData()->hasUrls()) event->acceptProposedAction();
}

void MainWindow::dropEvent(QDropEvent* event) {
    const QMimeData* mimeData = event->mimeData();
    if (mimeData->hasUrls()) {
        openFile(mimeData->urls().first().toLocalFile());
        event->acceptProposedAction();
    }
}

// ============== Find/Replace ==============

void MainWindow::onFindReplace() {
    if (!m_findReplaceDialog) {
        m_findReplaceDialog = new FindReplaceDialog(this);
        connect(m_findReplaceDialog, &FindReplaceDialog::findNext, this, &MainWindow::onFindNext);
        connect(m_findReplaceDialog, &FindReplaceDialog::findPrevious, this, &MainWindow::onFindPrevious);
        connect(m_findReplaceDialog, &FindReplaceDialog::replaceOne, this, &MainWindow::onReplaceOne);
        connect(m_findReplaceDialog, &FindReplaceDialog::replaceAll, this, &MainWindow::onReplaceAll);
    }
    m_findReplaceDialog->show();
    m_findReplaceDialog->raise();
    m_findReplaceDialog->activateWindow();
}

bool MainWindow::cellMatchesSearch(int row, int col, const QString& searchText, bool matchCase, bool wholeCell) const {
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (!sheet) return false;

    auto val = sheet->getCellValue(CellAddress(row, col));
    QString cellText = val.toString();

    if (wholeCell) {
        return matchCase ? (cellText == searchText)
                         : (cellText.compare(searchText, Qt::CaseInsensitive) == 0);
    } else {
        return cellText.contains(searchText, matchCase ? Qt::CaseSensitive : Qt::CaseInsensitive);
    }
}

void MainWindow::onFindNext() {
    if (!m_findReplaceDialog || !m_spreadsheetView) return;

    QString searchText = m_findReplaceDialog->findText();
    if (searchText.isEmpty()) return;

    bool matchCase = m_findReplaceDialog->matchCase();
    bool wholeCell = m_findReplaceDialog->matchWholeCell();
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (!sheet) return;

    int maxRow = sheet->getMaxRow();
    int maxCol = sheet->getMaxColumn();
    if (maxRow < 0 || maxCol < 0) {
        m_findReplaceDialog->setStatus("No data to search.");
        return;
    }

    QModelIndex current = m_spreadsheetView->currentIndex();
    int startRow = current.isValid() ? current.row() : 0;
    int startCol = current.isValid() ? current.column() + 1 : 0;

    // Search forward: row by row, column by column
    for (int r = startRow; r <= maxRow; ++r) {
        int cStart = (r == startRow) ? startCol : 0;
        for (int c = cStart; c <= maxCol; ++c) {
            if (cellMatchesSearch(r, c, searchText, matchCase, wholeCell)) {
                auto model = m_spreadsheetView->getModel();
                QModelIndex idx = model->index(r, c);
                m_spreadsheetView->setCurrentIndex(idx);
                m_spreadsheetView->scrollTo(idx);
                m_findReplaceDialog->setStatus(QString("Found at %1").arg(CellAddress(r, c).toString()));
                return;
            }
        }
    }

    // Wrap around from top
    for (int r = 0; r <= startRow; ++r) {
        int cEnd = (r == startRow) ? startCol - 1 : maxCol;
        for (int c = 0; c <= cEnd; ++c) {
            if (cellMatchesSearch(r, c, searchText, matchCase, wholeCell)) {
                auto model = m_spreadsheetView->getModel();
                QModelIndex idx = model->index(r, c);
                m_spreadsheetView->setCurrentIndex(idx);
                m_spreadsheetView->scrollTo(idx);
                m_findReplaceDialog->setStatus(QString("Found at %1 (wrapped)").arg(CellAddress(r, c).toString()));
                return;
            }
        }
    }

    m_findReplaceDialog->setStatus("Not found.");
}

void MainWindow::onFindPrevious() {
    if (!m_findReplaceDialog || !m_spreadsheetView) return;

    QString searchText = m_findReplaceDialog->findText();
    if (searchText.isEmpty()) return;

    bool matchCase = m_findReplaceDialog->matchCase();
    bool wholeCell = m_findReplaceDialog->matchWholeCell();
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (!sheet) return;

    int maxRow = sheet->getMaxRow();
    int maxCol = sheet->getMaxColumn();
    if (maxRow < 0 || maxCol < 0) return;

    QModelIndex current = m_spreadsheetView->currentIndex();
    int startRow = current.isValid() ? current.row() : maxRow;
    int startCol = current.isValid() ? current.column() - 1 : maxCol;

    // Search backward
    for (int r = startRow; r >= 0; --r) {
        int cStart = (r == startRow) ? startCol : maxCol;
        for (int c = cStart; c >= 0; --c) {
            if (cellMatchesSearch(r, c, searchText, matchCase, wholeCell)) {
                auto model = m_spreadsheetView->getModel();
                QModelIndex idx = model->index(r, c);
                m_spreadsheetView->setCurrentIndex(idx);
                m_spreadsheetView->scrollTo(idx);
                m_findReplaceDialog->setStatus(QString("Found at %1").arg(CellAddress(r, c).toString()));
                return;
            }
        }
    }

    // Wrap around from bottom
    for (int r = maxRow; r >= startRow; --r) {
        int cStart = (r == startRow) ? startCol + 1 : 0;
        for (int c = maxCol; c >= cStart; --c) {
            if (cellMatchesSearch(r, c, searchText, matchCase, wholeCell)) {
                auto model = m_spreadsheetView->getModel();
                QModelIndex idx = model->index(r, c);
                m_spreadsheetView->setCurrentIndex(idx);
                m_spreadsheetView->scrollTo(idx);
                m_findReplaceDialog->setStatus(QString("Found at %1 (wrapped)").arg(CellAddress(r, c).toString()));
                return;
            }
        }
    }

    m_findReplaceDialog->setStatus("Not found.");
}

void MainWindow::onReplaceOne() {
    if (!m_findReplaceDialog || !m_spreadsheetView) return;

    QString searchText = m_findReplaceDialog->findText();
    QString replaceText = m_findReplaceDialog->replaceText();
    if (searchText.isEmpty()) return;

    bool matchCase = m_findReplaceDialog->matchCase();
    bool wholeCell = m_findReplaceDialog->matchWholeCell();

    QModelIndex current = m_spreadsheetView->currentIndex();
    if (!current.isValid()) return;

    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (!sheet) return;

    if (cellMatchesSearch(current.row(), current.column(), searchText, matchCase, wholeCell)) {
        auto model = m_spreadsheetView->getModel();
        if (wholeCell) {
            model->setData(current, replaceText);
        } else {
            auto val = sheet->getCellValue(CellAddress(current.row(), current.column()));
            QString cellText = val.toString();
            cellText.replace(searchText, replaceText, matchCase ? Qt::CaseSensitive : Qt::CaseInsensitive);
            model->setData(current, cellText);
        }
        m_findReplaceDialog->setStatus("Replaced. Finding next...");
        onFindNext();
    } else {
        onFindNext();
    }
}

void MainWindow::onReplaceAll() {
    if (!m_findReplaceDialog || !m_spreadsheetView) return;

    QString searchText = m_findReplaceDialog->findText();
    QString replaceText = m_findReplaceDialog->replaceText();
    if (searchText.isEmpty()) return;

    bool matchCase = m_findReplaceDialog->matchCase();
    bool wholeCell = m_findReplaceDialog->matchWholeCell();
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (!sheet) return;

    auto model = m_spreadsheetView->getModel();
    int maxRow = sheet->getMaxRow();
    int maxCol = sheet->getMaxColumn();
    int count = 0;

    std::vector<CellSnapshot> before, after;
    model->setSuppressUndo(true);

    for (int r = 0; r <= maxRow; ++r) {
        for (int c = 0; c <= maxCol; ++c) {
            if (cellMatchesSearch(r, c, searchText, matchCase, wholeCell)) {
                CellAddress addr(r, c);
                before.push_back(sheet->takeCellSnapshot(addr));

                QModelIndex idx = model->index(r, c);
                if (wholeCell) {
                    model->setData(idx, replaceText);
                } else {
                    auto val = sheet->getCellValue(addr);
                    QString cellText = val.toString();
                    cellText.replace(searchText, replaceText, matchCase ? Qt::CaseSensitive : Qt::CaseInsensitive);
                    model->setData(idx, cellText);
                }
                after.push_back(sheet->takeCellSnapshot(addr));
                count++;
            }
        }
    }

    model->setSuppressUndo(false);

    if (!before.empty()) {
        sheet->getUndoManager().pushCommand(
            std::make_unique<MultiCellEditCommand>(before, after, "Replace All"));
    }

    m_findReplaceDialog->setStatus(QString("Replaced %1 occurrence(s).").arg(count));
    m_spreadsheetView->refreshView();
}

// ============== Go To ==============

void MainWindow::onGoTo() {
    GoToDialog dialog(this);
    if (dialog.exec() == QDialog::Accepted) {
        CellAddress addr = dialog.getAddress();
        if (addr.row >= 0 && addr.col >= 0) {
            auto model = m_spreadsheetView->getModel();
            if (model) {
                QModelIndex idx = model->index(addr.row, addr.col);
                m_spreadsheetView->setCurrentIndex(idx);
                m_spreadsheetView->scrollTo(idx, QAbstractItemView::PositionAtCenter);
                statusBar()->showMessage(QString("Navigated to %1").arg(addr.toString()));
            }
        } else {
            QMessageBox::warning(this, "Go To", "Invalid cell reference.");
        }
    }
}

// ============== Conditional Formatting ==============

void MainWindow::onConditionalFormat() {
    if (!m_spreadsheetView) return;
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (!sheet) return;

    QModelIndexList selected = m_spreadsheetView->selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    // Get selection bounding box
    int minRow = INT_MAX, maxRow = 0, minCol = INT_MAX, maxCol = 0;
    for (const auto& idx : selected) {
        minRow = qMin(minRow, idx.row());
        maxRow = qMax(maxRow, idx.row());
        minCol = qMin(minCol, idx.column());
        maxCol = qMax(maxCol, idx.column());
    }

    CellRange range(CellAddress(minRow, minCol), CellAddress(maxRow, maxCol));
    ConditionalFormatDialog dialog(range, sheet->getConditionalFormatting(), this);
    if (dialog.exec() == QDialog::Accepted) {
        m_spreadsheetView->refreshView();
        statusBar()->showMessage("Conditional formatting applied");
    }
}

// ============== Data Validation ==============

void MainWindow::onDataValidation() {
    if (!m_spreadsheetView) return;
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (!sheet) return;

    QModelIndexList selected = m_spreadsheetView->selectionModel()->selectedIndexes();
    if (selected.isEmpty()) return;

    int minRow = INT_MAX, maxRow = 0, minCol = INT_MAX, maxCol = 0;
    for (const auto& idx : selected) {
        minRow = qMin(minRow, idx.row());
        maxRow = qMax(maxRow, idx.row());
        minCol = qMin(minCol, idx.column());
        maxCol = qMax(maxCol, idx.column());
    }

    CellRange range(CellAddress(minRow, minCol), CellAddress(maxRow, maxCol));
    DataValidationDialog dialog(range, this);

    // Load existing rule if present
    const auto* existing = sheet->getValidationAt(minRow, minCol);
    if (existing) {
        dialog.setRule(*existing);
    }

    if (dialog.exec() == QDialog::Accepted) {
        auto rule = dialog.getRule();
        // Remove old rules for this range
        auto& rules = sheet->getValidationRules();
        for (int i = static_cast<int>(rules.size()) - 1; i >= 0; --i) {
            if (rules[i].range.intersects(range)) {
                sheet->removeValidationRule(i);
            }
        }
        sheet->addValidationRule(rule);
        statusBar()->showMessage("Data validation applied");
    }
}

// ============== Freeze Panes ==============

void MainWindow::onFreezePane() {
    if (!m_spreadsheetView) return;

    QModelIndex current = m_spreadsheetView->currentIndex();
    if (!current.isValid()) return;

    if (m_frozenPanes) {
        // Unfreeze
        m_spreadsheetView->setFrozenRow(-1);
        m_spreadsheetView->setFrozenColumn(-1);
        m_frozenPanes = false;
        statusBar()->showMessage("Panes unfrozen");
    } else {
        // Freeze at current cell position
        m_spreadsheetView->setFrozenRow(current.row());
        m_spreadsheetView->setFrozenColumn(current.column());
        m_frozenPanes = true;
        statusBar()->showMessage(QString("Panes frozen at %1")
            .arg(CellAddress(current.row(), current.column()).toString()));
    }
}

void MainWindow::onHighlightInvalidCells() {
    if (!m_spreadsheetView || !m_spreadsheetView->getModel()) return;
    auto* model = m_spreadsheetView->getModel();
    bool current = model->highlightInvalidCells();
    model->setHighlightInvalidCells(!current);
    model->resetModel();
    statusBar()->showMessage(current ? "Invalid cell highlighting off" : "Invalid cell highlighting on");
}

void MainWindow::updateStatusBarSummary() {
    if (!m_spreadsheetView) return;
    auto sheet = m_spreadsheetView->getSpreadsheet();
    if (!sheet) return;

    QModelIndexList selected = m_spreadsheetView->selectionModel()->selectedIndexes();
    if (selected.size() <= 1) {
        statusBar()->showMessage("Ready");
        return;
    }

    // Compute SUM, AVERAGE, COUNT for numeric values (like Excel status bar)
    double sum = 0;
    int numericCount = 0;
    int nonEmptyCount = 0;

    // Limit calculation to avoid lag on huge selections
    int limit = qMin(selected.size(), static_cast<qsizetype>(50000));
    for (int i = 0; i < limit; ++i) {
        const auto& idx = selected[i];
        auto val = sheet->getCellValue(CellAddress(idx.row(), idx.column()));
        QString text = val.toString();
        if (!text.isEmpty()) {
            nonEmptyCount++;
            bool ok;
            double num = text.toDouble(&ok);
            if (ok) {
                sum += num;
                numericCount++;
            }
        }
    }

    if (numericCount > 0) {
        double avg = sum / numericCount;
        statusBar()->showMessage(
            QString("Average: %1   Count: %2   Sum: %3")
                .arg(QString::number(avg, 'f', 2))
                .arg(nonEmptyCount)
                .arg(QString::number(sum, 'f', 2)));
    } else if (nonEmptyCount > 0) {
        statusBar()->showMessage(QString("Count: %1").arg(nonEmptyCount));
    } else {
        statusBar()->showMessage("Ready");
    }
}

// ============== Chat NLP Actions ==============

static CellAddress parseCellRef(const QString& ref) {
    int col = 0;
    int i = 0;
    while (i < ref.length() && ref[i].isLetter()) {
        col = col * 26 + (ref[i].toUpper().unicode() - 'A' + 1);
        i++;
    }
    col--; // 0-indexed
    int row = ref.mid(i).toInt() - 1; // 0-indexed
    return CellAddress(qMax(0, row), qMax(0, col));
}

static CellAddress parseRangeStart(const QString& rangeStr) {
    QStringList parts = rangeStr.split(':');
    return parseCellRef(parts[0]);
}

static CellAddress parseRangeEnd(const QString& rangeStr) {
    QStringList parts = rangeStr.split(':');
    return (parts.size() > 1) ? parseCellRef(parts[1]) : parseCellRef(parts[0]);
}

static int parseColLetter(const QString& col) {
    int result = 0;
    for (int i = 0; i < col.length(); ++i) {
        result = result * 26 + (col[i].toUpper().unicode() - 'A' + 1);
    }
    return result - 1;
}

void MainWindow::onChatActions(const QJsonArray& actions) {
    if (!m_spreadsheetView || m_sheets.empty()) return;

    auto sheet = m_sheets[m_activeSheetIndex];
    if (!sheet) return;

    for (const auto& item : actions) {
        QJsonObject action = item.toObject();
        QString type = action["action"].toString();

        if (type == "set_cell") {
            QString cellRef = action["cell"].toString();
            QJsonValue val = action["value"];
            CellAddress addr = parseCellRef(cellRef);
            auto cell = sheet->getCell(addr);
            if (val.isDouble()) {
                cell->setValue(val.toDouble());
            } else {
                cell->setValue(val.toString());
            }

        } else if (type == "set_formula") {
            QString cellRef = action["cell"].toString();
            QString formula = action["formula"].toString();
            CellAddress addr = parseCellRef(cellRef);
            sheet->setCellFormula(addr, formula);

        } else if (type == "format") {
            CellAddress start = parseRangeStart(action["range"].toString());
            CellAddress end = parseRangeEnd(action["range"].toString());

            for (int r = start.row; r <= end.row; ++r) {
                for (int c = start.col; c <= end.col; ++c) {
                    CellAddress addr(r, c);
                    auto cell = sheet->getCell(addr);
                    CellStyle style = cell->getStyle();

                    if (action.contains("bold")) style.bold = action["bold"].toBool();
                    if (action.contains("italic")) style.italic = action["italic"].toBool();
                    if (action.contains("underline")) style.underline = action["underline"].toBool();
                    if (action.contains("strikethrough")) style.strikethrough = action["strikethrough"].toBool();
                    if (action.contains("bg_color")) style.backgroundColor = action["bg_color"].toString();
                    if (action.contains("fg_color")) style.foregroundColor = action["fg_color"].toString();
                    if (action.contains("font_size")) style.fontSize = action["font_size"].toInt();
                    if (action.contains("font_name")) style.fontName = action["font_name"].toString();
                    if (action.contains("h_align")) {
                        QString align = action["h_align"].toString();
                        if (align == "left") style.hAlign = HorizontalAlignment::Left;
                        else if (align == "center") style.hAlign = HorizontalAlignment::Center;
                        else if (align == "right") style.hAlign = HorizontalAlignment::Right;
                    }
                    if (action.contains("v_align")) {
                        QString align = action["v_align"].toString();
                        if (align == "top") style.vAlign = VerticalAlignment::Top;
                        else if (align == "middle") style.vAlign = VerticalAlignment::Middle;
                        else if (align == "bottom") style.vAlign = VerticalAlignment::Bottom;
                    }

                    cell->setStyle(style);
                }
            }

        } else if (type == "merge") {
            CellAddress start = parseRangeStart(action["range"].toString());
            CellAddress end = parseRangeEnd(action["range"].toString());
            CellRange range(start, end);
            sheet->mergeCells(range);
            int rowSpan = end.row - start.row + 1;
            int colSpan = end.col - start.col + 1;
            m_spreadsheetView->setSpan(start.row, start.col, rowSpan, colSpan);
            // Center merged content
            auto cell = sheet->getCell(start);
            CellStyle style = cell->getStyle();
            style.hAlign = HorizontalAlignment::Center;
            style.vAlign = VerticalAlignment::Middle;
            cell->setStyle(style);

        } else if (type == "unmerge") {
            CellAddress start = parseRangeStart(action["range"].toString());
            CellAddress end = parseRangeEnd(action["range"].toString());
            CellRange range(start, end);
            m_spreadsheetView->setSpan(start.row, start.col, 1, 1);
            sheet->unmergeCells(range);

        } else if (type == "border") {
            QString borderType = action["type"].toString();
            CellAddress start = parseRangeStart(action["range"].toString());
            CellAddress end = parseRangeEnd(action["range"].toString());

            BorderStyle on;
            on.enabled = true;
            on.color = "#000000";
            on.width = (borderType == "thick_outside") ? 2 : 1;

            BorderStyle off;
            off.enabled = false;

            for (int r = start.row; r <= end.row; ++r) {
                for (int c = start.col; c <= end.col; ++c) {
                    CellAddress addr(r, c);
                    auto cell = sheet->getCell(addr);
                    CellStyle style = cell->getStyle();

                    if (borderType == "none") {
                        style.borderTop = off; style.borderBottom = off;
                        style.borderLeft = off; style.borderRight = off;
                    } else if (borderType == "all") {
                        style.borderTop = on; style.borderBottom = on;
                        style.borderLeft = on; style.borderRight = on;
                    } else if (borderType == "outside" || borderType == "thick_outside") {
                        if (r == start.row) style.borderTop = on;
                        if (r == end.row) style.borderBottom = on;
                        if (c == start.col) style.borderLeft = on;
                        if (c == end.col) style.borderRight = on;
                    } else if (borderType == "bottom") {
                        if (r == end.row) style.borderBottom = on;
                    } else if (borderType == "top") {
                        if (r == start.row) style.borderTop = on;
                    } else if (borderType == "left") {
                        if (c == start.col) style.borderLeft = on;
                    } else if (borderType == "right") {
                        if (c == end.col) style.borderRight = on;
                    }

                    cell->setStyle(style);
                }
            }

        } else if (type == "table") {
            CellAddress start = parseRangeStart(action["range"].toString());
            CellAddress end = parseRangeEnd(action["range"].toString());
            int themeIdx = action["theme"].toInt();
            auto themes = getBuiltinTableThemes();
            if (themeIdx >= 0 && themeIdx < static_cast<int>(themes.size())) {
                SpreadsheetTable table;
                table.range = CellRange(start, end);
                table.theme = themes[themeIdx];
                table.hasHeaderRow = true;
                table.bandedRows = true;
                int tableNum = static_cast<int>(sheet->getTables().size()) + 1;
                table.name = QString("Table%1").arg(tableNum);
                for (int c = start.col; c <= end.col; ++c) {
                    auto val = sheet->getCellValue(CellAddress(start.row, c));
                    QString name = val.toString();
                    if (name.isEmpty()) name = QString("Column%1").arg(c - start.col + 1);
                    table.columnNames.append(name);
                }
                sheet->addTable(table);
            }

        } else if (type == "number_format") {
            CellAddress start = parseRangeStart(action["range"].toString());
            CellAddress end = parseRangeEnd(action["range"].toString());
            QString fmt = action["format"].toString();
            for (int r = start.row; r <= end.row; ++r) {
                for (int c = start.col; c <= end.col; ++c) {
                    auto cell = sheet->getCell(CellAddress(r, c));
                    CellStyle style = cell->getStyle();
                    style.numberFormat = fmt;
                    cell->setStyle(style);
                }
            }

        } else if (type == "set_row_height") {
            int row = action["row"].toInt() - 1; // 1-based to 0-based
            int height = action["height"].toInt();
            if (row >= 0 && height > 0) {
                m_spreadsheetView->setRowHeight(row, height);
            }

        } else if (type == "set_col_width") {
            QString colStr = action["col"].toString();
            int col = parseColLetter(colStr);
            int width = action["width"].toInt();
            if (col >= 0 && width > 0) {
                m_spreadsheetView->setColumnWidth(col, width);
            }

        } else if (type == "clear") {
            CellAddress start = parseRangeStart(action["range"].toString());
            CellAddress end = parseRangeEnd(action["range"].toString());
            CellRange range(start, end);
            sheet->clearRange(range);
        }
    }

    // Refresh the view
    m_spreadsheetView->refreshView();
    if (m_spreadsheetView->getModel())
        m_spreadsheetView->getModel()->resetModel();
    statusBar()->showMessage(QString("Claude applied %1 action(s)").arg(actions.size()), 5000);
}

bool MainWindow::saveCurrentDocument() {
    auto doc = DocumentService::instance().getCurrentDocument();
    if (doc) return DocumentService::instance().saveDocument();
    return true;
}
