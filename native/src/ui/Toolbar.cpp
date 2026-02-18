#include "Toolbar.h"
#include <QAction>
#include <QToolButton>
#include <QSpinBox>
#include <QFontComboBox>
#include <QColorDialog>
#include <QLabel>
#include <QPainter>
#include <QPainterPath>
#include <QPixmap>
#include <QMenu>
#include <QPolygon>
#include "../core/TableStyle.h"

// Create horizontal alignment icon: lines of varying length positioned left/center/right
static QIcon createHAlignIcon(const QString& align) {
    QPixmap pix(16, 16);
    pix.fill(Qt::transparent);
    QPainter p(&pix);
    p.setRenderHint(QPainter::Antialiasing, false);
    p.setPen(QPen(QColor("#444"), 1.5));

    int widths[] = {12, 8, 10, 6};
    for (int i = 0; i < 4; ++i) {
        int y = 3 + i * 3;
        int w = widths[i];
        int x;
        if (align == "left") x = 2;
        else if (align == "center") x = (16 - w) / 2;
        else x = 14 - w;
        p.drawLine(x, y, x + w, y);
    }
    return QIcon(pix);
}

// Create sort icon: arrow with A/Z labels
static QIcon createSortIcon(bool ascending) {
    QPixmap pix(16, 16);
    pix.fill(Qt::transparent);
    QPainter p(&pix);
    p.setRenderHint(QPainter::Antialiasing, true);

    // Draw arrow
    p.setPen(QPen(QColor("#217346"), 1.5));
    p.setBrush(QColor("#217346"));
    if (ascending) {
        // Down arrow on right side
        p.drawLine(11, 3, 11, 13);
        QPolygon arrowHead;
        arrowHead << QPoint(8, 10) << QPoint(11, 14) << QPoint(14, 10);
        p.drawPolygon(arrowHead);
    } else {
        // Up arrow on right side
        p.drawLine(11, 3, 11, 13);
        QPolygon arrowHead;
        arrowHead << QPoint(8, 6) << QPoint(11, 2) << QPoint(14, 6);
        p.drawPolygon(arrowHead);
    }

    // Draw A/Z text
    p.setPen(QColor("#333"));
    p.setBrush(Qt::NoBrush);
    QFont f("Arial", 6, QFont::Bold);
    p.setFont(f);
    if (ascending) {
        p.drawText(QRect(0, 0, 9, 9), Qt::AlignCenter, "A");
        p.drawText(QRect(0, 7, 9, 9), Qt::AlignCenter, "Z");
    } else {
        p.drawText(QRect(0, 0, 9, 9), Qt::AlignCenter, "Z");
        p.drawText(QRect(0, 7, 9, 9), Qt::AlignCenter, "A");
    }

    return QIcon(pix);
}

// Create undo/redo icon: clean curved arrow (Excel-style)
static QIcon createUndoRedoIcon(bool isUndo) {
    QPixmap pix(20, 20);
    pix.fill(Qt::transparent);
    QPainter p(&pix);
    p.setRenderHint(QPainter::Antialiasing, true);

    QColor color("#217346");

    if (isUndo) {
        // Draw a smooth counter-clockwise curved arrow
        // Arc: from right side curving up and around to the left
        QPainterPath arc;
        arc.moveTo(4, 10);
        arc.cubicTo(4, 4, 10, 3, 15, 5);
        arc.cubicTo(17, 6, 17, 10, 15, 12);
        p.setPen(QPen(color, 2.0, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin));
        p.setBrush(Qt::NoBrush);
        p.drawPath(arc);

        // Arrowhead at the left end pointing left-down
        p.setPen(Qt::NoPen);
        p.setBrush(color);
        QPolygonF arrow;
        arrow << QPointF(1, 10) << QPointF(5.5, 7) << QPointF(5.5, 13);
        p.drawPolygon(arrow);
    } else {
        // Draw a smooth clockwise curved arrow (mirror of undo)
        QPainterPath arc;
        arc.moveTo(16, 10);
        arc.cubicTo(16, 4, 10, 3, 5, 5);
        arc.cubicTo(3, 6, 3, 10, 5, 12);
        p.setPen(QPen(color, 2.0, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin));
        p.setBrush(Qt::NoBrush);
        p.drawPath(arc);

        // Arrowhead at the right end pointing right-down
        p.setPen(Qt::NoPen);
        p.setBrush(color);
        QPolygonF arrow;
        arrow << QPointF(19, 10) << QPointF(14.5, 7) << QPointF(14.5, 13);
        p.drawPolygon(arrow);
    }

    return QIcon(pix);
}

// Create vertical alignment icon: lines at top/middle/bottom within a cell box
static QIcon createVAlignIcon(const QString& align) {
    QPixmap pix(16, 16);
    pix.fill(Qt::transparent);
    QPainter p(&pix);
    p.setRenderHint(QPainter::Antialiasing, false);

    // Draw cell border
    p.setPen(QPen(QColor("#AAA"), 0.5));
    p.drawRect(1, 1, 13, 13);

    // Draw lines representing text
    p.setPen(QPen(QColor("#444"), 1.5));
    int startY;
    if (align == "top") startY = 3;
    else if (align == "middle") startY = 5;
    else startY = 8;

    p.drawLine(3, startY, 12, startY);
    p.drawLine(3, startY + 3, 10, startY + 3);
    return QIcon(pix);
}

static const char* TOOLBAR_STYLE = R"(
    QToolBar {
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #217346, stop:0.06 #217346,
            stop:0.061 #F3F3F3, stop:1 #F3F3F3);
        border-bottom: 1px solid #C8C8C8;
        spacing: 1px;
        padding: 6px 6px 3px 6px;
    }
    QToolButton {
        background: transparent;
        border: 1px solid transparent;
        border-radius: 2px;
        padding: 3px 5px;
        margin: 0px 1px;
        font-size: 12px;
        color: #333;
    }
    QToolButton:hover {
        background-color: #E5E5E5;
        border-color: #CCCCCC;
    }
    QToolButton:pressed {
        background-color: #CDCDCD;
    }
    QToolButton:checked {
        background-color: #D0DEE8;
        border-color: #9CB8CC;
    }
    QFontComboBox {
        max-width: 140px;
        min-width: 100px;
        height: 24px;
        border: 1px solid #C8C8C8;
        border-radius: 2px;
        padding: 1px 4px;
        background: white;
        font-size: 12px;
    }
    QSpinBox {
        max-width: 46px;
        height: 24px;
        border: 1px solid #C8C8C8;
        border-radius: 2px;
        padding: 1px 4px;
        background: white;
        font-size: 12px;
    }
    QToolBar::separator {
        width: 1px;
        background-color: #D0D0D0;
        margin: 2px 3px;
    }
)";

Toolbar::Toolbar(QWidget* parent)
    : QToolBar("Standard Toolbar", parent) {

    setMovable(false);
    setFloatable(false);
    setIconSize(QSize(16, 16));
    setStyleSheet(TOOLBAR_STYLE);

    createActions();
}

void Toolbar::createActions() {
    // ===== File =====
    QToolButton* newBtn = new QToolButton(this);
    newBtn->setText("New");
    newBtn->setToolTip("New Document (Ctrl+N)");
    newBtn->setShortcut(QKeySequence::New);
    addWidget(newBtn);
    connect(newBtn, &QToolButton::clicked, this, &Toolbar::newDocument);

    QToolButton* saveBtn = new QToolButton(this);
    saveBtn->setText("Save");
    saveBtn->setToolTip("Save Document (Ctrl+S)");
    saveBtn->setShortcut(QKeySequence::Save);
    addWidget(saveBtn);
    connect(saveBtn, &QToolButton::clicked, this, &Toolbar::saveDocument);

    addSeparator();

    // ===== Undo/Redo =====
    QToolButton* undoBtn = new QToolButton(this);
    undoBtn->setIcon(createUndoRedoIcon(true));
    undoBtn->setIconSize(QSize(20, 20));
    undoBtn->setToolTip("Undo (Ctrl+Z)");
    undoBtn->setShortcut(QKeySequence::Undo);
    undoBtn->setFixedSize(30, 26);
    addWidget(undoBtn);
    connect(undoBtn, &QToolButton::clicked, this, &Toolbar::undo);

    QToolButton* redoBtn = new QToolButton(this);
    redoBtn->setIcon(createUndoRedoIcon(false));
    redoBtn->setIconSize(QSize(20, 20));
    redoBtn->setToolTip("Redo (Ctrl+Y)");
    redoBtn->setShortcut(QKeySequence::Redo);
    redoBtn->setFixedSize(30, 26);
    addWidget(redoBtn);
    connect(redoBtn, &QToolButton::clicked, this, &Toolbar::redo);

    addSeparator();

    // ===== Format Painter =====
    QToolButton* formatPainterBtn = new QToolButton(this);
    formatPainterBtn->setText("\u2702");  // ✂ scissors-like icon
    formatPainterBtn->setFont(QFont("Apple Symbols", 14));
    formatPainterBtn->setToolTip("Format Painter");
    formatPainterBtn->setCheckable(true);
    formatPainterBtn->setFixedSize(30, 26);
    formatPainterBtn->setStyleSheet(
        "QToolButton { font-size: 15px; color: #217346; }"
        "QToolButton:checked { background-color: #C8E6C9; border: 1px solid #217346; }");
    addWidget(formatPainterBtn);
    connect(formatPainterBtn, &QToolButton::clicked, this, &Toolbar::formatPainterToggled);

    addSeparator();

    // ===== Font =====
    m_fontCombo = new QFontComboBox(this);
    addWidget(m_fontCombo);
    connect(m_fontCombo, &QFontComboBox::currentFontChanged, this, [this](const QFont& font) {
        emit fontFamilyChanged(font.family());
    });

    m_fontSizeSpinBox = new QSpinBox(this);
    m_fontSizeSpinBox->setRange(6, 72);
    m_fontSizeSpinBox->setValue(11);
    addWidget(m_fontSizeSpinBox);
    connect(m_fontSizeSpinBox, &QSpinBox::valueChanged, this, &Toolbar::fontSizeChanged);

    addSeparator();

    // ===== B I U S =====
    QToolButton* boldBtn = new QToolButton(this);
    boldBtn->setText("B");
    boldBtn->setFont(QFont("Arial", 12, QFont::Bold));
    boldBtn->setShortcut(QKeySequence::Bold);
    boldBtn->setCheckable(true);
    boldBtn->setToolTip("Bold (Ctrl+B)");
    boldBtn->setFixedSize(26, 26);
    addWidget(boldBtn);
    connect(boldBtn, &QToolButton::clicked, this, [this]() { emit bold(); });

    QToolButton* italicBtn = new QToolButton(this);
    italicBtn->setText("I");
    QFont italicFont("Arial", 12);
    italicFont.setItalic(true);
    italicBtn->setFont(italicFont);
    italicBtn->setShortcut(QKeySequence::Italic);
    italicBtn->setCheckable(true);
    italicBtn->setToolTip("Italic (Ctrl+I)");
    italicBtn->setFixedSize(26, 26);
    addWidget(italicBtn);
    connect(italicBtn, &QToolButton::clicked, this, [this]() { emit italic(); });

    QToolButton* underlineBtn = new QToolButton(this);
    underlineBtn->setText("U");
    QFont underlineFont("Arial", 12);
    underlineFont.setUnderline(true);
    underlineBtn->setFont(underlineFont);
    underlineBtn->setShortcut(QKeySequence::Underline);
    underlineBtn->setCheckable(true);
    underlineBtn->setToolTip("Underline (Ctrl+U)");
    underlineBtn->setFixedSize(26, 26);
    addWidget(underlineBtn);
    connect(underlineBtn, &QToolButton::clicked, this, [this]() { emit underline(); });

    QToolButton* strikeBtn = new QToolButton(this);
    strikeBtn->setText("S");
    QFont strikeFont("Arial", 12);
    strikeFont.setStrikeOut(true);
    strikeBtn->setFont(strikeFont);
    strikeBtn->setCheckable(true);
    strikeBtn->setToolTip("Strikethrough");
    strikeBtn->setFixedSize(26, 26);
    addWidget(strikeBtn);
    connect(strikeBtn, &QToolButton::clicked, this, [this]() { emit strikethrough(); });

    addSeparator();

    // ===== Colors =====
    QToolButton* fgColorBtn = new QToolButton(this);
    fgColorBtn->setText("A");
    fgColorBtn->setFont(QFont("Arial", 12, QFont::Bold));
    fgColorBtn->setStyleSheet(
        "QToolButton { color: #C00000; font-weight: bold; border-bottom: 3px solid #C00000; }");
    fgColorBtn->setToolTip("Font Color");
    fgColorBtn->setFixedSize(26, 26);
    addWidget(fgColorBtn);
    connect(fgColorBtn, &QToolButton::clicked, this, [this, fgColorBtn]() {
        QColor color = QColorDialog::getColor(m_lastFgColor, this, "Font Color");
        if (color.isValid()) {
            m_lastFgColor = color;
            fgColorBtn->setStyleSheet(
                QString("QToolButton { color: %1; font-weight: bold; border-bottom: 3px solid %1; }")
                    .arg(color.name()));
            emit foregroundColorChanged(color);
        }
    });

    QToolButton* bgColorBtn = new QToolButton(this);
    bgColorBtn->setText("ab");
    bgColorBtn->setFont(QFont("Arial", 10));
    bgColorBtn->setStyleSheet(
        "QToolButton { background-color: #FFFF00; border-bottom: 3px solid #FFFF00; padding: 2px 4px; }");
    bgColorBtn->setToolTip("Fill Color");
    bgColorBtn->setFixedSize(26, 26);
    addWidget(bgColorBtn);
    connect(bgColorBtn, &QToolButton::clicked, this, [this, bgColorBtn]() {
        QColor color = QColorDialog::getColor(m_lastBgColor, this, "Fill Color");
        if (color.isValid()) {
            m_lastBgColor = color;
            bgColorBtn->setStyleSheet(
                QString("QToolButton { background-color: %1; border-bottom: 3px solid %1; padding: 2px 4px; }")
                    .arg(color.name()));
            emit backgroundColorChanged(color);
        }
    });

    addSeparator();

    // ===== Horizontal Alignment =====
    m_alignLeftBtn = new QToolButton(this);
    m_alignLeftBtn->setIcon(createHAlignIcon("left"));
    m_alignLeftBtn->setToolTip("Align Left");
    m_alignLeftBtn->setCheckable(true);
    m_alignLeftBtn->setFixedSize(26, 26);
    addWidget(m_alignLeftBtn);

    m_alignCenterBtn = new QToolButton(this);
    m_alignCenterBtn->setIcon(createHAlignIcon("center"));
    m_alignCenterBtn->setToolTip("Center");
    m_alignCenterBtn->setCheckable(true);
    m_alignCenterBtn->setFixedSize(26, 26);
    addWidget(m_alignCenterBtn);

    m_alignRightBtn = new QToolButton(this);
    m_alignRightBtn->setIcon(createHAlignIcon("right"));
    m_alignRightBtn->setToolTip("Align Right");
    m_alignRightBtn->setCheckable(true);
    m_alignRightBtn->setFixedSize(26, 26);
    addWidget(m_alignRightBtn);

    auto uncheckHAligns = [this](QToolButton* except) {
        if (m_alignLeftBtn != except) m_alignLeftBtn->setChecked(false);
        if (m_alignCenterBtn != except) m_alignCenterBtn->setChecked(false);
        if (m_alignRightBtn != except) m_alignRightBtn->setChecked(false);
    };

    connect(m_alignLeftBtn, &QToolButton::clicked, this, [this, uncheckHAligns]() {
        uncheckHAligns(m_alignLeftBtn);
        emit hAlignChanged(HorizontalAlignment::Left);
    });
    connect(m_alignCenterBtn, &QToolButton::clicked, this, [this, uncheckHAligns]() {
        uncheckHAligns(m_alignCenterBtn);
        emit hAlignChanged(HorizontalAlignment::Center);
    });
    connect(m_alignRightBtn, &QToolButton::clicked, this, [this, uncheckHAligns]() {
        uncheckHAligns(m_alignRightBtn);
        emit hAlignChanged(HorizontalAlignment::Right);
    });

    addSeparator();

    // ===== Vertical Alignment =====
    m_vAlignTopBtn = new QToolButton(this);
    m_vAlignTopBtn->setIcon(createVAlignIcon("top"));
    m_vAlignTopBtn->setToolTip("Top Align");
    m_vAlignTopBtn->setCheckable(true);
    m_vAlignTopBtn->setFixedSize(26, 26);
    addWidget(m_vAlignTopBtn);

    m_vAlignMiddleBtn = new QToolButton(this);
    m_vAlignMiddleBtn->setIcon(createVAlignIcon("middle"));
    m_vAlignMiddleBtn->setToolTip("Middle Align");
    m_vAlignMiddleBtn->setCheckable(true);
    m_vAlignMiddleBtn->setChecked(true);
    m_vAlignMiddleBtn->setFixedSize(26, 26);
    addWidget(m_vAlignMiddleBtn);

    m_vAlignBottomBtn = new QToolButton(this);
    m_vAlignBottomBtn->setIcon(createVAlignIcon("bottom"));
    m_vAlignBottomBtn->setToolTip("Bottom Align");
    m_vAlignBottomBtn->setCheckable(true);
    m_vAlignBottomBtn->setFixedSize(26, 26);
    addWidget(m_vAlignBottomBtn);

    auto uncheckVAligns = [this](QToolButton* except) {
        if (m_vAlignTopBtn != except) m_vAlignTopBtn->setChecked(false);
        if (m_vAlignMiddleBtn != except) m_vAlignMiddleBtn->setChecked(false);
        if (m_vAlignBottomBtn != except) m_vAlignBottomBtn->setChecked(false);
    };

    connect(m_vAlignTopBtn, &QToolButton::clicked, this, [this, uncheckVAligns]() {
        uncheckVAligns(m_vAlignTopBtn);
        emit vAlignChanged(VerticalAlignment::Top);
    });
    connect(m_vAlignMiddleBtn, &QToolButton::clicked, this, [this, uncheckVAligns]() {
        uncheckVAligns(m_vAlignMiddleBtn);
        emit vAlignChanged(VerticalAlignment::Middle);
    });
    connect(m_vAlignBottomBtn, &QToolButton::clicked, this, [this, uncheckVAligns]() {
        uncheckVAligns(m_vAlignBottomBtn);
        emit vAlignChanged(VerticalAlignment::Bottom);
    });

    addSeparator();

    // ===== Number formatting =====
    // Currency
    QToolButton* currencyBtn = new QToolButton(this);
    currencyBtn->setText("$");
    currencyBtn->setFont(QFont("Arial", 12, QFont::Bold));
    currencyBtn->setToolTip("Currency Format");
    currencyBtn->setFixedSize(26, 26);
    addWidget(currencyBtn);
    connect(currencyBtn, &QToolButton::clicked, this, [this]() { emit numberFormatChanged("Currency"); });

    // Percent
    QToolButton* percentBtn = new QToolButton(this);
    percentBtn->setText("%");
    percentBtn->setFont(QFont("Arial", 12, QFont::Bold));
    percentBtn->setToolTip("Percentage Format");
    percentBtn->setFixedSize(26, 26);
    addWidget(percentBtn);
    connect(percentBtn, &QToolButton::clicked, this, [this]() { emit numberFormatChanged("Percentage"); });

    // Thousand separator
    QToolButton* thousandBtn = new QToolButton(this);
    thousandBtn->setText(",");
    thousandBtn->setFont(QFont("Arial", 14, QFont::Bold));
    thousandBtn->setToolTip("Thousand Separator");
    thousandBtn->setCheckable(true);
    thousandBtn->setFixedSize(26, 26);
    addWidget(thousandBtn);
    connect(thousandBtn, &QToolButton::clicked, this, [this]() { emit thousandSeparatorToggled(); });

    QToolButton* formatCellsBtn = new QToolButton(this);
    formatCellsBtn->setText("Format\nCells");
    formatCellsBtn->setToolTip("Format Cells (Ctrl+1)");
    formatCellsBtn->setStyleSheet("QToolButton { font-size: 9px; padding: 1px 4px; }");
    addWidget(formatCellsBtn);
    connect(formatCellsBtn, &QToolButton::clicked, this, &Toolbar::formatCellsRequested);

    addSeparator();

    // ===== Data =====
    QToolButton* sortAscBtn = new QToolButton(this);
    sortAscBtn->setIcon(createSortIcon(true));
    sortAscBtn->setIconSize(QSize(16, 16));
    sortAscBtn->setToolTip("Sort A to Z");
    sortAscBtn->setFixedSize(30, 26);
    addWidget(sortAscBtn);
    connect(sortAscBtn, &QToolButton::clicked, this, &Toolbar::sortAscending);

    QToolButton* sortDescBtn = new QToolButton(this);
    sortDescBtn->setIcon(createSortIcon(false));
    sortDescBtn->setIconSize(QSize(16, 16));
    sortDescBtn->setToolTip("Sort Z to A");
    sortDescBtn->setFixedSize(30, 26);
    addWidget(sortDescBtn);
    connect(sortDescBtn, &QToolButton::clicked, this, &Toolbar::sortDescending);

    QToolButton* filterBtn = new QToolButton(this);
    filterBtn->setText("\u2BD7");  // ⯗ funnel/filter
    filterBtn->setFont(QFont("Apple Symbols", 13));
    filterBtn->setToolTip("Toggle Auto Filter");
    filterBtn->setCheckable(true);
    filterBtn->setFixedSize(26, 26);
    filterBtn->setStyleSheet("QToolButton { font-size: 14px; color: #217346; }");
    addWidget(filterBtn);
    connect(filterBtn, &QToolButton::clicked, this, &Toolbar::filterToggled);

    addSeparator();

    // ===== Table =====
    QToolButton* tableBtn = new QToolButton(this);
    tableBtn->setText("Table");
    tableBtn->setToolTip("Format as Table");
    tableBtn->setPopupMode(QToolButton::InstantPopup);
    tableBtn->setStyleSheet("QToolButton { font-size: 10px; font-weight: bold; color: #217346; padding: 2px 6px; }");

    QMenu* tableMenu = new QMenu(tableBtn);
    tableMenu->setStyleSheet(
        "QMenu { background: #FFFFFF; border: 1px solid #D0D0D0; padding: 4px; }"
        "QMenu::item { padding: 5px 12px 5px 8px; }"
        "QMenu::item:selected { background-color: #E8F0FE; border-radius: 3px; }"
    );

    auto themes = getBuiltinTableThemes();
    for (int i = 0; i < static_cast<int>(themes.size()); ++i) {
        const auto& theme = themes[i];
        // Create a rich color swatch: header row + 3 banded rows
        QPixmap swatch(48, 28);
        swatch.fill(Qt::transparent);
        QPainter sp(&swatch);
        sp.setRenderHint(QPainter::Antialiasing, true);

        // Rounded rect clip
        QPainterPath clip;
        clip.addRoundedRect(0, 0, 48, 28, 3, 3);
        sp.setClipPath(clip);

        // Header row
        sp.fillRect(0, 0, 48, 8, theme.headerBg);
        // Banded rows
        sp.fillRect(0, 8, 48, 7, theme.bandedRow1);
        sp.fillRect(0, 15, 48, 6, theme.bandedRow2);
        sp.fillRect(0, 21, 48, 7, theme.bandedRow1);

        // Border
        sp.setClipping(false);
        sp.setPen(QPen(QColor("#D0D0D0"), 0.5));
        sp.drawRoundedRect(QRectF(0.25, 0.25, 47.5, 27.5), 3, 3);
        sp.end();

        QAction* action = tableMenu->addAction(QIcon(swatch), theme.name);
        connect(action, &QAction::triggered, this, [this, i]() {
            emit tableStyleSelected(i);
        });
    }

    tableBtn->setMenu(tableMenu);
    addWidget(tableBtn);

    addSeparator();

    // ===== Conditional Formatting — icon with color bars =====
    QToolButton* condFmtBtn = new QToolButton(this);
    {
        QPixmap pix(20, 20);
        pix.fill(Qt::transparent);
        QPainter p(&pix);
        p.setRenderHint(QPainter::Antialiasing, true);
        // Three colored bars representing conditional format rules
        p.setPen(Qt::NoPen);
        p.setBrush(QColor("#4CAF50")); // Green bar
        p.drawRoundedRect(2, 2, 16, 4, 1, 1);
        p.setBrush(QColor("#FFC107")); // Yellow bar
        p.drawRoundedRect(2, 8, 16, 4, 1, 1);
        p.setBrush(QColor("#F44336")); // Red bar
        p.drawRoundedRect(2, 14, 16, 4, 1, 1);
        p.end();
        condFmtBtn->setIcon(QIcon(pix));
        condFmtBtn->setIconSize(QSize(20, 20));
    }
    condFmtBtn->setToolTip("Conditional Formatting");
    condFmtBtn->setFixedSize(30, 26);
    addWidget(condFmtBtn);
    connect(condFmtBtn, &QToolButton::clicked, this, &Toolbar::conditionalFormatRequested);

    // ===== Data Validation — icon with checkmark in circle =====
    QToolButton* validationBtn = new QToolButton(this);
    {
        QPixmap pix(20, 20);
        pix.fill(Qt::transparent);
        QPainter p(&pix);
        p.setRenderHint(QPainter::Antialiasing, true);
        // Circle outline
        p.setPen(QPen(QColor("#217346"), 1.5));
        p.setBrush(Qt::NoBrush);
        p.drawEllipse(2, 2, 16, 16);
        // Checkmark
        p.setPen(QPen(QColor("#217346"), 2.0, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin));
        p.drawLine(6, 10, 9, 14);
        p.drawLine(9, 14, 15, 5);
        p.end();
        validationBtn->setIcon(QIcon(pix));
        validationBtn->setIconSize(QSize(20, 20));
    }
    validationBtn->setToolTip("Data Validation");
    validationBtn->setFixedSize(30, 26);
    addWidget(validationBtn);
    connect(validationBtn, &QToolButton::clicked, this, &Toolbar::dataValidationRequested);
}
