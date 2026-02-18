#include "FormatCellsDialog.h"
#include "../core/NumberFormat.h"
#include <QTabWidget>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QGroupBox>
#include <QListWidget>
#include <QSpinBox>
#include <QCheckBox>
#include <QComboBox>
#include <QLineEdit>
#include <QLabel>
#include <QFontComboBox>
#include <QPushButton>
#include <QDialogButtonBox>
#include <QColorDialog>
#include <QStackedWidget>

FormatCellsDialog::FormatCellsDialog(const CellStyle& style, QWidget* parent)
    : QDialog(parent), m_style(style) {
    setWindowTitle("Format Cells");
    setMinimumSize(520, 420);

    QVBoxLayout* mainLayout = new QVBoxLayout(this);

    QTabWidget* tabs = new QTabWidget(this);

    QWidget* numberTab = new QWidget();
    createNumberTab(numberTab);
    tabs->addTab(numberTab, "Number");

    QWidget* fontTab = new QWidget();
    createFontTab(fontTab);
    tabs->addTab(fontTab, "Font");

    QWidget* alignTab = new QWidget();
    createAlignmentTab(alignTab);
    tabs->addTab(alignTab, "Alignment");

    QWidget* fillTab = new QWidget();
    createFillTab(fillTab);
    tabs->addTab(fillTab, "Fill");

    mainLayout->addWidget(tabs);

    QDialogButtonBox* buttons = new QDialogButtonBox(
        QDialogButtonBox::Ok | QDialogButtonBox::Cancel, this);
    connect(buttons, &QDialogButtonBox::accepted, this, &QDialog::accept);
    connect(buttons, &QDialogButtonBox::rejected, this, &QDialog::reject);
    mainLayout->addWidget(buttons);

    loadStyle(style);
}

void FormatCellsDialog::createNumberTab(QWidget* tab) {
    QHBoxLayout* layout = new QHBoxLayout(tab);

    // Category list
    m_categoryList = new QListWidget();
    m_categoryList->addItem("General");
    m_categoryList->addItem("Number");
    m_categoryList->addItem("Currency");
    m_categoryList->addItem("Accounting");
    m_categoryList->addItem("Percentage");
    m_categoryList->addItem("Date");
    m_categoryList->addItem("Time");
    m_categoryList->addItem("Text");
    m_categoryList->addItem("Custom");
    m_categoryList->setMaximumWidth(120);
    layout->addWidget(m_categoryList);

    // Options panel
    QVBoxLayout* optionsLayout = new QVBoxLayout();

    // Preview
    QGroupBox* previewBox = new QGroupBox("Preview");
    QVBoxLayout* previewLayout = new QVBoxLayout(previewBox);
    m_previewLabel = new QLabel("General");
    m_previewLabel->setStyleSheet("QLabel { padding: 8px; background: white; border: 1px solid #ccc; }");
    previewLayout->addWidget(m_previewLabel);
    optionsLayout->addWidget(previewBox);

    // Decimal places
    QHBoxLayout* decimalRow = new QHBoxLayout();
    decimalRow->addWidget(new QLabel("Decimal places:"));
    m_decimalSpin = new QSpinBox();
    m_decimalSpin->setRange(0, 10);
    m_decimalSpin->setValue(2);
    decimalRow->addWidget(m_decimalSpin);
    optionsLayout->addLayout(decimalRow);

    // Thousand separator
    m_thousandCheck = new QCheckBox("Use 1000 separator (,)");
    optionsLayout->addWidget(m_thousandCheck);

    // Currency
    QHBoxLayout* currencyRow = new QHBoxLayout();
    currencyRow->addWidget(new QLabel("Currency:"));
    m_currencyCombo = new QComboBox();
    for (const auto& c : NumberFormat::currencies()) {
        m_currencyCombo->addItem(c.label, c.code);
    }
    currencyRow->addWidget(m_currencyCombo);
    optionsLayout->addLayout(currencyRow);

    // Date format
    QHBoxLayout* dateRow = new QHBoxLayout();
    dateRow->addWidget(new QLabel("Date format:"));
    m_dateFormatCombo = new QComboBox();
    m_dateFormatCombo->addItem("MM/DD/YYYY", "mm/dd/yyyy");
    m_dateFormatCombo->addItem("DD/MM/YYYY", "dd/mm/yyyy");
    m_dateFormatCombo->addItem("YYYY-MM-DD", "yyyy-mm-dd");
    m_dateFormatCombo->addItem("MMM D, YYYY", "mmm d, yyyy");
    m_dateFormatCombo->addItem("MMMM D, YYYY", "mmmm d, yyyy");
    m_dateFormatCombo->addItem("D-MMM-YY", "d-mmm-yy");
    m_dateFormatCombo->addItem("MM/DD", "mm/dd");
    dateRow->addWidget(m_dateFormatCombo);
    optionsLayout->addLayout(dateRow);

    // Custom format
    QHBoxLayout* customRow = new QHBoxLayout();
    customRow->addWidget(new QLabel("Custom:"));
    m_customFormatEdit = new QLineEdit();
    m_customFormatEdit->setPlaceholderText("#,##0.00");
    customRow->addWidget(m_customFormatEdit);
    optionsLayout->addLayout(customRow);

    optionsLayout->addStretch();
    layout->addLayout(optionsLayout);

    // Connect category change
    connect(m_categoryList, &QListWidget::currentRowChanged, this, [this](int row) {
        QStringList types = {"General", "Number", "Currency", "Accounting",
                             "Percentage", "Date", "Time", "Text", "Custom"};
        if (row >= 0 && row < types.size()) {
            m_style.numberFormat = types[row];
            updatePreview();
        }
    });

    connect(m_decimalSpin, &QSpinBox::valueChanged, this, [this](int val) {
        m_style.decimalPlaces = val;
        updatePreview();
    });

    connect(m_thousandCheck, &QCheckBox::toggled, this, [this](bool checked) {
        m_style.useThousandsSeparator = checked;
        updatePreview();
    });

    connect(m_currencyCombo, &QComboBox::currentIndexChanged, this, [this](int idx) {
        m_style.currencyCode = m_currencyCombo->itemData(idx).toString();
        updatePreview();
    });

    connect(m_dateFormatCombo, &QComboBox::currentIndexChanged, this, [this](int idx) {
        m_style.dateFormatId = m_dateFormatCombo->itemData(idx).toString();
        updatePreview();
    });

    connect(m_customFormatEdit, &QLineEdit::textChanged, this, [this](const QString&) {
        updatePreview();
    });
}

void FormatCellsDialog::createFontTab(QWidget* tab) {
    QGridLayout* layout = new QGridLayout(tab);

    layout->addWidget(new QLabel("Font:"), 0, 0);
    m_fontFamilyCombo = new QFontComboBox();
    layout->addWidget(m_fontFamilyCombo, 0, 1, 1, 2);

    layout->addWidget(new QLabel("Size:"), 1, 0);
    m_fontSizeSpin = new QSpinBox();
    m_fontSizeSpin->setRange(6, 72);
    layout->addWidget(m_fontSizeSpin, 1, 1);

    QGroupBox* styleGroup = new QGroupBox("Style");
    QVBoxLayout* styleLayout = new QVBoxLayout(styleGroup);
    m_boldCheck = new QCheckBox("Bold");
    m_italicCheck = new QCheckBox("Italic");
    m_underlineCheck = new QCheckBox("Underline");
    m_strikethroughCheck = new QCheckBox("Strikethrough");
    styleLayout->addWidget(m_boldCheck);
    styleLayout->addWidget(m_italicCheck);
    styleLayout->addWidget(m_underlineCheck);
    styleLayout->addWidget(m_strikethroughCheck);
    layout->addWidget(styleGroup, 2, 0, 1, 3);

    QHBoxLayout* colorRow = new QHBoxLayout();
    colorRow->addWidget(new QLabel("Color:"));
    m_fontColorBtn = new QPushButton();
    m_fontColorBtn->setFixedSize(60, 24);
    colorRow->addWidget(m_fontColorBtn);
    colorRow->addStretch();
    layout->addLayout(colorRow, 3, 0, 1, 3);

    connect(m_fontColorBtn, &QPushButton::clicked, this, [this]() {
        QColor c = QColorDialog::getColor(m_fontColor, this, "Font Color");
        if (c.isValid()) {
            m_fontColor = c;
            m_fontColorBtn->setStyleSheet(
                QString("background-color: %1;").arg(c.name()));
        }
    });

    layout->setRowStretch(4, 1);
}

void FormatCellsDialog::createAlignmentTab(QWidget* tab) {
    QGridLayout* layout = new QGridLayout(tab);

    layout->addWidget(new QLabel("Horizontal:"), 0, 0);
    m_hAlignCombo = new QComboBox();
    m_hAlignCombo->addItem("General", static_cast<int>(HorizontalAlignment::General));
    m_hAlignCombo->addItem("Left", static_cast<int>(HorizontalAlignment::Left));
    m_hAlignCombo->addItem("Center", static_cast<int>(HorizontalAlignment::Center));
    m_hAlignCombo->addItem("Right", static_cast<int>(HorizontalAlignment::Right));
    layout->addWidget(m_hAlignCombo, 0, 1);

    layout->addWidget(new QLabel("Vertical:"), 1, 0);
    m_vAlignCombo = new QComboBox();
    m_vAlignCombo->addItem("Top", static_cast<int>(VerticalAlignment::Top));
    m_vAlignCombo->addItem("Middle", static_cast<int>(VerticalAlignment::Middle));
    m_vAlignCombo->addItem("Bottom", static_cast<int>(VerticalAlignment::Bottom));
    layout->addWidget(m_vAlignCombo, 1, 1);

    layout->setRowStretch(2, 1);
}

void FormatCellsDialog::createFillTab(QWidget* tab) {
    QVBoxLayout* layout = new QVBoxLayout(tab);

    QHBoxLayout* colorRow = new QHBoxLayout();
    colorRow->addWidget(new QLabel("Background color:"));
    m_fillColorBtn = new QPushButton();
    m_fillColorBtn->setFixedSize(60, 24);
    colorRow->addWidget(m_fillColorBtn);
    colorRow->addStretch();
    layout->addLayout(colorRow);

    connect(m_fillColorBtn, &QPushButton::clicked, this, [this]() {
        QColor c = QColorDialog::getColor(m_fillColor, this, "Fill Color");
        if (c.isValid()) {
            m_fillColor = c;
            m_fillColorBtn->setStyleSheet(
                QString("background-color: %1;").arg(c.name()));
        }
    });

    layout->addStretch();
}

void FormatCellsDialog::loadStyle(const CellStyle& style) {
    // Number tab
    QStringList types = {"General", "Number", "Currency", "Accounting",
                         "Percentage", "Date", "Time", "Text", "Custom"};
    int idx = types.indexOf(style.numberFormat);
    if (idx >= 0) m_categoryList->setCurrentRow(idx);
    m_decimalSpin->setValue(style.decimalPlaces);
    m_thousandCheck->setChecked(style.useThousandsSeparator);
    for (int i = 0; i < m_currencyCombo->count(); ++i) {
        if (m_currencyCombo->itemData(i).toString() == style.currencyCode) {
            m_currencyCombo->setCurrentIndex(i);
            break;
        }
    }
    for (int i = 0; i < m_dateFormatCombo->count(); ++i) {
        if (m_dateFormatCombo->itemData(i).toString() == style.dateFormatId) {
            m_dateFormatCombo->setCurrentIndex(i);
            break;
        }
    }

    // Font tab
    m_fontFamilyCombo->setCurrentFont(QFont(style.fontName));
    m_fontSizeSpin->setValue(style.fontSize);
    m_boldCheck->setChecked(style.bold);
    m_italicCheck->setChecked(style.italic);
    m_underlineCheck->setChecked(style.underline);
    m_strikethroughCheck->setChecked(style.strikethrough);
    m_fontColor = QColor(style.foregroundColor);
    m_fontColorBtn->setStyleSheet(
        QString("background-color: %1;").arg(m_fontColor.name()));

    // Alignment
    m_hAlignCombo->setCurrentIndex(static_cast<int>(style.hAlign));
    m_vAlignCombo->setCurrentIndex(static_cast<int>(style.vAlign));

    // Fill
    m_fillColor = QColor(style.backgroundColor);
    m_fillColorBtn->setStyleSheet(
        QString("background-color: %1;").arg(m_fillColor.name()));

    updatePreview();
}

CellStyle FormatCellsDialog::getStyle() const {
    CellStyle style = m_style;

    // Number
    style.decimalPlaces = m_decimalSpin->value();
    style.useThousandsSeparator = m_thousandCheck->isChecked();
    style.currencyCode = m_currencyCombo->currentData().toString();
    style.dateFormatId = m_dateFormatCombo->currentData().toString();

    // Font
    style.fontName = m_fontFamilyCombo->currentFont().family();
    style.fontSize = m_fontSizeSpin->value();
    style.bold = m_boldCheck->isChecked();
    style.italic = m_italicCheck->isChecked();
    style.underline = m_underlineCheck->isChecked();
    style.strikethrough = m_strikethroughCheck->isChecked();
    style.foregroundColor = m_fontColor.name();

    // Alignment
    style.hAlign = static_cast<HorizontalAlignment>(m_hAlignCombo->currentData().toInt());
    style.vAlign = static_cast<VerticalAlignment>(m_vAlignCombo->currentData().toInt());

    // Fill
    style.backgroundColor = m_fillColor.name();

    return style;
}

void FormatCellsDialog::updatePreview() {
    NumberFormatOptions opts;
    opts.type = NumberFormat::typeFromString(m_style.numberFormat);
    opts.decimalPlaces = m_decimalSpin->value();
    opts.useThousandsSeparator = m_thousandCheck->isChecked();
    opts.currencyCode = m_currencyCombo->currentData().toString();
    opts.dateFormatId = m_dateFormatCombo->currentData().toString();
    opts.customFormat = m_customFormatEdit->text();

    QString sample = "1234.56";
    if (opts.type == NumberFormatType::Percentage) sample = "0.1234";
    if (opts.type == NumberFormatType::Date) sample = "2026-02-17";

    QString formatted = NumberFormat::format(sample, opts);
    m_previewLabel->setText(formatted);
}
