#include "ConditionalFormatDialog.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGroupBox>
#include <QFormLayout>
#include <QDialogButtonBox>
#include <QMessageBox>

ConditionalFormatDialog::ConditionalFormatDialog(const CellRange& range,
                                                   ConditionalFormatting& formatting,
                                                   QWidget* parent)
    : QDialog(parent), m_range(range), m_formatting(formatting) {
    setWindowTitle("Conditional Formatting");
    setMinimumSize(520, 480);

    QVBoxLayout* mainLayout = new QVBoxLayout(this);

    // Range label
    QLabel* rangeLabel = new QLabel(QString("Applies to: %1").arg(range.toString()), this);
    rangeLabel->setStyleSheet("font-weight: bold; padding: 4px;");
    mainLayout->addWidget(rangeLabel);

    // Rules list
    QGroupBox* rulesGroup = new QGroupBox("Rules", this);
    QVBoxLayout* rulesLayout = new QVBoxLayout(rulesGroup);

    m_ruleList = new QListWidget(this);
    m_ruleList->setMaximumHeight(120);
    rulesLayout->addWidget(m_ruleList);

    QHBoxLayout* ruleButtons = new QHBoxLayout();
    QPushButton* addBtn = new QPushButton("Add Rule", this);
    QPushButton* deleteBtn = new QPushButton("Delete Rule", this);
    ruleButtons->addWidget(addBtn);
    ruleButtons->addWidget(deleteBtn);
    ruleButtons->addStretch();
    rulesLayout->addLayout(ruleButtons);

    mainLayout->addWidget(rulesGroup);

    // Condition setup
    QGroupBox* condGroup = new QGroupBox("Condition", this);
    QFormLayout* condLayout = new QFormLayout(condGroup);

    m_conditionType = new QComboBox(this);
    m_conditionType->addItem("Cell Value Equal To", static_cast<int>(ConditionType::Equal));
    m_conditionType->addItem("Cell Value Not Equal To", static_cast<int>(ConditionType::NotEqual));
    m_conditionType->addItem("Cell Value Greater Than", static_cast<int>(ConditionType::GreaterThan));
    m_conditionType->addItem("Cell Value Less Than", static_cast<int>(ConditionType::LessThan));
    m_conditionType->addItem("Cell Value Greater Than or Equal", static_cast<int>(ConditionType::GreaterThanOrEqual));
    m_conditionType->addItem("Cell Value Less Than or Equal", static_cast<int>(ConditionType::LessThanOrEqual));
    m_conditionType->addItem("Cell Value Between", static_cast<int>(ConditionType::Between));
    m_conditionType->addItem("Cell Contains", static_cast<int>(ConditionType::CellContains));
    m_conditionType->addItem("Use a Formula", static_cast<int>(ConditionType::Formula));
    condLayout->addRow("Format cells if:", m_conditionType);

    m_value1Label = new QLabel("Value:", this);
    m_value1Edit = new QLineEdit(this);
    condLayout->addRow(m_value1Label, m_value1Edit);

    m_value2Label = new QLabel("And:", this);
    m_value2Edit = new QLineEdit(this);
    condLayout->addRow(m_value2Label, m_value2Edit);

    m_formulaLabel = new QLabel("Formula:", this);
    m_formulaEdit = new QLineEdit(this);
    m_formulaEdit->setPlaceholderText("e.g. =A1>100");
    condLayout->addRow(m_formulaLabel, m_formulaEdit);

    mainLayout->addWidget(condGroup);

    // Format style
    QGroupBox* styleGroup = new QGroupBox("Format Style", this);
    QHBoxLayout* styleLayout = new QHBoxLayout(styleGroup);

    m_boldCheck = new QCheckBox("Bold", this);
    m_italicCheck = new QCheckBox("Italic", this);
    m_underlineCheck = new QCheckBox("Underline", this);
    styleLayout->addWidget(m_boldCheck);
    styleLayout->addWidget(m_italicCheck);
    styleLayout->addWidget(m_underlineCheck);

    m_fgColorBtn = new QPushButton("Font Color", this);
    m_fgColorBtn->setStyleSheet("QPushButton { border-bottom: 3px solid #000000; }");
    m_fgColorBtn->setFixedWidth(90);
    styleLayout->addWidget(m_fgColorBtn);

    m_bgColorBtn = new QPushButton("Fill Color", this);
    m_bgColorBtn->setStyleSheet("QPushButton { background-color: #FFFFFF; border: 1px solid #CCC; }");
    m_bgColorBtn->setFixedWidth(90);
    styleLayout->addWidget(m_bgColorBtn);

    mainLayout->addWidget(styleGroup);

    // Dialog buttons
    QDialogButtonBox* buttons = new QDialogButtonBox(
        QDialogButtonBox::Ok | QDialogButtonBox::Cancel, this);
    mainLayout->addWidget(buttons);

    // Connections
    connect(addBtn, &QPushButton::clicked, this, &ConditionalFormatDialog::onAddRule);
    connect(deleteBtn, &QPushButton::clicked, this, &ConditionalFormatDialog::onDeleteRule);
    connect(m_ruleList, &QListWidget::currentRowChanged, this, &ConditionalFormatDialog::onRuleSelected);
    connect(m_conditionType, &QComboBox::currentIndexChanged, this, &ConditionalFormatDialog::onConditionTypeChanged);

    connect(m_fgColorBtn, &QPushButton::clicked, this, [this]() {
        QColor color = QColorDialog::getColor(m_selectedFgColor, this, "Font Color");
        if (color.isValid()) {
            m_selectedFgColor = color;
            m_fgColorBtn->setStyleSheet(
                QString("QPushButton { border-bottom: 3px solid %1; }").arg(color.name()));
        }
    });

    connect(m_bgColorBtn, &QPushButton::clicked, this, [this]() {
        QColor color = QColorDialog::getColor(m_selectedBgColor, this, "Fill Color");
        if (color.isValid()) {
            m_selectedBgColor = color;
            m_bgColorBtn->setStyleSheet(
                QString("QPushButton { background-color: %1; border: 1px solid #CCC; }").arg(color.name()));
        }
    });

    connect(buttons, &QDialogButtonBox::accepted, this, &ConditionalFormatDialog::onApply);
    connect(buttons, &QDialogButtonBox::rejected, this, &QDialog::reject);

    // Initialize
    populateRuleList();
    updateValueFieldsVisibility();

    setStyleSheet(
        "QGroupBox { font-weight: bold; border: 1px solid #D0D0D0; border-radius: 4px; "
        "margin-top: 8px; padding-top: 16px; }"
        "QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }"
    );
}

void ConditionalFormatDialog::populateRuleList() {
    m_ruleList->clear();
    const auto& rules = m_formatting.getAllRules();
    for (const auto& rule : rules) {
        if (rule->getRange().intersects(m_range)) {
            QString desc;
            switch (rule->getType()) {
                case ConditionType::Equal: desc = "Equal to"; break;
                case ConditionType::NotEqual: desc = "Not equal to"; break;
                case ConditionType::GreaterThan: desc = "Greater than"; break;
                case ConditionType::LessThan: desc = "Less than"; break;
                case ConditionType::GreaterThanOrEqual: desc = "Greater than or equal"; break;
                case ConditionType::LessThanOrEqual: desc = "Less than or equal"; break;
                case ConditionType::Between: desc = "Between"; break;
                case ConditionType::CellContains: desc = "Contains"; break;
                case ConditionType::Formula: desc = "Formula"; break;
            }
            const CellStyle& s = rule->getStyle();
            QString styleDesc;
            if (s.bold) styleDesc += "B ";
            if (s.italic) styleDesc += "I ";
            if (s.backgroundColor != "#FFFFFF") styleDesc += "Fill ";
            if (s.foregroundColor != "#000000") styleDesc += "Color ";
            m_ruleList->addItem(QString("%1 — %2 [%3]")
                .arg(rule->getRange().toString(), desc, styleDesc.trimmed()));
        }
    }
}

void ConditionalFormatDialog::onAddRule() {
    auto rule = buildRuleFromUI();
    if (rule) {
        m_formatting.addRule(rule);
        populateRuleList();
    }
}

void ConditionalFormatDialog::onDeleteRule() {
    int row = m_ruleList->currentRow();
    if (row < 0) return;

    // Find the actual rule index that matches this range
    const auto& rules = m_formatting.getAllRules();
    int matchIdx = 0;
    for (size_t i = 0; i < rules.size(); ++i) {
        if (rules[i]->getRange().intersects(m_range)) {
            if (matchIdx == row) {
                m_formatting.removeRule(i);
                populateRuleList();
                return;
            }
            matchIdx++;
        }
    }
}

void ConditionalFormatDialog::onRuleSelected(int row) {
    if (row < 0) return;

    const auto& rules = m_formatting.getAllRules();
    int matchIdx = 0;
    for (const auto& rule : rules) {
        if (rule->getRange().intersects(m_range)) {
            if (matchIdx == row) {
                // Populate UI from rule
                int typeIdx = m_conditionType->findData(static_cast<int>(rule->getType()));
                if (typeIdx >= 0) m_conditionType->setCurrentIndex(typeIdx);

                const CellStyle& s = rule->getStyle();
                m_boldCheck->setChecked(s.bold);
                m_italicCheck->setChecked(s.italic);
                m_underlineCheck->setChecked(s.underline);
                m_selectedBgColor = QColor(s.backgroundColor);
                m_selectedFgColor = QColor(s.foregroundColor);
                m_bgColorBtn->setStyleSheet(
                    QString("QPushButton { background-color: %1; border: 1px solid #CCC; }").arg(s.backgroundColor));
                m_fgColorBtn->setStyleSheet(
                    QString("QPushButton { border-bottom: 3px solid %1; }").arg(s.foregroundColor));
                return;
            }
            matchIdx++;
        }
    }
}

void ConditionalFormatDialog::onConditionTypeChanged(int) {
    updateValueFieldsVisibility();
}

void ConditionalFormatDialog::updateValueFieldsVisibility() {
    int typeData = m_conditionType->currentData().toInt();
    ConditionType type = static_cast<ConditionType>(typeData);

    bool showValue1 = (type != ConditionType::Formula);
    bool showValue2 = (type == ConditionType::Between);
    bool showFormula = (type == ConditionType::Formula);

    m_value1Label->setVisible(showValue1);
    m_value1Edit->setVisible(showValue1);
    m_value2Label->setVisible(showValue2);
    m_value2Edit->setVisible(showValue2);
    m_formulaLabel->setVisible(showFormula);
    m_formulaEdit->setVisible(showFormula);
}

std::shared_ptr<ConditionalFormat> ConditionalFormatDialog::buildRuleFromUI() {
    int typeData = m_conditionType->currentData().toInt();
    ConditionType type = static_cast<ConditionType>(typeData);

    auto rule = std::make_shared<ConditionalFormat>(m_range, type);

    if (type == ConditionType::Formula) {
        rule->setFormula(m_formulaEdit->text());
    } else {
        rule->setValue1(m_value1Edit->text());
        if (type == ConditionType::Between) {
            rule->setValue2(m_value2Edit->text());
        }
    }

    CellStyle style;
    style.bold = m_boldCheck->isChecked();
    style.italic = m_italicCheck->isChecked();
    style.underline = m_underlineCheck->isChecked();
    style.foregroundColor = m_selectedFgColor.name();
    style.backgroundColor = m_selectedBgColor.name();
    rule->setStyle(style);

    return rule;
}

void ConditionalFormatDialog::onApply() {
    // If no rules exist yet and the user fills in the form, add the rule
    if (m_ruleList->count() == 0) {
        auto rule = buildRuleFromUI();
        if (rule) {
            m_formatting.addRule(rule);
        }
    }
    accept();
}
