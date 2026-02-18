#include "FormulaBar.h"
#include <QHBoxLayout>
#include <QLabel>

FormulaBar::FormulaBar(QWidget* parent)
    : QWidget(parent) {
    
    QHBoxLayout* layout = new QHBoxLayout(this);
    layout->setContentsMargins(5, 5, 5, 5);

    // Cell address label
    m_cellAddressLabel = new QLabel("A1", this);
    m_cellAddressLabel->setMinimumWidth(50);
    m_cellAddressLabel->setStyleSheet("border: 1px solid #d0d0d0; padding: 2px;");
    layout->addWidget(m_cellAddressLabel);

    // Formula/Content input
    m_formulaEdit = new QLineEdit(this);
    m_formulaEdit->setPlaceholderText("Enter formula or value...");
    layout->addWidget(m_formulaEdit);

    // Setup connections
    connect(m_formulaEdit, &QLineEdit::textChanged, this, &FormulaBar::onTextChanged);
    connect(m_formulaEdit, &QLineEdit::textEdited, this, &FormulaBar::onTextEdited);

    setStyleSheet(
        "QWidget {"
        "   background-color: #ffffff;"
        "   border-bottom: 1px solid #e0e0e0;"
        "}"
        "QLineEdit {"
        "   border: 1px solid #d0d0d0;"
        "   padding: 3px;"
        "   border-radius: 3px;"
        "}"
    );
}

void FormulaBar::setCellAddress(const QString& address) {
    m_cellAddressLabel->setText(address);
}

void FormulaBar::setCellContent(const QString& content) {
    m_formulaEdit->blockSignals(true);
    m_formulaEdit->setText(content);
    m_formulaEdit->blockSignals(false);
}

QString FormulaBar::getContent() const {
    return m_formulaEdit->text();
}

void FormulaBar::onTextChanged(const QString& text) {
    emit contentChanged(text);
}

void FormulaBar::onTextEdited(const QString& text) {
    emit contentEdited(text);
    emit formulaEditModeChanged(m_formulaEdit->hasFocus() && text.startsWith("="));
}

bool FormulaBar::isFormulaEditing() const {
    return m_formulaEdit->hasFocus() && m_formulaEdit->text().startsWith("=");
}

void FormulaBar::insertText(const QString& text) {
    m_formulaEdit->insert(text);
}
