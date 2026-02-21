#ifndef CONDITIONALFORMATDIALOG_H
#define CONDITIONALFORMATDIALOG_H

#include <QDialog>
#include <QComboBox>
#include <QLineEdit>
#include <QListWidget>
#include <QPushButton>
#include <QLabel>
#include <QColorDialog>
#include <QCheckBox>
#include "../core/ConditionalFormatting.h"
#include "../core/CellRange.h"

class ConditionalFormatDialog : public QDialog {
    Q_OBJECT

public:
    explicit ConditionalFormatDialog(const CellRange& defaultRange,
                                     ConditionalFormatting& formatting,
                                     QWidget* parent = nullptr);

private slots:
    void onAddRule();
    void onUpdateRule();
    void onDeleteRule();
    void onRuleSelected(int row);
    void onConditionTypeChanged(int index);
    void onApply();

private:
    void populateRuleList();
    std::shared_ptr<ConditionalFormat> buildRuleFromUI();
    void updateValueFieldsVisibility();

    CellRange m_defaultRange;  // default range for new rules (from current selection)
    ConditionalFormatting& m_formatting;

    QListWidget* m_ruleList;
    QLineEdit* m_rangeEdit;    // editable range for new rules
    QComboBox* m_conditionType;
    QLineEdit* m_value1Edit;
    QLineEdit* m_value2Edit;
    QLabel* m_value1Label;
    QLabel* m_value2Label;
    QLineEdit* m_formulaEdit;
    QLabel* m_formulaLabel;

    // Action buttons
    QPushButton* m_updateBtn;

    // Style preview
    QPushButton* m_bgColorBtn;
    QPushButton* m_fgColorBtn;
    QCheckBox* m_boldCheck;
    QCheckBox* m_italicCheck;
    QCheckBox* m_underlineCheck;

    QColor m_selectedBgColor = QColor("#FFFFFF");
    QColor m_selectedFgColor = QColor("#000000");
};

#endif // CONDITIONALFORMATDIALOG_H
