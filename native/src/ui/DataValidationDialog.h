#ifndef DATAVALIDATIONDIALOG_H
#define DATAVALIDATIONDIALOG_H

#include <QDialog>
#include <QComboBox>
#include <QLineEdit>
#include <QTextEdit>
#include <QCheckBox>
#include <QTabWidget>
#include <QLabel>
#include "../core/Spreadsheet.h"

class DataValidationDialog : public QDialog {
    Q_OBJECT

public:
    explicit DataValidationDialog(const CellRange& range, QWidget* parent = nullptr);

    Spreadsheet::DataValidationRule getRule() const;
    void setRule(const Spreadsheet::DataValidationRule& rule);

private slots:
    void onTypeChanged(int index);
    void onOperatorChanged(int index);

private:
    void updateFieldVisibility();

    CellRange m_range;

    // Settings tab
    QComboBox* m_typeCombo;
    QComboBox* m_operatorCombo;
    QLineEdit* m_value1Edit;
    QLineEdit* m_value2Edit;
    QLabel* m_value1Label;
    QLabel* m_value2Label;
    QTextEdit* m_listEdit;
    QLabel* m_listLabel;
    QLineEdit* m_formulaEdit;
    QLabel* m_formulaLabel;
    QCheckBox* m_ignoreBlank;

    // Input Message tab
    QCheckBox* m_showInputMsg;
    QLineEdit* m_inputTitle;
    QTextEdit* m_inputMessage;

    // Error Alert tab
    QCheckBox* m_showErrorAlert;
    QComboBox* m_errorStyle;
    QLineEdit* m_errorTitle;
    QTextEdit* m_errorMessage;
};

#endif // DATAVALIDATIONDIALOG_H
