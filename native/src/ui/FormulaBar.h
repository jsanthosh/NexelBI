#ifndef FORMULABAR_H
#define FORMULABAR_H

#include <QWidget>
#include <QLineEdit>
#include <QLabel>

class FormulaBar : public QWidget {
    Q_OBJECT

public:
    explicit FormulaBar(QWidget* parent = nullptr);

    void setCellAddress(const QString& address);
    void setCellContent(const QString& content);
    QString getContent() const;
    bool isFormulaEditing() const;
    void insertText(const QString& text);

signals:
    void contentChanged(const QString& content);
    void contentEdited(const QString& content);
    void formulaEditModeChanged(bool active);

private slots:
    void onTextChanged(const QString& text);
    void onTextEdited(const QString& text);

private:
    QLabel* m_cellAddressLabel;
    QLineEdit* m_formulaEdit;
};

#endif // FORMULABAR_H
