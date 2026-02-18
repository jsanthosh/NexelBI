#ifndef FINDREPLACEDIALOG_H
#define FINDREPLACEDIALOG_H

#include <QDialog>
#include <QLineEdit>
#include <QCheckBox>
#include <QPushButton>
#include <QLabel>

class FindReplaceDialog : public QDialog {
    Q_OBJECT

public:
    explicit FindReplaceDialog(QWidget* parent = nullptr);

    QString findText() const;
    QString replaceText() const;
    bool matchCase() const;
    bool matchWholeCell() const;

signals:
    void findNext();
    void findPrevious();
    void replaceOne();
    void replaceAll();

private:
    QLineEdit* m_findEdit;
    QLineEdit* m_replaceEdit;
    QCheckBox* m_matchCaseCheck;
    QCheckBox* m_wholeCellCheck;
    QLabel* m_statusLabel;

public:
    void setStatus(const QString& text);
};

#endif // FINDREPLACEDIALOG_H
