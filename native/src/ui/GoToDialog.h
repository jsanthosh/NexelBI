#ifndef GOTODIALOG_H
#define GOTODIALOG_H

#include <QDialog>
#include <QLineEdit>
#include "../core/CellRange.h"

class GoToDialog : public QDialog {
    Q_OBJECT

public:
    explicit GoToDialog(QWidget* parent = nullptr);

    CellAddress getAddress() const;

private:
    QLineEdit* m_cellRefEdit;
};

#endif // GOTODIALOG_H
