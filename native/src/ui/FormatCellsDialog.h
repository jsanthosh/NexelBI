#ifndef FORMATCELLSDIALOG_H
#define FORMATCELLSDIALOG_H

#include <QDialog>
#include "../core/Cell.h"

class QTabWidget;
class QListWidget;
class QSpinBox;
class QCheckBox;
class QComboBox;
class QLineEdit;
class QLabel;
class QFontComboBox;
class QColorDialog;
class QPushButton;

class FormatCellsDialog : public QDialog {
    Q_OBJECT

public:
    FormatCellsDialog(const CellStyle& style, QWidget* parent = nullptr);
    CellStyle getStyle() const;

private:
    void createNumberTab(QWidget* tab);
    void createFontTab(QWidget* tab);
    void createAlignmentTab(QWidget* tab);
    void createFillTab(QWidget* tab);
    void updatePreview();
    void loadStyle(const CellStyle& style);

    CellStyle m_style;

    // Number tab
    QListWidget* m_categoryList = nullptr;
    QSpinBox* m_decimalSpin = nullptr;
    QCheckBox* m_thousandCheck = nullptr;
    QComboBox* m_currencyCombo = nullptr;
    QComboBox* m_dateFormatCombo = nullptr;
    QLineEdit* m_customFormatEdit = nullptr;
    QLabel* m_previewLabel = nullptr;

    // Font tab
    QFontComboBox* m_fontFamilyCombo = nullptr;
    QSpinBox* m_fontSizeSpin = nullptr;
    QCheckBox* m_boldCheck = nullptr;
    QCheckBox* m_italicCheck = nullptr;
    QCheckBox* m_underlineCheck = nullptr;
    QCheckBox* m_strikethroughCheck = nullptr;
    QPushButton* m_fontColorBtn = nullptr;
    QColor m_fontColor;

    // Alignment tab
    QComboBox* m_hAlignCombo = nullptr;
    QComboBox* m_vAlignCombo = nullptr;

    // Fill tab
    QPushButton* m_fillColorBtn = nullptr;
    QColor m_fillColor;
};

#endif // FORMATCELLSDIALOG_H
