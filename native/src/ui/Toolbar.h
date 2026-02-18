#ifndef TOOLBAR_H
#define TOOLBAR_H

#include <QToolBar>
#include <QColor>
#include "../core/Cell.h"

class QFontComboBox;
class QSpinBox;
class QToolButton;

class Toolbar : public QToolBar {
    Q_OBJECT

public:
    explicit Toolbar(QWidget* parent = nullptr);

signals:
    // File
    void newDocument();
    void saveDocument();
    // Edit
    void undo();
    void redo();
    // Format painter
    void formatPainterToggled();
    // Font formatting
    void bold();
    void italic();
    void underline();
    void strikethrough();
    void fontFamilyChanged(const QString& family);
    void fontSizeChanged(int size);
    // Colors
    void foregroundColorChanged(const QColor& color);
    void backgroundColorChanged(const QColor& color);
    // Alignment
    void hAlignChanged(HorizontalAlignment align);
    void vAlignChanged(VerticalAlignment align);
    // Number formatting
    void thousandSeparatorToggled();
    void numberFormatChanged(const QString& format);
    void formatCellsRequested();
    // Data
    void sortAscending();
    void sortDescending();
    void filterToggled();
    // Tables
    void tableStyleSelected(int themeIndex);
    // Conditional formatting & validation
    void conditionalFormatRequested();
    void dataValidationRequested();

private:
    void createActions();

    QFontComboBox* m_fontCombo = nullptr;
    QSpinBox* m_fontSizeSpinBox = nullptr;
    QColor m_lastFgColor = QColor("#000000");
    QColor m_lastBgColor = QColor("#FFFF00");

    // Alignment button refs for checked state management
    QToolButton* m_alignLeftBtn = nullptr;
    QToolButton* m_alignCenterBtn = nullptr;
    QToolButton* m_alignRightBtn = nullptr;
    QToolButton* m_vAlignTopBtn = nullptr;
    QToolButton* m_vAlignMiddleBtn = nullptr;
    QToolButton* m_vAlignBottomBtn = nullptr;
};

#endif // TOOLBAR_H
