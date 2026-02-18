#include "CellDelegate.h"
#include <QLineEdit>
#include <QPainter>
#include <QStyleOptionViewItem>
#include <QApplication>
#include <QCompleter>
#include <QStringListModel>
#include <QAbstractItemView>
#include <QListView>

// All supported formula function names
static const QStringList s_formulaNames = {
    "SUM", "AVERAGE", "COUNT", "COUNTA", "MIN", "MAX",
    "IF", "IFERROR", "AND", "OR", "NOT",
    "CONCAT", "LEN", "UPPER", "LOWER", "TRIM",
    "LEFT", "RIGHT", "MID", "FIND", "SUBSTITUTE", "TEXT",
    "ROUND", "ABS", "SQRT", "POWER", "MOD", "INT", "CEILING", "FLOOR",
    "COUNTIF", "SUMIF",
    "NOW", "TODAY", "YEAR", "MONTH", "DAY",
};

CellDelegate::CellDelegate(QObject* parent)
    : QStyledItemDelegate(parent) {
}

QWidget* CellDelegate::createEditor(QWidget* parent, const QStyleOptionViewItem& option,
                                   const QModelIndex& index) const {
    QLineEdit* editor = new QLineEdit(parent);
    editor->setFrame(false);
    editor->setStyleSheet(
        "QLineEdit { background: white; padding: 1px 2px; selection-background-color: #0078D4; }");

    // Formula autocomplete
    auto* completer = new QCompleter(s_formulaNames, editor);
    completer->setWidget(editor);
    completer->setCaseSensitivity(Qt::CaseInsensitive);
    completer->setCompletionMode(QCompleter::PopupCompletion);
    completer->setFilterMode(Qt::MatchContains);
    completer->popup()->setStyleSheet(
        "QListView { background: white; border: 1px solid #C8C8C8; font-size: 12px; }"
        "QListView::item { padding: 3px 8px; }"
        "QListView::item:selected { background: #E8F0FE; color: #333; }"
    );

    // When a formula is selected from the completer, append parentheses
    connect(completer, QOverload<const QString&>::of(&QCompleter::activated),
            editor, [editor](const QString& funcName) {
        // The completer replaces text after '=', so we need to set it with parens
        QString text = editor->text();
        // Find where the function name starts (after '=' and any existing content)
        int eqPos = text.lastIndexOf('=');
        if (eqPos >= 0) {
            editor->setText(text.left(eqPos + 1) + funcName + "()");
            editor->setCursorPosition(editor->text().length() - 1); // Position between parens
        }
    });

    // Show completer only when typing after '='
    connect(editor, &QLineEdit::textChanged, this, [this, editor, completer](const QString& text) {
        emit formulaEditModeChanged(text.startsWith("="));

        if (text.startsWith("=") && text.length() > 1) {
            // Extract the token being typed (after last operator/paren/comma)
            QString afterEq = text.mid(1);
            int lastDelim = -1;
            for (int i = afterEq.length() - 1; i >= 0; --i) {
                QChar ch = afterEq[i];
                if (ch == '(' || ch == ')' || ch == ',' || ch == '+' || ch == '-' ||
                    ch == '*' || ch == '/' || ch == ':' || ch == ' ') {
                    lastDelim = i;
                    break;
                }
            }
            QString prefix = afterEq.mid(lastDelim + 1);
            if (!prefix.isEmpty() && prefix[0].isLetter()) {
                completer->setCompletionPrefix(prefix);
                if (completer->completionCount() > 0) {
                    completer->complete();
                } else {
                    completer->popup()->hide();
                }
            } else {
                completer->popup()->hide();
            }
        } else {
            completer->popup()->hide();
        }
    });

    return editor;
}

void CellDelegate::setEditorData(QWidget* editor, const QModelIndex& index) const {
    QLineEdit* lineEdit = qobject_cast<QLineEdit*>(editor);
    if (lineEdit) {
        lineEdit->setText(index.data(Qt::EditRole).toString());
    }
}

void CellDelegate::setModelData(QWidget* editor, QAbstractItemModel* model,
                               const QModelIndex& index) const {
    QLineEdit* lineEdit = qobject_cast<QLineEdit*>(editor);
    if (lineEdit) {
        model->setData(index, lineEdit->text(), Qt::EditRole);
    }
}

void CellDelegate::updateEditorGeometry(QWidget* editor, const QStyleOptionViewItem& option,
                                       const QModelIndex& index) const {
    editor->setGeometry(option.rect);
}

void CellDelegate::paint(QPainter* painter, const QStyleOptionViewItem& option,
                        const QModelIndex& index) const {
    painter->save();
    painter->setRenderHint(QPainter::Antialiasing, false);

    QRect rect = option.rect;
    bool isSelected = option.state & QStyle::State_Selected;
    bool hasFocus = option.state & QStyle::State_HasFocus;

    // --- Background ---
    QColor bgColor(Qt::white);
    QVariant bgData = index.data(Qt::BackgroundRole);
    if (bgData.isValid()) {
        QColor cellBg = bgData.value<QColor>();
        if (cellBg.isValid() && cellBg != QColor("#FFFFFF") && cellBg != QColor(Qt::white)) {
            bgColor = cellBg;
        }
    }

    if (isSelected && !hasFocus) {
        // Multi-select: light blue tint over cell background
        painter->fillRect(rect, bgColor);
        painter->fillRect(rect, QColor(198, 217, 240, 60));
    } else {
        painter->fillRect(rect, bgColor);
    }

    // --- Text ---
    QString text = index.data(Qt::DisplayRole).toString();
    if (!text.isEmpty()) {
        QFont font = option.font;
        QVariant fontData = index.data(Qt::FontRole);
        if (fontData.isValid()) {
            font = fontData.value<QFont>();
        }
        painter->setFont(font);

        QColor fgColor(Qt::black);
        QVariant fgData = index.data(Qt::ForegroundRole);
        if (fgData.isValid()) {
            QColor c = fgData.value<QColor>();
            if (c.isValid()) fgColor = c;
        }
        painter->setPen(fgColor);

        int alignment = Qt::AlignVCenter | Qt::AlignLeft;
        QVariant alignData = index.data(Qt::TextAlignmentRole);
        if (alignData.isValid()) {
            alignment = alignData.toInt();
        }

        QRect textRect = rect.adjusted(4, 1, -4, -1);
        painter->drawText(textRect, alignment, text);
    }

    // --- Gridlines: single thin line on right and bottom edges ---
    painter->setPen(QPen(QColor(218, 220, 224), 1, Qt::SolidLine));
    painter->drawLine(rect.right(), rect.top(), rect.right(), rect.bottom());
    painter->drawLine(rect.left(), rect.bottom(), rect.right(), rect.bottom());

    // --- Focus border: green rectangle for the current cell ---
    if (hasFocus) {
        QPen focusPen(QColor(16, 124, 16), 2, Qt::SolidLine);
        painter->setPen(focusPen);
        painter->drawRect(rect.adjusted(0, 0, -1, -1));
    }

    painter->restore();
}

QSize CellDelegate::sizeHint(const QStyleOptionViewItem& option,
                             const QModelIndex& index) const {
    return QStyledItemDelegate::sizeHint(option, index);
}
