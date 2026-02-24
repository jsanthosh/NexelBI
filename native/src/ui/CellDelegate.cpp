#include "CellDelegate.h"
#include "SpreadsheetView.h"
#include "../core/FormulaMetadata.h"
#include <QLineEdit>
#include <QPainter>
#include <QPainterPath>
#include <QStyleOptionViewItem>
#include <QApplication>
#include <QAbstractItemView>
#include <QListWidget>
#include <QKeyEvent>
#include <QTableView>
#include <QTimer>
#include <QLabel>
#include <QVBoxLayout>
#include <QFrame>
#include <QScrollBar>

// Extract the token currently being typed (after last delimiter)
static QString extractCurrentToken(const QString& text) {
    if (!text.startsWith("=") || text.length() <= 1) return {};
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
    return afterEq.mid(lastDelim + 1);
}

CellDelegate::CellDelegate(QObject* parent)
    : QStyledItemDelegate(parent) {
}

bool CellDelegate::eventFilter(QObject* object, QEvent* event) {
    // During formula edit mode, block FocusOut from closing the editor
    // (user is clicking on the grid to select cell references)
    if (m_formulaEditMode && event->type() == QEvent::FocusOut) {
        return true;  // consume the event, keep editor open
    }

    if (event->type() == QEvent::KeyPress) {
        QKeyEvent* keyEvent = static_cast<QKeyEvent*>(event);
        QLineEdit* editor = qobject_cast<QLineEdit*>(object);

        // Handle popup keyboard navigation
        if (editor) {
            auto* popup = qobject_cast<QListWidget*>(
                editor->property("_formulaPopup").value<QObject*>());
            if (popup && popup->isVisible()) {
                if (keyEvent->key() == Qt::Key_Down) {
                    int next = popup->currentRow() + 1;
                    if (next < popup->count()) popup->setCurrentRow(next);
                    return true;
                }
                if (keyEvent->key() == Qt::Key_Up) {
                    int prev = popup->currentRow() - 1;
                    if (prev >= 0) popup->setCurrentRow(prev);
                    return true;
                }
                if (keyEvent->key() == Qt::Key_Return || keyEvent->key() == Qt::Key_Enter ||
                    keyEvent->key() == Qt::Key_Tab) {
                    auto* item = popup->currentItem();
                    if (item) {
                        QString funcName = item->data(Qt::UserRole).toString();
                        popup->hide();
                        QString text = editor->text();
                        QString token = extractCurrentToken(text);
                        if (!token.isEmpty()) {
                            int tokenStart = text.length() - token.length();
                            editor->setText(text.left(tokenStart) + funcName + "(");
                            editor->setCursorPosition(editor->text().length());
                        }
                    }
                    return true;
                }
                if (keyEvent->key() == Qt::Key_Escape) {
                    popup->hide();
                    return true;
                }
            }
        }

        // Arrow keys during editing: commit and move (like Excel)
        // But NOT during formula edit mode — arrows navigate in the formula text
        if (!m_formulaEditMode &&
            (keyEvent->key() == Qt::Key_Up || keyEvent->key() == Qt::Key_Down ||
             keyEvent->key() == Qt::Key_Left || keyEvent->key() == Qt::Key_Right)) {
            if (editor) {
                // Left/Right: only commit if cursor is at boundary
                if (keyEvent->key() == Qt::Key_Left && editor->cursorPosition() > 0)
                    return QStyledItemDelegate::eventFilter(object, event);
                if (keyEvent->key() == Qt::Key_Right && editor->cursorPosition() < editor->text().length())
                    return QStyledItemDelegate::eventFilter(object, event);

                emit commitData(editor);
                emit closeEditor(editor, QAbstractItemDelegate::NoHint);

                // Navigate to adjacent cell after editor closes
                int key = keyEvent->key();
                QWidget* viewport = editor->parentWidget();
                QTableView* view = viewport ? qobject_cast<QTableView*>(viewport->parentWidget()) : nullptr;
                if (view) {
                    QTimer::singleShot(0, view, [view, key]() {
                        QModelIndex cur = view->currentIndex();
                        if (!cur.isValid() || !view->model()) return;
                        int row = cur.row(), col = cur.column();
                        if (key == Qt::Key_Up) row = qMax(0, row - 1);
                        else if (key == Qt::Key_Down) row = qMin(view->model()->rowCount() - 1, row + 1);
                        else if (key == Qt::Key_Left) col = qMax(0, col - 1);
                        else if (key == Qt::Key_Right) col = qMin(view->model()->columnCount() - 1, col + 1);
                        QModelIndex next = view->model()->index(row, col);
                        if (next.isValid()) {
                            view->setCurrentIndex(next);
                        }
                    });
                }
                return true; // consume the event
            }
        }
    }
    return QStyledItemDelegate::eventFilter(object, event);
}

QWidget* CellDelegate::createEditor(QWidget* parent, const QStyleOptionViewItem& option,
                                   const QModelIndex& index) const {
    QLineEdit* editor = new QLineEdit(parent);
    editor->setFrame(false);
    editor->setStyleSheet(
        "QLineEdit { background: white; padding: 1px 2px; "
        "border: 2px solid #107C10; selection-background-color: #0078D4; }");

    // --- Custom formula autocomplete popup ---
    auto* popup = new QListWidget(parent->window());
    popup->setWindowFlags(Qt::Popup | Qt::FramelessWindowHint);
    popup->setAttribute(Qt::WA_ShowWithoutActivating);
    popup->setFocusPolicy(Qt::NoFocus);
    popup->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    popup->setStyleSheet(
        "QListWidget { background: white; border: 1px solid #C8C8C8; font-size: 12px; outline: none; }"
        "QListWidget::item { padding: 4px 8px; border: none; }"
        "QListWidget::item:selected { background: #E8F0FE; }"
    );
    popup->setIconSize(QSize(0, 0));
    popup->hide();

    // --- Parameter hint tooltip ---
    auto* paramHint = new QLabel(parent->window());
    paramHint->setWindowFlags(Qt::ToolTip | Qt::FramelessWindowHint);
    paramHint->setAttribute(Qt::WA_ShowWithoutActivating);
    paramHint->setStyleSheet(
        "QLabel { background: #FFF8DC; border: 1px solid #E0D8B0; padding: 4px 8px; "
        "font-size: 12px; color: #333; border-radius: 3px; }");
    paramHint->setTextFormat(Qt::RichText);
    paramHint->hide();

    // Populate popup items from registry
    auto populatePopup = [popup](const QString& prefix) {
        popup->clear();
        const auto& reg = formulaRegistry();
        QString upper = prefix.toUpper();
        for (auto it = reg.begin(); it != reg.end(); ++it) {
            if (it->name.contains(upper, Qt::CaseInsensitive)) {
                auto* item = new QListWidgetItem(popup);
                // Store function name in UserRole for retrieval
                item->setData(Qt::UserRole, it->name);
                // Rich text display: bold name + gray description
                QString display = it->name;
                // Pad to align descriptions
                while (display.length() < 14) display += ' ';
                item->setText(display + it->description);

                // Custom font: name bold, rest normal
                QFont f = popup->font();
                f.setPointSize(11);
                item->setFont(f);
            }
        }
        return popup->count() > 0;
    };

    // Insert selected function into editor
    auto insertFunction = [editor, popup, paramHint](const QString& funcName) {
        popup->hide();
        QString text = editor->text();
        // Find the start of the token we're replacing
        QString token = extractCurrentToken(text);
        if (!token.isEmpty()) {
            int tokenStart = text.length() - token.length();
            editor->setText(text.left(tokenStart) + funcName + "(");
            editor->setCursorPosition(editor->text().length());
        }
        editor->setFocus();
    };

    // When item clicked in popup
    QObject::connect(popup, &QListWidget::itemClicked, editor,
        [insertFunction](QListWidgetItem* item) {
            insertFunction(item->data(Qt::UserRole).toString());
        });

    // Handle keyboard in editor for popup navigation
    editor->installEventFilter(const_cast<CellDelegate*>(this));

    // Store popup and paramHint on editor for access in eventFilter
    editor->setProperty("_formulaPopup", QVariant::fromValue(static_cast<QObject*>(popup)));
    editor->setProperty("_paramHint", QVariant::fromValue(static_cast<QObject*>(paramHint)));

    // Show popup / param hint on text change
    QObject::connect(editor, &QLineEdit::textChanged, editor,
        [this, editor, popup, paramHint, populatePopup, insertFunction](const QString& text) {
        emit formulaEditModeChanged(text.startsWith("="));

        if (text.startsWith("=") && text.length() > 1) {
            QString token = extractCurrentToken(text);

            // Show autocomplete if typing a function name token
            if (!token.isEmpty() && token[0].isLetter()) {
                if (populatePopup(token)) {
                    // Position below the editor
                    QPoint pos = editor->mapToGlobal(QPoint(0, editor->height()));
                    popup->setFixedWidth(qMax(380, editor->width()));
                    int visibleItems = qMin(popup->count(), 8);
                    popup->setFixedHeight(visibleItems * 26 + 4);
                    popup->move(pos);
                    popup->show();
                    popup->setCurrentRow(0);
                } else {
                    popup->hide();
                }
            } else {
                popup->hide();
            }

            // Show param hint if inside a function call
            int cursorPos = editor->cursorPosition();
            FormulaContext ctx = findFormulaContext(text, cursorPos);
            const auto& reg = formulaRegistry();
            if (ctx.paramIndex >= 0 && reg.contains(ctx.funcName)) {
                const auto& info = reg[ctx.funcName];
                paramHint->setText(buildParamHintHtml(info, ctx.paramIndex));
                paramHint->adjustSize();
                QPoint hintPos = editor->mapToGlobal(QPoint(0, editor->height() + 2));
                // If popup is visible, place hint below it
                if (popup->isVisible()) {
                    hintPos.setY(popup->mapToGlobal(QPoint(0, popup->height())).y() + 2);
                }
                paramHint->move(hintPos);
                paramHint->show();
            } else {
                paramHint->hide();
            }
        } else {
            popup->hide();
            paramHint->hide();
        }
    });

    // Also update param hint on cursor position change
    QObject::connect(editor, &QLineEdit::cursorPositionChanged, editor,
        [editor, paramHint, popup](int, int newPos) {
        QString text = editor->text();
        if (!text.startsWith("=")) { paramHint->hide(); return; }
        FormulaContext ctx = findFormulaContext(text, newPos);
        const auto& reg = formulaRegistry();
        if (ctx.paramIndex >= 0 && reg.contains(ctx.funcName)) {
            const auto& info = reg[ctx.funcName];
            paramHint->setText(buildParamHintHtml(info, ctx.paramIndex));
            paramHint->adjustSize();
            QPoint hintPos = editor->mapToGlobal(QPoint(0, editor->height() + 2));
            if (popup->isVisible()) {
                hintPos.setY(popup->mapToGlobal(QPoint(0, popup->height())).y() + 2);
            }
            paramHint->move(hintPos);
            paramHint->show();
        } else {
            paramHint->hide();
        }
    });

    // Clean up on editor destruction
    QObject::connect(editor, &QObject::destroyed, popup, &QObject::deleteLater);
    QObject::connect(editor, &QObject::destroyed, paramHint, &QObject::deleteLater);

    return editor;
}

void CellDelegate::setEditorData(QWidget* editor, const QModelIndex& index) const {
    QLineEdit* lineEdit = qobject_cast<QLineEdit*>(editor);
    if (lineEdit) {
        lineEdit->setText(index.data(Qt::EditRole).toString());
        // Place cursor at end (not selecting all text) — like Excel
        // Deferred to after Qt finishes initializing the editor, which otherwise re-selects all text
        QTimer::singleShot(0, lineEdit, [lineEdit]() {
            lineEdit->deselect();
            lineEdit->setCursorPosition(lineEdit->text().length());
        });
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
    // Editor fills the full cell — its own green border replaces the delegate focus border
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
        if (cellBg.isValid() && cellBg.rgb() != QColor(Qt::white).rgb()) {
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

    // --- Formula recalc flash overlay (bright yellow, fade-in then hold then fade-out) ---
    if (m_spreadsheetView) {
        double flashProgress = m_spreadsheetView->cellAnimationProgress(index.row(), index.column());
        if (flashProgress > 0.01) {
            int alpha = static_cast<int>(130 * flashProgress);
            painter->fillRect(rect, QColor(255, 243, 100, alpha));
        }
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

        // Indent support: add left padding per indent level
        int indentPx = 0;
        QVariant indentData = index.data(Qt::UserRole + 10);
        if (indentData.isValid()) {
            indentPx = indentData.toInt() * 12; // 12px per indent level
        }

        // Text rotation support
        int rotation = 0;
        QVariant rotData = index.data(Qt::UserRole + 16);
        if (rotData.isValid()) {
            rotation = rotData.toInt();
        }

        QRect textRect = rect.adjusted(4 + indentPx, 1, -4, -1);

        if (rotation == 0) {
            painter->drawText(textRect, alignment, text);
        } else if (rotation == 270) {
            // Vertical stacked text: draw each character on its own line
            QFontMetrics fm(font);
            int charH = fm.height();
            int maxCharW = 0;
            for (const QChar& ch : text) {
                maxCharW = qMax(maxCharW, fm.horizontalAdvance(ch));
            }
            int totalH = charH * text.length();
            int startY = textRect.top() + qMax(0, (textRect.height() - totalH) / 2);
            int centerX = textRect.left() + (textRect.width() - maxCharW) / 2;
            for (int i = 0; i < text.length(); ++i) {
                QRect charRect(centerX, startY + i * charH, maxCharW, charH);
                painter->drawText(charRect, Qt::AlignCenter, QString(text[i]));
            }
        } else {
            // Angled text: rotate painter
            painter->save();
            QFontMetrics fm(font);
            int textW = fm.horizontalAdvance(text);
            int textH = fm.height();
            QPointF center = textRect.center();
            painter->translate(center);
            painter->rotate(-rotation); // Qt rotates clockwise, we want CCW for positive angles
            QRectF rotRect(-textW / 2.0, -textH / 2.0, textW, textH);
            painter->drawText(rotRect, Qt::AlignCenter, text);
            painter->restore();
        }
    }

    // --- Sparkline rendering ---
    QVariant sparkData = index.data(Qt::UserRole + 15); // SparklineRole
    if (sparkData.isValid() && sparkData.canConvert<SparklineRenderData>()) {
        auto rd = sparkData.value<SparklineRenderData>();
        if (!rd.values.isEmpty()) {
            QRect sparkRect = rect.adjusted(3, 3, -3, -3);
            drawSparkline(painter, sparkRect, rd);
        }
    }

    // --- Gridlines: single thin line on right and bottom edges ---
    if (m_showGridlines) {
        painter->setPen(QPen(QColor(218, 220, 224), 1, Qt::SolidLine));
        painter->drawLine(rect.right(), rect.top(), rect.right(), rect.bottom());
        painter->drawLine(rect.left(), rect.bottom(), rect.right(), rect.bottom());
    }

    // --- Cell borders (user-defined) ---
    auto drawBorder = [&](const QVariant& borderData, int x1, int y1, int x2, int y2) {
        if (!borderData.isValid()) return;
        QStringList parts = borderData.toString().split(',');
        if (parts.size() >= 2) {
            int w = parts[0].toInt();
            QColor c(parts[1]);
            if (c.isValid() && w > 0) {
                painter->setPen(QPen(c, w, Qt::SolidLine));
                painter->drawLine(x1, y1, x2, y2);
            }
        }
    };
    drawBorder(index.data(Qt::UserRole + 11), rect.left(), rect.top(), rect.right(), rect.top());      // top
    drawBorder(index.data(Qt::UserRole + 12), rect.left(), rect.bottom(), rect.right(), rect.bottom()); // bottom
    drawBorder(index.data(Qt::UserRole + 13), rect.left(), rect.top(), rect.left(), rect.bottom());     // left
    drawBorder(index.data(Qt::UserRole + 14), rect.right(), rect.top(), rect.right(), rect.bottom());   // right

    // --- Focus border: green rectangle for the current cell ---
    if (hasFocus) {
        painter->setClipRect(rect.adjusted(-1, -1, 1, 1)); // expand clip to avoid top/left clipping
        QPen focusPen(QColor(16, 124, 16), 2, Qt::SolidLine);
        painter->setPen(focusPen);
        painter->drawRect(rect.adjusted(1, 1, -1, -1));
    }

    painter->restore();
}

QSize CellDelegate::sizeHint(const QStyleOptionViewItem& option,
                             const QModelIndex& index) const {
    return QStyledItemDelegate::sizeHint(option, index);
}

void CellDelegate::drawSparkline(QPainter* painter, const QRect& rect,
                                  const SparklineRenderData& data) const {
    if (data.values.isEmpty() || rect.width() < 4 || rect.height() < 4) return;

    double range = data.maxVal - data.minVal;
    if (range == 0) range = 1.0;
    int n = data.values.size();

    switch (data.type) {
        case SparklineType::Line: {
            painter->setRenderHint(QPainter::Antialiasing, true);
            double stepX = static_cast<double>(rect.width()) / qMax(1, n - 1);

            QPainterPath path;
            for (int i = 0; i < n; ++i) {
                double x = rect.left() + i * stepX;
                double y = rect.bottom() - ((data.values[i] - data.minVal) / range) * rect.height();
                if (i == 0) path.moveTo(x, y);
                else path.lineTo(x, y);
            }
            painter->setPen(QPen(data.lineColor, data.lineWidth));
            painter->setBrush(Qt::NoBrush);
            painter->drawPath(path);

            if (data.showHighPoint && data.highIndex >= 0) {
                double x = rect.left() + data.highIndex * stepX;
                double y = rect.bottom() - ((data.values[data.highIndex] - data.minVal) / range) * rect.height();
                painter->setPen(Qt::NoPen);
                painter->setBrush(data.highPointColor);
                painter->drawEllipse(QPointF(x, y), 3, 3);
            }
            if (data.showLowPoint && data.lowIndex >= 0) {
                double x = rect.left() + data.lowIndex * stepX;
                double y = rect.bottom() - ((data.values[data.lowIndex] - data.minVal) / range) * rect.height();
                painter->setPen(Qt::NoPen);
                painter->setBrush(data.lowPointColor);
                painter->drawEllipse(QPointF(x, y), 3, 3);
            }
            painter->setRenderHint(QPainter::Antialiasing, false);
            break;
        }
        case SparklineType::Column: {
            double barW = static_cast<double>(rect.width()) / (n * 1.4);
            double zeroY = rect.bottom();
            if (data.minVal < 0) {
                zeroY = rect.bottom() - ((-data.minVal) / range) * rect.height();
            }
            for (int i = 0; i < n; ++i) {
                double x = rect.left() + i * (static_cast<double>(rect.width()) / n) +
                           (static_cast<double>(rect.width()) / n - barW) / 2;
                double barH = (std::abs(data.values[i]) / range) * rect.height();
                if (data.values[i] >= 0) {
                    QRectF bar(x, zeroY - barH, barW, barH);
                    painter->fillRect(bar, data.lineColor);
                } else {
                    QRectF bar(x, zeroY, barW, barH);
                    painter->fillRect(bar, data.negativeColor);
                }
            }
            break;
        }
        case SparklineType::WinLoss: {
            double barW = static_cast<double>(rect.width()) / (n * 1.4);
            double midY = rect.top() + rect.height() / 2.0;
            double halfH = rect.height() / 2.0 - 2;
            for (int i = 0; i < n; ++i) {
                double x = rect.left() + i * (static_cast<double>(rect.width()) / n) +
                           (static_cast<double>(rect.width()) / n - barW) / 2;
                if (data.values[i] > 0) {
                    QRectF bar(x, midY - halfH, barW, halfH);
                    painter->fillRect(bar, data.lineColor);
                } else if (data.values[i] < 0) {
                    QRectF bar(x, midY, barW, halfH);
                    painter->fillRect(bar, data.negativeColor);
                }
            }
            break;
        }
    }
}
