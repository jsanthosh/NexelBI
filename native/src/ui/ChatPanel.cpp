#include "ChatPanel.h"
#include "../core/Spreadsheet.h"
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QLabel>
#include <QScrollBar>
#include <QSettings>
#include <QInputDialog>
#include <QRegularExpression>
#include <QHBoxLayout>
#include <QGraphicsDropShadowEffect>
#include <QTextBrowser>

ChatPanel::ChatPanel(QWidget* parent)
    : QWidget(parent) {

    m_networkManager = new QNetworkAccessManager(this);
    connect(m_networkManager, &QNetworkAccessManager::finished, this, &ChatPanel::onApiResponse);

    // Layout
    m_mainLayout = new QVBoxLayout(this);
    m_mainLayout->setContentsMargins(0, 0, 0, 0);
    m_mainLayout->setSpacing(0);

    // ---- Header with gradient ----
    auto* header = new QWidget(this);
    header->setFixedHeight(48);
    header->setStyleSheet(
        "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, "
        "stop:0 #1B5E3B, stop:1 #2E8B57);"
        "border: none;");
    auto* headerLayout = new QHBoxLayout(header);
    headerLayout->setContentsMargins(14, 0, 10, 0);

    auto* sparkleLabel = new QLabel(header);
    sparkleLabel->setText("\u2728");
    sparkleLabel->setStyleSheet("font-size: 18px; background: transparent; border: none;");
    headerLayout->addWidget(sparkleLabel);

    auto* headerLabel = new QLabel("Claude Assistant", header);
    headerLabel->setStyleSheet(
        "color: white; font-weight: 600; font-size: 14px; "
        "letter-spacing: 0.3px; background: transparent; border: none;");
    headerLayout->addWidget(headerLabel);
    headerLayout->addStretch();

    auto* apiKeyBtn = new QPushButton(header);
    apiKeyBtn->setText("\u2699");  // gear icon
    apiKeyBtn->setToolTip("Set API Key");
    apiKeyBtn->setFixedSize(30, 30);
    apiKeyBtn->setStyleSheet(
        "QPushButton { background: rgba(255,255,255,0.15); color: white; border: none; "
        "border-radius: 15px; font-size: 16px; }"
        "QPushButton:hover { background: rgba(255,255,255,0.3); }");
    headerLayout->addWidget(apiKeyBtn);
    connect(apiKeyBtn, &QPushButton::clicked, this, [this]() {
        QString key = QInputDialog::getText(this, "Claude API Key",
            "Enter your Anthropic API key:", QLineEdit::Password, m_apiKey);
        if (!key.isEmpty()) {
            setApiKey(key);
            QSettings settings("NativeSpreadsheet", "NativeSpreadsheet");
            settings.setValue("claude_api_key", key);
        }
    });

    m_mainLayout->addWidget(header);

    // ---- Chat display ----
    m_chatDisplay = new QTextEdit(this);
    m_chatDisplay->setReadOnly(true);
    m_chatDisplay->setStyleSheet(
        "QTextEdit { background: #F8FAFB; border: none; padding: 12px; font-size: 13px; "
        "font-family: -apple-system, 'Segoe UI', system-ui, sans-serif; }"
        "QScrollBar:vertical { width: 6px; background: transparent; margin: 2px; }"
        "QScrollBar::handle:vertical { background: rgba(0,0,0,0.15); border-radius: 3px; min-height: 30px; }"
        "QScrollBar::handle:vertical:hover { background: rgba(0,0,0,0.25); }"
        "QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }");

    // Welcome message
    m_chatDisplay->setHtml(
        "<div style='padding: 16px 8px; text-align: center;'>"
        "<div style='font-size: 32px; margin-bottom: 8px;'>\u2728</div>"
        "<p style='font-size: 15px; font-weight: 600; color: #1B5E3B; margin: 4px 0;'>Claude Spreadsheet Assistant</p>"
        "<p style='font-size: 12px; color: #64748B; margin: 4px 0 12px 0;'>I can modify your spreadsheet directly. Try asking:</p>"
        "<div style='text-align: left; background: #F0FDF4; border: 1px solid #BBF7D0; "
        "border-radius: 8px; padding: 10px 14px; margin: 0 4px;'>"
        "<p style='font-size: 12px; color: #15803D; margin: 3px 0;'>\u25B6 \"Set A1 to Name, B1 to Age\"</p>"
        "<p style='font-size: 12px; color: #15803D; margin: 3px 0;'>\u25B6 \"Make row 1 bold with blue background\"</p>"
        "<p style='font-size: 12px; color: #15803D; margin: 3px 0;'>\u25B6 \"Create a monthly budget table\"</p>"
        "<p style='font-size: 12px; color: #15803D; margin: 3px 0;'>\u25B6 \"Merge cells A1:D1 and center\"</p>"
        "<p style='font-size: 12px; color: #15803D; margin: 3px 0;'>\u25B6 \"Apply Ocean Blue table theme\"</p>"
        "</div>"
        "<p style='font-size: 11px; color: #94A3B8; margin-top: 10px;'>"
        "Click \u2699 above to set your API key</p>"
        "</div>");

    m_mainLayout->addWidget(m_chatDisplay);

    // ---- Thinking indicator (hidden by default) ----
    m_thinkingWidget = new QWidget(this);
    m_thinkingWidget->setFixedHeight(44);
    m_thinkingWidget->setStyleSheet(
        "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, "
        "stop:0 #F0FDF4, stop:1 #ECFDF5); "
        "border-top: 1px solid #D1FAE5;");
    auto* thinkingLayout = new QHBoxLayout(m_thinkingWidget);
    thinkingLayout->setContentsMargins(14, 6, 14, 6);

    auto* thinkingIcon = new QLabel(m_thinkingWidget);
    thinkingIcon->setText("\u2728");
    thinkingIcon->setStyleSheet("font-size: 14px; border: none; background: transparent;");
    thinkingLayout->addWidget(thinkingIcon);

    m_thinkingLabel = new QLabel("Claude is thinking", m_thinkingWidget);
    m_thinkingLabel->setStyleSheet(
        "color: #16A34A; font-size: 12px; font-weight: 500; "
        "border: none; background: transparent;");
    thinkingLayout->addWidget(m_thinkingLabel);
    thinkingLayout->addStretch();

    m_thinkingWidget->hide();
    m_mainLayout->addWidget(m_thinkingWidget);

    // Thinking animation timer
    m_thinkingTimer = new QTimer(this);
    m_thinkingTimer->setInterval(500);
    connect(m_thinkingTimer, &QTimer::timeout, this, &ChatPanel::onThinkingTick);

    // ---- Input area ----
    auto* inputContainer = new QWidget(this);
    inputContainer->setFixedHeight(56);
    inputContainer->setStyleSheet(
        "background: #FFFFFF; border-top: 1px solid #E2E8F0;");
    auto* inputLayout = new QHBoxLayout(inputContainer);
    inputLayout->setContentsMargins(10, 8, 10, 8);
    inputLayout->setSpacing(8);

    m_inputField = new QLineEdit(inputContainer);
    m_inputField->setPlaceholderText("Ask Claude to modify your spreadsheet...");
    m_inputField->setStyleSheet(
        "QLineEdit { background: #F8FAFB; border: 1.5px solid #E2E8F0; border-radius: 20px; "
        "padding: 8px 16px; font-size: 13px; color: #1E293B; "
        "font-family: -apple-system, 'Segoe UI', system-ui, sans-serif; }"
        "QLineEdit:focus { border-color: #22C55E; background: #FFFFFF; }");
    inputLayout->addWidget(m_inputField);

    m_sendBtn = new QPushButton(inputContainer);
    m_sendBtn->setText("\u27A4");  // arrow
    m_sendBtn->setFixedSize(36, 36);
    m_sendBtn->setStyleSheet(
        "QPushButton { background: #16A34A; color: white; border: none; border-radius: 18px; "
        "font-size: 16px; font-weight: bold; }"
        "QPushButton:hover { background: #15803D; }"
        "QPushButton:disabled { background: #D1D5DB; }");
    inputLayout->addWidget(m_sendBtn);

    m_mainLayout->addWidget(inputContainer);

    // Connections
    connect(m_sendBtn, &QPushButton::clicked, this, &ChatPanel::onSendMessage);
    connect(m_inputField, &QLineEdit::returnPressed, this, &ChatPanel::onSendMessage);

    // Load saved API key
    QSettings settings("NativeSpreadsheet", "NativeSpreadsheet");
    m_apiKey = settings.value("claude_api_key").toString();
}

void ChatPanel::setSpreadsheet(std::shared_ptr<Spreadsheet> spreadsheet) {
    m_spreadsheet = spreadsheet;
}

void ChatPanel::setApiKey(const QString& apiKey) {
    m_apiKey = apiKey;
}

void ChatPanel::showThinkingIndicator() {
    m_thinkingDots = 0;
    m_thinkingLabel->setText("\u2728 Claude is thinking");
    m_thinkingWidget->show();
    m_thinkingTimer->start();
    m_sendBtn->setEnabled(false);
    m_inputField->setEnabled(false);
}

void ChatPanel::hideThinkingIndicator() {
    m_thinkingTimer->stop();
    m_thinkingWidget->hide();
    m_sendBtn->setEnabled(true);
    m_inputField->setEnabled(true);
    m_inputField->setFocus();
}

void ChatPanel::onThinkingTick() {
    m_thinkingDots++;

    static const QString phases[] = {
        "Analyzing your spreadsheet",
        "Working on it",
        "Preparing changes",
        "Generating response",
    };

    int phase = (m_thinkingDots / 6) % 4;
    int dotCount = (m_thinkingDots % 3) + 1;
    QString dots;
    for (int i = 0; i < dotCount; i++) dots += ".";

    m_thinkingLabel->setText(phases[phase] + dots);
}

void ChatPanel::onSendMessage() {
    QString text = m_inputField->text().trimmed();
    if (text.isEmpty()) return;

    m_inputField->clear();
    addMessage("You", text, true);

    if (m_apiKey.isEmpty()) {
        addMessage("Claude", "Please set your API key first using the \u2699 button above.", false);
        return;
    }

    showThinkingIndicator();
    sendToApi(text);
}

void ChatPanel::addMessage(const QString& sender, const QString& text, bool isUser) {
    QString escapedText = text.toHtmlEscaped().replace("\n", "<br>");

    QString html;
    if (isUser) {
        html = QString(
            "<div style='text-align: right; margin: 6px 4px;'>"
            "<div style='display: inline-block; background: #16A34A; "
            "border-radius: 16px 16px 4px 16px; padding: 10px 14px; max-width: 85%%;'>"
            "<div style='font-size: 13px; color: #FFFFFF; line-height: 1.4;'>%1</div>"
            "</div></div>")
            .arg(escapedText);
    } else {
        html = QString(
            "<div style='text-align: left; margin: 6px 4px;'>"
            "<div style='display: inline-block; background: #FFFFFF; "
            "border: 1px solid #E2E8F0; "
            "border-radius: 16px 16px 16px 4px; padding: 10px 14px; max-width: 85%%;'>"
            "<div style='font-size: 13px; color: #1E293B; line-height: 1.4;'>%1</div>"
            "</div></div>")
            .arg(escapedText);
    }

    m_chatDisplay->append(html);
    m_chatDisplay->verticalScrollBar()->setValue(m_chatDisplay->verticalScrollBar()->maximum());
}

void ChatPanel::addSystemMessage(const QString& text) {
    QString html = QString(
        "<div style='text-align: center; margin: 4px 0;'>"
        "<div style='display: inline-block; background: #F0FDF4; border: 1px solid #BBF7D0; "
        "border-radius: 12px; padding: 6px 14px;'>"
        "<span style='font-size: 11px; color: #16A34A; font-weight: 500;'>\u2713 %1</span>"
        "</div></div>").arg(text.toHtmlEscaped());
    m_chatDisplay->append(html);
    m_chatDisplay->verticalScrollBar()->setValue(m_chatDisplay->verticalScrollBar()->maximum());
}

QString ChatPanel::buildContext() const {
    if (!m_spreadsheet) return "";

    QString context;
    context += "Sheet: " + m_spreadsheet->getSheetName() + "\n";
    context += "Rows: " + QString::number(m_spreadsheet->getMaxRow() + 1) + "\n";
    context += "Cols: " + QString::number(m_spreadsheet->getMaxColumn() + 1) + "\n";

    // Tables
    const auto& tables = m_spreadsheet->getTables();
    if (!tables.empty()) {
        context += "Tables: ";
        for (const auto& t : tables) {
            context += t.name + " (" + t.theme.name + ") ";
        }
        context += "\n";
    }

    // Merged regions
    const auto& merged = m_spreadsheet->getMergedRegions();
    if (!merged.empty()) {
        context += "Merged regions: " + QString::number(merged.size()) + "\n";
    }

    // Sample data
    int maxRow = qMin(m_spreadsheet->getMaxRow(), 14);
    int maxCol = qMin(m_spreadsheet->getMaxColumn(), 9);

    if (maxRow >= 0 && maxCol >= 0) {
        context += "\nData (first rows/cols):\n";
        // Column headers
        context += "\t";
        for (int c = 0; c <= maxCol; ++c) {
            context += QString(QChar('A' + c)) + "\t";
        }
        context += "\n";
        for (int r = 0; r <= maxRow; ++r) {
            context += QString::number(r + 1) + "\t";
            for (int c = 0; c <= maxCol; ++c) {
                CellAddress addr(r, c);
                auto val = m_spreadsheet->getCellValue(addr);
                QString cellStr = val.toString();
                if (cellStr.length() > 20) cellStr = cellStr.left(17) + "...";
                context += cellStr + "\t";
            }
            context += "\n";
        }
    }

    return context;
}

void ChatPanel::sendToApi(const QString& userMessage) {
    m_pendingUserMessage = userMessage;

    QNetworkRequest request(QUrl("https://api.anthropic.com/v1/messages"));
    request.setHeader(QNetworkRequest::ContentTypeHeader, "application/json");
    request.setRawHeader("x-api-key", m_apiKey.toUtf8());
    request.setRawHeader("anthropic-version", "2023-06-01");

    QString context = buildContext();

    // Available table themes for reference
    QString themeList;
    auto themes = getBuiltinTableThemes();
    for (int i = 0; i < static_cast<int>(themes.size()); ++i) {
        themeList += QString::number(i) + "=" + themes[i].name + ", ";
    }

    QString systemPrompt =
        "You are Claude, an AI spreadsheet assistant inside NativeSpreadsheet. "
        "You can explain things AND directly modify the spreadsheet by returning action blocks.\n\n"
        "Return actions using this EXACT format:\n"
        "[ACTIONS]\n"
        "[\n"
        "  {\"action\": \"set_cell\", \"cell\": \"A1\", \"value\": \"Hello\"},\n"
        "  {\"action\": \"set_formula\", \"cell\": \"B2\", \"formula\": \"=SUM(A1:A10)\"},\n"
        "  {\"action\": \"format\", \"range\": \"A1:D1\", \"bold\": true, \"bg_color\": \"#4472C4\", \"fg_color\": \"#FFFFFF\"},\n"
        "  {\"action\": \"merge\", \"range\": \"A1:D1\"},\n"
        "  {\"action\": \"unmerge\", \"range\": \"A1:D1\"},\n"
        "  {\"action\": \"border\", \"range\": \"A1:D10\", \"type\": \"all\"},\n"
        "  {\"action\": \"table\", \"range\": \"A1:D10\", \"theme\": 0},\n"
        "  {\"action\": \"set_row_height\", \"row\": 1, \"height\": 30},\n"
        "  {\"action\": \"set_col_width\", \"col\": \"A\", \"width\": 120},\n"
        "  {\"action\": \"number_format\", \"range\": \"B2:B10\", \"format\": \"Currency\"},\n"
        "  {\"action\": \"clear\", \"range\": \"A1:Z100\"}\n"
        "]\n"
        "[/ACTIONS]\n\n"
        "Available actions:\n"
        "- set_cell: Set cell value. Fields: cell, value (string or number)\n"
        "- set_formula: Set formula. Fields: cell, formula (starts with =)\n"
        "- format: Apply formatting. Fields: range, and any of: bold, italic, underline, strikethrough (bool), "
        "bg_color, fg_color (hex like \"#4472C4\"), font_size (int), font_name (string), "
        "h_align (\"left\"/\"center\"/\"right\"), v_align (\"top\"/\"middle\"/\"bottom\")\n"
        "- merge: Merge cells. Fields: range\n"
        "- unmerge: Unmerge cells. Fields: range\n"
        "- border: Apply borders. Fields: range, type (\"all\"/\"outside\"/\"none\"/\"bottom\"/\"top\"/\"left\"/\"right\"/\"thick_outside\")\n"
        "- table: Apply table theme with banded rows & header. Fields: range, theme (index 0-11)\n"
        "  Available themes: " + themeList + "\n"
        "- number_format: Set number format. Fields: range, format (\"General\"/\"Number\"/\"Currency\"/\"Percentage\"/\"Date\"/\"Text\")\n"
        "- set_row_height: Set row height. Fields: row (1-based), height (pixels)\n"
        "- set_col_width: Set column width. Fields: col (letter like \"A\"), width (pixels)\n"
        "- clear: Clear cell values and formatting. Fields: range\n\n"
        "Rules:\n"
        "- Always explain what you're doing in plain text BEFORE the [ACTIONS] block\n"
        "- Use cell references like A1, B2, AA1 etc.\n"
        "- Ranges use colon: A1:D10\n"
        "- For formulas, use standard Excel syntax starting with =\n"
        "- You can combine many actions in one response\n"
        "- Be concise but friendly\n";

    if (!context.isEmpty()) {
        systemPrompt += "\nCurrent spreadsheet state:\n" + context;
    }

    QJsonObject body;
    body["model"] = "claude-sonnet-4-5-20250929";
    body["max_tokens"] = 4096;
    body["system"] = systemPrompt;

    QJsonArray messages;
    QJsonObject msg;
    msg["role"] = "user";
    msg["content"] = userMessage;
    messages.append(msg);
    body["messages"] = messages;

    m_networkManager->post(request, QJsonDocument(body).toJson());
}

QString ChatPanel::extractAndProcessActions(const QString& responseText) {
    QRegularExpression re("\\[ACTIONS\\]\\s*(\\[.*?\\])\\s*\\[/ACTIONS\\]",
                          QRegularExpression::DotMatchesEverythingOption);
    QRegularExpressionMatch match = re.match(responseText);

    if (!match.hasMatch()) return responseText;

    QString jsonStr = match.captured(1);
    QJsonParseError parseError;
    QJsonDocument doc = QJsonDocument::fromJson(jsonStr.toUtf8(), &parseError);

    if (parseError.error != QJsonParseError::NoError || !doc.isArray()) {
        return responseText;
    }

    QJsonArray actions = doc.array();

    if (!actions.isEmpty()) {
        emit executeActions(actions);

        int setCells = 0, formulas = 0, formats = 0, merges = 0, borders = 0, tables = 0, other = 0;
        for (const auto& item : actions) {
            QJsonObject obj = item.toObject();
            QString type = obj["action"].toString();
            if (type == "set_cell") setCells++;
            else if (type == "set_formula") formulas++;
            else if (type == "format") formats++;
            else if (type == "merge" || type == "unmerge") merges++;
            else if (type == "border") borders++;
            else if (type == "table") tables++;
            else other++;
        }

        QStringList parts;
        if (setCells > 0) parts << QString::number(setCells) + " cell(s) set";
        if (formulas > 0) parts << QString::number(formulas) + " formula(s)";
        if (formats > 0) parts << QString::number(formats) + " format(s)";
        if (merges > 0) parts << QString::number(merges) + " merge(s)";
        if (borders > 0) parts << QString::number(borders) + " border(s)";
        if (tables > 0) parts << QString::number(tables) + " table(s)";
        if (other > 0) parts << QString::number(other) + " other";
        addSystemMessage("Applied: " + parts.join(", "));
    }

    QString cleanText = responseText;
    cleanText.replace(re, "");
    cleanText = cleanText.trimmed();
    return cleanText;
}

void ChatPanel::onApiResponse(QNetworkReply* reply) {
    hideThinkingIndicator();

    if (reply->error() != QNetworkReply::NoError) {
        QString errorMsg = reply->errorString();
        if (reply->error() == QNetworkReply::AuthenticationRequiredError) {
            errorMsg = "Invalid API key. Please check your key and try again.";
        }
        addMessage("Claude", "Error: " + errorMsg, false);
        reply->deleteLater();
        return;
    }

    QByteArray data = reply->readAll();
    QJsonDocument doc = QJsonDocument::fromJson(data);
    QJsonObject obj = doc.object();

    QJsonArray content = obj["content"].toArray();
    QString responseText;
    for (const auto& item : content) {
        QJsonObject block = item.toObject();
        if (block["type"].toString() == "text") {
            responseText += block["text"].toString();
        }
    }

    if (responseText.isEmpty()) {
        responseText = "Sorry, I couldn't process that request. Please try again.";
    }

    QString displayText = extractAndProcessActions(responseText);

    if (!displayText.isEmpty()) {
        addMessage("Claude", displayText, false);
    }

    reply->deleteLater();
}
