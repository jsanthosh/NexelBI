#ifndef CHATPANEL_H
#define CHATPANEL_H

#include <QWidget>
#include <QTextEdit>
#include <QLineEdit>
#include <QPushButton>
#include <QVBoxLayout>
#include <QNetworkAccessManager>
#include <QNetworkReply>
#include <QJsonArray>
#include <QTimer>
#include <QLabel>

class Spreadsheet;

class ChatPanel : public QWidget {
    Q_OBJECT

public:
    explicit ChatPanel(QWidget* parent = nullptr);

    void setSpreadsheet(std::shared_ptr<Spreadsheet> spreadsheet);
    void setApiKey(const QString& apiKey);

signals:
    void insertFormula(const QString& formula);
    void insertValue(const QString& value);
    void executeActions(const QJsonArray& actions);

private slots:
    void onSendMessage();
    void onApiResponse(QNetworkReply* reply);
    void onThinkingTick();

private:
    void addMessage(const QString& sender, const QString& text, bool isUser);
    void addSystemMessage(const QString& text);
    QString buildContext() const;
    void sendToApi(const QString& userMessage);
    QString extractAndProcessActions(const QString& responseText);
    void showThinkingIndicator();
    void hideThinkingIndicator();

    QTextEdit* m_chatDisplay;
    QLineEdit* m_inputField;
    QPushButton* m_sendBtn;
    QNetworkAccessManager* m_networkManager;
    std::shared_ptr<Spreadsheet> m_spreadsheet;
    QString m_apiKey;
    QString m_pendingUserMessage;

    // Thinking indicator
    QWidget* m_thinkingWidget = nullptr;
    QLabel* m_thinkingLabel = nullptr;
    QTimer* m_thinkingTimer = nullptr;
    int m_thinkingDots = 0;
    QVBoxLayout* m_mainLayout = nullptr;
};

#endif // CHATPANEL_H
