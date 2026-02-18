#ifndef CLAUDESERVICE_H
#define CLAUDESERVICE_H

#include <QString>
#include <QVariant>
#include <vector>

class ClaudeService {
public:
    static ClaudeService& instance();

    // Initialize with API key
    bool initialize(const QString& apiKey);

    // Query Claude for spreadsheet assistance
    QString queryAssistant(const QString& question, const QString& context = "");

    // Generate formula suggestions
    std::vector<QString> suggestFormulas(const QString& description);

    // Data analysis
    QString analyzeData(const QString& dataDescription);

    // Cell content suggestion
    QString suggestCellContent(const QString& cellContext);

    QString getLastError() const;
    bool hasError() const;

private:
    ClaudeService();
    ~ClaudeService() = default;

    QString m_apiKey;
    QString m_lastError;
    bool m_initialized;

    QString makeRequest(const QString& prompt);
};

#endif // CLAUDESERVICE_H
