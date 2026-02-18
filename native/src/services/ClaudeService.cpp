#include "ClaudeService.h"

ClaudeService& ClaudeService::instance() {
    static ClaudeService s_instance;
    return s_instance;
}

ClaudeService::ClaudeService() : m_initialized(false) {
}

bool ClaudeService::initialize(const QString& apiKey) {
    if (apiKey.isEmpty()) {
        m_lastError = "API key is empty";
        return false;
    }

    m_apiKey = apiKey;
    m_initialized = true;
    return true;
}

QString ClaudeService::queryAssistant(const QString& question, const QString& context) {
    if (!m_initialized) {
        m_lastError = "Claude service not initialized";
        return "";
    }

    QString prompt = context.isEmpty() ? question : "Context: " + context + "\n\nQuestion: " + question;
    return makeRequest(prompt);
}

std::vector<QString> ClaudeService::suggestFormulas(const QString& description) {
    if (!m_initialized) {
        m_lastError = "Claude service not initialized";
        return {};
    }

    QString prompt = "Suggest Excel formulas for the following task: " + description + "\nProvide 3-5 formula suggestions as a comma-separated list.";
    QString response = makeRequest(prompt);

    // Parse response to extract formulas
    std::vector<QString> formulas;
    QStringList parts = response.split(',');
    for (const auto& part : parts) {
        QString formula = part.trimmed();
        if (!formula.isEmpty()) {
            formulas.push_back(formula);
        }
    }

    return formulas;
}

QString ClaudeService::analyzeData(const QString& dataDescription) {
    if (!m_initialized) {
        m_lastError = "Claude service not initialized";
        return "";
    }

    QString prompt = "Analyze the following spreadsheet data and provide insights:\n" + dataDescription;
    return makeRequest(prompt);
}

QString ClaudeService::suggestCellContent(const QString& cellContext) {
    if (!m_initialized) {
        m_lastError = "Claude service not initialized";
        return "";
    }

    QString prompt = "Given this cell context: " + cellContext + "\nSuggest appropriate content or formula for this cell.";
    return makeRequest(prompt);
}

QString ClaudeService::getLastError() const {
    return m_lastError;
}

bool ClaudeService::hasError() const {
    return !m_lastError.isEmpty();
}

QString ClaudeService::makeRequest(const QString& prompt) {
    // TODO: Implement actual HTTP request to Claude API
    // For now, return a placeholder
    return "Claude API response pending implementation";
}
