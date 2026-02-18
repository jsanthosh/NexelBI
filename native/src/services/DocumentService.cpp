#include "DocumentService.h"

DocumentService& DocumentService::instance() {
    static DocumentService s_instance;
    return s_instance;
}

DocumentService::DocumentService() {
}

bool DocumentService::createNewDocument(const QString& name) {
    auto spreadsheet = std::make_shared<Spreadsheet>();
    return DocumentRepository::instance().createDocument(name, spreadsheet);
}

bool DocumentService::openDocument(const QString& id) {
    m_currentDocument = DocumentRepository::instance().getDocument(id);
    if (!m_currentDocument) {
        m_lastError = "Failed to open document: " + id;
        return false;
    }
    return true;
}

bool DocumentService::saveDocument() {
    if (!m_currentDocument) {
        m_lastError = "No document is currently open";
        return false;
    }

    return DocumentRepository::instance().updateDocument(
        m_currentDocument->id,
        m_currentDocument->name,
        m_currentDocument->spreadsheet
    );
}

bool DocumentService::saveDocumentAs(const QString& name) {
    if (!m_currentDocument) {
        m_lastError = "No document is currently open";
        return false;
    }

    m_currentDocument->name = name;
    return saveDocument();
}

bool DocumentService::closeDocument() {
    if (m_currentDocument) {
        m_currentDocument = nullptr;
        return true;
    }
    return false;
}

std::shared_ptr<Document> DocumentService::getCurrentDocument() const {
    return m_currentDocument;
}

std::shared_ptr<Spreadsheet> DocumentService::getCurrentSpreadsheet() const {
    if (!m_currentDocument) {
        return nullptr;
    }
    return m_currentDocument->spreadsheet;
}

bool DocumentService::importCSV(const QString& filePath) {
    // TODO: Implement CSV import
    return false;
}

bool DocumentService::importExcel(const QString& filePath) {
    // TODO: Implement Excel import
    return false;
}

bool DocumentService::exportCSV(const QString& filePath) {
    // TODO: Implement CSV export
    return false;
}

bool DocumentService::exportExcel(const QString& filePath) {
    // TODO: Implement Excel export
    return false;
}

QString DocumentService::getLastError() const {
    return m_lastError;
}
