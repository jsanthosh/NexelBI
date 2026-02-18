#ifndef DOCUMENTSERVICE_H
#define DOCUMENTSERVICE_H

#include <QString>
#include <QVector>
#include <memory>
#include "../core/Spreadsheet.h"
#include "../database/DocumentRepository.h"

class DocumentService {
public:
    static DocumentService& instance();

    // Document management
    bool createNewDocument(const QString& name);
    bool openDocument(const QString& id);
    bool saveDocument();
    bool saveDocumentAs(const QString& name);
    bool closeDocument();

    // Get current document
    std::shared_ptr<Document> getCurrentDocument() const;
    std::shared_ptr<Spreadsheet> getCurrentSpreadsheet() const;

    // Import/Export
    bool importCSV(const QString& filePath);
    bool importExcel(const QString& filePath);
    bool exportCSV(const QString& filePath);
    bool exportExcel(const QString& filePath);

    QString getLastError() const;

private:
    DocumentService();
    ~DocumentService() = default;

    std::shared_ptr<Document> m_currentDocument;
    QString m_lastError;
};

#endif // DOCUMENTSERVICE_H
