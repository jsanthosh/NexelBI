#ifndef DOCUMENTREPOSITORY_H
#define DOCUMENTREPOSITORY_H

#include <QString>
#include <QVector>
#include <QJsonObject>
#include <memory>
#include "../core/Spreadsheet.h"

struct Document {
    QString id;
    QString name;
    QString createdAt;
    QString updatedAt;
    std::shared_ptr<Spreadsheet> spreadsheet;
};

class DocumentRepository {
public:
    static DocumentRepository& instance();

    // CRUD operations
    bool createDocument(const QString& name, std::shared_ptr<Spreadsheet> spreadsheet);
    std::shared_ptr<Document> getDocument(const QString& id);
    QVector<std::shared_ptr<Document>> getAllDocuments();
    bool updateDocument(const QString& id, const QString& name, std::shared_ptr<Spreadsheet> spreadsheet);
    bool deleteDocument(const QString& id);

    // Sheet operations
    bool addSheet(const QString& documentId, const QString& sheetName, int index);
    bool removeSheet(const QString& documentId, int index);

    // Save/Load operations
    bool saveDocument(const QString& id);
    bool loadDocument(const QString& id);

    // Version control
    bool saveVersion(const QString& documentId);
    QVector<std::shared_ptr<Document>> getVersionHistory(const QString& documentId);
    bool restoreVersion(const QString& documentId, const QString& versionId);

    QString getLastError() const;

private:
    DocumentRepository();
    ~DocumentRepository() = default;

    QString m_lastError;

    QJsonObject serializeSpreadsheet(const std::shared_ptr<Spreadsheet>& spreadsheet);
    std::shared_ptr<Spreadsheet> deserializeSpreadsheet(const QJsonObject& json);
};

#endif // DOCUMENTREPOSITORY_H
