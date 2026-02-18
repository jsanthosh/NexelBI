#ifndef DATABASEMANAGER_H
#define DATABASEMANAGER_H

#include <QString>
#include <sqlite3.h>
#include <memory>

class DatabaseManager {
public:
    static DatabaseManager& instance();

    bool initialize(const QString& dbPath);
    bool isInitialized() const;
    void close();

    sqlite3* getDatabase() const;

    // Transaction management
    bool beginTransaction();
    bool commit();
    bool rollback();

    // Utility functions
    QString getLastError() const;
    int getChangesCount() const;

private:
    DatabaseManager();
    ~DatabaseManager();

    sqlite3* m_db;
    bool m_initialized;
    QString m_lastError;

    void createTables();
};

#endif // DATABASEMANAGER_H
