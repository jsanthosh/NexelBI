#include "DatabaseManager.h"
#include <QFileInfo>
#include <iostream>

DatabaseManager::DatabaseManager() : m_db(nullptr), m_initialized(false) {
}

DatabaseManager::~DatabaseManager() {
    close();
}

DatabaseManager& DatabaseManager::instance() {
    static DatabaseManager s_instance;
    return s_instance;
}

bool DatabaseManager::initialize(const QString& dbPath) {
    if (m_initialized) {
        return true;
    }

    int rc = sqlite3_open(dbPath.toStdString().c_str(), &m_db);
    if (rc != SQLITE_OK) {
        m_lastError = "Cannot open database: " + QString::fromStdString(std::string(sqlite3_errmsg(m_db)));
        return false;
    }

    // Enable foreign keys
    sqlite3_exec(m_db, "PRAGMA foreign_keys = ON", nullptr, nullptr, nullptr);

    // Set journal mode to WAL for better concurrency
    sqlite3_exec(m_db, "PRAGMA journal_mode = WAL", nullptr, nullptr, nullptr);

    // Enable memory-mapped I/O for faster access
    sqlite3_exec(m_db, "PRAGMA mmap_size = 30000000", nullptr, nullptr, nullptr);

    // Increase cache size
    sqlite3_exec(m_db, "PRAGMA cache_size = 10000", nullptr, nullptr, nullptr);

    createTables();
    m_initialized = true;
    return true;
}

bool DatabaseManager::isInitialized() const {
    return m_initialized;
}

void DatabaseManager::close() {
    if (m_db) {
        sqlite3_close(m_db);
        m_db = nullptr;
        m_initialized = false;
    }
}

sqlite3* DatabaseManager::getDatabase() const {
    return m_db;
}

bool DatabaseManager::beginTransaction() {
    return sqlite3_exec(m_db, "BEGIN TRANSACTION", nullptr, nullptr, nullptr) == SQLITE_OK;
}

bool DatabaseManager::commit() {
    return sqlite3_exec(m_db, "COMMIT", nullptr, nullptr, nullptr) == SQLITE_OK;
}

bool DatabaseManager::rollback() {
    return sqlite3_exec(m_db, "ROLLBACK", nullptr, nullptr, nullptr) == SQLITE_OK;
}

QString DatabaseManager::getLastError() const {
    return m_lastError;
}

int DatabaseManager::getChangesCount() const {
    return sqlite3_changes(m_db);
}

void DatabaseManager::createTables() {
    const char* sql = R"(
        CREATE TABLE IF NOT EXISTS documents (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            createdAt DATETIME DEFAULT CURRENT_TIMESTAMP,
            updatedAt DATETIME DEFAULT CURRENT_TIMESTAMP,
            content BLOB NOT NULL
        );

        CREATE TABLE IF NOT EXISTS sheets (
            id TEXT PRIMARY KEY,
            documentId TEXT NOT NULL,
            name TEXT NOT NULL,
            index INTEGER NOT NULL,
            FOREIGN KEY(documentId) REFERENCES documents(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS cells (
            id TEXT PRIMARY KEY,
            sheetId TEXT NOT NULL,
            row INTEGER NOT NULL,
            col INTEGER NOT NULL,
            type TEXT NOT NULL,
            value TEXT,
            formula TEXT,
            FOREIGN KEY(sheetId) REFERENCES sheets(id) ON DELETE CASCADE,
            UNIQUE(sheetId, row, col)
        );

        CREATE TABLE IF NOT EXISTS cellStyles (
            id TEXT PRIMARY KEY,
            cellId TEXT NOT NULL UNIQUE,
            fontName TEXT,
            fontSize INTEGER,
            bold INTEGER,
            italic INTEGER,
            underline INTEGER,
            foregroundColor TEXT,
            backgroundColor TEXT,
            hAlign TEXT,
            vAlign TEXT,
            numberFormat TEXT,
            FOREIGN KEY(cellId) REFERENCES cells(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS versions (
            id TEXT PRIMARY KEY,
            documentId TEXT NOT NULL,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
            content BLOB NOT NULL,
            FOREIGN KEY(documentId) REFERENCES documents(id) ON DELETE CASCADE
        );

        CREATE INDEX IF NOT EXISTS idx_sheets_documentId ON sheets(documentId);
        CREATE INDEX IF NOT EXISTS idx_cells_sheetId ON cells(sheetId);
        CREATE INDEX IF NOT EXISTS idx_cells_position ON cells(sheetId, row, col);
        CREATE INDEX IF NOT EXISTS idx_versions_documentId ON versions(documentId);
    )";

    char* errMsg = nullptr;
    int rc = sqlite3_exec(m_db, sql, nullptr, nullptr, &errMsg);
    if (rc != SQLITE_OK) {
        m_lastError = "SQL error: " + QString::fromUtf8(errMsg);
        sqlite3_free(errMsg);
    }
}
