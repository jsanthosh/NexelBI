#include <QApplication>
#include <QStandardPaths>
#include <QDir>
#include "ui/MainWindow.h"
#include "database/DatabaseManager.h"
#include "services/DocumentService.h"

int main(int argc, char* argv[]) {
    QApplication app(argc, argv);

    app.setApplicationName("NativeSpreadsheet");
    app.setApplicationVersion("1.0.0");
    app.setApplicationDisplayName("Native Spreadsheet - Excel Alternative");

    // Initialize database
    QString appDataPath = QStandardPaths::writableLocation(QStandardPaths::AppDataLocation);
    QDir().mkpath(appDataPath);
    QString dbPath = appDataPath + "/documents.db";

    if (!DatabaseManager::instance().initialize(dbPath)) {
        qWarning() << "Failed to initialize database:" << DatabaseManager::instance().getLastError();
        return 1;
    }

    // Create main window
    MainWindow window;
    window.show();

    return app.exec();
}
