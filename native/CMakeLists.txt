cmake_minimum_required(VERSION 3.24)
project(NativeSpreadsheet
    VERSION 1.0.0
    LANGUAGES CXX
    HOMEPAGE_URL "https://spreadsheet.local"
    DESCRIPTION "High-performance native spreadsheet application")

# C++ Standard
set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -fPIC -Wall -Wextra -O3")

# Force native architecture (avoid universal binary issues with Homebrew Qt)
if(APPLE)
    set(CMAKE_OSX_ARCHITECTURES "${CMAKE_SYSTEM_PROCESSOR}")
endif()

# Threading support
find_package(Threads REQUIRED)

# Qt Configuration
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_AUTOUIC ON)

find_package(Qt6 COMPONENTS
    Core
    Gui
    Widgets
    Sql
    Network
    Concurrent
    REQUIRED
)
find_package(Qt6CorePrivate QUIET)

find_package(SQLite3 REQUIRED)

# Core Engine Sources
set(CORE_SOURCES
    src/core/Cell.cpp
    src/core/Cell.h
    src/core/Spreadsheet.cpp
    src/core/Spreadsheet.h
    src/core/FormulaEngine.cpp
    src/core/FormulaEngine.h
    src/core/CellRange.cpp
    src/core/CellRange.h
    src/core/ConditionalFormatting.cpp
    src/core/ConditionalFormatting.h
    src/core/UndoManager.cpp
    src/core/UndoManager.h
    src/core/DependencyGraph.cpp
    src/core/DependencyGraph.h
    src/core/NumberFormat.cpp
    src/core/NumberFormat.h
    src/core/FillSeries.cpp
    src/core/FillSeries.h
    src/core/TableStyle.h
)

# Database Sources
set(DATABASE_SOURCES
    src/database/DatabaseManager.cpp
    src/database/DatabaseManager.h
    src/database/DocumentRepository.cpp
    src/database/DocumentRepository.h
)

# Services Sources
set(SERVICES_SOURCES
    src/services/ClaudeService.cpp
    src/services/ClaudeService.h
    src/services/DocumentService.cpp
    src/services/DocumentService.h
    src/services/CsvService.cpp
    src/services/CsvService.h
    src/services/XlsxService.cpp
    src/services/XlsxService.h
)

# UI Sources
set(UI_SOURCES
    src/ui/MainWindow.cpp
    src/ui/MainWindow.h
    src/ui/SpreadsheetView.cpp
    src/ui/SpreadsheetView.h
    src/ui/CellDelegate.cpp
    src/ui/CellDelegate.h
    src/ui/SpreadsheetModel.cpp
    src/ui/SpreadsheetModel.h
    src/ui/FormulaBar.cpp
    src/ui/FormulaBar.h
    src/ui/Toolbar.cpp
    src/ui/Toolbar.h
    src/ui/FormatCellsDialog.cpp
    src/ui/FormatCellsDialog.h
    src/ui/FindReplaceDialog.cpp
    src/ui/FindReplaceDialog.h
    src/ui/GoToDialog.cpp
    src/ui/GoToDialog.h
    src/ui/ConditionalFormatDialog.cpp
    src/ui/ConditionalFormatDialog.h
    src/ui/DataValidationDialog.cpp
    src/ui/DataValidationDialog.h
)

# Main Application
set(APP_SOURCES
    src/main.cpp
)

# Create executable
add_executable(NativeSpreadsheet
    ${APP_SOURCES}
    ${CORE_SOURCES}
    ${DATABASE_SOURCES}
    ${SERVICES_SOURCES}
    ${UI_SOURCES}
)

# Link libraries
target_link_libraries(NativeSpreadsheet
    Qt6::Core
    Qt6::CorePrivate
    Qt6::Gui
    Qt6::Widgets
    Qt6::Sql
    Qt6::Network
    Qt6::Concurrent
    SQLite::SQLite3
    Threads::Threads
)

# Include directories
target_include_directories(NativeSpreadsheet PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/src/core
    ${CMAKE_CURRENT_SOURCE_DIR}/src/database
    ${CMAKE_CURRENT_SOURCE_DIR}/src/services
    ${CMAKE_CURRENT_SOURCE_DIR}/src/ui
)

# macOS specific settings
if(APPLE)
    set_target_properties(NativeSpreadsheet PROPERTIES
        MACOSX_BUNDLE TRUE
        MACOSX_BUNDLE_GUI_IDENTIFIER "com.spreadsheet.native"
        MACOSX_BUNDLE_BUNDLE_NAME "NativeSpreadsheet"
        MACOSX_BUNDLE_BUNDLE_VERSION "1.0.0"
        MACOSX_BUNDLE_SHORT_VERSION_STRING "1.0"
    )
endif()

# Windows specific settings
if(WIN32)
    set_target_properties(NativeSpreadsheet PROPERTIES
        WIN32_EXECUTABLE TRUE
    )
endif()

# Install
install(TARGETS NativeSpreadsheet
    BUNDLE DESTINATION . COMPONENT Runtime
    RUNTIME DESTINATION bin COMPONENT Runtime
)
