# Quick Start Guide - Native Spreadsheet

## 60-Second Setup

### 1. Install Dependencies

**macOS (Homebrew):**
```bash
brew install qt6 sqlite cmake
```

**Ubuntu:**
```bash
sudo apt-get install -y qt6-base-dev libsqlite3-dev cmake build-essential
```

### 2. Build

```bash
cd native
chmod +x build.sh
./build.sh --debug --clean
```

### 3. Run

**macOS:**
```bash
open install/NativeSpreadsheet.app
```

**Linux/Windows:**
```bash
./install/bin/NativeSpreadsheet
```

## Understanding the Codebase

### Start Here

1. **[ARCHITECTURE.md](ARCHITECTURE.md)** - System design
2. **[README.md](README.md)** - Full documentation
3. **Source code structure:**

```
src/
├── core/
│   ├── Spreadsheet.h         ← Main API
│   ├── Cell.h                ← Cell data
│   ├── FormulaEngine.h       ← Formula evaluation
│   └── CellRange.h           ← Range operations
│
├── database/
│   ├── DatabaseManager.h     ← SQLite setup
│   └── DocumentRepository.h  ← CRUD operations
│
├── services/
│   ├── DocumentService.h     ← Document management
│   └── ClaudeService.h       ← AI integration
│
└── ui/
    ├── MainWindow.h          ← Main window
    ├── SpreadsheetView.h     ← Grid display
    └── ...other UI components
```

## Building from Scratch

### Debug Build
```bash
cd native
mkdir build-debug
cd build-debug
cmake -DCMAKE_BUILD_TYPE=Debug ..
make -j4
./NativeSpreadsheet  # Run directly
```

### Release Build
```bash
cd native
mkdir build-release
cd build-release
cmake -DCMAKE_BUILD_TYPE=Release ..
make -j4
```

### Clean Build
```bash
cd native
rm -rf build-*
./build.sh --clean --debug
```

## Key Concepts

### The Core Engine

Everything lives in `src/core/`. No Qt dependencies here:

```cpp
// Create a spreadsheet
Spreadsheet sheet;

// Set cell values
sheet.setCellValue(CellAddress(0, 0), 42);      // Cell A1 = 42
sheet.setCellValue(CellAddress(0, 1), 8);       // Cell B1 = 8

// Use formulas
sheet.setCellFormula(CellAddress(0, 2), "=A1+B1");  // Cell C1 = A1+B1

// Get results
QVariant result = sheet.getCellValue(CellAddress(0, 2));  // Result: 50
```

### The UI Layer

Connects the engine to Qt:

```cpp
// In MainWindow::MainWindow()
m_spreadsheetView = new SpreadsheetView(this);
m_spreadsheetView->setSpreadsheet(m_spreadsheet);

// User edits in formula bar
connect(m_formulaBar, &FormulaBar::contentEdited,
        this, [this](const QString& content) {
    // Update spreadsheet
    auto addr = m_currentCell;
    if (content.startsWith("=")) {
        m_spreadsheet->setCellFormula(addr, content);
    } else {
        m_spreadsheet->setCellValue(addr, content);
    }
});
```

### Data Persistence

SQLite database stores everything:

```cpp
// Save a document
DocumentRepository::instance().createDocument("MySheet", spreadsheet);

// Load it back
auto doc = DocumentRepository::instance().getDocument(docId);
```

## Testing

### Run a Simple Test

Add this to `src/main.cpp`:

```cpp
#include "core/Spreadsheet.h"

int main(int argc, char* argv[]) {
    // Test core engine
    Spreadsheet sheet;
    sheet.setCellValue(CellAddress(0, 0), 10);
    sheet.setCellValue(CellAddress(0, 1), 20);
    sheet.setCellFormula(CellAddress(0, 2), "=A1+B1");
    
    QVariant result = sheet.getCellValue(CellAddress(0, 2));
    qDebug() << "Result: " << result.toDouble();  // Should print 30
    
    // Continue with Qt app...
    QApplication app(argc, argv);
    MainWindow window;
    window.show();
    return app.exec();
}
```

## Performance Tips

### For Large Spreadsheets (10K+ cells)

1. **Disable auto-recalculate** during batch operations:
```cpp
sheet.setAutoRecalculate(false);
// ... many updates ...
sheet.setAutoRecalculate(true);
```

2. **Use transactions**:
```cpp
sheet.startTransaction();
for (auto& cell : largeBatch) {
    sheet.setCellValue(cell.addr, cell.value);
}
sheet.commitTransaction();
```

3. **Batch database operations**:
```cpp
db.beginTransaction();
for (int i = 0; i < 10000; ++i) {
    db.insert(cell);
}
db.commit();
```

## Debugging Tips

### Print Cell Values
```cpp
auto cell = sheet.getCell(CellAddress(0, 0));
qDebug() << "Value:" << cell->getValue();
qDebug() << "Formula:" << cell->getFormula();
qDebug() << "Type:" << (int)cell->getType();
```

### Check Formula Errors
```cpp
QVariant result = engine.evaluate("=SUM(A1:B10)");
if (engine.hasError()) {
    qWarning() << "Formula error:" << engine.getLastError();
}
```

### Database Debugging
```cpp
qDebug() << "DB Path:" << dbPath;
qDebug() << "DB initialized:" << DatabaseManager::instance().isInitialized();
qDebug() << "Last error:" << DatabaseManager::instance().getLastError();
```

## Common Issues & Solutions

### Issue: Qt6 not found
**Solution:**
```bash
# Find Qt installation
brew --cellar qt6
# Tell CMake where Qt is
export Qt6_DIR=/usr/local/opt/qt6/lib/cmake/Qt6
```

### Issue: Cannot find sqlite3.h
**Solution:**
```bash
# macOS
brew install sqlite
# Ubuntu
sudo apt-get install libsqlite3-dev
```

### Issue: Build fails with many errors
**Solution:**
```bash
# Clean everything and rebuild
cd native
rm -rf build-* install
./build.sh --debug --clean
```

### Issue: Application won't start
**Solution:**
```bash
# Check for runtime errors
lldb ./build/NativeSpreadsheet
(lldb) run
# Check console output for errors
```

## Next Steps

1. **Get it building** - Follow build instructions
2. **Test the core** - Add test code in main.cpp
3. **Connect the UI** - Update signal/slot connections
4. **Add features** - Implement formulas, I/O, etc.
5. **Optimize** - Profile and improve performance

## Resources

- **Qt Documentation**: https://doc.qt.io/
- **SQLite Guide**: https://www.sqlite.org/
- **CMake Docs**: https://cmake.org/
- **C++20 Reference**: https://cppreference.com/

## Need Help?

1. Check **ARCHITECTURE.md** for design overview
2. Review **README.md** for detailed documentation
3. Look at similar code in the codebase
4. Check Qt documentation for UI issues
5. Use debugger (lldb on macOS, gdb on Linux)

---

**Happy developing! 🚀**
