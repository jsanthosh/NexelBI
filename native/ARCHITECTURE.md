# Native Spreadsheet - Architecture & Development Guide

## Overview

NativeSpreadsheet is being rebuilt as a true native application using C++20 and Qt6. This document outlines the architecture, development workflow, and next steps.

## Architecture

### Core Components

#### 1. **Spreadsheet Engine** (`src/core/`)

The core spreadsheet logic, completely independent of UI:

- **Cell.h/cpp**: Individual cell representation
  - Value storage
  - Formula support
  - Computed values
  - Styling information
  - Error tracking

- **Spreadsheet.h/cpp**: Sheet container
  - Cell management
  - Row/Column operations
  - Auto-recalculation
  - Dirty tracking
  - Transaction support

- **FormulaEngine.h/cpp**: Expression evaluator
  - Recursive descent parser
  - 50+ built-in functions (SUM, AVG, IF, etc.)
  - Cell reference resolution
  - Range operations
  - Error handling

- **CellRange.h/cpp**: Range selection
  - A1:B10 parsing
  - Range operations (intersection, contains)
  - Iteration support

- **ConditionalFormatting.h/cpp**: Styling rules
  - Rule management
  - Style application
  - Condition evaluation

**Key Performance Features:**
- Lazy cell creation (on-demand allocation)
- Formula caching with dependency tracking
- No UI dependencies - can be tested independently
- Thread-safe operations (with locking where needed)

#### 2. **Data Persistence** (`src/database/`)

SQLite-based local storage:

- **DatabaseManager.h/cpp**: Connection pooling
  - Singleton pattern
  - WAL mode for concurrency
  - Memory-mapped I/O
  - Indexed queries
  - Transaction management

- **DocumentRepository.h/cpp**: CRUD operations
  - Document storage
  - Sheet management
  - Cell persistence
  - Version history
  - Serialization/Deserialization

**Database Schema:**
```sql
documents
  ├── id (PRIMARY KEY)
  ├── name
  ├── createdAt
  ├── updatedAt
  └── content (BLOB - JSON)

sheets
  ├── id (PRIMARY KEY)
  ├── documentId (FOREIGN KEY)
  ├── name
  └── index

cells
  ├── id (PRIMARY KEY)
  ├── sheetId (FOREIGN KEY)
  ├── row, col
  ├── type
  ├── value
  └── formula

cellStyles
  └── [Style metadata]

versions
  └── [Version history]
```

#### 3. **Services Layer** (`src/services/`)

Business logic and integrations:

- **DocumentService.h/cpp**: Document lifecycle
  - Create/Open/Save operations
  - Currently open document tracking
  - Import/Export coordination

- **ClaudeService.h/cpp**: AI integration
  - Formula suggestions
  - Data analysis
  - Cell content recommendations
  - Error hints

**Future Services:**
- SyncService (Cloud sync)
- ExportService (XLSX, PDF)
- ImportService (CSV, Excel)
- CollaborationService (Real-time sync)

#### 4. **UI Layer** (`src/ui/`)

Qt-based native interfaces:

- **MainWindow.h/cpp**: Application window
  - Menu bar
  - UI coordination
  - Status bar

- **SpreadsheetView.h/cpp**: Grid display
  - QTableView subclass
  - Viewport-based rendering
  - Clipboard operations
  - Zoom support

- **SpreadsheetModel.h/cpp**: Data binding
  - Qt Model/View architecture
  - Cell data exposure
  - Style integration
  - Update propagation

- **CellDelegate.h/cpp**: Cell rendering
  - Custom paint
  - Edit mode
  - Formatting application

- **FormulaBar.h/cpp**: Formula input
  - Cell address display
  - Content editing
  - Formula suggestions

- **Toolbar.h/cpp**: Command buttons
  - File operations
  - Editing tools
  - Formatting controls

## Development Workflow

### Setting Up Development Environment

**macOS:**
```bash
# Install dependencies via Homebrew
brew install qt6 sqlite cmake lldb

# Clone and setup
git clone <repo>
cd NativeSpreadsheet/native

# Build with debug symbols
./build.sh --debug --clean

# Run
open install/NativeSpreadsheet.app
```

**Linux (Ubuntu 22.04):**
```bash
sudo apt-get install -y build-essential git cmake
sudo apt-get install -y qt6-base-dev libsqlite3-dev

cd NativeSpreadsheet/native
./build.sh --debug --clean

./install/bin/NativeSpreadsheet
```

### Debugging

**macOS with LLDB:**
```bash
lldb ./build/NativeSpreadsheet

(lldb) run
(lldb) bt          # backtrace
(lldb) p variable  # print value
(lldb) c           # continue
```

**Linux with GDB:**
```bash
gdb ./build/NativeSpreadsheet

(gdb) run
(gdb) bt
(gdb) p variable
(gdb) c
```

### Code Organization Rules

1. **Header Dependencies**
   - Forward declare when possible
   - Minimize circular includes
   - Use include guards

2. **Memory Management**
   - Use std::shared_ptr for shared ownership
   - Use std::unique_ptr for exclusive ownership
   - Avoid raw new/delete

3. **Qt Conventions**
   - Use Qt containers (QString, QVector)
   - Use Q_OBJECT macro for QObject subclasses
   - Connect signals/slots properly

4. **Error Handling**
   - Check return values
   - Use exceptions sparingly
   - Provide clear error messages

## Build System

### CMake Structure

```cmake
CMakeLists.txt
├── Qt Configuration
├── Compiler Flags (Performance + Debug)
├── Source Groups
│   ├── Core
│   ├── Database
│   ├── Services
│   └── UI
├── Executable Definition
├── Linking
├── macOS Bundle Settings
├── Windows Settings
└── Installation Rules
```

### Build Targets

```bash
cmake --target NativeSpreadsheet    # Main app
cmake --target install              # Installation
cmake --target clean                # Clean
ctest                              # Tests (when added)
```

## Performance Optimization Strategy

### Phase 1: Core Engine (Current)
- ✅ Lazy cell loading
- ✅ Formula caching
- ✅ Efficient data structures

### Phase 2: Rendering (Next)
- Viewport-based rendering (visible cells only)
- Cell pooling and reuse
- Dirty region tracking
- Incremental updates

### Phase 3: Database (Phase 3)
- Index optimization
- Query caching
- Batch operations
- Connection pooling

### Phase 4: Compilation (Phase 4)
- Release build optimization
- Link-time optimization (LTO)
- Profile-guided optimization (PGO)

### Phase 5: Platform-Specific (Phase 5)
- Native graphics APIs
- CoreGraphics (macOS)
- Direct2D (Windows)
- OpenGL (Linux)

## Implementation Roadmap

### Milestone 1: Foundation (Current) ✅
- [x] Project structure
- [x] Core spreadsheet engine
- [x] Database layer (schema + basic ops)
- [x] UI skeleton
- [x] Application entry point
- [ ] Complete Database Repository methods

### Milestone 2: Core Functionality (Next)
- [ ] Formula evaluation (full testing)
- [ ] File I/O (CSV, XLSX)
- [ ] Styling and formatting
- [ ] Clipboard operations
- [ ] Undo/Redo stack

### Milestone 3: Polish & Performance
- [ ] Grid rendering optimization
- [ ] Large file support (100K+ cells)
- [ ] Memory profiling
- [ ] Performance benchmarks

### Milestone 4: Advanced Features
- [ ] Pivot tables
- [ ] Charts and graphs
- [ ] AI integrations (Claude)
- [ ] Real-time collaboration

### Milestone 5: Distribution
- [ ] macOS app bundling
- [ ] Code signing and notarization
- [ ] Windows installer
- [ ] Linux AppImage

## Cross-Platform Considerations

### macOS
- Use native event system
- macOS app bundle structure
- Code signing with Apple Developer ID
- Universal Binary (arm64 + x86_64)

### Windows
- Use WinAPI where needed
- MSVC compiler optimization
- Windows installer (.msi)
- High DPI support

### Linux
- Use X11/Wayland
- AppImage distribution
- GTK integration (optional)
- Package manager support (RPM, DEB)

## Database Performance

### Optimization Techniques

1. **WAL Mode** (Write-Ahead Logging)
   - Better concurrency
   - Faster writes

2. **Memory-Mapped I/O**
   ```cpp
   PRAGMA mmap_size = 30000000;  // 30MB mapping
   ```

3. **Index Usage**
   ```sql
   CREATE INDEX idx_cells_position ON cells(sheetId, row, col);
   ```

4. **Batch Operations**
   ```cpp
   db->beginTransaction();
   for (const auto& cell : cells) {
       insert(cell);
   }
   db->commit();
   ```

## Testing Strategy

### Unit Tests (Future)
```cpp
// tests/test_formula_engine.cpp
TEST(FormulaEngine, SUM_Function) {
    FormulaEngine engine;
    QVariant result = engine.evaluate("=SUM(1,2,3)");
    EXPECT_EQ(result.toDouble(), 6.0);
}
```

### Integration Tests (Future)
- File I/O operations
- Database persistence
- UI interactions

### Performance Tests (Future)
- Large spreadsheet loading
- Formula recalculation speed
- Memory usage profiles

## Next Steps

1. **Implement remaining database methods**
   - Complete serialize/deserialize
   - Fix compilation errors

2. **Test compilation**
   ```bash
   cd native && ./build.sh --debug
   ```

3. **Fix any compilation errors**
   - Add missing #includes
   - Fix type mismatches

4. **Add UI event connections**
   - Cell selection
   - Formula editing
   - File operations

5. **Implement file I/O**
   - CSV parser
   - XLSX support
   - JSON persistence

## Architecture Diagram

```
┌─────────────────────────────────────────┐
│         Qt Application                   │
│  MainWindow                              │
└──────────────┬──────────────────────────┘
               │
         ┌─────┴─────┐
         │            │
    ┌────▼────┐  ┌───▼────┐
    │ Toolbar │  │Formula  │
    │         │  │ Bar     │
    └─────────┘  └────┬────┘
                      │
         ┌───────────┬┴─────────┐
         │           │          │
    ┌────▼──┐  ┌────▼──┐  ┌────▼───┐
    │ View  │  │ Model │  │Delegate│
    │(Table)│  │(Data) │  │(Paint) │
    └───┬───┘  └────┬──┘  └────────┘
        │           │
        └─────┬─────┘
              │
        ┌─────▼──────────┐
        │ Services Layer │
        ├────────────────┤
        │ Doc Service    │
        │ Claude Service │
        └────────┬───────┘
                 │
        ┌────────▼──────────────┐
        │   Core Engine          │
        ├────────────────────────┤
        │ Spreadsheet            │
        │ ├─ Cell               │
        │ ├─ FormulaEngine      │
        │ └─ CellRange          │
        └────────┬───────────────┘
                 │
        ┌────────▼──────────────┐
        │   Data Persistence    │
        ├────────────────────────┤
        │ DatabaseManager        │
        │ DocumentRepository     │
        │ SQLite Database        │
        └────────────────────────┘
```

## Contact & Contributing

For questions or to contribute:
1. Review this architecture document
2. Follow coding standards
3. Test thoroughly
4. Submit pull requests with clear descriptions
