# Native Spreadsheet Application

This is a high-performance native spreadsheet application for macOS, Windows, and Linux, built with C++20 and Qt6.

## Features

- **Native Performance**: Built with C++ for native-speed spreadsheet operations
- **Cross-Platform**: Works on macOS, Windows, and Linux with native UIs
- **Excel Compatible**: Supports XLSX, XLS, CSV import/export
- **AI-Powered**: Integrated Claude AI for formula suggestions and data analysis
- **Real-Time Collaboration**: Optional real-time sync with local/cloud database
- **Advanced Formulas**: Comprehensive formula engine with 100+ functions
- **Conditional Formatting**: Visual data analysis with conditional styles
- **Pivot Tables**: Built-in pivot table support
- **Version Control**: Automatic version history and recovery

## System Requirements

### macOS
- macOS 10.15+
- Intel or Apple Silicon (M1/M2/etc.)
- 200 MB disk space

### Windows
- Windows 10 or newer
- 200 MB disk space

### Linux
- Ubuntu 20.04+ or equivalent
- Qt6 libraries
- 200 MB disk space

## Building

### macOS

```bash
cd native
mkdir build
cd build
cmake -DCMAKE_BUILD_TYPE=Release ..
make
make install
```

The application bundle will be created in `build/NativeSpreadsheet.app`

### Windows

```bash
cd native
mkdir build
cd build
cmake -DCMAKE_BUILD_TYPE=Release ..
cmake --build . --config Release
cmake --install . --config Release
```

### Linux

```bash
cd native
mkdir build
cd build
cmake -DCMAKE_BUILD_TYPE=Release ..
make
sudo make install
```

## Dependencies

- Qt 6.5+
- SQLite3
- CMake 3.24+
- C++20 compatible compiler (GCC 11+, Clang 13+, MSVC 2022+)

### macOS Installation (Homebrew)

```bash
brew install qt6 sqlite cmake
```

### Ubuntu Installation

```bash
sudo apt-get install -y qt6-base-dev libsqlite3-dev cmake build-essential
```

### Windows Installation (vcpkg)

```bash
vcpkg install qt6:x64-windows sqlite3:x64-windows
```

## Project Structure

```
native/
├── CMakeLists.txt              # Build configuration
├── src/
│   ├── core/                   # Core spreadsheet engine
│   │   ├── Cell.h/cpp          # Cell data structure
│   │   ├── Spreadsheet.h/cpp   # Spreadsheet container
│   │   ├── FormulaEngine.h/cpp # Formula evaluator
│   │   ├── CellRange.h/cpp     # Range operations
│   │   └── ConditionalFormatting.h/cpp
│   ├── database/               # Data persistence
│   │   ├── DatabaseManager.h/cpp
│   │   └── DocumentRepository.h/cpp
│   ├── services/               # Business logic
│   │   ├── ClaudeService.h/cpp # AI integration
│   │   └── DocumentService.h/cpp
│   ├── ui/                     # Qt UI components
│   │   ├── MainWindow.h/cpp
│   │   ├── SpreadsheetView.h/cpp
│   │   ├── SpreadsheetModel.h/cpp
│   │   ├── CellDelegate.h/cpp
│   │   ├── FormulaBar.h/cpp
│   │   └── Toolbar.h/cpp
│   └── main.cpp                # Application entry point
└── resources/                  # Icons and resources
```

## Performance Optimizations

1. **Lazy Cell Loading**: Cells are created on-demand, not upfront
2. **Efficient Viewport Rendering**: Only visible cells are rendered
3. **Formula Caching**: Formula results are cached until dependent cells change
4. **Database Indexes**: Optimized SQLite queries with proper indexing
5. **Memory-Mapped I/O**: SQLite uses memory-mapped I/O for faster access
6. **WAL Mode**: Write-Ahead Logging for better concurrency
7. **Native Rendering**: Direct use of platform-specific graphics APIs via Qt

## Architecture

### Three-Layer Architecture

1. **Core Engine** (`src/core/`)
   - Pure C++ spreadsheet logic
   - No UI dependencies
   - High-performance formula evaluation
   - Cell management and recalculation

2. **Services** (`src/services/`)
   - Business logic
   - Document management
   - AI integration
   - Import/Export operations

3. **UI Layer** (`src/ui/`)
   - Qt-based native interfaces
   - Responsive grid rendering
   - User interactions
   - Theme and styling

### Data Flow

```
User Input → UI Layer → Services → Core Engine ↔ Database
                                  ↔ Claude API
```

## Building for macOS

### Development Build

```bash
cd native
mkdir build-debug
cd build-debug
cmake -DCMAKE_BUILD_TYPE=Debug -DCMAKE_OSX_ARCHITECTURES="arm64;x86_64" ..
make
./NativeSpreadsheet
```

### Release Build (Universal Binary)

```bash
cd native
mkdir build-release
cd build-release
cmake -DCMAKE_BUILD_TYPE=Release -DCMAKE_OSX_ARCHITECTURES="arm64;x86_64" ..
make
make install
```

### Creating DMG Distribution

```bash
cd native/build-release
cpack -G DragNDrop
```

## Development

### Code Style
- Follow C++20 modern conventions
- Use STL containers (no raw pointers where possible)
- Use Qt container types (QString, QVector, etc.)
- 4-space indentation

### Testing

```bash
cd native/build
ctest
```

### Debugging

macOS:
```bash
lldb ./build-debug/NativeSpreadsheet
```

Linux:
```bash
gdb ./build/NativeSpreadsheet
```

## License

MIT License - See LICENSE file for details

## Contributing

1. Create a feature branch
2. Follow code style guidelines
3. Add tests for new features
4. Submit pull request

## Support

For issues and feature requests, visit the GitHub issues page.
