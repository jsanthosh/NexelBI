#!/bin/bash

# Native Spreadsheet Build Script
# Supports macOS, Linux, and Windows (with MSVC)

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BUILD_DIR="${SCRIPT_DIR}/build"
INSTALL_PREFIX="${SCRIPT_DIR}/install"

# Detect OS
if [[ "$OSTYPE" == "darwin"* ]]; then
    OS="macOS"
elif [[ "$OSTYPE" == "linux-gnu"* ]]; then
    OS="Linux"
elif [[ "$OSTYPE" == "msys" || "$OSTYPE" == "cygwin" ]]; then
    OS="Windows"
else
    OS="Unknown"
fi

echo "===== Native Spreadsheet Build Script ====="
echo "OS: $OS"
echo "Build Directory: $BUILD_DIR"

# Parse arguments
BUILD_TYPE="Release"
CLEAN_BUILD=false
RUN_TESTS=false

while [[ $# -gt 0 ]]; do
    case $1 in
        --debug)
            BUILD_TYPE="Debug"
            shift
            ;;
        --clean)
            CLEAN_BUILD=true
            shift
            ;;
        --test)
            RUN_TESTS=true
            shift
            ;;
        *)
            echo "Unknown option: $1"
            exit 1
            ;;
    esac
done

# Clean if requested
if [ "$CLEAN_BUILD" = true ]; then
    echo "Cleaning build directory..."
    rm -rf "$BUILD_DIR"
fi

# Create build directory
mkdir -p "$BUILD_DIR"
cd "$BUILD_DIR"

# Configure
echo "Configuring CMake..."
if [ "$OS" = "macOS" ]; then
    cmake -DCMAKE_BUILD_TYPE=$BUILD_TYPE \
          -DCMAKE_OSX_ARCHITECTURES="arm64;x86_64" \
          -DCMAKE_INSTALL_PREFIX="$INSTALL_PREFIX" \
          ..
else
    cmake -DCMAKE_BUILD_TYPE=$BUILD_TYPE \
          -DCMAKE_INSTALL_PREFIX="$INSTALL_PREFIX" \
          ..
fi

# Build
echo "Building..."
if [ "$OS" = "Windows" ]; then
    cmake --build . --config $BUILD_TYPE -j$(nproc)
else
    make -j$(nproc)
fi

# Install
echo "Installing..."
if [ "$OS" = "Windows" ]; then
    cmake --install . --config $BUILD_TYPE
else
    make install
fi

# Run tests if requested
if [ "$RUN_TESTS" = true ]; then
    echo "Running tests..."
    ctest -C $BUILD_TYPE --output-on-failure
fi

echo ""
echo "===== Build Complete ====="
echo "Build Type: $BUILD_TYPE"

if [ "$OS" = "macOS" ]; then
    echo "App Location: $INSTALL_PREFIX/NativeSpreadsheet.app/Contents/MacOS/NativeSpreadsheet"
    echo ""
    echo "To run the application:"
    echo "  open $INSTALL_PREFIX/NativeSpreadsheet.app"
else
    echo "App Location: $INSTALL_PREFIX/bin/NativeSpreadsheet"
    echo ""
    echo "To run the application:"
    echo "  $INSTALL_PREFIX/bin/NativeSpreadsheet"
fi

echo ""
echo "Build time: $(date)"
