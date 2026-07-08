#!/bin/bash

echo ""
echo "==================================================="
echo "OneNote MCP Server: Bootstrap Setup (POSIX)"
echo "==================================================="
echo ""

# Check Node.js
if ! command -v node &> /dev/null; then
    echo "[ERROR] Node.js is not installed or not in your PATH."
    echo ""
    echo "Please install Node.js (v18 or higher recommended):"
    echo "  - macOS (Homebrew): brew install node"
    echo "  - Linux (Ubuntu/Debian): sudo apt install nodejs npm"
    echo "  - Website: https://nodejs.org/"
    echo ""
    echo "After installing, restart your terminal and run this script again."
    exit 1
fi

# Check NPM
if ! command -v npm &> /dev/null; then
    echo "[ERROR] npm is not installed or not in your PATH."
    echo "Please ensure npm is installed alongside Node.js."
    exit 1
fi

# Delegate to setup.js
echo "[INFO] Node.js and npm found. Starting main setup..."
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
node "$SCRIPT_DIR/setup.js"
