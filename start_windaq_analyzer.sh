#!/bin/bash

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Function to print colored output
print_status() {
    echo -e "${GREEN}✓${NC} $1"
}

print_error() {
    echo -e "${RED}✗${NC} $1"
}

print_info() {
    echo -e "${BLUE}ℹ${NC} $1"
}

print_warning() {
    echo -e "${YELLOW}⚠${NC} $1"
}

# Clear screen and show header
clear
echo -e "${BLUE}"
echo "========================================"
echo "    WinDaq Analyzer - Starting Up"
echo "========================================"
echo -e "${NC}"

# Check if Node.js is installed
echo "[1/5] Checking Node.js installation..."
if command -v node >/dev/null 2>&1; then
    NODE_VERSION=$(node --version)
    print_status "Node.js is installed ($NODE_VERSION)"
else
    print_error "Node.js is not installed!"
    echo "Please install Node.js from https://nodejs.org/"
    echo "Or install via Homebrew: brew install node"
    exit 1
fi

# Check if Python is installed
echo "[2/5] Checking Python installation..."
if command -v python3 >/dev/null 2>&1; then
    PYTHON_VERSION=$(python3 --version)
    print_status "Python is installed ($PYTHON_VERSION)"
elif command -v python >/dev/null 2>&1; then
    PYTHON_VERSION=$(python --version)
    print_status "Python is installed ($PYTHON_VERSION)"
else
    print_error "Python is not installed!"
    echo "Please install Python from https://python.org/"
    echo "Or install via Homebrew: brew install python"
    exit 1
fi

# Install Node.js dependencies
echo "[3/5] Installing Node.js dependencies..."
if [ ! -d "node_modules" ]; then
    print_info "Installing npm packages..."
    npm install
    if [ $? -eq 0 ]; then
        print_status "Node.js dependencies installed"
    else
        print_error "Failed to install npm packages!"
        exit 1
    fi
else
    print_status "Node.js dependencies already installed"
fi

# Install Python dependencies
echo "[4/5] Installing Python dependencies..."
print_info "Installing Python packages..."

# Try pip3 first, then pip
if command -v pip3 >/dev/null 2>&1; then
    pip3 install pandas openpyxl numpy scipy matplotlib xlsxwriter >/dev/null 2>&1
elif command -v pip >/dev/null 2>&1; then
    pip install pandas openpyxl numpy scipy matplotlib xlsxwriter >/dev/null 2>&1
else
    print_warning "pip not found, trying with python -m pip..."
    python3 -m pip install pandas openpyxl numpy scipy matplotlib xlsxwriter >/dev/null 2>&1
fi

print_status "Python dependencies installed"

# Start the application
echo "[5/5] Starting WinDaq Analyzer..."
echo
echo -e "${BLUE}"
echo "========================================"
echo "   WinDaq Analyzer is starting..."
echo "   Opening browser in 3 seconds..."
echo "========================================"
echo -e "${NC}"

# Start the server in background
npm start &
SERVER_PID=$!

# Wait a moment for server to start
sleep 3

# Open browser
if command -v open >/dev/null 2>&1; then
    open http://localhost:3000
elif command -v xdg-open >/dev/null 2>&1; then
    xdg-open http://localhost:3000
else
    print_info "Please open http://localhost:3000 in your browser"
fi

echo
print_status "WinDaq Analyzer is now running!"
print_status "Browser should open automatically"
echo
echo -e "${YELLOW}Instructions:${NC}"
echo "- Upload your WinDaq files (.wdq, .wdh, .wdc)"
echo "- View automatic pulse analysis"
echo "- Download Excel files with charts"
echo
echo -e "${YELLOW}To stop the server:${NC}"
echo "- Press Ctrl+C in this terminal"
echo "- Or close this terminal window"
echo

# Function to handle cleanup on exit
cleanup() {
    echo
    print_info "Shutting down WinDaq Analyzer..."
    kill $SERVER_PID 2>/dev/null
    exit 0
}

# Set up signal handlers
trap cleanup SIGINT SIGTERM

# Wait for the server process
wait $SERVER_PID
