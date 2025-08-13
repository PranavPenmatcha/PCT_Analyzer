# Quick Start Scripts

## Windows Users

**Double-click:** `start_windaq_analyzer.bat`

This will:
- Check if Node.js and Python are installed
- Install all required dependencies automatically
- Start the WinDaq Analyzer server
- Open your browser to http://localhost:3000

## Mac Users

**Double-click:** `start_windaq_analyzer.sh`

Or run in Terminal:
```bash
./start_windaq_analyzer.sh
```

This will:
- Check if Node.js and Python are installed
- Install all required dependencies automatically
- Start the WinDaq Analyzer server
- Open your browser to http://localhost:3000

## Requirements

### Windows
- **Node.js**: Download from https://nodejs.org/
- **Python**: Download from https://python.org/

### Mac
- **Node.js**: Download from https://nodejs.org/ or `brew install node`
- **Python**: Download from https://python.org/ or `brew install python`

## Troubleshooting

### If the script fails:
1. **Install Node.js and Python** manually from the links above
2. **Run the script again** - it will detect the installations
3. **Check your internet connection** - scripts need to download packages

### If the browser doesn't open:
- Manually go to: **http://localhost:3000**

### To stop the application:
- **Windows**: Close the command window or press Ctrl+C
- **Mac**: Press Ctrl+C in the terminal or close the terminal window

## What the scripts do:

1. **Check prerequisites** (Node.js, Python)
2. **Install Node.js packages** (`npm install`)
3. **Install Python packages** (pandas, openpyxl, numpy, etc.)
4. **Start the server** (`npm start`)
5. **Open browser** to http://localhost:3000
6. **Keep running** until you close it

## First Time Setup

The first time you run the script, it may take a few minutes to download and install all the required packages. Subsequent runs will be much faster.
