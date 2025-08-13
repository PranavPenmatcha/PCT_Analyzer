# WinDaq Analyzer

A web application for analyzing WinDaq files and detecting current pulses with automatic Excel generation and interactive charts.

## How to Use

### 1. Start the Application

```bash
npm install
npm start
```

The application will start at: **http://localhost:3000**

### 2. Upload Your WinDaq File

- **Drag & Drop**: Drag your .wdq, .wdh, or .wdc file onto the upload area
- **Browse**: Click "Browse Files" to select a file from your computer
- **File Size**: Up to 100MB supported

### 3. Automatic Analysis

The application will automatically:
- Convert your WinDaq file to Excel format
- Detect current pulses using advanced algorithms
- Generate interactive charts
- Create a downloadable Excel file

### 4. View Results

**Summary Statistics:**
- Total number of pulses detected
- Peak current range (minimum to maximum)
- Highest peak current value

**Individual Pulse Data:**
- Peak current for each pulse
- Timing information for each pulse
- Complete pulse analysis

### 5. Download Excel File

Click "Download Excel" to get your analyzed data with:
- **Raw Data Sheet**: Time and current data with pulse summary table
- **Interactive Chart Sheet**: Professional Current vs Time visualization
- **Ready for Reports**: Professional formatting for presentations

## Requirements

- **Node.js** (v14 or higher)
- **Python 3.6+** with packages:
  ```bash
  pip install pandas openpyxl numpy scipy matplotlib xlsxwriter
  ```

## Supported File Types

- `.wdq` - WinDaq data files
- `.wdh` - WinDaq header files  
- `.wdc` - WinDaq compressed files

## Features

✅ **Automatic pulse detection**  
✅ **Peak current analysis**  
✅ **Interactive charts**  
✅ **Excel export with charts**  
✅ **Professional formatting**  
✅ **Drag & drop upload**  
✅ **Real-time progress tracking**
