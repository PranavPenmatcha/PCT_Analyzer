const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs-extra');
const { v4: uuidv4 } = require('uuid');
const { spawn } = require('child_process');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(express.static('public'));
app.use(express.json());

// Create uploads directory if it doesn't exist
const uploadsDir = path.join(__dirname, 'uploads');
const outputsDir = path.join(__dirname, 'outputs');
fs.ensureDirSync(uploadsDir);
fs.ensureDirSync(outputsDir);

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, uploadsDir);
    },
    filename: (req, file, cb) => {
        const uniqueId = uuidv4();
        const ext = path.extname(file.originalname);
        cb(null, `${uniqueId}${ext}`);
    }
});

const upload = multer({
    storage: storage,
    fileFilter: (req, file, cb) => {
        const allowedExtensions = ['.wdq', '.wdh', '.wdc'];
        const ext = path.extname(file.originalname).toLowerCase();
        if (allowedExtensions.includes(ext)) {
            cb(null, true);
        } else {
            cb(new Error('Only WinDaq files (.wdq, .wdh, .wdc) are allowed'));
        }
    },
    limits: {
        fileSize: 100 * 1024 * 1024 // 100MB limit
    }
});

// Routes
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Upload and analyze WinDaq file
app.post('/upload', upload.single('wdqFile'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        const uploadedFile = req.file;
        const sessionId = path.parse(uploadedFile.filename).name;
        const sessionDir = path.join(outputsDir, sessionId);
        
        // Create session directory
        await fs.ensureDir(sessionDir);
        
        // Copy uploaded file to session directory
        const sessionFilePath = path.join(sessionDir, uploadedFile.originalname);
        await fs.copy(uploadedFile.path, sessionFilePath);
        
        console.log(`Processing file: ${uploadedFile.originalname}`);
        
        // Step 1: Convert WinDaq to Excel
        const convertResult = await runPythonScript('windaq_to_excel_converter.py', [sessionFilePath], sessionDir);
        
        if (!convertResult.success) {
            throw new Error(`Conversion failed: ${convertResult.error}`);
        }
        
        // Find the generated Excel file
        const excelFiles = await fs.readdir(sessionDir);
        const excelFile = excelFiles.find(f => f.endsWith('.xlsx') && !f.includes('_with_chart'));
        
        if (!excelFile) {
            throw new Error('Excel file not generated');
        }
        
        const excelPath = path.join(sessionDir, excelFile);
        
        // Step 2: Analyze pulses
        const analyzeResult = await runPythonScript('pulse_analyzer.py', [], sessionDir, excelPath);

        if (!analyzeResult.success) {
            throw new Error(`Pulse analysis failed: ${analyzeResult.error}`);
        }

        // Step 3: Create chart
        const chartResult = await runPythonScript('add_chart_to_excel.py', [], sessionDir);

        if (!chartResult.success) {
            throw new Error(`Chart creation failed: ${chartResult.error}`);
        }
        
        // Parse analysis results
        const analysisData = parseAnalysisOutput(analyzeResult.output);
        
        // Find the chart Excel file
        const chartFiles = await fs.readdir(sessionDir);
        const chartFile = chartFiles.find(f => f.includes('_with_chart.xlsx'));
        
        res.json({
            success: true,
            sessionId: sessionId,
            filename: uploadedFile.originalname,
            analysis: analysisData,
            downloadUrl: chartFile ? `/download/${sessionId}/${encodeURIComponent(chartFile)}` : null
        });
        
        // Clean up uploaded file
        await fs.remove(uploadedFile.path);
        
    } catch (error) {
        console.error('Upload error:', error);
        res.status(500).json({ error: error.message });
    }
});

// Download processed file
app.get('/download/:sessionId/*', (req, res) => {
    const { sessionId } = req.params;
    const filename = req.params[0]; // Get the rest of the path
    const decodedFilename = decodeURIComponent(filename);
    const filePath = path.join(outputsDir, sessionId, decodedFilename);



    if (fs.existsSync(filePath)) {
        res.download(filePath, decodedFilename, (err) => {
            if (err && !res.headersSent) {
                res.status(500).json({ error: 'Download failed' });
            }
        });
    } else {

        // Try to find any chart file in the session directory
        try {
            const sessionDir = path.join(outputsDir, sessionId);
            const files = fs.readdirSync(sessionDir);
            const chartFile = files.find(f => f.includes('_with_chart.xlsx'));

            if (chartFile) {
                const chartPath = path.join(sessionDir, chartFile);
                res.download(chartPath, chartFile, (err) => {
                    if (err && !res.headersSent) {
                        res.status(500).json({ error: 'Download failed' });
                    }
                });
            } else {
                res.status(404).json({ error: 'File not found' });
            }
        } catch (error) {
            res.status(404).json({ error: 'File not found' });
        }
    }
});

// Helper function to detect Python command
function detectPythonCommand() {
    const { execSync } = require('child_process');
    const commands = process.platform === 'win32' ? ['python', 'py', 'python3'] : ['python3', 'python'];

    for (const cmd of commands) {
        try {
            execSync(`${cmd} --version`, { stdio: 'ignore' });
            return cmd;
        } catch (error) {
            // Command not found, try next
        }
    }
    return null;
}

// Helper function to run Python scripts
function runPythonScript(scriptName, args = [], workingDir = __dirname, targetFile = null) {
    return new Promise((resolve) => {
        const scriptPath = path.join(__dirname, scriptName);
        let pythonArgs = [scriptPath, ...args];
        
        // For pulse_analyzer.py, we need to target the specific Excel file
        if (scriptName === 'pulse_analyzer.py' && targetFile) {
            // Copy the script to working directory and modify it to target specific file
            const tempScriptPath = path.join(workingDir, 'temp_pulse_analyzer.py');
            const originalScript = fs.readFileSync(scriptPath, 'utf8');
            const modifiedScript = originalScript.replace(
                'excel_files = list(Path(\'.\').glob(\'*.xlsx\'))',
                `excel_files = [Path('${path.basename(targetFile)}')]`
            );
            fs.writeFileSync(tempScriptPath, modifiedScript);
            pythonArgs = [tempScriptPath];
        }
        
        // Detect the correct Python command
        const pythonCmd = detectPythonCommand();

        if (!pythonCmd) {
            resolve({
                success: false,
                output: '',
                error: 'Python not found. Please install Python and ensure it is in your PATH.',
                code: -1
            });
            return;
        }

        const python = spawn(pythonCmd, pythonArgs, {
            cwd: workingDir,
            stdio: ['pipe', 'pipe', 'pipe']
        });
        
        let output = '';
        let error = '';
        
        python.stdout.on('data', (data) => {
            output += data.toString();
        });
        
        python.stderr.on('data', (data) => {
            error += data.toString();
        });
        
        python.on('close', (code) => {
            if (code !== 0) {
                console.error(`Python script ${scriptName} failed with code ${code}`);
                console.error(`Python command used: ${pythonCmd}`);
                console.error(`Error output: ${error}`);
                console.error(`Standard output: ${output}`);
            }
            resolve({
                success: code === 0,
                output: output,
                error: error,
                code: code
            });
        });
    });
}

// Helper function to parse analysis output
function parseAnalysisOutput(output) {

    const lines = output.split('\n');
    const analysis = {
        totalPulses: 0,
        peakCurrentRange: '',
        highestPeak: 0,
        pulses: []
    };

    let inPulseSection = false;

    // Extract key information from output
    for (const line of lines) {
        const trimmedLine = line.trim();

        if (trimmedLine.includes('Found') && trimmedLine.includes('current pulses')) {
            const match = trimmedLine.match(/Found (\d+) current pulses/);
            if (match) analysis.totalPulses = parseInt(match[1]);
        }

        if (trimmedLine.includes('Peak current range:')) {
            const match = trimmedLine.match(/Peak current range: ([\d.-]+ - [\d.-]+ A)/);
            if (match) analysis.peakCurrentRange = match[1];
        }

        if (trimmedLine.includes('Highest peak:')) {
            const match = trimmedLine.match(/Highest peak: ([\d.-]+) A/);
            if (match) analysis.highestPeak = parseFloat(match[1]);
        }

        // Look for pulse summary section
        if (trimmedLine.includes('Individual Pulse Details:') || trimmedLine.includes('🔋 Individual Pulse Details:')) {
            inPulseSection = true;
            analysis.pulses = []; // Clear existing pulses to avoid duplicates
            continue;
        }

        // Parse individual pulse lines with time information
        if (inPulseSection && trimmedLine.includes('Pulse') && trimmedLine.includes('A at') && trimmedLine.includes(':')) {
            const match = trimmedLine.match(/Pulse (\d+): ([\d.-]+) A at ([\d.-]+)s/);
            if (match) {
                analysis.pulses.push({
                    number: parseInt(match[1]),
                    peakCurrent: parseFloat(match[2]),
                    time: parseFloat(match[3])
                });
            }
        }

        // Alternative parsing for different output format (only if not in detailed section)
        if (!inPulseSection && trimmedLine.match(/Pulse \d+\s+\d+/)) {
            const match = trimmedLine.match(/Pulse (\d+)\s+([\d.-]+)/);
            if (match) {
                analysis.pulses.push({
                    number: parseInt(match[1]),
                    peakCurrent: parseFloat(match[2]),
                    time: 0 // Will be updated if we find time info
                });
            }
        }
    }

    return analysis;
}

// Clean up old files (run every hour)
setInterval(async () => {
    try {
        const sessions = await fs.readdir(outputsDir);
        const now = Date.now();
        const maxAge = 24 * 60 * 60 * 1000; // 24 hours
        
        for (const session of sessions) {
            const sessionPath = path.join(outputsDir, session);
            const stats = await fs.stat(sessionPath);
            
            if (now - stats.mtime.getTime() > maxAge) {
                await fs.remove(sessionPath);
                console.log(`Cleaned up old session: ${session}`);
            }
        }
    } catch (error) {
        console.error('Cleanup error:', error);
    }
}, 60 * 60 * 1000); // Run every hour

app.listen(PORT, () => {
    console.log(`🚀 WinDaq Analyzer server running on http://localhost:${PORT}`);
    console.log(`📊 Ready to analyze WinDaq files!`);
});
