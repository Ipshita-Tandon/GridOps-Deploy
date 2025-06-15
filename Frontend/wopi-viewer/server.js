const express = require("express");
const fs = require("fs");
const path = require("path");
const axios = require("axios");
const cors = require("cors");
const { spawn } = require("child_process");

const app = express();
const port = 3000;
const ACCESS_TOKEN = "12345";
let HOST_URL = `http://localhost:${port}`;

app.use(express.json());
app.use(cors());
app.use(express.static(__dirname));

// --- DUMMY MERGED CELLS JSON DATA ---
const dummyMergedCellsData = {
  "merged_cells": [
    {
      "sheet_name": "Monthly budget report",
      "merged_cells": [
        { "range": "C3:F3", "value": "Budget\nOverview" }
      ]
    },
    {
      "sheet_name": "Sheet1", // IMPORTANT: Match your test.xlsx sheet name
      "merged_cells": [
        { "range": "B5:C6", "value": "Merged B5:C6" },
        { "range": "F5:G6", "value": "Merged F5:G6" }
      ]
    }
  ]
};
const dummyMergedCellsDataJson = JSON.stringify(dummyMergedCellsData);
// ------------------------------------

// --- Read Charts and Images JSON Data ---
let chartsDataJson = "{}";
let imagesDataJson = "{}";

try {
  const chartsFilePath = path.join(__dirname, 'charts.json');
  if (fs.existsSync(chartsFilePath)) {
    chartsDataJson = fs.readFileSync(chartsFilePath, 'utf8');
    console.log('charts.json loaded successfully.');
  } else {
    console.warn('charts.json not found. Charts highlighting will not work.');
  }

  const imagesFilePath = path.join(__dirname, 'images.json');
  if (fs.existsSync(imagesFilePath)) {
    imagesDataJson = fs.readFileSync(imagesFilePath, 'utf8');
    console.log('images.json loaded successfully.');
  } else {
    console.warn('images.json not found. Images highlighting will not work.');
  }
} catch (e) {
  console.error('Error loading charts/images JSON files:', e);
}
// ---------------------------------------------


async function getNgrokUrl() {
  try {
    const res = await axios.get("http://127.0.0.1:4040/api/tunnels");
    const tunnels = res.data.tunnels;
    const httpsTunnel = tunnels.find(t => t.public_url.startsWith("https"));
    if (httpsTunnel) {
      HOST_URL = httpsTunnel.public_url;
      console.log("HOST_URL:", HOST_URL);
    }
  } catch (e) {
    console.log("Ngrok not ready yet (or Ngrok API not accessible). Make sure Ngrok is running.)");
  }
}

app.get("/host-url", async (req, res) => {
  await getNgrokUrl();
  res.json({ hostUrl: HOST_URL });
});

app.get(`/wopi/files/:fileId`, (req, res) => {
  const fileId = req.params.fileId;
  const filePath = path.join(__dirname, fileId);

  if (!fs.existsSync(filePath)) {
    console.error(`CheckFileInfo: File not found: ${filePath}`);
    return res.status(404).send("File Not Found");
  }

  const stat = fs.statSync(filePath);
  res.json({
    BaseFileName: fileId,
    Size: stat.size,
    OwnerId: "user",
    UserId: "user",
    Version: Date.now().toString(), // Crucial for WOPI to detect changes
    UserFriendlyName: "Local User",
    FileUrl: `${HOST_URL}/wopi/files/${fileId}/contents`,
    SupportsUpdate: false
  });
});

app.get(`/wopi/files/:fileId/contents`, (req, res) => {
  const fileId = req.params.fileId;
  const filePath = path.join(__dirname, fileId);

  if (!fs.existsSync(filePath)) {
    console.error(`GetFile: File not found: ${filePath}`);
    return res.status(404).send("File Not Found");
  }

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  fs.createReadStream(filePath).pipe(res);
});

// Function to clean up previous highlighted files
function cleanupPreviousHighlights(originalFileNameWithoutExt) {
  const files = fs.readdirSync(__dirname);
  files.forEach(file => {
    if (file.startsWith(`${originalFileNameWithoutExt}_highlighted`) && file.endsWith('.xlsx')) {
      try {
        fs.unlinkSync(path.join(__dirname, file));
        console.log(`Cleaned up previous highlight file: ${file}`);
      } catch (err) {
        console.error(`Error cleaning up file ${file}:`, err);
      }
    }
  });
}

app.post("/highlight-cells", async (req, res) => {
  const { fileName, sheetName, cellRanges } = req.body;

  let baseOriginalFileName = fileName;
  const highlightedSuffix = '_highlighted_';
  if (fileName.includes(highlightedSuffix)) {
    const parts = fileName.split(highlightedSuffix);
    if (parts.length > 1) {
      const originalBaseName = parts[0];
      const originalExt = path.extname(fileName);
      baseOriginalFileName = originalBaseName + originalExt;
    }
  }
  console.log(`Received request for: ${fileName}. Derived original base file: ${baseOriginalFileName}`);

  const originalFilePath = path.join(__dirname, baseOriginalFileName);
  const pythonScriptPath = path.join(__dirname, 'app.py');

  const originalFileNameWithoutExt = path.parse(baseOriginalFileName).name;
  const timestamp = Date.now();
  const highlightedFileName = `${originalFileNameWithoutExt}_highlighted_${timestamp}.xlsx`;
  const highlightedFilePath = path.join(__dirname, highlightedFileName);

  // Clean up previous highlighted files
  cleanupPreviousHighlights(originalFileNameWithoutExt);

  if (!fs.existsSync(originalFilePath)) {
    console.error(`Original file not found at: ${originalFilePath}`);
    return res.status(404).send("Original Excel file not found.");
  }
  if (!fs.existsSync(pythonScriptPath)) {
    console.error(`Python script not found: ${pythonScriptPath}`);
    return res.status(500).send("Python script (app.py) not found on server.");
  }

  try {
    const cellRangesArray = cellRanges.split(',').map(s => s.trim()).filter(s => s.length > 0);
    const cellAddressesJson = JSON.stringify(cellRangesArray);

    console.log(`Spawning Python script with:`);
    console.log(`  Input File: ${originalFilePath}`);
    console.log(`  Sheet Name: ${sheetName}`);
    console.log(`  Cell Ranges: ${cellAddressesJson}`);
    console.log(`  Merged Cells Data: (first 100 chars) ${dummyMergedCellsDataJson.substring(0, 100)}...`);
    console.log(`  Charts Data: (first 100 chars) ${chartsDataJson.substring(0, 100)}...`);
    console.log(`  Images Data: (first 100 chars) ${imagesDataJson.substring(0, 100)}...`);
    console.log(`  Output File: ${highlightedFilePath}`);

    const pythonProcess = spawn('python3', [
      pythonScriptPath,
      originalFilePath,
      sheetName,
      cellAddressesJson,
      dummyMergedCellsDataJson,
      chartsDataJson,
      imagesDataJson,
      highlightedFilePath
    ]);

    let pythonOutput = '';
    let pythonError = '';

    pythonProcess.stdout.on('data', (data) => {
      pythonOutput += data.toString();
      console.log('Python stdout:', data.toString());
    });

    pythonProcess.stderr.on('data', (data) => {
      pythonError += data.toString();
      console.error('Python stderr:', data.toString());
    });

    await new Promise((resolve, reject) => {
      pythonProcess.on('close', (code) => {
        if (code === 0) {
          console.log(`Python script exited successfully. Output file: ${highlightedFileName}`);
          resolve();
        } else {
          const errorMessage = `Python script exited with code ${code}.\nError:\n${pythonError}\nOutput:\n${pythonOutput}`;
          console.error(errorMessage);
          reject(new Error(errorMessage));
        }
      });

      pythonProcess.on('error', (err) => {
        console.error('Failed to start Python child process:', err);
        reject(err);
      });
    });

    res.json({ highlightedFileName: highlightedFileName });

  } catch (error) {
    console.error('Error in /highlight-cells endpoint:', error);
    res.status(500).send(`Error processing request: ${error.message}`);
  }
});

app.post('/highlight-cells', async (req, res) => {
  try {
    const { fileName, sheetName, cellRanges } = req.body;
    if (!fileName || !sheetName || !cellRanges) {
      return res.status(400).json({ error: 'Missing required parameters' });
    }

    // Generate a unique name for the highlighted file
    const highlightedFileName = `highlighted_${Date.now()}_${fileName}`;
    
    // Call the Python script to highlight cells
    const pythonProcess = spawn('python3', [
      'app.py',
      '--input', fileName,
      '--output', highlightedFileName,
      '--sheet', sheetName,
      '--cells', cellRanges
    ]);

    let scriptOutput = '';
    let scriptError = '';

    pythonProcess.stdout.on('data', (data) => {
      scriptOutput += data.toString();
    });

    pythonProcess.stderr.on('data', (data) => {
      scriptError += data.toString();
    });

    await new Promise((resolve, reject) => {
      pythonProcess.on('close', (code) => {
        if (code === 0) {
          resolve();
        } else {
          reject(new Error(`Python script failed with code ${code}: ${scriptError}`));
        }
      });
    });

    res.json({ highlightedFileName });
  } catch (error) {
    console.error('Error highlighting cells:', error);
    res.status(500).json({ error: error.message });
  }
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});