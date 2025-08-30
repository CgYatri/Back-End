const express = require('express');
const ExcelJS = require('exceljs');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(cors());
const PORT = 3000;

const excelFilePath = path.join(__dirname, 'fare_data.xlsx');

// Verify file exists
if (!fs.existsSync(excelFilePath)) {
  console.error(`Error: Excel file not found at ${excelFilePath}`);
  process.exit(1);
}

// Cache for fare data
let fareDataCache = null;

async function loadFareData() {
  try {
    console.log('Loading fare data from Excel...');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelFilePath);

    // Get the first worksheet (Table 2 in your case)
    const worksheet = workbook.worksheets[0];
    
    // Extract stops from first row (skip first cell)
    const stops = [];
    const firstRow = worksheet.getRow(1);
    firstRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      if (colNumber > 1) { // Skip first column (BRT Bus Shelter)
        stops.push(cell.text.replace(/\n/g, ' ')); // Replace newlines with spaces
      }
    });

    // Extract fares
    const fares = [];
    for (let i = 2; i <= worksheet.rowCount; i++) {
      const row = worksheet.getRow(i);
      const rowData = [];
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        if (colNumber > 1) { // Skip first column (stop names)
          rowData.push(Number(cell.text) || 0);
        }
      });
      fares.push(rowData);
    }

    // Validate data
    if (stops.length === 0 || fares.length === 0) {
      throw new Error('No data found in worksheet');
    }

    if (stops.length !== fares.length) {
      throw new Error('Stops and fares data mismatch in size');
    }

    console.log('Successfully loaded fare data');
    return { stops, fares };
  } catch (error) {
    console.error('Error loading fare data:', error.message);
    throw error;
  }
}

// Middleware to load fare data
app.use(async (req, res, next) => {
  try {
    if (!fareDataCache) {
      fareDataCache = await loadFareData();
    }
    req.fareData = fareDataCache;
    next();
  } catch (error) {
    res.status(500).json({ error: 'Failed to load fare data' });
  }
});

// API endpoints
app.get('/api/stops', (req, res) => {
  res.json(req.fareData.stops);
});

app.get('/api/fare', (req, res) => {
  const { from, to } = req.query;
  const { stops, fares } = req.fareData;

  if (!from || !to) {
    return res.status(400).json({ error: 'Both from and to parameters are required' });
  }

  const fromIndex = stops.indexOf(from);
  const toIndex = stops.indexOf(to);

  if (fromIndex === -1 || toIndex === -1) {
    return res.status(404).json({ error: 'Invalid stop name' });
  }

  const fare = fares[fromIndex][toIndex];
  res.json({ from, to, fare });
});

app.get('/api/download-excel', async (req, res) => {
  try {
    const fileStream = fs.createReadStream(excelFilePath);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=brt_fare_chart.xlsx');
    fileStream.pipe(res);
  } catch (error) {
    res.status(500).json({ error: 'Failed to download file' });
  }
});

// Start server
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
  console.log(`Fare data loaded from: ${excelFilePath}`);
});