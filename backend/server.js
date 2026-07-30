const express = require('express');
const cors = require('cors');
const multer = require('multer');
const xlsx = require('xlsx');

const app = express();
const port = 3001;

app.use(cors());

app.get('/', (req, res) => {
    res.send('Report Reconciliation API is running!');
});

// Configure multer for in-memory file storage
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.post('/api/match', upload.fields([
    { name: 'masterFile', maxCount: 1 },
    { name: 'smallFiles' }
]), async (req, res) => {
    try {
        if (!req.files || !req.files['masterFile']) {
            return res.status(400).json({ error: 'Master file is required.' });
        }
        if (!req.files['smallFiles'] || req.files['smallFiles'].length === 0) {
            return res.status(400).json({ error: 'At least one small file is required.' });
        }

        const masterFile = req.files['masterFile'][0];
        const smallFiles = req.files['smallFiles'];
        const labelsInput = req.body.labels;
        const labels = Array.isArray(labelsInput) ? labelsInput : [labelsInput];

        // Aggregate small files data into a Map
        const callerMap = new Map();
        for (let i = 0; i < smallFiles.length; i++) {
            const file = smallFiles[i];
            const label = labels[i] || `File ${i + 1}`;

            const smallWorkbook = xlsx.read(file.buffer, { type: 'buffer' });
            const smallSheetName = smallWorkbook.SheetNames[0];
            const smallSheet = smallWorkbook.Sheets[smallSheetName];
            const smallData = xlsx.utils.sheet_to_json(smallSheet, { defval: "" });

            for (const row of smallData) {
                // Heuristic column matching
                let callerIdKey = Object.keys(row).find(k => k.toLowerCase().replace(/[^a-z0-9]/g, '') === 'callerid');
                if (!callerIdKey) callerIdKey = Object.keys(row).find(k => k.toLowerCase().includes('caller'));

                let dispositionKey = Object.keys(row).find(k => k.toLowerCase().includes('disposition') || k.toLowerCase().includes('status'));
                let timeKey = Object.keys(row).find(k => k.toLowerCase().includes('time') || k.toLowerCase().includes('duration'));

                // Fallbacks if columns are not found exactly
                if (!callerIdKey) callerIdKey = 'Caller ID';
                if (!dispositionKey) dispositionKey = 'Disposition';
                if (!timeKey) timeKey = 'Total Time';

                // We use String() to convert raw numbers to strings, which handles Excel's scientific notation display issues natively for standard phone numbers
                const callerId = String(row[callerIdKey]).trim();
                if (callerId && callerId !== "undefined") {
                    callerMap.set(callerId, {
                        disposition: row[dispositionKey],
                        totalTime: row[timeKey],
                        label: label
                    });
                }
            }
        }

        // Read master workbook
        const masterWorkbook = xlsx.read(masterFile.buffer, { type: 'buffer' });
        const masterSheetName = masterWorkbook.SheetNames[0];
        const masterSheet = masterWorkbook.Sheets[masterSheetName];
        const masterData = xlsx.utils.sheet_to_json(masterSheet, { defval: "" });

        let matchedCount = 0;
        let notFoundCount = 0;
        const notFoundList = [];

        // Match with master data
        if (masterData.length > 0) {
            const firstRow = masterData[0];
            let masterCallerIdKey = Object.keys(firstRow).find(k => k.toLowerCase().replace(/[^a-z0-9]/g, '') === 'callerid') ||
                Object.keys(firstRow).find(k => k.toLowerCase().includes('caller')) || 'Caller ID';

            let masterDispositionKey = Object.keys(firstRow).find(k => k.toLowerCase().includes('disposition') || k.toLowerCase().includes('status')) || 'Disposition';
            let masterTimeKey = Object.keys(firstRow).find(k => k.toLowerCase().includes('time') || k.toLowerCase().includes('duration')) || 'Total Time';

            for (let i = 0; i < masterData.length; i++) {
                const row = masterData[i];
                const callerId = String(row[masterCallerIdKey]).trim();

                if (callerId && callerId !== "undefined" && callerMap.has(callerId)) {
                    const matchData = callerMap.get(callerId);
                    // Update the master row with data from small files
                    row[masterDispositionKey] = matchData.disposition;
                    row[masterTimeKey] = matchData.totalTime;
                    row['Source Label'] = matchData.label;
                    row['Match Status'] = 'MATCHED ✅';
                    matchedCount++;
                } else if (callerId && callerId !== "undefined" && callerId !== "") {
                    // Update match status to "Not Found"
                    row['Match Status'] = 'NOT FOUND ❌';
                    notFoundCount++;
                    notFoundList.push(row);
                } else {
                    row['Match Status'] = '';
                }
            }

            // Sort masterData so MATCHED ✅ appears at the top
            masterData.sort((a, b) => {
                if (a['Match Status'] === 'MATCHED ✅' && b['Match Status'] !== 'MATCHED ✅') return -1;
                if (a['Match Status'] !== 'MATCHED ✅' && b['Match Status'] === 'MATCHED ✅') return 1;
                return 0;
            });
        }

        // Create updated master sheet using exceljs
        const ExcelJS = require('exceljs');
        const newWorkbook = new ExcelJS.Workbook();
        const updatedMasterSheet = newWorkbook.addWorksheet('Updated Master');

        if (masterData.length > 0) {
            const headers = Object.keys(masterData[0]);
            updatedMasterSheet.columns = headers.map(header => ({ header: header, key: header, width: 20 }));

            masterData.forEach((rowData) => {
                const row = updatedMasterSheet.addRow(rowData);
                
                // Highlight row based on Match Status
                if (rowData['Match Status'] === 'MATCHED ✅') {
                    row.eachCell((cell) => {
                        cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFC6EFCE' } // Light green
                        };
                        cell.font = { color: { argb: 'FF006100' } }; // Dark green text
                    });
                } else if (rowData['Match Status'] === 'NOT FOUND ❌') {
                    row.eachCell((cell) => {
                        cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFFFC7CE' } // Light red
                        };
                        cell.font = { color: { argb: 'FF9C0006' } }; // Dark red text
                    });
                }
            });
            
            // Make header bold
            updatedMasterSheet.getRow(1).font = { bold: true };
        }

        // Create Not Found sheet
        if (notFoundList.length > 0) {
            const notFoundSheet = newWorkbook.addWorksheet('Not Found');
            const headers = Object.keys(notFoundList[0]);
            notFoundSheet.columns = headers.map(header => ({ header: header, key: header, width: 20 }));
            notFoundList.forEach(rowData => {
                notFoundSheet.addRow(rowData);
            });
            notFoundSheet.getRow(1).font = { bold: true };
        }

        // Generate buffer
        const excelBuffer = await newWorkbook.xlsx.writeBuffer();

        // Send response with base64 encoded file and stats
        res.json({
            success: true,
            stats: {
                totalMaster: masterData.length,
                matched: matchedCount,
                notFound: notFoundCount
            },
            fileBase64: excelBuffer.toString('base64'),
            fileName: 'Updated_Master_Report.xlsx'
        });

    } catch (error) {
        console.error('Error processing files:', error);
        res.status(500).json({ error: 'Failed to process files.' });
    }
});

app.listen(port, () => {
    console.log(`Backend server running on http://localhost:${port}`);
});

module.exports = app;
