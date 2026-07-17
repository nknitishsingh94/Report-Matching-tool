const express = require('express');
const cors = require('cors');
const multer = require('multer');
const xlsx = require('xlsx');

const app = express();
const port = 3001;

app.use(cors());

// Configure multer for in-memory file storage
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.post('/api/match', upload.fields([
    { name: 'masterFile', maxCount: 1 },
    { name: 'smallFiles' }
]), (req, res) => {
    try {
        if (!req.files || !req.files['masterFile']) {
            return res.status(400).json({ error: 'Master file is required.' });
        }
        if (!req.files['smallFiles'] || req.files['smallFiles'].length === 0) {
            return res.status(400).json({ error: 'At least one small file is required.' });
        }

        const masterFile = req.files['masterFile'][0];
        const smallFiles = req.files['smallFiles'];

        // Read master workbook and store Caller IDs in a Set
        const masterWorkbook = xlsx.read(masterFile.buffer, { type: 'buffer' });
        const masterSheetName = masterWorkbook.SheetNames[0];
        const masterSheet = masterWorkbook.Sheets[masterSheetName];
        const masterData = xlsx.utils.sheet_to_json(masterSheet, { defval: "" });

        const masterCallerIds = new Set();
        
        if (masterData.length > 0) {
            const firstRow = masterData[0];
            let masterCallerIdKey = Object.keys(firstRow).find(k => k.toLowerCase().replace(/[^a-z0-9]/g, '') === 'callerid') || 
                                    Object.keys(firstRow).find(k => k.toLowerCase().includes('caller')) || 'Caller ID';

            for (const row of masterData) {
                const callerId = String(row[masterCallerIdKey]).trim();
                if (callerId && callerId !== "undefined" && callerId !== "") {
                    masterCallerIds.add(callerId);
                }
            }
        }

        // Process Small Files
        const resultData = [];
        let matchedCount = 0;
        let notFoundCount = 0;

        for (const file of smallFiles) {
            const smallWorkbook = xlsx.read(file.buffer, { type: 'buffer' });
            const smallSheetName = smallWorkbook.SheetNames[0];
            const smallSheet = smallWorkbook.Sheets[smallSheetName];
            const smallData = xlsx.utils.sheet_to_json(smallSheet, { defval: "" });

            for (const row of smallData) {
                // Heuristic column matching for Caller ID
                let callerIdKey = Object.keys(row).find(k => k.toLowerCase().replace(/[^a-z0-9]/g, '') === 'callerid');
                if (!callerIdKey) callerIdKey = Object.keys(row).find(k => k.toLowerCase().includes('caller'));
                if (!callerIdKey) callerIdKey = 'Caller ID'; // Fallback

                const callerId = String(row[callerIdKey]).trim();
                
                if (callerId && callerId !== "undefined" && callerId !== "") {
                    if (masterCallerIds.has(callerId)) {
                        row['Status'] = 'MATCHED ✅';
                        matchedCount++;
                    } else {
                        row['Status'] = 'NOT FOUND ❌';
                        notFoundCount++;
                    }
                    resultData.push(row);
                }
            }
        }

        // Create updated master sheet
        const resultSheet = xlsx.utils.json_to_sheet(resultData);
        const newWorkbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(newWorkbook, resultSheet, 'Result Report');

        // Generate buffer
        const excelBuffer = xlsx.write(newWorkbook, { type: 'buffer', bookType: 'xlsx' });

        // Send response with base64 encoded file and stats
        res.json({
            success: true,
            stats: {
                totalMaster: resultData.length, // Updating stat to represent Total Processed from small files
                matched: matchedCount,
                notFound: notFoundCount
            },
            fileBase64: excelBuffer.toString('base64'),
            fileName: 'Result_Report.xlsx'
        });

    } catch (error) {
        console.error('Error processing files:', error);
        res.status(500).json({ error: 'Failed to process files.' });
    }
});

app.listen(port, () => {
    console.log(`Backend server running on http://localhost:${port}`);
});
