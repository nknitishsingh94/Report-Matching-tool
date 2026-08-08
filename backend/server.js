const express = require('express');
const cors = require('cors');
const multer = require('multer');
const xlsx = require('xlsx');
const JSZip = require('jszip');

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

        const normalizeDate = (val) => {
            if (!val) return '';
            if (typeof val === 'number') {
                // Excel serial date to JS Date
                const date = new Date(Math.round((val - 25569) * 86400 * 1000));
                return date.toISOString().split('T')[0];
            }
            const str = String(val).trim();
            // Just take the date part before any space or 'T'
            return str.split(' ')[0].split('T')[0];
        };

        const sanitizeId = (id) => {
            if (id === undefined || id === null) return '';
            // Remove all non-alphanumeric chars and lowercase
            return String(id).toLowerCase().replace(/[^a-z0-9]/g, '');
        };

        // Aggregate small files data into a Map
        const smallFilesData = [];
        const callerMap = new Map();
        for (let i = 0; i < smallFiles.length; i++) {
            const file = smallFiles[i];
            const label = labels[i] || `File ${i + 1}`;
            const originalName = file.originalname || label;

            const smallWorkbook = xlsx.read(file.buffer, { type: 'buffer', cellDates: true });
            const smallSheetName = smallWorkbook.SheetNames[0];
            const smallSheet = smallWorkbook.Sheets[smallSheetName];
            const smallData = xlsx.utils.sheet_to_json(smallSheet, { defval: "" });
            
            smallFilesData.push({
                originalName: originalName,
                data: smallData
            });

            for (const row of smallData) {
                // Heuristic column matching
                let callerIdKey = Object.keys(row).find(k => k.toLowerCase().replace(/[^a-z0-9]/g, '') === 'callerid');
                if (!callerIdKey) callerIdKey = Object.keys(row).find(k => k.toLowerCase().includes('caller') || k.toLowerCase().includes('number'));

                let dateKey = Object.keys(row).find(k => k.toLowerCase().includes('date') && !k.toLowerCase().includes('update'));

                let dispositionKey = Object.keys(row).find(k => k.toLowerCase().includes('disposition') || k.toLowerCase().includes('status'));
                let timeKey = Object.keys(row).find(k => k.toLowerCase().includes('time') || k.toLowerCase().includes('duration'));

                // Fallbacks if columns are not found exactly
                if (!callerIdKey) callerIdKey = 'Number';
                if (!dispositionKey) dispositionKey = 'Status';
                if (!timeKey) timeKey = 'Duration time';
                if (!dateKey) dateKey = 'Date';

                const callerId = sanitizeId(row[callerIdKey]);
                const normDate = normalizeDate(row[dateKey]);
                
                const compositeKey = `${callerId}_${normDate}`;

                if (callerId && callerId !== "undefined") {
                    if (callerMap.has(compositeKey)) {
                        const existing = callerMap.get(compositeKey);
                        if (!existing.label.includes(label)) {
                            existing.label += `, ${label}`;
                        }
                        existing.sourceRows.push(row);
                    } else {
                        callerMap.set(compositeKey, {
                            disposition: row[dispositionKey],
                            totalTime: row[timeKey],
                            label: label,
                            sourceRows: [row]
                        });
                    }
                }
            }
        }

        // Read master workbook
        const masterWorkbook = xlsx.read(masterFile.buffer, { type: 'buffer', cellDates: true });
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
                Object.keys(firstRow).find(k => k.toLowerCase().includes('caller') || k.toLowerCase().includes('number')) || 'Number';

            let masterDateKey = Object.keys(firstRow).find(k => k.toLowerCase().includes('date') && !k.toLowerCase().includes('update')) || 'Date';

            let masterDispositionKey = Object.keys(firstRow).find(k => k.toLowerCase().includes('disposition') || k.toLowerCase().includes('status')) || 'Status';
            let masterTimeKey = Object.keys(firstRow).find(k => k.toLowerCase().includes('time') || k.toLowerCase().includes('duration')) || 'Duration time';

            for (let i = 0; i < masterData.length; i++) {
                const row = masterData[i];
                const callerId = sanitizeId(row[masterCallerIdKey]);
                const normDate = normalizeDate(row[masterDateKey]);
                
                const compositeKey = `${callerId}_${normDate}`;

                if (callerId && callerId !== "undefined" && callerMap.has(compositeKey)) {
                    const matchData = callerMap.get(compositeKey);
                    // Update the master row with data from small files
                    row[masterDispositionKey] = matchData.disposition;
                    row[masterTimeKey] = matchData.totalTime;
                    row['Source Label'] = matchData.label;
                    row['Match Status'] = 'MATCHED ✅';
                    matchedCount++;
                    
                    // Mark the source small file rows as matched
                    matchData.sourceRows.forEach(sr => sr['Match Status'] = 'MATCHED ✅');
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

        // Mark unmatched small file rows
        smallFilesData.forEach(sf => {
            sf.data.forEach(row => {
                if (!row['Match Status']) {
                    row['Match Status'] = 'NOT FOUND ❌';
                }
            });
        });

        // Create updated master sheet using exceljs
        const ExcelJS = require('exceljs');
        const newWorkbook = new ExcelJS.Workbook();
        const updatedMasterSheet = newWorkbook.addWorksheet('Updated Master');

        const highlightRows = (sheet, data) => {
            if (data.length > 0) {
                const headers = Object.keys(data[0]);
                sheet.columns = headers.map(header => ({ header: header, key: header, width: 20 }));

                data.forEach((rowData) => {
                    const row = sheet.addRow(rowData);
                    
                    if (rowData['Match Status'] === 'MATCHED ✅') {
                        row.eachCell((cell) => {
                            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
                            cell.font = { color: { argb: 'FF006100' } };
                        });
                    } else if (rowData['Match Status'] === 'NOT FOUND ❌') {
                        row.eachCell((cell) => {
                            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
                            cell.font = { color: { argb: 'FF9C0006' } };
                        });
                    }
                });
                sheet.getRow(1).font = { bold: true };
            }
        };

        highlightRows(updatedMasterSheet, masterData);

        if (notFoundList.length > 0) {
            const notFoundSheet = newWorkbook.addWorksheet('Not Found');
            highlightRows(notFoundSheet, notFoundList);
        }

        const excelBuffer = await newWorkbook.xlsx.writeBuffer();
        
        const masterFileResponse = {
            fileName: 'Updated_Master_Report.xlsx',
            fileBase64: excelBuffer.toString('base64')
        };
        
        const zip = new JSZip();
        const smallFilesStats = [];

        // Generate buffers for small files
        for (const sf of smallFilesData) {
            let matched = 0;
            let notFound = 0;
            
            sf.data.forEach(r => {
                if (r['Match Status'] === 'MATCHED ✅') matched++;
                else notFound++;
            });

            smallFilesStats.push({
                fileName: sf.originalName,
                matched: matched,
                notFound: notFound
            });

            const sfWorkbook = new ExcelJS.Workbook();
            const sfSheet = sfWorkbook.addWorksheet('Annotated');
            
            // Sort so MATCHED are at top (optional)
            sf.data.sort((a, b) => {
                if (a['Match Status'] === 'MATCHED ✅' && b['Match Status'] !== 'MATCHED ✅') return -1;
                if (a['Match Status'] !== 'MATCHED ✅' && b['Match Status'] === 'MATCHED ✅') return 1;
                return 0;
            });
            
            highlightRows(sfSheet, sf.data);
            const sfBuffer = await sfWorkbook.xlsx.writeBuffer();
            
            // Add prefix to original name to avoid confusion
            let original = sf.originalName;
            if (!original.endsWith('.xlsx') && !original.endsWith('.xls') && !original.endsWith('.csv')) {
                original += '.xlsx'; // fallback
            } else {
                original = original.replace(/\.(csv|xls)$/i, '.xlsx'); // save as xlsx since exceljs
            }
            
            zip.file(`Annotated_${original}`, sfBuffer);
        }

        const zipBuffer = await zip.generateAsync({ type: 'nodebuffer' });
        
        const zipFileResponse = {
            fileName: 'Annotated_Small_Reports.zip',
            fileBase64: zipBuffer.toString('base64')
        };

        // Send response with files array and stats
        res.json({
            success: true,
            stats: {
                totalMaster: masterData.length,
                matched: matchedCount,
                notFound: notFoundCount,
                smallFilesStats: smallFilesStats
            },
            masterFile: masterFileResponse,
            zipFile: zipFileResponse
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
