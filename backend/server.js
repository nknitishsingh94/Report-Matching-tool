const express = require('express');
const cors = require('cors');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const app = express();
const port = 3001;

app.use(cors());

app.get('/', (req, res) => {
    res.send('Report Reconciliation API is running!');
});

// Configure multer for in-memory file storage
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });
app.use(express.json());

// Helper functions for duration removed as requested
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
            
            // If it's a JS Date object (parsed by cellDates: true)
            if (val instanceof Date) {
                if (!isNaN(val)) return val.toISOString().split('T')[0];
                return '';
            }

            // If it's an Excel serial date number
            if (typeof val === 'number') {
                const date = new Date(Math.round((val - 25569) * 86400 * 1000));
                if (!isNaN(date)) return date.toISOString().split('T')[0];
                return '';
            }
            
            // It's a string. Try parsing it first.
            const str = String(val).trim();
            const d = new Date(str);
            if (!isNaN(d)) {
                return d.toISOString().split('T')[0];
            }

            // Fallback: just take the date part before any space or 'T'
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
                
                // Match strictly on callerId as requested
                const compositeKey = callerId;

                if (callerId && callerId !== "undefined") {
                    const mappedRow = {
                        row: row,
                        dispositionKey: dispositionKey,
                        timeKey: timeKey,
                        label: label
                    };

                    if (callerMap.has(compositeKey)) {
                        callerMap.get(compositeKey).push(mappedRow);
                    } else {
                        callerMap.set(compositeKey, [mappedRow]);
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
                
                // Match strictly on callerId
                const compositeKey = callerId;

                const mappedRows = callerMap.get(compositeKey);
                if (callerId && callerId !== "undefined" && mappedRows && mappedRows.length > 0) {
                    // Allow one small file record to match multiple master records (no shift)
                    const matchData = mappedRows[0];
                    const sr = matchData.row;

                    // Update the master row with data from small files
                    row[masterDispositionKey] = sr[matchData.dispositionKey];
                    row[masterTimeKey] = sr[matchData.timeKey];
                    row['Source Label'] = matchData.label;
                    row['Match Status'] = 'MATCHED ✅';
                    matchedCount++;
                    
                    // Mark the source small file row as matched
                    sr['Match Status'] = 'MATCHED ✅';
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
                // Compute headers dynamically to capture newly added columns
                const headerSet = new Set();
                data.forEach(r => Object.keys(r).forEach(k => headerSet.add(k)));
                const headers = Array.from(headerSet);
                
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
            fileBase64: excelBuffer.toString('base64'),
            previewData: [...masterData, ...notFoundList]
        };
        
        const smallFilesResponse = [];

        // Generate buffers and stats for small files
        for (const sf of smallFilesData) {
            let matched = 0;
            let notFound = 0;
            
            sf.data.forEach(r => {
                if (r['Match Status'] === 'MATCHED ✅') matched++;
                else notFound++;
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
            
            smallFilesResponse.push({
                fileName: `Annotated_${original}`,
                fileBase64: sfBuffer.toString('base64'),
                matched: matched,
                notFound: notFound,
                previewData: sf.data
            });
        }

        // Send response with files array and stats
        res.json({
            success: true,
            stats: {
                totalMaster: masterData.length,
                matched: matchedCount,
                notFound: notFoundCount
            },
            masterFile: masterFileResponse,
            smallFiles: smallFilesResponse
        });

    } catch (error) {
        console.error('Error processing files:', error);
        res.status(500).json({ error: 'Failed to process files.' });
    }
});

// Helper for Mock Veriphone
const getMockVeriphoneResponse = (phone) => {
    const types = ['mobile', 'voip', 'fixed_line'];
    const type = types[Math.floor(Math.random() * types.length)];
    return {
        status: "success",
        phone: phone,
        phone_valid: true,
        phone_type: type,
        carrier: "Mock Carrier",
        country: "United States"
    };
};

const LOG_FILE = path.join(__dirname, 'validation_log.json');

app.post('/api/validate-numbers', upload.single('sheet'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'Sheet file is required.' });
        }

        const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(sheet, { defval: "" });

        if (data.length === 0) {
            return res.status(400).json({ error: 'Sheet is empty.' });
        }

        let phoneKey = Object.keys(data[0]).find(k => k.toLowerCase().includes('phone') || k.toLowerCase().includes('number') || k.toLowerCase().includes('caller'));
        if (!phoneKey) phoneKey = Object.keys(data[0])[0]; // Fallback to first column

        let voipCount = 0;
        let mobileCount = 0;
        let landlineCount = 0;
        let validCount = 0;

        const logEntries = [];

        for (let row of data) {
            const phone = row[phoneKey];
            if (phone) {
                // Mock API Call
                const apiResponse = getMockVeriphoneResponse(phone);
                
                row['Valid'] = apiResponse.phone_valid ? 'Yes' : 'No';
                row['Line Type'] = apiResponse.phone_type;
                row['Carrier'] = apiResponse.carrier;

                if (apiResponse.phone_valid) validCount++;
                if (apiResponse.phone_type === 'voip') voipCount++;
                else if (apiResponse.phone_type === 'mobile') mobileCount++;
                else if (apiResponse.phone_type === 'fixed_line') landlineCount++;

                logEntries.push({
                    timestamp: new Date().toISOString(),
                    phone: phone,
                    type: apiResponse.phone_type,
                    carrier: apiResponse.carrier
                });
            }
        }

        // Save to Active Log
        let existingLog = [];
        if (fs.existsSync(LOG_FILE)) {
            existingLog = JSON.parse(fs.readFileSync(LOG_FILE, 'utf8'));
        }
        existingLog = [...logEntries, ...existingLog].slice(0, 500); // Keep last 500
        fs.writeFileSync(LOG_FILE, JSON.stringify(existingLog, null, 2));

        // Create Response Excel
        const newWorkbook = new ExcelJS.Workbook();
        const updatedSheet = newWorkbook.addWorksheet('Validated Numbers');
        
        if (data.length > 0) {
            const headers = Object.keys(data[0]);
            updatedSheet.columns = headers.map(header => ({ header: header, key: header, width: 20 }));
            
            data.forEach(rowData => {
                const row = updatedSheet.addRow(rowData);
                if (rowData['Line Type'] === 'voip') {
                    row.eachCell(cell => {
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } }; // Light red for VoIP
                        cell.font = { color: { argb: 'FF9C0006' } };
                    });
                } else {
                    row.eachCell(cell => {
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } }; // Light green for others
                        cell.font = { color: { argb: 'FF006100' } };
                    });
                }
            });
            updatedSheet.getRow(1).font = { bold: true };
        }

        const excelBuffer = await newWorkbook.xlsx.writeBuffer();
        
        let originalName = req.file.originalname || 'Numbers';
        originalName = originalName.replace(/\.(csv|xls)$/i, '.xlsx');

        res.json({
            success: true,
            stats: {
                total: data.length,
                valid: validCount,
                voip: voipCount,
                mobile: mobileCount,
                landline: landlineCount
            },
            file: {
                fileName: `Validated_${originalName}`,
                fileBase64: excelBuffer.toString('base64'),
                previewData: data
            },
            activeLog: existingLog.slice(0, 50) // Return top 50 for UI
        });

    } catch (error) {
        console.error('Error validating numbers:', error);
        res.status(500).json({ error: 'Failed to validate numbers.' });
    }
});

app.listen(port, () => {
    console.log(`Backend server running on http://localhost:${port}`);
});

module.exports = app;
\napp.post('/api/match', upload.fields([
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
            if (val instanceof Date) {
                if (!isNaN(val)) return val.toISOString().split('T')[0];
                return '';
            }
            if (typeof val === 'number') {
                const date = new Date(Math.round((val - 25569) * 86400 * 1000));
                if (!isNaN(date)) return date.toISOString().split('T')[0];
                return '';
            }
            const str = String(val).trim();
            const d = new Date(str);
            if (!isNaN(d)) return d.toISOString().split('T')[0];
            return str.split(' ')[0].split('T')[0] || str;
        };

        const sanitizeId = (id) => {
            if (id === undefined || id === null) return '';
            return String(id).toLowerCase().replace(/[^a-z0-9]/g, '');
        };

        const getKeys = (row) => {
            let callerIdKey = Object.keys(row).find(k => k.toLowerCase().replace(/[^a-z0-9]/g, '') === 'callerid');
            if (!callerIdKey) callerIdKey = Object.keys(row).find(k => k.toLowerCase().includes('caller') || k.toLowerCase().includes('number') || k.toLowerCase().includes('phone'));
            if (!callerIdKey) callerIdKey = 'Number';

            let dateKey = Object.keys(row).find(k => k.toLowerCase().includes('date') && !k.toLowerCase().includes('update'));
            if (!dateKey) dateKey = 'Date';

            let extKey = Object.keys(row).find(k => k.toLowerCase().includes('ext') || k.toLowerCase().includes('extension'));
            if (!extKey) extKey = 'Extension';
            
            let statusKey = Object.keys(row).find(k => k.toLowerCase().includes('disposition') || k.toLowerCase().includes('status'));
            if (!statusKey) statusKey = 'Status';

            return { callerIdKey, dateKey, extKey, statusKey };
        };

        // Parse Master File
        const masterWorkbook = xlsx.read(masterFile.buffer, { type: 'buffer', cellDates: true });
        const masterSheetName = masterWorkbook.SheetNames[0];
        const masterData = xlsx.utils.sheet_to_json(masterWorkbook.Sheets[masterSheetName], { defval: "" });

        const masterMap = new Map();
        for (const row of masterData) {
            const keys = getKeys(row);
            const num = sanitizeId(row[keys.callerIdKey]);
            if (num) {
                if (!masterMap.has(num)) masterMap.set(num, []);
                masterMap.get(num).push({
                    original: row,
                    date: normalizeDate(row[keys.dateKey]),
                    ext: row[keys.extKey] || '',
                    status: row[keys.statusKey] || '',
                    numberRaw: row[keys.callerIdKey]
                });
            }
        }

        // Parse Small Files
        const smallMap = new Map();
        const smallFilesData = [];

        for (let i = 0; i < smallFiles.length; i++) {
            const file = smallFiles[i];
            const label = labels[i] || `File ${i + 1}`;
            const originalName = file.originalname || label;

            const smallWorkbook = xlsx.read(file.buffer, { type: 'buffer', cellDates: true });
            const smallData = xlsx.utils.sheet_to_json(smallWorkbook.Sheets[smallWorkbook.SheetNames[0]], { defval: "" });
            
            smallFilesData.push({ originalName, data: smallData, label });

            for (const row of smallData) {
                const keys = getKeys(row);
                const num = sanitizeId(row[keys.callerIdKey]);
                if (num) {
                    if (!smallMap.has(num)) smallMap.set(num, []);
                    smallMap.get(num).push({
                        original: row,
                        date: normalizeDate(row[keys.dateKey]),
                        ext: row[keys.extKey] || '',
                        status: row[keys.statusKey] || '',
                        numberRaw: row[keys.callerIdKey],
                        source: label
                    });
                }
            }
        }

        const auditData = [];
        let matchedCount = 0;
        let notFoundCount = 0; // Missing
        let extraCount = 0;

        // Process Master -> Mathced & Missing
        for (const [num, mRows] of masterMap.entries()) {
            const sRows = smallMap.get(num);
            const isDuplicateMaster = mRows.length > 1;
            const isDuplicateSmall = sRows && sRows.length > 1;
            
            let duplicateNote = '';
            if (isDuplicateMaster && isDuplicateSmall) duplicateNote = 'Duplicate in Both';
            else if (isDuplicateMaster) duplicateNote = 'Duplicate in Master';
            else if (isDuplicateSmall) duplicateNote = 'Duplicate in Small File';

            if (sRows && sRows.length > 0) {
                // MATCHED
                matchedCount += mRows.length;
                for (let i = 0; i < mRows.length; i++) {
                    const mRow = mRows[i];
                    const sRow = sRows[i % sRows.length]; // cyclical if mismatch in counts
                    
                    auditData.push({
                        'Date': mRow.date || sRow.date,
                        'Extension': mRow.ext || sRow.ext,
                        'Number': mRow.numberRaw || sRow.numberRaw,
                        'Match Status': 'MATCHED ✅',
                        'Source': `Master + ${sRow.source}`,
                        'Master Status': mRow.status,
                        'Small File Status': sRow.status,
                        'Duplicate Note': duplicateNote
                    });
                }
            } else {
                // MISSING (In Master but not in Small)
                notFoundCount += mRows.length;
                for (const mRow of mRows) {
                    auditData.push({
                        'Date': mRow.date,
                        'Extension': mRow.ext,
                        'Number': mRow.numberRaw,
                        'Match Status': 'MISSING ❌',
                        'Source': 'Master Only',
                        'Master Status': mRow.status,
                        'Small File Status': 'N/A',
                        'Duplicate Note': duplicateNote
                    });
                }
            }
        }

        // Process Small -> Extra
        for (const [num, sRows] of smallMap.entries()) {
            if (!masterMap.has(num)) {
                // EXTRA (In Small but not in Master)
                extraCount += sRows.length;
                const isDuplicateSmall = sRows.length > 1;
                const duplicateNote = isDuplicateSmall ? 'Duplicate in Small File' : '';
                
                for (const sRow of sRows) {
                    auditData.push({
                        'Date': sRow.date,
                        'Extension': sRow.ext,
                        'Number': sRow.numberRaw,
                        'Match Status': 'EXTRA ⚠️',
                        'Source': sRow.source,
                        'Master Status': 'N/A',
                        'Small File Status': sRow.status,
                        'Duplicate Note': duplicateNote
                    });
                }
            }
        }

        // Create Excel with ExcelJS
        const newWorkbook = new ExcelJS.Workbook();
        
        // 1. Create Audit Report Sheet
        const auditSheet = newWorkbook.addWorksheet('Audit Report');
        if (auditData.length > 0) {
            const headers = Object.keys(auditData[0]);
            auditSheet.columns = headers.map(header => ({ header: header, key: header, width: 22 }));
            
            auditData.forEach(rowData => {
                const row = auditSheet.addRow(rowData);
                if (rowData['Match Status'] === 'MATCHED ✅') {
                    row.eachCell((cell) => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } }; cell.font = { color: { argb: 'FF006100' } }; });
                } else if (rowData['Match Status'] === 'MISSING ❌') {
                    row.eachCell((cell) => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } }; cell.font = { color: { argb: 'FF9C0006' } }; });
                } else if (rowData['Match Status'] === 'EXTRA ⚠️') {
                    row.eachCell((cell) => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } }; cell.font = { color: { argb: 'FF9C6500' } }; });
                }
            });
            auditSheet.getRow(1).font = { bold: true };
            
            // Add AutoFilter
            auditSheet.autoFilter = {
                from: { row: 1, column: 1 },
                to: { row: auditData.length + 1, column: headers.length }
            };
        }

        const excelBuffer = await newWorkbook.xlsx.writeBuffer();
        
        const masterFileResponse = {
            fileName: 'Master_Audit_Report.xlsx',
            fileBase64: excelBuffer.toString('base64'),
            previewData: auditData
        };

        res.json({
            success: true,
            stats: {
                totalMaster: masterData.length,
                matched: matchedCount,
                notFound: notFoundCount,
                extra: extraCount
            },
            masterFile: masterFileResponse,
            smallFiles: [] // Deprecated annotated small files for audit mode
        });

    } catch (error) {
        console.error('Error processing files:', error);
        res.status(500).json({ error: 'Failed to process files.' });
    }
});\n