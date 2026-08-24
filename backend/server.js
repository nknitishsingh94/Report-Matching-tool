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

        const normalizeTime = (val) => {
            if (!val) return '';
            if (val instanceof Date) {
                if (!isNaN(val)) return val.toISOString().split('T')[1].substring(0, 5);
                return '';
            }
            if (typeof val === 'number') {
                const date = new Date(Math.round((val - 25569) * 86400 * 1000));
                if (!isNaN(date)) return date.toISOString().split('T')[1].substring(0, 5);
                return '';
            }
            const str = String(val).trim();
            const d = new Date(str);
            if (!isNaN(d)) return d.toISOString().split('T')[1].substring(0, 5);
            
            const timeMatch = str.match(/\b([01]\d|2[0-3]):([0-5]\d)/);
            if (timeMatch) return timeMatch[0];
            return '';
        };

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

        const mapTimeZone = (tz) => {
            if (!tz) return '';
            const upperTz = String(tz).toUpperCase().trim();
            
            const tzMap = {
                // North America (US & Canada)
                'EDT': 'EDT (USA/Canada - Eastern)',
                'EST': 'EST (USA/Canada - Eastern)',
                'CDT': 'CDT (USA/Canada - Central)',
                'CST': 'CST (USA/Canada - Central)',
                'MDT': 'MDT (USA/Canada - Mountain)',
                'MST': 'MST (USA/Canada - Mountain)',
                'PDT': 'PDT (USA/Canada - Pacific)',
                'PST': 'PST (USA/Canada - Pacific)',
                'AST': 'AST (Canada - Atlantic)',
                'ADT': 'ADT (Canada - Atlantic)',
                'AKST': 'AKST (USA - Alaska)',
                'AKDT': 'AKDT (USA - Alaska)',
                'HST': 'HST (USA - Hawaii)',
                
                // Europe
                'GMT': 'GMT (United Kingdom/Europe)',
                'BST': 'BST (United Kingdom)',
                'CET': 'CET (Central Europe)',
                'CEST': 'CEST (Central Europe)',
                'EET': 'EET (Eastern Europe)',
                'EEST': 'EEST (Eastern Europe)',
                'WET': 'WET (Western Europe)',
                'WEST': 'WEST (Western Europe)',
                
                // Asia
                'IST': 'IST (India)',
                'JST': 'JST (Japan)',
                'KST': 'KST (South Korea)',
                'SGT': 'SGT (Singapore)',
                'HKT': 'HKT (Hong Kong)',
                'GST': 'GST (Gulf/UAE)',
                'PKT': 'PKT (Pakistan)',
                'PHT': 'PHT (Philippines)',
                'WIB': 'WIB (Indonesia - Western)',
                
                // Oceania
                'AEST': 'AEST (Australia - Eastern)',
                'AEDT': 'AEDT (Australia - Eastern)',
                'ACST': 'ACST (Australia - Central)',
                'ACDT': 'ACDT (Australia - Central)',
                'AWST': 'AWST (Australia - Western)',
                'NZST': 'NZST (New Zealand)',
                'NZDT': 'NZDT (New Zealand)',
                
                // Africa
                'SAST': 'SAST (South Africa)',
                'EAT': 'EAT (East Africa)',
                'CAT': 'CAT (Central Africa)',
                'WAT': 'WAT (West Africa)',
                
                // South America
                'ART': 'ART (Argentina)',
                'BRT': 'BRT (Brazil)',
                'PET': 'PET (Peru)',
                'COT': 'COT (Colombia)',
                'CLT': 'CLT (Chile)',
                'VET': 'VET (Venezuela)'
            };

            for (const [key, val] of Object.entries(tzMap)) {
                if (upperTz === key) return val;
                if (upperTz.includes(key)) return upperTz.replace(new RegExp(key, 'ig'), val);
            }
            return tz;
        };

        const sanitizeId = (id) => {
            if (id === undefined || id === null) return '';
            let cleaned = String(id).toLowerCase().replace(/[^a-z0-9]/g, '');
            if (/^\d+$/.test(cleaned) && cleaned.length > 10) {
                return cleaned.slice(-10);
            }
            return cleaned;
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
            
            let tzKey = Object.keys(row).find(k => k.toLowerCase().includes('zone') || k.toLowerCase().includes('tz'));
            if (!tzKey) tzKey = 'Time Zone';

            let timeKey = Object.keys(row).find(k => k.toLowerCase() === 'time' || k.toLowerCase().includes(' time'));
            return { callerIdKey, dateKey, extKey, statusKey, tzKey, timeKey };
        };

        const masterWorkbook = xlsx.read(masterFile.buffer, { type: 'buffer', cellDates: true });
        const masterSheetName = masterWorkbook.SheetNames[0];
        const masterData = xlsx.utils.sheet_to_json(masterWorkbook.Sheets[masterSheetName], { defval: "" });

        const masterMap = new Map();
        if (masterData.length > 0) {
            const keys = getKeys(masterData[0]);
            for (const row of masterData) {
                const num = sanitizeId(row[keys.callerIdKey]);
                if (num) {
                    if (!masterMap.has(num)) masterMap.set(num, []);
                    masterMap.get(num).push({
                        original: row,
                        date: normalizeDate(row[keys.dateKey]),
                        time: (keys.timeKey && row[keys.timeKey]) ? normalizeTime(row[keys.timeKey]) : normalizeTime(row[keys.dateKey]),
                        ext: row[keys.extKey] || '',
                        status: row[keys.statusKey] || '',
                        numberRaw: row[keys.callerIdKey],
                        tz: mapTimeZone(row[keys.tzKey] || '')
                    });
                }
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

            if (smallData.length > 0) {
                const keys = getKeys(smallData[0]);
                for (const row of smallData) {
                    const num = sanitizeId(row[keys.callerIdKey]);
                    if (num) {
                        if (!smallMap.has(num)) smallMap.set(num, []);
                        smallMap.get(num).push({
                            original: row,
                            date: normalizeDate(row[keys.dateKey]),
                            ext: row[keys.extKey] || '',
                            status: row[keys.statusKey] || '',
                            numberRaw: row[keys.callerIdKey],
                            source: label,
                            tz: mapTimeZone(row[keys.tzKey] || '')
                        });
                    }
                }
            }
        }

        // Track small file reports
        const smallFileReports = {};
        for (let i = 0; i < smallFiles.length; i++) {
            const label = labels[i] || `File ${i + 1}`;
            smallFileReports[label] = {
                originalName: smallFiles[i].originalname || label,
                matched: 0,
                notFound: 0, // Unmatched
                data: []
            };
        }

        const auditData = [];
        let matchedCount = 0;
        let notFoundCount = 0; // Missing
        let extraCount = 0;

        const allNumbers = new Set([...masterMap.keys(), ...smallMap.keys()]);

        for (const num of allNumbers) {
            const mRows = masterMap.get(num) || [];
            const sRows = smallMap.get(num) || [];
            
            const isDuplicateMaster = mRows.length > 1;
            const isDuplicateSmall = sRows.length > 1;
            
            let duplicateNote = '';
            if (isDuplicateMaster && isDuplicateSmall) duplicateNote = 'Duplicate in Both';
            else if (isDuplicateMaster) duplicateNote = 'Duplicate in Master';
            else if (isDuplicateSmall) duplicateNote = 'Duplicate in Small File';

            const matchLimit = Math.min(mRows.length, sRows.length);
            const missingLimit = Math.max(0, mRows.length - sRows.length);
            const extraLimit = Math.max(0, sRows.length - mRows.length);

            // MATCHED
            for (let i = 0; i < matchLimit; i++) {
                matchedCount++;
                const mRow = mRows[i];
                const sRow = sRows[i];
                
                auditData.push({
                    'Date': mRow.date || sRow.date,
                    'Time': mRow.time || sRow.time,
                    'Time Zone': mRow.tz || sRow.tz,
                    'Extension': mRow.ext || sRow.ext,
                    'Number': mRow.numberRaw || sRow.numberRaw,
                    'Match Status': 'MATCHED ✅',
                    'Source': `Master + ${sRow.source}`,
                    'Master Status': mRow.status,
                    'Small File Status': sRow.status,
                    'Duplicate Note': duplicateNote
                });

                if (smallFileReports[sRow.source]) {
                    smallFileReports[sRow.source].matched++;
                    smallFileReports[sRow.source].data.push({
                        ...sRow.original,
                        'Match Status': 'MATCHED ✅',
                        'Duplicate Note': duplicateNote
                    });
                }
            }

            // MISSING (In Master but not paired with Small)
            for (let i = matchLimit; i < matchLimit + missingLimit; i++) {
                notFoundCount++;
                const mRow = mRows[i];
                
                auditData.push({
                    'Date': mRow.date,
                    'Time': mRow.time,
                    'Time Zone': mRow.tz,
                    'Extension': mRow.ext,
                    'Number': mRow.numberRaw,
                    'Match Status': 'UNMATCHED ❌',
                    'Source': 'Master Only',
                    'Master Status': mRow.status,
                    'Small File Status': 'N/A',
                    'Duplicate Note': duplicateNote
                });
            }

            // EXTRA (In Small but not paired with Master)
            for (let i = matchLimit; i < matchLimit + extraLimit; i++) {
                extraCount++;
                const sRow = sRows[i];
                
                auditData.push({
                    'Date': sRow.date,
                    'Time': sRow.time,
                    'Time Zone': sRow.tz,
                    'Extension': sRow.ext,
                    'Number': sRow.numberRaw,
                    'Match Status': 'EXTRA ⚠️',
                    'Source': sRow.source,
                    'Master Status': 'N/A',
                    'Small File Status': sRow.status,
                    'Duplicate Note': duplicateNote
                });

                if (smallFileReports[sRow.source]) {
                    smallFileReports[sRow.source].notFound++;
                    smallFileReports[sRow.source].data.push({
                        ...sRow.original,
                        'Match Status': 'UNMATCHED ❌',
                        'Duplicate Note': duplicateNote
                    });
                }
            }
        }

        // Create Excel with ExcelJS
        const newWorkbook = new ExcelJS.Workbook();
        
        // 1. Create Audit Report Sheet
        const auditSheet = newWorkbook.addWorksheet('Audit Report');
        
        const fillMatched = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
        const fontMatched = { color: { argb: 'FF006100' } };
        const fillUnmatched = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
        const fontUnmatched = { color: { argb: 'FF9C0006' } };
        const fillExtra = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } };
        const fontExtra = { color: { argb: 'FF9C6500' } };

        if (auditData.length > 0) {
            const headers = Object.keys(auditData[0]);
            auditSheet.columns = headers.map(header => ({ header: header, key: header, width: 22 }));
            
            auditData.forEach(rowData => {
                const row = auditSheet.addRow(rowData);
                if (rowData['Match Status'] === 'MATCHED ✅') {
                    row.eachCell((cell) => { cell.fill = fillMatched; cell.font = fontMatched; });
                } else if (rowData['Match Status'] === 'UNMATCHED ❌') {
                    row.eachCell((cell) => { cell.fill = fillUnmatched; cell.font = fontUnmatched; });
                } else if (rowData['Match Status'] === 'EXTRA ⚠️') {
                    row.eachCell((cell) => { cell.fill = fillExtra; cell.font = fontExtra; });
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

        // Generate small files reports
        const smallFilesResponse = [];
        for (const [label, report] of Object.entries(smallFileReports)) {
            const sfWorkbook = new ExcelJS.Workbook();
            const sfSheet = sfWorkbook.addWorksheet('Report');
            
            if (report.data.length > 0) {
                const headers = Object.keys(report.data[0]);
                sfSheet.columns = headers.map(header => ({ header: header, key: header, width: 22 }));
                
                report.data.forEach(rowData => {
                    const row = sfSheet.addRow(rowData);
                    if (rowData['Match Status'] === 'MATCHED ✅') {
                        row.eachCell((cell) => { cell.fill = fillMatched; cell.font = fontMatched; });
                    } else if (rowData['Match Status'] === 'UNMATCHED ❌') {
                        row.eachCell((cell) => { cell.fill = fillUnmatched; cell.font = fontUnmatched; });
                    }
                });
                sfSheet.getRow(1).font = { bold: true };
                
                sfSheet.autoFilter = {
                    from: { row: 1, column: 1 },
                    to: { row: report.data.length + 1, column: headers.length }
                };
            }
            
            const sfBuffer = await sfWorkbook.xlsx.writeBuffer();
            const originalNameExt = report.originalName.includes('.') ? report.originalName.substring(0, report.originalName.lastIndexOf('.')) : report.originalName;
            
            smallFilesResponse.push({
                fileName: `Audit_${originalNameExt}.xlsx`,
                fileBase64: sfBuffer.toString('base64'),
                previewData: report.data,
                matched: report.matched,
                notFound: report.notFound
            });
        }

        res.json({
            success: true,
            stats: {
                totalMaster: masterData.length,
                matched: matchedCount,
                notFound: notFoundCount,
                extra: extraCount
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
            
            const fillVoip = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
            const fontVoip = { color: { argb: 'FF9C0006' } };
            const fillOther = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
            const fontOther = { color: { argb: 'FF006100' } };

            data.forEach(rowData => {
                const row = updatedSheet.addRow(rowData);
                if (rowData['Line Type'] === 'voip') {
                    row.eachCell(cell => {
                        cell.fill = fillVoip;
                        cell.font = fontVoip;
                    });
                } else {
                    row.eachCell(cell => {
                        cell.fill = fillOther;
                        cell.font = fontOther;
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
