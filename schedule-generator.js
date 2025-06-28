// schedule-generator.js
// --- מודול לניהול ויצירת לוחות זמנים שבועיים ---

const { google } = require('googleapis');
const fs = require('fs');
const { performance } = require('perf_hooks');
const path = require('path');

// =================================================================
// פונקציית ה"מתאם" (Adapter)
// =================================================================

function transformDbDataToScheduleConfig(rawData) {
    const daysConfig = [
        { name: "ראשון", key: "sunday" }, { name: "שני", key: "monday" },
        { name: "שלישי", key: "tuesday" }, { name: "רביעי", key: "wednesday" },
        { name: "חמישי", key: "thursday" }, { name: "שישי", key: "friday" }, { name: "שבת", key: "saturday" },
    ];
    const shiftsConfig = { morning: "בוקר", afternoon: "צהריים", evening: "ערב" };

    const stationIds = new Set();
    if (rawData.schedule) {
        Object.values(rawData.schedule).forEach(dayData => {
            Object.values(dayData).forEach(shiftAssignments => {
                if (Array.isArray(shiftAssignments)) {
                    shiftAssignments.forEach(assignment => stationIds.add(assignment.position));
                }
            });
        });
    }

    const stations = Array.from(stationIds).sort().map(id => {
        const station = {
            id,
            slotsPerShift: rawData.stationSlots?.[id] || 3,
            assignments: {}
        };
        const startDate = new Date(rawData.weekStartDate);

        daysConfig.forEach((day, index) => {
            const currentDate = new Date(startDate);
            currentDate.setDate(startDate.getDate() + index);
            const dateKey = currentDate.toISOString().split('T')[0];
            station.assignments[day.key] = {};
            const dayDataFromDb = rawData.schedule?.[dateKey] || {};

            for (const shiftKey in shiftsConfig) {
                const shiftName = shiftsConfig[shiftKey];
                const assignmentsInShift = dayDataFromDb[shiftName] || [];
                station.assignments[day.key][shiftKey] = assignmentsInShift
                    .filter(a => a.position === id)
                    .map(a => ({ name: a.guard.name, from: a.start, to: a.end || '' }));
            }
        });
        return station;
    });

    return {
        weekTitle: rawData.title,
        weekStartDate: rawData.weekStartDate,
        days: daysConfig,
        shifts: shiftsConfig,
        stations: stations
    };
}

// =================================================================
// פונקציות עזר, בנייה והרכבה
// =================================================================

// (כאן נמצאות כל שאר הפונקציות: getFormat, createCell, generateDatedWeek, buildMainTitle, buildStationBlock וכו')
// ... העתק לכאן את כל הפונקציות מחלקים ג', ד' ו-ה' מהתשובות הקודמות ...
// הנה הן שוב במלואן לנוחיותך:

function getFormat(styleName) {
    const formats = {
        title: { backgroundColor: { red: 0.85, green: 0.85, blue: 0.85 }, textFormat: { fontSize: 28, bold: true }, horizontalAlignment: "CENTER", verticalAlignment: "MIDDLE" },
        dayHeader: { backgroundColor: { red: 0.24, green: 0.24, blue: 0.24 }, textFormat: { fontSize: 11, bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } }, horizontalAlignment: "CENTER", verticalAlignment: "MIDDLE", wrapStrategy: "WRAP" },
        subHeader: { backgroundColor: { red: 0.64, green: 0.64, blue: 0.64 }, textFormat: { fontSize: 11, bold: true }, horizontalAlignment: "CENTER", verticalAlignment: "MIDDLE" },
        stationTitle: { backgroundColor: { red: 0.61, green: 0.76, blue: 0.89 }, textFormat: { fontSize: 18, bold: true }, horizontalAlignment: "CENTER", verticalAlignment: "MIDDLE" },
        shiftName: { backgroundColor: { red: 0.64, green: 0.64, blue: 0.64 }, textFormat: { fontSize: 11, bold: true }, horizontalAlignment: "CENTER", verticalAlignment: "MIDDLE" },
        personName: { backgroundColor: { red: 0.95, green: 0.69, blue: 0.51 }, textFormat: { fontSize: 11, bold: true }, horizontalAlignment: "CENTER", verticalAlignment: "MIDDLE" },
        timeCell: { textFormat: { fontSize: 11, bold: true }, horizontalAlignment: "CENTER", verticalAlignment: "MIDDLE" },
        emptySlot: { backgroundColor: { red: 0.88, green: 0.88, blue: 0.88 } },
        separator: { backgroundColor: { red: 0.64, green: 0.64, blue: 0.64 } }
    };
    const baseFormat = formats[styleName] || {};
    return { ...baseFormat, borders: { top: { style: "SOLID", width: 1 }, bottom: { style: "SOLID", width: 1 }, left: { style: "SOLID", width: 1 }, right: { style: "SOLID", width: 1 } } };
}

function createCell(value, styleName, formatOverrides = {}) {
    const baseFormat = getFormat(styleName);
    const finalFormat = { ...baseFormat, ...formatOverrides };
    const cell = { userEnteredFormat: finalFormat };
    if (value !== null && value !== undefined && value !== "") {
        cell.userEnteredValue = { [typeof value === 'number' ? 'numberValue' : 'stringValue']: value };
    }
    return cell;
}

function generateDatedWeek(startDateString, daysConfig) {
    const startDate = new Date(startDateString);
    return daysConfig.map((day, index) => {
        const currentDate = new Date(startDate);
        currentDate.setDate(startDate.getDate() + index);
        const date = String(currentDate.getDate()).padStart(2, '0');
        const month = String(currentDate.getMonth() + 1).padStart(2, '0');
        return { ...day, date: `${date}.${month}` };
    });
}

function buildMainTitle(fullTitle, config, startRowIndex) {
    const numColumns = 1 + (config.days.length * 3);
    const rows = [{ values: [createCell(fullTitle, 'title'), ...Array(numColumns - 1).fill(createCell(null, 'title'))] }];
    const merges = [{ startRowIndex, endRowIndex: startRowIndex + 1, startColumnIndex: 0, endColumnIndex: numColumns }];
    return { rows, merges, newRowIndex: startRowIndex + 1 };
}

function buildStationTitle(station, config, startRowIndex) {
    const numColumns = 1 + (config.days.length * 3);
    const rows = [{ values: [createCell(`עמדה ${station.id}`, 'stationTitle'), ...Array(numColumns - 1).fill(createCell(null, 'stationTitle'))] }];
    const merges = [{ startRowIndex, endRowIndex: startRowIndex + 1, startColumnIndex: 0, endColumnIndex: numColumns }];
    return { rows, merges, newRowIndex: startRowIndex + 1 };
}

function buildScheduleGridHeaders(config, startRowIndex) {
    const merges = [];
    const subHeaderTexts = ['שם המאבטח', 'מ', 'עד'];
    const dayHeaderRow = [createCell(null, 'dayHeader')];
    const subHeaderRow = [createCell(null, 'subHeader')];
    config.datedWeek.forEach((day, i) => {
        const headerText = `${day.name}\n${day.date}`;
        dayHeaderRow.push(createCell(headerText, 'dayHeader'), createCell(null, 'dayHeader'), createCell(null, 'dayHeader'));
        merges.push({ startRowIndex, endRowIndex: startRowIndex + 1, startColumnIndex: 1 + (i * 3), endColumnIndex: 1 + (i * 3) + 3 });
        subHeaderRow.push(...subHeaderTexts.map(text => createCell(text, 'subHeader')));
    });
    const rows = [{ values: dayHeaderRow }, { values: subHeaderRow }];
    return { rows, merges, newRowIndex: startRowIndex + 2 };
}

function buildSeparatorRow(config, startRowIndex) {
    const numColumns = 1 + (config.days.length * 3);
    const rows = [{ values: [createCell(null, 'separator')] }];
    const merges = [{ startRowIndex, endRowIndex: startRowIndex + 1, startColumnIndex: 0, endColumnIndex: numColumns }];
    const dimensionRequests = [{ updateDimensionProperties: { range: { dimension: 'ROWS', startIndex: startRowIndex, endIndex: startRowIndex + 1 }, properties: { pixelSize: 20 }, fields: 'pixelSize' } }];
    return { rows, merges, dimensionRequests, newRowIndex: startRowIndex + 1 };
}

function buildAssignmentSlotCells(person) {
    const emptySlotColor = { backgroundColor: { red: 0.88, green: 0.88, blue: 0.88 } };
    const colorOverride = person ? {} : emptySlotColor;
    return [
        createCell(person?.name || null, person ? 'personName' : 'emptySlot', colorOverride),
        createCell(person?.from || null, 'timeCell', colorOverride),
        createCell(person?.to || null, 'timeCell', colorOverride)
    ];
}

function buildAssignmentRows(station, config, startRowIndex) {
    let allRows = [], allMerges = [], allDimensionRequests = [], currentRowIndex = startRowIndex;
    const shiftKeys = Object.keys(config.shifts);
    shiftKeys.forEach((shiftKey, shiftIndex) => {
        const assignmentsByDay = config.datedWeek.map(day => station.assignments[day.key]?.[shiftKey] || []);
        const configuredSlots = station.slotsPerShift || 3;
        const maxActualAssignments = Math.max(0, ...assignmentsByDay.map(arr => arr.length));
        const numRowsForShift = Math.max(configuredSlots, maxActualAssignments);
        
        if (numRowsForShift > 0) {
            if (numRowsForShift > 1) {
                allMerges.push({ startRowIndex: currentRowIndex, endRowIndex: currentRowIndex + numRowsForShift, startColumnIndex: 0, endColumnIndex: 1 });
            }
            for (let i = 0; i < numRowsForShift; i++) {
                const shiftName = config.shifts[shiftKey];
                const rowCells = [i === 0 ? createCell(shiftName, 'shiftName') : createCell(null, 'shiftName')];
                assignmentsByDay.forEach(peopleOnDay => rowCells.push(...buildAssignmentSlotCells(peopleOnDay[i])));
                allRows.push({ values: rowCells });
            }
            currentRowIndex += numRowsForShift;
        }

        if (shiftIndex < shiftKeys.length - 1) {
            const separatorResult = buildSeparatorRow(config, currentRowIndex);
            allRows.push(...separatorResult.rows);
            allMerges.push(...separatorResult.merges);
            allDimensionRequests.push(...separatorResult.dimensionRequests);
            currentRowIndex = separatorResult.newRowIndex;
        }
    });
    return { rows: allRows, merges: allMerges, dimensionRequests: allDimensionRequests, newRowIndex: currentRowIndex };
}

function buildStationBlock(station, config, startRowIndex) {
    let allRows = [], allMerges = [], allDimensionRequests = [], currentRow = startRowIndex;
    const process = (result) => {
        if (!result || !result.rows) return;
        allRows.push(...result.rows);
        allMerges.push(...result.merges);
        if (result.dimensionRequests) allDimensionRequests.push(...result.dimensionRequests);
        currentRow = result.newRowIndex;
    };
    process(buildStationTitle(station, config, currentRow));
    process(buildScheduleGridHeaders(config, currentRow));
    process(buildAssignmentRows(station, config, currentRow));
    return { rows: allRows, merges: allMerges, dimensionRequests: allDimensionRequests, newRowIndex: currentRow };
}

async function prepareTargetSheet(sheets, spreadsheetId, datedWeek) {
    const newSheetName = `לוז ${datedWeek[0].date} - ${datedWeek[datedWeek.length - 1].date}`;
    console.log(`מכין גיליון יעד: "${newSheetName}"`);
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const requests = [];
    let targetSheetId = null;

    const spreadsheetInfo = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets/properties(sheetId,title)' });
    const year = today.getFullYear();

    for (const sheet of spreadsheetInfo.data.sheets) {
        const title = sheet.properties.title;
        const match = title.match(/(\d{2})\.(\d{2})-(\d{2})\.(\d{2})/);

        if (match) {
            const [, , , dayEnd, monthEnd] = match.map(Number);
            const sheetEndDate = new Date(year, monthEnd - 1, dayEnd);
            if (sheetEndDate < today) {
                console.log(`מזהה גיליון ישן למחיקה: "${title}"`);
                requests.push({ deleteSheet: { sheetId: sheet.properties.sheetId } });
            }
        }
        if (title === newSheetName) {
            targetSheetId = sheet.properties.sheetId;
        }
    }

    if (!targetSheetId) {
        requests.push({ addSheet: { properties: { title: newSheetName, index: 0 } } });
    }

    if (requests.length > 0) {
        const response = await sheets.spreadsheets.batchUpdate({ spreadsheetId, resource: { requests } });
        if (!targetSheetId && response.data.replies.find(r => r.addSheet)) {
            targetSheetId = response.data.replies.find(r => r.addSheet).addSheet.properties.sheetId;
        }
    }

    console.log(`מנקה תוכן מגיליון יעד (ID: ${targetSheetId})...`);
    await sheets.spreadsheets.values.clear({ spreadsheetId, range: newSheetName });
    await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        resource: { requests: [{ unmergeCells: { range: { sheetId: targetSheetId } } }] }
    });
    
    console.log("גיליון היעד מוכן לבנייה.");
    return targetSheetId;
}


// =================================================================
// פונקציית הכניסה הראשית (Entry Point)
// =================================================================
async function generateScheduleFromData(dbData) {
    const startTime = performance.now();
    const targetSpreadsheetId = '1Gj7XvZLzrLudAOdL9flItxto76w39B9XyvdYZLYe33c'; 
    // const keyFilePath = 'config/credentials.json';
    // const keyFilePath = 'credentials.json';

const keyFilePath = process.env.RENDER === 'true'
  ? '/etc/secrets/credentials.json'
  : path.join(__dirname, 'config', 'credentials.json');
    console.log('Looking for credentials at:', keyFilePath);
    
if (!fs.existsSync(keyFilePath)) {
  throw new Error(`Credentials file not found at ${keyFilePath}`);
}

  // ה-try/catch עבר לשרת, כאן אנחנו רוצים שהשגיאה "תיזרק" למעלה
    const scheduleConfig = transformDbDataToScheduleConfig(dbData);
    
    const auth = new google.auth.GoogleAuth({ keyFile: keyFilePath, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
    const sheets = google.sheets({ version: 'v4', auth });
    
    // חישוב תאריך התחלה מתוקן
    const userStartDate = new Date(scheduleConfig.weekStartDate);
    const dayOfWeek = userStartDate.getDay();
    const weekStartDate = new Date(userStartDate);
    weekStartDate.setDate(userStartDate.getDate() - dayOfWeek);
    const correctedStartDateStr = weekStartDate.toISOString().split('T')[0];
    
    const datedWeek = generateDatedWeek(correctedStartDateStr, scheduleConfig.days);
    
    const actualSheetId = await prepareTargetSheet(sheets, targetSpreadsheetId, datedWeek);
    
        const dynamicWeekTitle = ` \u200E( ${datedWeek[0].date} - ${datedWeek[datedWeek.length - 1].date} ) ${scheduleConfig.weekTitle}`;
    const processedConfig = { ...scheduleConfig, datedWeek };
    
    let allRows = [], allMerges = [], allOtherRequests = [], currentRow = 0;
    
    const processBuildResult = (buildResult) => {
        if (!buildResult || !buildResult.rows) return;
        allRows.push(...buildResult.rows);
        allMerges.push(...buildResult.merges);
        if (buildResult.dimensionRequests) allOtherRequests.push(...buildResult.dimensionRequests);
        currentRow = buildResult.newRowIndex;
    };

    allOtherRequests.push({ updateDimensionProperties: { range: { dimension: 'ROWS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 50 }, fields: 'pixelSize' } });
    processBuildResult(buildMainTitle(dynamicWeekTitle, processedConfig, currentRow));
    for (const station of processedConfig.stations) {
        processBuildResult(buildStationBlock(station, processedConfig, currentRow));
    }
    
    const requests = [
        { updateCells: { rows: allRows, fields: 'userEnteredValue,userEnteredFormat', start: { sheetId: actualSheetId, rowIndex: 0, columnIndex: 0 } } },
        ...allMerges.map(merge => ({ mergeCells: { range: { ...merge, sheetId: actualSheetId }, mergeType: 'MERGE_ALL' } })),
        ...allOtherRequests.map(req => {
            req.updateDimensionProperties.range.sheetId = actualSheetId;
            return req;
        })
    ];

    console.log(`בניית הבקשה הסתיימה. שולח לגוגל...`);
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: targetSpreadsheetId, resource: { requests } });
    
    const endTime = performance.now();
    console.log(`הטבלה נוצרה בהצלחה. זמן ביצוע: ${((endTime - startTime) / 1000).toFixed(2)} שניות.`);
}

// ייצוא הפונקציה כדי שהשרת יוכל להשתמש בה
module.exports = {
    generateScheduleFromData
};