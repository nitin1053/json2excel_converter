const fs = require('fs');
const XLSX = require('xlsx');

function flattenObject(obj, parentKey = '') {
    let result = {};
    for (const key in obj) {
        if (obj.hasOwnProperty(key)) {
            const newKey = parentKey ? `${parentKey}_${key}` : key;
            if (typeof obj[key] === 'object' && !Array.isArray(obj[key])) {
                result = { ...result, ...flattenObject(obj[key], newKey) };
            } else {
                result[newKey] = obj[key];
            }
        }
    }
    return result;
}

function addHyperlink(sheet, cell, sheetName) {
    const hyperlinkFormula = `=HYPERLINK("#'${sheetName}'!A1", "SHEET::${sheetName}")`;
    sheet[cell] = { f: hyperlinkFormula };
}

function jsonToExcel(jsonFile, excelFile) {
    try {
        const jsonData = JSON.parse(fs.readFileSync(jsonFile, 'utf8'));
        const workbook = XLSX.utils.book_new();

        for (const key in jsonData) {
            if (jsonData.hasOwnProperty(key)) {
                const sheetName = key.replace(/\s+/g, '_');
                const flattenedData = flattenObject(jsonData[key]);

                const ws = XLSX.utils.json_to_sheet([flattenedData]);
                XLSX.utils.book_append_sheet(workbook, ws, sheetName);

                // Add hyperlink in cell A1 to link to the corresponding sheet
                addHyperlink(ws, 'A1', sheetName);
            }
        }

        XLSX.writeFile(workbook, excelFile);

        console.log(`Conversion successful. Excel file saved as ${excelFile}`);
    } catch (error) {
        if (error.code === 'ENOENT') {
            console.error(`Error: File '${jsonFile}' not found.`);
        } else {
            console.error(`Error: ${error.message}`);
        }
    }
}

// Replace 'input.json' and 'output.xlsx' with your file names
jsonToExcel('input.json', 'output.xlsx');
