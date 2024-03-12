const fs = require('fs');
const readline = require('readline').createInterface({
  input: process.stdin,
  output: process.stdout
});
const path = require('path');
const xlsx = require('xlsx');
const dayjs = require('dayjs');

function askQuestion(query) {
    return new Promise((resolve) => readline.question(query, resolve));
}

function readFile(filePath) {
    const resolvedPath = path.resolve(filePath);
    const fileExtension = path.extname(resolvedPath).toLowerCase();

    switch (fileExtension) {
        case '.csv':
            return parseCsvToJSON(resolvedPath);
        case '.xlsx':
        case '.xls':
            return parseExcelToJSON(resolvedPath);
        default:
            console.log('Unsupported file format');
            return [];
    }
}

function parseCsvToJSON(filePath) {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    return xlsx.utils.sheet_to_json(worksheet, { defval: "" });
}

function parseExcelToJSON(filePath) {
    return parseCsvToJSON(filePath); // xlsx.readFile works for both Excel and CSV formats
}

async function main() {
    const filePaths = [
        "Bone Reduction Instruments - All Brands",
        "Chairside Pick-Up Materials - All Brands",
        "Drill Kit and Drill Sequences - All Brands",
        "Drivers (Restorative, Multi-Unit Abutments) - All Brands",
        "Healing Abutments - All Brands ",
        "Implant Analogs - All Brands",
        "Implant Drivers (Restorative, Direct to Implant) - All Brands",
        "Implant Screws - All Brands",
        "Implants - All Brands ",
        "Impression Copings (Direct to Implant) - All Brands ",
        "Impression Copings (MUAs) - All Brands ",
        "MUAs - All Brands ",
        "Scanbodies (Multi-Unit Abutments) - All Brands",
        "Scanbodies (Single Unit) - All Brands",
        "Scanbody Drivers (Direct to Implant) - All Brands",
        "Scanbody Drivers (MUAs) - All Brands ",
        "Stock Abutments - All Brands ",
        "Temporary Copings (Direct to Implant) All Brands",
        "Temporary Copings (MUAs) - All Brands",
        "Ti Base Calculator (Direct to Implant) - All Brands"
    ]
    // const filePath = await askQuestion('Input the path to the file (.csv, .xlsx, .xls): ');
    // readline.close();
    // filePaths.forEach((filePath) => {
        const jsonData = readFile(`data/Bone Reduction Instruments - All Brands.xlsx`);
        const outputPath = `output/Bone Reduction Instruments - All Brands.json`;
        fs.writeFileSync(outputPath, JSON.stringify(jsonData, null, 2));
        console.info(`Result is saved in ${outputPath}`);
        readline.close();
    // })
}

main().catch(console.error);
