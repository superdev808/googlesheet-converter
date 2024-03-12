const fs = require("fs");
const recursive = require("recursive-readdir");
const makeDir = require("make-dir");
const path = require("path");
const xlsx = require("xlsx");

require("dotenv").config();

function readFile(filePath) {
    const resolvedPath = path.resolve(filePath);
    const fileExtension = path.extname(resolvedPath).toLowerCase();

    switch (fileExtension) {
        case ".csv":
            return parseCsvToJSON(resolvedPath);
        case ".xlsx":
        case ".xls":
            return parseExcelToJSON(resolvedPath);
        default:
            console.log("Unsupported file format");
            return [];
    }
}

function isNotExcelFile(filePath) {
    return (
        [".csv", ".xlsx", ".xls"].indexOf(
            path.extname(filePath).toLowerCase()
        ) < 0
    );
}

function parseExcelToJSON(filePath) {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    return xlsx.utils.sheet_to_json(worksheet, { defval: "" });
}

async function main() {
    const filePaths = await recursive(process.env.XLSX_DIR, [isNotExcelFile]);
    await makeDir(process.env.JSON_DIR);
    console.log(filePaths);
    filePaths.forEach((filePath) => {
        parseExcelToJSON(filePath);
        const fileName = path.basename(filePath);
        const outputPath = path.join(
            process.env.JSON_DIR,
            fileName.slice(0, fileName.lastIndexOf(".")) + ".json"
        );
        fs.writeFileSync(
            outputPath,
            JSON.stringify(parseExcelToJSON(filePath), null, 2)
        );
        console.log(`${fileName} writing done`);
    });
}

main().catch(console.error);
