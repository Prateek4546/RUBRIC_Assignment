const fs = require('fs');
const json2xls = require('json2xls');

function convertToExcel(data, parentFileName = 'output') {
  const xls = json2xls(data);
  const outputFileName = `${parentFileName}.xlsx`;
  fs.writeFileSync(outputFileName, xls, 'binary');
  console.log(`Excel file '${outputFileName}' created successfully!`);
}

function processNestedObjects(obj, parentFileName) {
  for (const key in obj) {
    if (typeof obj[key] === 'object' && obj[key] !== null) {
      const nestedFileName = `${parentFileName}_${key}`;
      convertToExcel(obj[key], nestedFileName);
      processNestedObjects(obj[key], nestedFileName);
    }
  }
}

function main() {
  const jsonFilePath = 'sample.json';

  try {
    const jsonData = JSON.parse(fs.readFileSync(jsonFilePath, 'utf8'));
    convertToExcel(jsonData);

    // Process nested objects recursively
    processNestedObjects(jsonData, 'output');
  } catch (error) {
    console.error('Error processing JSON:', error.message);
  }
}

main();
