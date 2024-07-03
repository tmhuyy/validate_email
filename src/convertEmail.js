const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const _ = require('lodash');
const xlsxToJson = (fileName, jsonFileName) =>
{
    let workbook = XLSX.readFile(path.resolve(__dirname, fileName));
    let Emails = [];
  
    workbook.SheetNames.forEach((sheetName) => {
      let sheet = workbook.Sheets[sheetName];
      let contents = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  
      let headerRow = contents[0];
      let EmailIndex = headerRow.indexOf('Email');
  
      if (EmailIndex !== -1) {
        for (let row = 1; row < contents.length; row++) {
          let Email = contents[row][EmailIndex];
          if (Email) {
            Emails.push(Email);
          }
        }
      }
    });
  
    fs.writeFileSync(jsonFileName, JSON.stringify(Emails, null, 2));
  };
  
xlsxToJson('Export-All-Users-Data.xlsx', 'email.json');