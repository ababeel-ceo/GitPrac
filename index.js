//  another Modification by Abd
// modifield by sam->abd
// const fs=require("fs"
);
// const xlsx=require("xlsx");
// const jsontoxml=require("jsontoxml");
// const excel_file=xlsx.readFile("demo.xlsx");
// let worksheet={};
// for(const $sheetName of excel_file.SheetNames)
// {
//     worksheet[$sheetName]=xlsx.utils.sheet_to_json(excel_file.Sheets[$sheetName]);
// }
// console.log(worksheet.firstsheet[0].name);
const fs=require('fs')
const {google}=require('googleapis');
const {GoogleAuth}=require('google-auth-library');
const { get } = require('http');
const xlsx=require("xlsx");
const readXlsxFile = require('read-excel-file/node');
const SCOPE=["https://www.googleapis.com/auth/spreadsheets"]

async function transferExcelData() {
    // Replace the variables in this block with real values.
    //const spreadsheetId = '1i91PxHALnlrmhuq0eimWs7URWJL281klyl3gqd6EsMo';
    const spreadsheetId = '1dNdeDvfHKQT9vH3zU2Xo8hAw2TRnvDMgQaGWNleRMfI';
    const range = 'Sheet1!A:C';
    const valueInputOption = 'RAW';
  
    // Read the Excel file
    const rows = await readXlsxFile('demo.xlsx');
    console.log(rows);
    const auth=new GoogleAuth({
                keyFile:"eng-scene-374011-f19eab70b8e3.json",
                scopes:SCOPE
            });
    const sheets = google.sheets({version: 'v4', auth});
    
    try {
      const resource = {
        values: rows,
      };
      const response = await sheets.spreadsheets.values.update({
        spreadsheetId,
        range,
        valueInputOption,
        resource,
      });
      console.log(`Success: ${response.status}`);
    } catch (error) {
      console.error(error);
    }
  }
  
  transferExcelData();

// async function showAutomate()
// {
//     const auth=new GoogleAuth({
//         keyFile:"eng-scene-374011-f19eab70b8e3.json",
//         scopes:SCOPE
//     });
//     const client=await auth.getClient();
//     const sheets=google.sheets({version:"v4",auth:client});
//     const getRows=await sheets.spreadsheets.values.get({
//         auth,
//         spreadsheetId:"1dNdeDvfHKQT9vH3zU2Xo8hAw2TRnvDMgQaGWNleRMfI",
//         range:"Sheet1!A:D",
//     });
//     const keyInSheet=[...getRows.data.values.flat()];
//     console.log(getRows.data.values)
// }
// async function insertData() {
//     // Replace the variables in this block with real values.
//     const spreadsheetId = '1dNdeDvfHKQT9vH3zU2Xo8hAw2TRnvDMgQaGWNleRMfI';
//     const range = 'Sheet1!A1:B2';
//     const valueInputOption = 'RAW';
//     const insertDataOption = 'INSERT_ROWS'; 
//     const auth=new GoogleAuth({
//         keyFile:"eng-scene-374011-f19eab70b8e3.json",
//         scopes:SCOPE
//     });
//     const sheets = google.sheets({version: 'v4', auth});
//     const values = xlsx.readFile("demo.xlsx");
  
//     const resource = {
//       values,
//     };
  
//     try {
//       const response = await sheets.spreadsheets.values.append({
//         spreadsheetId,
//         range,
//         valueInputOption,
//         insertDataOption,
//         resource,
//       });
//       console.log(`Success: ${response.status}`);
//     } catch (error) {
//       console.error(error);
//     }
//   }
  
//   insertData();
