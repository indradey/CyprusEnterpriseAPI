var EXCEL = require('exceljs');
var fs = require('fs');

var currentDirectory = __dirname;
/*
sample execution
var uuid = require('uuid/v4');
var userData = {};
userData.RefNo = uuid();
userData.Name = 'Harish';
userData.ContactNo = '9844875454';
userData.Email = 'harishchand@gmail.com';
userData.Query = 'This is a sample Query';
userData.Response = 'This is sample Response';

writeToFile(userData);
*/

async function readFromFile() {
    var workbook = new EXCEL.Workbook();
    var filelocation = currentDirectory + '\\Queries.xlsx';
    var responseData = [];
    await workbook.xlsx.readFile(filelocation)
        .then(exceldata => {
            var worksheet = workbook.getWorksheet(1);
            var rowcount = worksheet.rowCount;
            for (var i = 2; i <= rowcount; i++) {
                var row = worksheet.getRow(i);
                var respjson = {};
                respjson.Reqid = row.getCell(1).value;
                respjson.RefNo = row.getCell(2).value;
                respjson.Name = row.getCell(3).value;
                respjson.ContactNo = row.getCell(4).value;
                respjson.Email = row.getCell(5).value;
                respjson.Query = row.getCell(6).value;
                respjson.Response = row.getCell(7).value;
                responseData.push(respjson);
            }
        });
    return responseData;
}

async function writeToFile(userData) {
    var workbook = new EXCEL.Workbook();
    var filelocation = currentDirectory + '\\Queries.xlsx';
    await workbook.xlsx.readFile(filelocation)
        .then(function () {
            var worksheet = workbook.getWorksheet(1);
            var rowcount = worksheet.rowCount;
            var row = worksheet.getRow(rowcount + 1);
            row.getCell(1).value = rowcount; // A5's value set to 5
            row.getCell(2).value = userData.RefNo;
            row.getCell(3).value = userData.Name;
            row.getCell(4).value = userData.ContactNo;
            row.getCell(5).value = userData.Email;
            row.getCell(6).value = userData.Query;
            row.getCell(7).value = userData.Response;
            row.commit();
            return workbook.xlsx.writeFile(filelocation);
        })
}

module.exports.writeToFile = writeToFile;
module.exports.readFromFile = readFromFile;