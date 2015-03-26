var fs = require('fs');
var Excel = require('exceljs');
var workbook = new Excel.Workbook();

// console.time('excel');
// workbook.xlsx.readFile('input/target_50000.xlsx')
// 	.then(function() {
// 		var worksheet = workbook.getWorksheet(1);
// 		worksheet.eachRow(function(row, rowNumber) {
// 			console.log("row " + rowNumber + "=" + JSON.stringify(row.values));
// 		});
// 		console.timeEnd('excel');
// 	}); 
var stream = fs.createReadStream('input/target_50000.xlsx');
stream.pipe(workbook.xlsx.createInputStream());