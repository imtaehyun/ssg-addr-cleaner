var request = require('request'),
	Promise = require('promise'),
	Sequence = require('sequence').Sequence,
	sequence = Sequence.create(),
	SpreadsheetReader = require('pyspreadsheet').SpreadsheetReader,
	SpreadsheetWriter = require('pyspreadsheet').SpreadsheetWriter;

// Reader / Writer
var INPUT_FILE = 'input/input_20150326.xlsx',
	OUTPUT_FILE = 'output/output_20150326',
	writer = new SpreadsheetWriter(OUTPUT_FILE + '_0.xlsx');

var FILE_LENGTH = 50000;
var currentRow, totalCount;

var requestUAPI = function (queryAddr) {
	return new Promise(function (resolve, reject) {
		request({
			method: 'POST',
			headers: {
				'Accept': 'application/json',
				'Content-Type': 'application/json'
			},
			uri: 'http://dev-uapi.ssglocal.com/zip/send.json',
			json: {
				'sendZip': { 'addr': queryAddr }
			}
		}, function (err, res, body) {
			if (err || body.resultZip.resultCode !== '00') {
				console.error('request err: ' + err);
				reject(err);
			} else {
				resolve(body.resultZip.data);
			}
		});
	});
}

var processSync = function (currentRow, row) {
	sequence
	.then(function (next) {
		var queryAddr = row[6].value + ' ' + row[7].value;
		// console.log(currentRow + '/' + totalCount + '\t' + queryAddr);
		requestUAPI(queryAddr)
		.then(function (result) {
			next(result);
		});
	})
	.then(function (next, result) {
		if (result == null) console.error('request error');
		writer.write(((currentRow - 1) % FILE_LENGTH), 0, [currentRow, row[0].value, row[1].value, row[2].value, row[3].value, row[4].value, row[5].value, row[6].value, row[7].value, row[8].value, result.CL_ZIPCODE]);
		next(currentRow);
	})
	.then(function (next, currentRow) {
		if (currentRow % FILE_LENGTH == 0) {
			writer.save(function (err) {
				if (err) console.error('excel save error');
				console.timeEnd('excel');
				console.time('excel');
				console.log(Math.floor(currentRow/FILE_LENGTH) + '/' + Math.ceil(totalCount/FILE_LENGTH) + ' excel file saved');
				writer = new SpreadsheetWriter(OUTPUT_FILE + '_' + Math.floor(currentRow/FILE_LENGTH) + '.xlsx');
				next();
			});
		} else if (currentRow == totalCount) {
			writer.save(function (err) {
				if (err) console.error('excel save error');
				console.log('excel file saved');
				console.timeEnd('excel');
				next();
			});
		} else {
			next();
		}
	});
}

SpreadsheetReader.read(INPUT_FILE, function (err, workbook) {
	if (err) throw err;
	console.time('excel');
	console.log('excel file read finish');
	workbook.sheets.forEach(function (sheet) {
		currentRow = 0;
		totalCount = sheet.rows.length;

		sheet.rows.forEach(function (row) {
			processSync(++currentRow, row);
		});
	});
});
