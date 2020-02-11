const Excel = require('exceljs')
const fs = require('fs');
const excelfile = "test.xlsx";
var workbook = new Excel.Workbook();

workbook.xlsx.readFile(excelfile).then(function () {
    var worksheet = workbook.getWorksheet(1); //获取第一个worksheet

    var results = new Array();
    for (var i = 0; i < 12; i++) {
        results[i] = new Array();
        for (var j = 0; j < 12; j++) {
            results[i][j] = 0;
        }
    }

    var startDate = convertToDate(worksheet.getCell('A2').value)
    var endDate = convertToDate(worksheet.getCell(worksheet.actualRowCount, 1).value)

    worksheet.eachRow(function (row, rowNumber) {

        if (rowNumber == 1) {
            //Do nothing
        } else {
            var activatedDate = convertToDate(worksheet.getCell(rowNumber, 1).value)
            var offset1 = monthDiff(startDate, activatedDate)
            row.eachCell(function (cell, colNumber) {

                if (colNumber < 4) {
                    //Do nothing
                } else {
                    var payDate = activatedDate.addDays((colNumber - 4))
                    var offset2 = monthDiff(activatedDate, payDate)
                    results[offset1][offset2] = results[offset1][offset2] + cell.value
                }
            });
        }
    });

    for (var i = 0; i < 12; i++) {
        var newline = ""
        for (var j = 0; j < 12; j++) {
            newline = newline + " " + results[i][j]
        }
        console.log(newline)
    }

});

function convertToDate(days) {
    var result = new Date(1900, 0, 0);
    result.setDate(result.getDate() + days);
    return result;
}

Date.prototype.addDays = function(days) {
    var date = new Date(this.valueOf());
    date.setDate(date.getDate() + days);
    return date;
}

function monthDiff(d1, d2) {
    var months;
    months = (d2.getFullYear() - d1.getFullYear()) * 12;
    months -= d1.getMonth() + 1;
    months += d2.getMonth();
    return months <= 0 ? 0 : months;
}