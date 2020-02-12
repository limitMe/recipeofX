const Excel = require('exceljs')
const fs = require('fs');
const excelfile = "test.xlsx";
var workbook = new Excel.Workbook();

workbook.xlsx.readFile(excelfile).then(function () {
    //获取第一个worksheet，如果数据在其他worksheet会失效
    var worksheet = workbook.getWorksheet(1); 

    //存储结果的地方，最多处理12个月的激活用户，每日用户数据不超过365天
    var results = new Array();
    for (var i = 0; i < 12; i++) {
        results[i] = new Array();
        for (var j = 0; j < 12; j++) {
            results[i][j] = 0;
        }
    }

    var startDate = convertToDate(worksheet.getCell('A2').value)
    worksheet.eachRow(function (row, rowNumber) {

        if (rowNumber < 2) {
            //第一行 标题省略
        } else {
            var activatedDate = convertToDate(worksheet.getCell(rowNumber, 1).value)
            var offset1 = monthDiff(startDate, activatedDate)
            row.eachCell(function (cell, colNumber) {

                if (colNumber < 4) {
                    //前三行，不需要在这个过程中计算
                } else {
                    //核心算法，依次读每个单元格，把它加到该加的位置就可以了
                    var payDate = activatedDate.addDays((colNumber - 4))
                    var offset2 = monthDiff(activatedDate, payDate)
                    results[offset1][offset2] = results[offset1][offset2] + cell.value
                    /*
                    var output = "现在读到第"+colNumber+"列第"+rowNumber+"行，数据是"+cell.value+
                        "。目前用户激活于"+(activatedDate.getMonth()+1)+"月"+
                        ",目前付费时间"+(payDate.getMonth()+1)+"月"+
                        "。故写入数组"+offset1+","+offset2
                    var output = "现在激活日期是"+activatedDate+",付费日期是"+payDate
                    console.log(output)
                    */
                }
            });
        }
    });

    //展示结果
    for (var i = 0; i < 12; i++) {
        var newline = ""
        for (var j = 0; j < 12; j++) {
            newline = newline + "\t" + results[i][j]
        }
        console.log(newline)
    }

});

function convertToDate(days) {
    var result = new Date(1900, 0, 0);
    //Why we have to minors 1?
    result.setDate(result.getDate() + days - 1);
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
    months -= d1.getMonth();
    months += d2.getMonth();
    return months
}