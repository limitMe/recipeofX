const Excel = require('exceljs')
const fs = require('fs');
const readline = require('readline')

//----configs-----
var costColumn = 2
var startRow = 4
var startColumn = 8
var comparedDate = new Date('2021-02-10')
//----configs-----
var isProcessing = true
var csvOutput = "空行,项目,渠道,投放周期,投放周期总支出,总回收,ROI,第1天回收,第1天ROI,第7天回收,第7天ROI,第14天回收,第14天ROI,第30天回收,第30天ROI,第60天回收,第60天ROI,第90天回收,第90天ROI,第120天回收,第120天ROI,第150天回收,第150天ROI,第190天回收,第190天ROI,"
var excelFilename = "";
var files = fs.readdirSync('./')
files = files.filter(element => (element.endsWith(".xlsx") && !element.startsWith("~$")))
files.forEach(excelFilename => {
    var workbook = new Excel.Workbook();
    workbook.xlsx.readFile(excelFilename).then(function () {
        workbook.eachSheet(function(worksheet, sheetId) {
            var sheetName = worksheet.name

            if (typeof worksheet.getCell(startRow,1).value.getMonth === 'function') {

                //存储结果的地方，最多处理12个月的激活用户，每月最多190天
                var monthlyProfit = new Array();
                for (var i = 0; i < 12; i++) {
                    monthlyProfit[i] = new Array();
                    for (var j = 0; j < 9; j++) {
                        monthlyProfit[i][j] = 0;
                    }
                }

                var monthlyCost = new Array()
                for (var k = 0; k < 12; k++) {
                    monthlyCost[k] = 0
                }

                var totalProfit = new Array()
                for (var k = 0; k < 12; k++) {
                    totalProfit[k] = 0
                }

                var startDate = worksheet.getCell(startRow, 1).value
                if( !startDate.getMonth || typeof startDate.getMonth != "function") {
                    startDate = convertToDate(worksheet.getCell(startRow, 1).value)
                }
                var lastStartingRow = startRow
                var needToCount = [1,7,14,30,60,90,120,150,190]
                worksheet.eachRow(function (row, rowNumber) {
                    isProcessing = true

                    if (rowNumber < startRow) {
                        //第一行 标题省略
                    } else {
                        var currentActivatedDate = worksheet.getCell(rowNumber, 1).value
                        var nextActivatedDate = worksheet.getCell(rowNumber+1, 1).value

                        if(nextActivatedDate == null || currentActivatedDate.getMonth() != nextActivatedDate.getMonth()){
                            var offset1 = monthDiff(startDate, currentActivatedDate)
                            for(var k=0; k<9; k++){
                                if(currentActivatedDate.addDays(needToCount[k]) < comparedDate) {
                                    var totalGain = 0
                                    for(var i=lastStartingRow; i<=rowNumber; i++){
                                        totalGain = totalGain + worksheet.getCell(i, startColumn + needToCount[k] - 1).value
                                    }
                                    monthlyProfit[offset1][k] = totalGain
                                }
                            }

                            for(var i=lastStartingRow; i<=rowNumber; i++){
                                //总消耗
                                if(worksheet.getCell(i, costColumn).value == undefined){
                                    monthlyCost[offset1] = monthlyCost[offset1] + parseFloat(worksheet.getCell(i, costColumn).value)
                                } else {
                                    monthlyCost[offset1] = monthlyCost[offset1] + parseFloat(worksheet.getCell(i, costColumn).value)
                                }

                                //总回收
                                totalProfit[offset1] = totalProfit[offset1] + worksheet.getCell(i, worksheet.columnCount).value
                            }
                        
                            lastStartingRow = rowNumber + 1
                        }
                    }
                })

                for (var i = 0; i < 12; i++) {
                    if(monthlyProfit[i][0] == 0){
                        continue
                    }
                    var tempStr = sheetName.split('】')
                    var nameStr = tempStr[1].split('-')
                    var month = (startDate.getMonth() + i + 1)
                    var newline = "," + nameStr[0] + "," + nameStr[1] + "," + month + "月" + "," + monthlyCost[i].round(2) + "," + totalProfit[i].round(2) + "," + (totalProfit[i]/monthlyCost[i]*100).round(2)
                    for (var j = 0; j < 9; j++) {
                        newline = newline + "," + monthlyProfit[i][j].round(2) + "," + (monthlyProfit[i][j]/monthlyCost[i]*100).round(2)
                    }
                    console.log(newline)
                    csvOutput = csvOutput + "\n" + newline + ","
                }

                isProcessing = false
            }
            else {
                //console.log(sheetName+"被识别为非数据表，跳过")
            }
        });
    })
})

var task = setTimeout(() => {  
    if(!isProcessing) {
        fs.writeFile("output.csv", csvOutput, function(err) {
            clearTimeout(task)
            if(err) {
                return console.log(err);
            }
            console.log("已将上述内容写入到output.csv，请在excel中导入");  
        }); 
    }
    
}, 10000);

function convertToDate(days) {
    var result = new Date(1900, 0, 0);
    //Time zone issue to minors 1
    result.setDate(result.getDate() + days - 1);
    return result;
}

Date.prototype.addDays = function (days) {
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

Number.prototype.round = function(p) {
    p = p || 10;
    return parseFloat( this.toFixed(p) );
  };