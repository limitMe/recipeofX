const Excel = require('exceljs')
const fs = require('fs');
const readline = require('readline')

handleFileName()

//----------处理文件名------------
function handleFileName() {
    var excelFilename = "";
    var files = fs.readdirSync('./')
    files = files.filter(element => (element.endsWith(".xlsx") && !element.startsWith("~$")))
    if (files.length == 0) {
        console.log("请把要处理的.xlsx格式工作表放在本程序的相同目录下")
        process.exit()
    } else if (files.length == 1) {
        excelFilename = files[0]
        loadWorkbook(excelFilename)
    } else {
        for (var o = 0; o < files.length; o++) {
            console.log(o + " : " + files[o])
        }
        var rl1 = readline.createInterface({
            input: process.stdin,
            output: process.stdout
        });
        var questionFilename = function () {
            rl1.question('输入你想要处理文件前面的数字：', (answer) => {
                if (answer == "x" || answer == "X") {
                    console.log("拜拜甜甜圈")
                    process.exit()
                }

                var userInput = parseInt(answer)
                if (userInput >= 0 && userInput < files.length) {
                    excelFilename = files[userInput]
                    loadWorkbook(excelFilename)
                    rl1.close();
                } else {
                    console.log("输入错误。输入正确的数字以继续。输入x以退出。")
                    questionFilename()
                }
            });
        }
        questionFilename()
    }
}

function loadWorkbook(excelFilename) {

    var workbook = new Excel.Workbook();
    workbook.xlsx.readFile(excelFilename).then(function () {

        var sheetNames = new Object()
        //----------处理Worksheet------------
        workbook.eachSheet(function(worksheet, sheetId) {
            sheetNames[sheetId] = worksheet.name
        });
        console.log(sheetNames)

        var rl2 = readline.createInterface({
            input: process.stdin,
            output: process.stdout
        });
        var questionWorksheet = function () {
            rl2.question('输入你想要处理的工作表前面的数字： ', (answer) => {
                if (answer == "x" || answer == "X") {
                    console.log("拜拜甜甜圈")
                    process.exit()
                }
                if( !isNaN(answer) && 
                    parseInt(answer) != undefined &&
                    sheetNames[answer] != undefined ){
                        rl2.close();
                        handleWorksheet(parseInt(answer))
                } else {
                    console.log("输入错误。输入正确的数字以继续。输入x以退出。")
                    questionWorksheet()
                }
            })
        }
        questionWorksheet()

        var handleWorksheet = function(worksheetNo) {
            //获取第一个worksheet，如果数据在其他worksheet会失效
            var worksheet = workbook.getWorksheet(worksheetNo)
            var startRow = 0
            var startColumn = 0
            var costColumn = 0

            //----------处理数据格式------------
            var rl3 = readline.createInterface({
                input: process.stdin,
                output: process.stdout
            })
            var rl4
            var rl5
            var questionCostColumn = function () {
                rl3.question('每日消耗数据在哪一列： ', (answer) => {
                    if (answer == "x" || answer == "X") {
                        console.log("拜拜甜甜圈")
                        process.exit()
                    }
                    if( !isNaN(answer) && 
                        parseInt(answer) != undefined){
                            rl3.close()
                            rl4 = readline.createInterface({
                                input: process.stdin,
                                output: process.stdout
                            })
                            costColumn = parseInt(answer)
                            questionStartRow()
                    } else {
                        console.log("输入错误。输入正确的数字以继续。输入x以退出。")
                        questionCostColumn()
                    }
                })
            }
            questionCostColumn()
            var questionStartRow = function () {
                rl4.question('要处理的数据，起始于多少行： ', (answer) => {
                    if (answer == "x" || answer == "X") {
                        console.log("拜拜甜甜圈")
                        process.exit()
                    }
                    if( !isNaN(answer) && 
                        parseInt(answer) != undefined){
                            rl4.close()
                            rl5 = readline.createInterface({
                                input: process.stdin,
                                output: process.stdout
                            })
                            startRow = parseInt(answer)
                            questionStartColumn()
                    } else {
                        console.log("输入错误。输入正确的数字以继续。输入x以退出。")
                        questionStartRow()
                    }
                })
            }
            var questionStartColumn = function () {
                rl5.question('要处理的每日回收数据，起始于多少列： ', (answer) => {
                    if (answer == "x" || answer == "X") {
                        console.log("拜拜甜甜圈")
                        process.exit()
                    }
                    if( !isNaN(answer) && 
                        parseInt(answer) != undefined){
                            rl5.close()
                            startColumn = parseInt(answer)
                            handleData()
                    } else {
                        console.log("输入错误。输入正确的数字以继续。输入x以退出。")
                        questionStartColumn()
                    }
                })
            }

            var handleData = function() {
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

                var startDate = worksheet.getCell(startRow, 1).value
                if( !startDate.getMonth || typeof startDate.getMonth != "function") {
                    startDate = convertToDate(worksheet.getCell(startRow, 1).value)
                }
                var lastStartingRow = startRow
                var needToCount = [1,7,14,30,60,90,120,150,190]
                var comparedDate = new Date('2021-02-10')
                worksheet.eachRow(function (row, rowNumber) {

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
                        
                            lastStartingRow = rowNumber + 1
                        }
                    }
                })

                for (var i = 0; i < 12; i++) {
                    if(monthlyProfit[i][0] == 0){
                        //continue
                    }
                    var month = (startDate.getMonth() + i + 1)
                    var newline = month + "月"
                    for (var j = 0; j < 9; j++) {
                        newline = newline + "\t" + monthlyProfit[i][j].round(2)
                    }
                    console.log(newline)
                }
                
            }
        }
    })
}


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