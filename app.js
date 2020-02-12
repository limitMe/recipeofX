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
                //存储结果的地方，最多处理12个月的激活用户，每日用户数据不超过365天
                var monthlyProfit = new Array();
                for (var i = 0; i < 12; i++) {
                    monthlyProfit[i] = new Array();
                    for (var j = 0; j < 12; j++) {
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
                worksheet.eachRow(function (row, rowNumber) {

                    if (rowNumber < startRow) {
                        //第一行 标题省略
                    } else {
                        activatedDate = worksheet.getCell(rowNumber, 1).value
                        
                        if( !activatedDate.getMonth || typeof activatedDate.getMonth != "function") {
                            activatedDate = convertToDate(worksheet.getCell(rowNumber, 1).value)
                        }
                        var offset1 = monthDiff(startDate, activatedDate)

                        row.eachCell(function (cell, colNumber) {

                            if (colNumber == costColumn) {
                                if(cell.result == undefined){
                                    monthlyCost[offset1] = monthlyCost[offset1] + parseFloat(cell.value)
                                } else {
                                    monthlyCost[offset1] = monthlyCost[offset1] + parseFloat(cell.result)
                                }
                            }

                            if (colNumber < startColumn) {
                                //前三行，不需要在这个过程中计算
                            } else {
                                //核心算法，依次读每个单元格，把它加到该加的位置就可以了
                                var payDate = activatedDate.addDays((colNumber - startColumn))
                                var offset2 = monthDiff(activatedDate, payDate)

                                if(colNumber == startColumn){
                                    monthlyProfit[offset1][offset2] = monthlyProfit[offset1][offset2] + parseInt(cell.value)
                                } else {
                                    monthlyProfit[offset1][offset2] = monthlyProfit[offset1][offset2] + parseInt(cell.value) - parseInt(worksheet.getCell(rowNumber,colNumber-1).value)
                                }
                                
                                var output = "现在读到第"+colNumber+"列第"+rowNumber+"行，数据是"+cell.value+
                                    "。目前用户激活于"+(activatedDate.getMonth()+1)+"月"+
                                    ",目前付费时间"+(payDate.getMonth()+1)+"月"+
                                    "。故写入数组"+offset1+","+offset2
                                //var output = "现在激活日期是"+activatedDate+",付费日期是"+payDate
                                //console.log(output)
                                
                                
                            }
                        });
                    }
                });

                //如果变每月新增收益为每月累积收益
                for (var i = 0; i < 12; i++) {
                    for (var j = 0; j < 12; j++) {
                        if(j!=0){
                            monthlyProfit[i][j] = monthlyProfit[i][j-1] + monthlyProfit[i][j]
                        }
                    }
                }

                //展示结果
                for (var i = 0; i < 12; i++) {
                    var firstline = "本月消耗总量"
                    var secondline = "   " + monthlyCost[i].round(2)
                    for (var j = 0; j < 12; j++) {
                        firstline = firstline + "\t" + monthlyProfit[i][j]
                        secondline = secondline + "\t" + (monthlyProfit[i][j]/monthlyCost[i]).round(3)
                    }
                    if (!firstline.endsWith("\t0")) {
                        console.log(firstline)
                        console.log(secondline)
                    }
                }
            }
        }
    });
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