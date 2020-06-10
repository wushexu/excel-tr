const path = require('path')
const Excel = require('exceljs')

const { IN, OUT, MAP } = require('./tr-config')


async function load() {

    const workbook = new Excel.Workbook()
    await workbook.xlsx.readFile(path.join(__dirname, IN))
    console.log('文件: ' + IN)

    const SUBS = []

    for (let fromStr in MAP) {
        SUBS.push([fromStr, MAP[fromStr], 0])
    }

    let sheetCount = 0, rowCount = 0, cellCount = 0, replaceCount = 0

    workbook.eachSheet(function (worksheet, sheetId) {
        console.log('Sheet: ' + worksheet.name)
        sheetCount++

        worksheet.eachRow(function (row, rowNumber) {
            // console.log('Row ' + rowNumber)

            rowCount++;

            row.eachCell(function (cell, colNumber) {
                let value = cell.value
                if (typeof value !== 'string') {
                    return
                }

                let occur = false

                for (let sub of SUBS) {
                    const [fromStr, toStr, count] = sub
                    if (value.indexOf(fromStr) === -1) {
                        continue
                    }
                    let newValue = value.split(fromStr).join(toStr)
                    if (newValue !== value) {
                        occur = true
                        sub[2]++
                        replaceCount++
                        console.log('\t' + rowNumber + '.' + colNumber + ' ' + fromStr + ' -> ' + toStr)
                        value = newValue
                    }
                }

                if (occur) {
                    cellCount++
                    cell.value = value
                }
                // console.log('   Cell ' + colNumber + ' = ' + value)

            })
        })
    })


    console.log('')

    console.log('Sheet数: ' + sheetCount)
    console.log('总行数: ' + rowCount)
    console.log('替换单元格数: ' + cellCount)
    console.log('替换总次数: ' + replaceCount)


    console.log('替换: ')
    for (let [fromStr, toStr, count] of SUBS) {
        console.log('\t' + fromStr + ' -> ' + toStr + '\t:' + count)
    }

    await workbook.xlsx.writeFile(path.join(__dirname, OUT))
    console.log('保存到: ' + OUT)
}


load().catch(console.log)
