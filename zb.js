const path = require('path')
const Excel = require('exceljs')

// 输入Excel
const IN = 'data.xlsx'


async function load() {

    const workbook = new Excel.Workbook()
    await workbook.xlsx.readFile(path.join(__dirname, IN))
    console.log('文件: ' + IN)

    const ZB_MAP = {}


    // [B02020204011流动资金贷款]_PA01发生地区:<境内>_CU04客户所属行业:<农业,林业,畜牧业,渔业>
    const patt = /\[([^\[\]]+)\](_[^:]+\:\<[^>]+\>)+/g
    // _PA01发生地区:<境内>
    const wdPatt = /_([^:]+)\:\<[^>]+\>/g


    workbook.eachSheet(function (worksheet, sheetId) {
        console.log('Sheet: ' + worksheet.name)

        worksheet.eachRow(function (row, rowNumber) {

            row.eachCell(function (cell, colNumber) {
                let value = cell.value
                if (typeof value !== 'string') {
                    return
                }

                for (let zbm of value.matchAll(patt)) {
                    let [text, zb] = zbm
                    let wdText = text.substr(2 + zb.length)
                    // console.log('  ' + rowNumber + '.' + colNumber + ', ' + zb + ' ' + wdText)

                    let wds = ZB_MAP[zb]
                    if (!wds) {
                        wds = []
                        ZB_MAP[zb] = wds
                    }

                    for (let wdm of wdText.matchAll(wdPatt)) {
                        let [_, wd] = wdm
                        // console.log('    ' + wd)
                        if (!wds.includes(wd)) {
                            wds.push(wd)
                        }
                    }
                }

            })
        })
    })

    console.log()
    console.log('指标维度：')
    for (let zb in ZB_MAP) {
        console.log(zb)
        for (let wd of ZB_MAP[zb]) {
            console.log('\t' + wd)
        }
    }

}


load().catch(console.log)
