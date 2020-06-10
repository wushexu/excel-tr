// 输入Excel
const IN = 'data.xlsx'

// 输出Excel
const OUT = 'data-zb-out.xlsx'


const path = require('path')
const Excel = require('exceljs')


async function load() {

    const workbook = new Excel.Workbook()
    await workbook.xlsx.readFile(path.join(__dirname, IN))
    console.log('文件: ' + IN)

    const REP_MAP = {}
    const ZB_MAP = {}


    // [B02020204011流动资金贷款]_PA01发生地区:<境内>_CU04客户所属行业:<农业,林业,畜牧业,渔业>
    const patt = /\[([^\[\]]+)\](_[^:]+\:\<[^>]+\>)+/g
    // _PA01发生地区:<境内>
    const wdPatt = /_([^:]+)\:\<[^>]+\>/g


    workbook.eachSheet(function (worksheet, sheetId) {
        const repName = worksheet.name
        console.log('Sheet: ' + repName)


        const ZBS = []
        REP_MAP[repName] = ZBS

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

                    if (!ZBS.includes(zb)) {
                        ZBS.push(zb)
                    }

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

    const workbookOut = new Excel.Workbook()

    const worksheet1 = workbookOut.addWorksheet('报表-指标')

    worksheet1.columns = [
        { header: '报表', key: 'rep', style: { font: { bold: true } }, width: 20 },
        { header: '指标', key: 'zb', width: 40 }
    ]

    // console.log()
    // console.log('报表指标：')
    for (let rep in REP_MAP) {
        // console.log(rep)
        worksheet1.addRow({ rep, zb: '' })
        for (let zb of REP_MAP[rep]) {
            // console.log('\t' + zb)
            worksheet1.addRow({ rep: '', zb })
        }
    }

    worksheet1.columns.forEach(col => col.alignment = { vertical: 'middle' })
    worksheet1.getRow(1).font = { bold: true }
    worksheet1.eachRow(row => row.height = 16)


    const worksheet2 = workbookOut.addWorksheet('指标-维度')

    worksheet2.columns = [
        { header: '指标', key: 'zb', width: 40 },
        { header: '维度', key: 'wd', width: 40 }
    ]

    // console.log()
    // console.log('指标维度：')
    for (let zb in ZB_MAP) {
        // console.log(zb)
        worksheet2.addRow({ zb, wd: '' })
        for (let wd of ZB_MAP[zb]) {
            // console.log('\t' + wd)
            worksheet2.addRow({ zb: '', wd })
        }
    }

    worksheet2.getRow(1).font = { bold: true }
    worksheet2.eachRow(row => row.height = 16)
    worksheet2.columns.forEach(col => col.alignment = { vertical: 'middle' })

    await workbookOut.xlsx.writeFile(path.join(__dirname, OUT))

    console.log('保存到: ' + OUT)
}


load().catch(console.log)
