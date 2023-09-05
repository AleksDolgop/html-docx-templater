const ExcelJS = require('exceljs');

/**
 * @param xlsxFile
 * @param {number} sheetIndex
 * @return {Promise<{workbook: Workbook, foundReplaceCells: import('exceljs').Cell[]}>}
 */
module.exports.prepareXlsx = async function prepareXlsx(xlsxFile, sheetIndex) {
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(xlsxFile)
    console.log('workbook', workbook.title)
    const sheet = workbook.worksheets[sheetIndex]

    const foundReplaceCells = []
    const valuesSet = new Set()

    const matchRegex = /\{([а-я_-]+)}/i

    sheet.eachRow((row) => {
        row.eachCell(cell => {
            if (typeof cell.value === 'string') {
                const [matched] = cell.value.match(matchRegex) ?? []
                if (matched && !valuesSet.has(cell.value)) {
                    valuesSet.add(cell.value)
                    foundReplaceCells.push(cell)
                }
            }
        })
    })

    return {
        workbook,
        foundReplaceCells,
    }
}