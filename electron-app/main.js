require('./bootstrap')
const { app, BrowserWindow, ipcMain, dialog } = require('electron')
const ProgressBar = require('electron-progressbar')
const path = require('path')

const createWindow = () => {
    const mainWindow = new BrowserWindow({
        resizable: false,
        fullscreenable: false,
        width: 800,
        height: 600,
        title: 'Documents templater',
        icon: path.join(__dirname, 'favicon.ico'),
        webPreferences: {
            nodeIntegration: true,
            preload: path.join(__dirname, 'preload.js'),
        }
    })

    // and load the index.html of the app.
    mainWindow.loadFile(path.join(__dirname, 'index.html'))

    // Open the DevTools.
    // mainWindow.webContents.openDevTools({
    //     mode: 'undocked',
    //     activate: process.env.ENV === 'development'
    // })
}

app.whenReady().then(() => {
    createWindow()
    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) createWindow()
    })
})

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit()
})

const fs = require('fs')
const ExcelJS = require('exceljs')
const JSZip = require('jszip');

ipcMain.handle('readTitles', async (e, params) => {
    const file = await fs.promises.readFile(params.path)
    const data = new TextDecoder(params.encoding).decode(file)
    const splitData = data.split('\n')
    return splitData[0]
})

ipcMain.handle('readExcelSheets', async (e, params) => {
    const file = await fs.promises.readFile(params.path)
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(file)
    return workbook.worksheets.map(i => i.name)
})

const { prepareCsv, prepareDocx, prepareNowDate, handleTemplateText } = require('./utils')
ipcMain.handle('generate', async (e, params) => {
    const { docType, docPath, sheet: sheetName, csvPath, encoding, documentTemplate } = params
    const docFile = await fs.promises.readFile(docPath)
    const csvFile = await fs.promises.readFile(csvPath)
    const preparedCsv = prepareCsv(csvFile, encoding)

    const progressBar = new ProgressBar({
        indeterminate: false,
        text: 'Preparing data...',
        detail: 'Wait...',
        maxValue: preparedCsv.length,
    })

    progressBar
        .on('completed', function() {
            progressBar.detail = 'Task completed. Exiting...';
        })
        .on('progress', function(value) {
            progressBar.detail = `${value} of ${progressBar.getOptions().maxValue}...`;
        });

    let stop = false
    app.on('quit', () => {
        progressBar.setCompleted()
        stop = true
    })

    const zip = new JSZip()
    if (docType === 'docx') {
        for (const k in preparedCsv) {
            const item = preparedCsv[k]
            const buff = prepareDocx(docFile, item)
            zip.file(`doc${+k + 1}.docx`, buff)
        }
    } else if(docType === 'excel') {
        if (!sheetName) {
            return
        }

        const workbook = new ExcelJS.Workbook()
        await workbook.xlsx.load(docFile)

        // Поиск ячеек с шаблонами, для получения их адресов
        const sheet = workbook.worksheets.find(i => i.name === sheetName)
        const foundReplaceCells = []
        const valuesSet = new Set()
        const matchRegex = /\{([а-я_-]+)}/i
        sheet.eachRow((row) => {
            row.eachCell(cell => {
                if (typeof cell.value === 'string') {
                    const [matched] = cell.value.match(matchRegex) ?? []
                    if (matched && !valuesSet.has(cell.value)) {
                        valuesSet.add(cell.value)
                        foundReplaceCells.push(cell.address)
                    }
                }
            })
        })


        for (const k in preparedCsv) {
            if (stop) return
            progressBar.value += 1
            const tempBook = await new ExcelJS.Workbook().xlsx.load(docFile)
            const tempBookSheet = tempBook.worksheets.find(i => i.name === sheetName)

            const item = preparedCsv[k]
            foundReplaceCells.forEach((address) => {
                const cell = tempBookSheet.getCell(address)
                cell.value = handleTemplateText(cell.value, item)
            })

            const outBf = await tempBook.xlsx.writeBuffer()
            const outFileName = handleTemplateText(
                documentTemplate,
                {...item, 'Порядковый_номер_документа': +k + 1}
            )
            zip.file(`${outFileName}.xlsx`, outBf, { createFolders: true })
        }

        progressBar.setCompleted()

        const buffer = await zip.generateAsync({ type: 'nodebuffer'})

        dialog.showSaveDialog(null, {
            title: "Save file",
            defaultPath : `out_archive_${prepareNowDate()}`,
            buttonLabel : "Save",
            filters :[
                {name: 'zip', extensions: ['zip']},
            ]
        }).then(({ filePath }) => {
            fs.writeFileSync(filePath, buffer, 'utf-8');
        });
    }
})