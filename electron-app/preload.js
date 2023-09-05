const { contextBridge, ipcRenderer } = require('electron')
require('file-saver')

let docFilePath
let docMimeType
let csvFilePath
let documentTemplate = 'document{Порядковый_номер_документа}'

const encodersList = ['cp1251','cp1252',  'utf-8', 'utf-16', 'macintosh']
let encoding = 'cp1251'
let docType = 'docx'
let sheet

contextBridge.exposeInMainWorld('electron',
    {
        send: (channel, payload) => ipcRenderer.send(channel, payload),
        docFilePath: async (path) => {
            docFilePath = path
            if(docType === 'excel') {
                const excelSheet = document.getElementById("excel-sheet")
                const sheets = await ipcRenderer.invoke('readExcelSheets', {
                    path: docFilePath,
                    type: docMimeType
                })
                console.log('sheets', sheets)
                sheets.forEach(sheetName => {
                    const option = document.createElement('option')
                    option.value = option.text = sheetName
                    excelSheet.appendChild(option)
                })
                sheet = sheets[0]
            }
        },
        docMimeType: (mimetype) => docMimeType = mimetype,
        csvFilePath: (path) => csvFilePath = path,
        selectEncoding: (enc) => encoding = enc,
        selectTable: (sh) => sheet = sh,
        selectDocType:  (type) => docType = type,
        changeOutputTemplate:  (template) => documentTemplate = template,
        readTitles: async () => {
            const csvPreRaw = document.getElementById("csv-raw")
            if (!csvFilePath) return

            const line = await ipcRenderer.invoke('readTitles', { path: csvFilePath, encoding })
            csvPreRaw.style.display = 'block'
            csvPreRaw.innerText = line

            console.log('readTitles', line)
        },
        generate: async (first) => {
            const timer = setInterval(() => {
                console.log('generate...')
            }, 1000)
            ipcRenderer.invoke(
                'generate', {
                    first,
                    documentTemplate,
                    docMimeType,
                    docType,
                    docPath: docFilePath,
                    sheet,
                    csvPath: csvFilePath,
                    encoding,
                }).then(() => {
                    clearInterval(timer)
            })


        },
    }
)

window.addEventListener('DOMContentLoaded', () => {
    const selectEncoding = document.getElementById("csv-encoding");
    encodersList.map(enc => {
        const option = document.createElement('option')
        option.value = enc
        option.text = enc
        if (enc === encoding) option.selected = true
        selectEncoding.appendChild(option)
    })

    const outputFileTemplate = document.getElementById('output-file-template')
    outputFileTemplate.value = documentTemplate

    const labelDocx = document.getElementById('label-doc-docx')
    const labelExcel = document.getElementById('label-doc-excel')
    const excelControl = document.getElementById('excel-control')
    labelDocx.style.display = 'block'
    labelExcel.style.display = 'none'
    excelControl.style.display = 'none'

})

