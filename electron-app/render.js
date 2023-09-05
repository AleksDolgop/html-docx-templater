const docs = document.getElementById("doc");
const csv = document.getElementById("csv");
const selectEncoding = document.getElementById("csv-encoding");
const labelDocx = document.getElementById('label-doc-docx')
const labelExcel = document.getElementById('label-doc-excel')
const excelControl = document.getElementById('excel-control')
const outputFileTemplate = document.getElementById('output-file-template')

docs.onchange = csv.onchange = function onchangeFileSelect(e) {
    const files = e.target.files
    if (files.length === 0) return

    if(e.target.id === 'doc') {
        window.electron.docMimeType(files[0].type)
        window.electron.docFilePath(files[0].path)
    } else if(e.target.id === 'csv') {
        window.electron.csvFilePath(files[0].path)
    }
}

selectEncoding.onchange = function changeEncoding(e) {
    window.electron.selectEncoding(e.target.value)
}

excelControl.onchange = function changeSheet(e) {
    window.electron.selectTable(e.target.value)
}

outputFileTemplate.onchange = function changeOutputFileTemplate(e) {
    window.electron.changeOutputTemplate(e.target.value)
}


window.selectDocType = function selectDocType(type) {
    if (type === 'docx') {
        labelDocx.style.display = 'block'
        labelExcel.style.display = 'none'
        excelControl.style.display = 'none'
        window.electron.selectDocType(type)
    } else if(type === 'excel'){
        labelDocx.style.display = 'none'
        labelExcel.style.display = 'block'
        excelControl.style.display = 'block'
        window.electron.selectDocType(type)
    }
}


window.readTitles = function readTitles() {
    window.electron.readTitles().then()
}

window.generate = function generate(first = false) {
    window.electron.generate(first).then()
}
