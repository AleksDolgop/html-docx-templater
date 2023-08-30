function decodeUTF16LE(binaryStr) {
    var cp = [];
    for (var i = 0; i < binaryStr.length; i += 2) {
        cp.push(
            binaryStr.charCodeAt(i) |
            (binaryStr.charCodeAt(i + 1) << 8)
        );
    }

    return String.fromCharCode.apply(String, cp);
}

/**
 * @param {ArrayBuffer} textCsv 
 * @returns {{[key: string]: string}[]}
 */
function prepareCsv(csvBuffer) {
    let csvDataText = null
    const encoders = ['utf-8', 'cp1251']
    for (const encoding of encoders) {
        if (!csvDataText) {
            try {
                csvDataText = new TextDecoder(encoding, { fatal: true }).decode(csvFile)
            } catch {
                csvDataText = null
            }
        }

        if(typeof csvDataText === 'string') {
            break
        }
    }
    
    if (!csvDataText) {
        alert('Invalid CSV file')
        return
    }
    
    const csvParsed = csvDataText.split('\n').map(i => i.split(';'))
    const [csvTitle, ...csvItems] = csvParsed
    return csvItems.reduce((prev, cur) => {
        csvTitle
        prev.push({
            ...cur.reduce((p, c, i) => {
                return {
                    ...p,
                    [csvTitle[i].trim()]: c
                }
            }, {})
        })
        return prev
    }, [])
}

/**
 * 
 * @param {string} text 
 * @returns {string[]}
 */
function getTemplateFieldsFromText(text) {
    return text.match(/\{[а-я\w_]+\}/igm)
}


let docFile
let csvFile

let run = false
function handle() {
    console.log('Run handle')
    console.log('docFile', !!docFile)
    console.log('csvFile', !!csvFile)
    if (!docFile || !csvFile) {
        return
    }
    if (run) {
        return
    }
    run = true
    try {
        const preparedCsv = prepareCsv(csvFile)
        console.log('preparedCsv', preparedCsv)

        const zip = new PizZip(docFile);
        const templateDoc = new window.docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        const templatesFields = getTemplateFieldsFromText(templateDoc.getFullText())
        console.log('templatesFields', templatesFields)

        const outputZip = new JSZip()
        for (const index in preparedCsv) {
            const item = preparedCsv[index]
            
            const doc = new window.docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
            });

            doc.render(item);

            const blob = doc.getZip().generate({
                type: "blob",
                mimeType:
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                compression: "DEFLATE",
            });
            outputZip.file(`doc_${index}.docx`, blob)
        }

        const dateText = new Date().toISOString().split('.')[0].replace(/-|:/g, '_')

        outputZip.generateAsync({ type: "blob" })
            .then(function (content) {
                saveAs(content, `output${dateText}.zip`);
            });
    } finally {
        run = false
    }


}

const docs = document.getElementById("doc");
const csv = document.getElementById("csv");
window.generate = function generate() {
    const readerDoc = new FileReader();
    const readerCsv = new FileReader();
    if (docs.files.length === 0) {
        alert("No docs file selected");
    }
    readerDoc.readAsBinaryString(docs.files[0]);
    readerDoc.onerror = function (evt) {
        console.log("error reading file", evt);
        alert("error reading doc file" + evt);
    };
    readerDoc.onload = function (evt) {
        docFile = evt.target.result;
        handle()
    };

    if (csv.files.length === 0) {
        alert("No csvs file selected");
    }

    readerCsv.readAsArrayBuffer(csv.files[0]);
    readerCsv.onerror = function (evt) {
        console.log("error reading file", evt);
        alert("error reading csv file" + evt);
    };
    readerCsv.onload = function (evt) {
        csvFile = evt.target.result;
        handle()
    };

};
