
/**
 * @returns {{[key: string]: string}[]}
 * @param csvBuffer
 * @param encoding
 */
function prepareCsv(csvBuffer, encoding) {
  const csvDataText = new TextDecoder(encoding).decode(
      csvBuffer
  );
  const csvParsed = csvDataText.split("\n").map((i) => i.split(";"));
  const [csvTitle, ...csvItems] = csvParsed;
  return csvItems
    .reduce((prev, cur) => {
      csvTitle;
      prev.push({
        ...cur.reduce((p, c, i) => {
          return {
            ...p,
            [csvTitle[i].trim()]: c,
          };
        }, {}),
      });
      return prev;
    }, [])
    .slice(0, csvItems.length - 1);
}

let docFile;
let csvFile;

const encodersList = ['cp1252', 'cp1251', 'utf-8', 'utf-16', 'macintosh']
let encoding = 'cp1251'

function prepareDocx(docBlob, csvPatch) {
  const zip = new PizZip(docBlob);
  const doc = new window.docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });
  doc.render(csvPatch);
  return doc.getZip().generate({
    type: "blob",
    mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    compression: "DEFLATE",
  });
}

function prepareDate() {
  return new Date()
      .toISOString()
      .split(".")[0]
      .replace(/[-:]/g, "_");
}

const docs = document.getElementById("doc");
docs.onchange = function (e) {
  const files = e.target.files
  if (files.length === 0) {
    return
  }
  const readerDoc = new FileReader();
  readerDoc.readAsBinaryString(files[0]);
  readerDoc.onerror = function (evt) {
    alert("Error reading doc file" + evt);
  };
  readerDoc.onload = function (evt) {
    docFile = evt.target.result;
  };
}

const selectEncoding = document.getElementById("csv-encoding");
encodersList.map(enc => {
  const option = document.createElement('option')
  option.value = enc
  option.text = enc
  if (enc === encoding) {
    option.selected = true
  }

  selectEncoding.appendChild(option)
})
selectEncoding.onchange = function (e) {
  encoding = e.target.value
  console.log('Changed encoding to', encoding)
}

const csv = document.getElementById("csv");
csv.onchange = function (e) {
  const files = e.target.files
  if (files.length === 0) {
    return
  }
  const readerCsv = new FileReader();
  readerCsv.readAsArrayBuffer(files[0]);
  readerCsv.onerror = function (evt) {
    console.log("error reading file", evt);
    alert("error reading csv file" + evt);
  };
  readerCsv.onload = function (evt) {
    csvFile = evt.target.result;
  };
}

/**
 * @type {HTMLPreElement}
 */
const csvPreRaw = document.getElementById("csv-raw")
window.readTitles = function readTitles() {
  if (csv.files.length === 0) {
    alert("No CSV file selected");
    return
  }

  csvPreRaw.style.display = 'block'

  csvPreRaw.innerText = new TextDecoder(encoding).decode(
      csvFile
  ).split('\n')[0]
}

window.generate = function generate(first = false) {
  if (docs.files.length === 0) {
    alert("No docx file selected");
    return
  }
  if (csv.files.length === 0) {
    alert("No CSV file selected");
    return
  }
  const preparedCsv = prepareCsv(csvFile, encoding);
  console.log("preparedCsv", preparedCsv);

  if (preparedCsv.length === 0) {
    alert(`CSV document have 0 lines`)
    return;
  }

  if (first) {
    const blob = prepareDocx(docFile, preparedCsv[0])
    saveAs(blob, `output${prepareDate()}.docx`);
    return
  }

  const outputZip = new JSZip();
  for (const index in preparedCsv) {
    const blob = prepareDocx(docFile, preparedCsv[index])
    outputZip.file(`doc_${+index + 1}.docx`, blob);
  }

  outputZip.generateAsync({ type: "blob" }).then(function (content) {
    saveAs(content, `output${prepareDate()}.zip`);
  });
};
