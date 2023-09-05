const docxtemplater = require("docxtemplater");
const PizZip = require("pizzip");

/**
 * @param docBlob
 * @param csvPatch
 * @return {Buffer}
 */
module.exports.prepareDocx = function prepareDocx(docBlob, csvPatch) {
    const zip = new PizZip(docBlob);
    const doc = new docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });
    doc.render(csvPatch);
    return doc.getZip().generate({
        type: "nodebuffer",
        mimeType:
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        compression: "DEFLATE",
    });
}
