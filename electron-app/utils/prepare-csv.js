
/**
 * @returns {{[key: string]: string}[]}
 * @param csvBuffer
 * @param encoding
 */
module.exports.prepareCsv = function prepareCsv(csvBuffer, encoding) {
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
