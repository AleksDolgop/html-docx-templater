/**
 * @return {string}
 */
module.exports.prepareNowDate = function prepareNowDate() {
    return new Date()
        .toISOString()
        .split(".")[0]
        .replace(/[-:]/g, "_");
}