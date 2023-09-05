module.exports = {
    ...require('./prepare-csv'),
    ...require('./prepare-docx'),
    ...require('./prepare-now-date'),
    ...require('./prepare-xlsx'),
    handleTemplateText(template, obj) {
        const regexp = /\{([a-zа-я_-]+)}/gi
        return template.replace(regexp, (_, g1) => obj[g1])
    }
}