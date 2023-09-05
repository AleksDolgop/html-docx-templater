const { handleTemplateText } = require('./index')

console.log(handleTemplateText('test_{test}_{data}', {
    test: 'sukaa',
    data: 'atata'
}))