const fs = require('fs');
const path = require('path');
const jsreport = require('jsreport-core')();

jsreport.use(
    require('jsreport-handlebars')())
    .use(
        require('jsreport-xlsx')()
    ).use(
    require('jsreport-templates')());

const renderExcelReport = (template) => jsreport.init()
    .then(() => jsreport.documentStore.collection('xlsxTemplates')
        .remove({name: template.template.xlsxTemplate.name}))
    .then(() => jsreport.documentStore.collection('xlsxTemplates')
        .insert({
            contentRaw: template.template.xlsxTemplate.content,
            shortid: template.template.xlsxTemplate.shortid,
            name: template.template.xlsxTemplate.name
        }))
    .then(() => jsreport.render(template));



const template = {
    template: {
        content: fs.readFileSync(path.resolve('./template.hbs')).toString('utf-8'),
        recipe: 'xlsx',
        engine: 'handlebars',
        helpers: 'function add(a, b){return a + b;}',
        xlsxTemplate: {
            shortid: 'test',
            name: 'testTemplate',
            content: fs.readFileSync(path.resolve('./template.xlsx')).toString('base64')
        },
    },
    data: {
       cellOne: "NEW TEXT"
    }
};

renderExcelReport(template).then(() => console.log('RUNNING ONCE WAS OK')).catch((err) => console.log(err));
//renderExcelReport(template).then(() => console.log('RUNNING TWICE WAS NOT OK')).catch((err) => console.log(err));
