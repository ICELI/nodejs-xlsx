/**
 * Created by iceli on 17/5/5.
 */
const XLSX = require('xlsx')
const isString = (str) => typeof str === 'string'
let workBook = {
    SheetNames: [],
    Sheets: {}
}

exports.parse = function(mixed, options = {}) {
    const workSheet = XLSX[isString(mixed) ? 'readFile' : 'read'](mixed, options);
    return Object.keys(workSheet.Sheets).map((name) => {
        const sheet = workSheet.Sheets[name];
        return {name, data: XLSX.utils.sheet_to_json(sheet, {header: 1, raw: true})};
    });
}

exports.build = function(worksheets, filename) {
    worksheets.forEach((worksheet) => {
        const name = worksheet.name || 'Sheet';
        const data = worksheet.data || [];
        workBook.SheetNames.push(name);
        workBook.Sheets[name] = XLSX.utils.aoa_to_sheet(data); //将修改过的数据重新生成标准格式
    })

    XLSX.writeFile(workBook, filename)
}