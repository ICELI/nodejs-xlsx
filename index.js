/**
 * Created by iceli on 17/5/5.
 */
const fs = require('fs')
const XLSX = require('xlsx')
const isString = (str) => typeof str === 'string'

function parse(mixed, options = {}) {
    const workSheet = XLSX[isString(mixed) ? 'readFile' : 'read'](mixed, options);
    return Object.keys(workSheet.Sheets).map((name) => {
        const sheet = workSheet.Sheets[name];
        return {name, data: XLSX.utils.sheet_to_json(sheet, {header: 1, raw: true})};
    });
}
let data = parse('zx.xls')
console.dir("Data: " + JSON.stringify(data[1]));


