/**
 * Created by iceli on 17/5/5.
 */
const fs = require('fs')
const XLSX = require('xlsx')
const isString = (str) => typeof str === 'string'

let workBook = {
    SheetNames: [],
    Sheets: {}
}

function parse(mixed, options = {}) {
    const workSheet = XLSX[isString(mixed) ? 'readFile' : 'read'](mixed, options);
    return Object.keys(workSheet.Sheets).map((name) => {
        const sheet = workSheet.Sheets[name];
        return {name, data: XLSX.utils.sheet_to_json(sheet, {header: 1, raw: true})};
    });
}

function getWorkBook(name, wb) {
    workBook.SheetNames.push(name)
    workBook.Sheets[name] = wb

    return workBook
}

let data = parse('zx.xls')
let second = data[1] //选择第二张工作表
second.data.unshift(['新插入的数据'])
//console.dir("Data: " + JSON.stringify(second.data));
let name = '装修' //重新命名第二张工作表
let wb = XLSX.utils.aoa_to_sheet(second.data) //将修改过的数据重新生成标准格式
//console.dir("Data: " + JSON.stringify(wb));

XLSX.writeFile(getWorkBook(name, wb), 'out.xlsx')


