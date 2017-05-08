const xlsx = require('./index')

let data = xlsx.parse('zx.xls')

let second = data[1] //选择第二张工作表
second.name = '装修' //重新命名工作表
second.data.unshift(['新插入的数据']) //操作工作表
console.log(`Data: ${JSON.stringify(second)}`)

xlsx.build([second], 'out.xlsx')