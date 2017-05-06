> 数据导入导出经常会用到Excel，静态网站如何快速将Excel数据转为json数据，这里借助强大的 [xlsx库](https://www.npmjs.com/package/xlsx) 可以轻松搞定。基本上是一个熟悉Excel数据结构和xlsx API的过程

- `XLSX.readFile`首先读取xlsx文件可以获得整个数据结构，大概包含以下字段
` ["opts","Directory","SheetNames","Sheets","Preamble","Strings","SSF","Metadata","Workbook","Custprops","Props"]`
- 我们可以通过`SheetNames`，`Sheets`取到我们想要的数据
- `XLSX.utils.sheet_to_json` 将`Sheets`转化为json格式后即可方便的进行操作
- `XLSX.utils.aoa_to_sheet` 将json数组转为标准的工作表格式
- `XLSX.writeFile` 将一个至少包含`SheetNames`，`Sheets`字段的工作簿保存为新的文件，否则抛出异常
` if(!wb || !wb.SheetNames || !wb.Sheets) throw new Error("Invalid Workbook");`

### 注意工作簿的数据结构
```js
let workBook = {
    SheetNames: [],
    Sheets: {}
}
```