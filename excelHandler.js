
// import * as Excel from './ExcelGenerator'
var Excel = require('./ExcelGenerator')
const merge_input = [
  {
  from : "A1",
  to : "E1"
  }
]

const csv_path = './input/sheet.csv'

module.exports.excel = async (event, context, callback) => {
  // let csv = await Excel.read_csv_file(csv_path)
  let csv = await Excel.read_csv_string(event.csv)
  let wb = await Excel.generate(csv, merge_input, Excel.getStyle)
  // wb.write("output.xlsx", res)
  callback(null, wb.write('output.xlsx'))
}

