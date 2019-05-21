// import * as Excel from './ExcelGenerator'
var Excel = require('./ExcelGenerator')
var handler  = require('./excelHandler')
async function test(){
  console.log(Excel.getStyle())
  let event = {
    csv : 'a,b,c\nd,e,f'}

  handler.excel(event)
    // let csv_file = await read_csv_file('./input/sheet.csv');
    // console.log(csv_file)
  
    // var wb = new xl.Workbook();
    // var ws = wb.addWorksheet('Sheet 1');
  
    
    // applyCellFromCSV(ws, csv_file);
    // applyMerge(ws, merge_input);
    // // console.log(ws.cells)
  
    // console.log(ws)
    // write_excel_file(wb, "output_excel.xlsx")
  }
  
  
  test()
   
  
  
  
  
  