const xl = require('excel4node');
const fs = require('fs')
const csv = require('csv-parse')
const csv_string = require('csv-string')


// =================== READ FILE =========================
async function read_csv_file(file_path){
  return new Promise( (resolve, reject) => {
    let csv_file = []
    fs.createReadStream(file_path)  
    .pipe(csv())
    .on('data', (row) => {
      // console.log(row);
      csv_file.push(row)
    })
    .on('end', () => {
      console.log('CSV file successfully processed');
      resolve(csv_file);
    });
    // reject("Error")
  });
}

function read_csv_string(csv){
  return csv_string.parse(csv);
}
// =======================================================


 
function getStyle(){
  var wb = new xl.Workbook();
  let style = wb.createStyle({
    font: {
      color: '#FF0800',
      size: 12,
    },
    numberFormat: '$#,##0.00; ($#,##0.00); -',
  });
  console.log(style)
  return style;
}

 
function applyData(worksheet, row, col, type ,value){
  // console.log(type, value)
  type = type.toLowerCase()
  switch(type)
  {
      case "num":
      case "number":
          worksheet.cell(row, col).number(Number(value));
          break;
      case "str":
      case "string":
          worksheet.cell(row, col).string(String(value));
          break;
      case "formula":
          worksheet.cell(row, col).formula((String(value)).substring(1));
          break;
      default:
          // callback("400 Invalid Operator");
          console.log("Invalid type of data")
          break;
  }
}

function getType(value){
  if(value === '' || isNaN(value)){
    if(value[0] === '=')
      return "formula"
    else
      return "string" 
  }
  return "number"
}

function applyCellFromCSV(worksheet, cells){
    console.log("applying cell ")
    // console.log(cells.length)

    for(let i = 0 ; i <cells.length; i++){
        // console.log(cells[i])
        for(let j = 0 ; j < cells[i].length; j++){
          // console.log('Applying cell ', i+1, j+1)
          applyData(worksheet, i+1, j+1, getType(cells[i][j]), cells[i][j]);
        }
        
    }
  }

function applyStyle(worksheet, styles){
  console.log("applying styles")
  // console.log(styles.length)
  for(let i = 0 ; i <styles.length; i++){
      console.log(styles[i])
      worksheet.cell(styles[i].from.row, styles[i].from.col, styles[i].to.row, styles[i].to.col).style(styles[i].style);
  }
}

function applyMerge(worksheet, merge_input){
  console.log("Applying merge ", merge_input)
  for(let i = 0; i<merge_input.length; i++){
    let from  = xl.getExcelRowCol(merge_input[i].from);
    let to = xl.getExcelRowCol(merge_input[i].to);
    console.log(from, to)
    worksheet.cell(from.row, from.col, to.row, to.col, true);
  }
}

async function write_excel_file(wb, fileName){
    wb.write(fileName, (err, stats) => {
      if (err) {
        console.error(err);
      } else {
        // console.log(stats); // Prints out an instance of a node.js fs.Stats object
      }
  });
}

async function generate(csv_data, merge_input, style_input){
  console.log(csv_data, merge_input, style_input)
  
  var wb = new xl.Workbook();
  var ws = wb.addWorksheet('Sheet 1');
  
  applyCellFromCSV(ws, csv_data);
  if(merge_input) applyMerge(ws, merge_input);
  
  try{
    applyStyle(ws,wb.createStyle(style_input))
  }
  catch(err){
    console.log("style", err)
  }
  return wb;
}

module.exports = { 
  read_csv_file, 
  read_csv_string, 
  generate, 
  getStyle, 
  write_excel_file
}
