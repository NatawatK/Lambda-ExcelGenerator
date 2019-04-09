var xl = require('excel4node');

var wb = new xl.Workbook();
var ws = wb.addWorksheet('Sheet 1');
 
var style = wb.createStyle({
  font: {
    color: '#FF0800',
    size: 12,
  },
  numberFormat: '$#,##0.00; ($#,##0.00); -',
});
 
var input = {
    cells : [
        {
		    row : 1,
		    col : 1,
		    type : "string",
		    value : "Hello Test",
        },
        {
          row : 2,
          col : 1,
          type : "number",
          value : 1,
        },
        {
          row : 2,
          col : 2,
          type : "number",
          value : 2,
        },
        {
          row : 2,
          col : 3,
          type : "formula",
          value : "A2+B2",
        },
    ],
    styles : [
        {
            from : {row : 1, col: 1},
            to : {row: 10, col : 10},
            style : style
        }
    ]
}
console.log(input)

applyData = function applyData(worksheet, row, col, type ,value){
  type = type.toLowerCase()
  switch(type)
  {
      case "num":
      case "number":
          worksheet.cell(row, col).number(value);
          break;
      case "str":
      case "string":
          worksheet.cell(row, col).string(value);
          break;
      case "formula":
          worksheet.cell(row, col).formula(value);
          break;
      default:
          // callback("400 Invalid Operator");
          console.log("Invalid type of data")
          break;
  }
}

function applyCell(worksheet, cells){
    console.log("applying cell ")
    console.log(cells.length)
    for(let i = 0 ; i <cells.length; i++){
        console.log(cells[i])
        applyData(worksheet, cells[i].row, cells[i].col, cells[i].type, cells[i].value);
    }
}

function applyStyle(worksheet, styles){
    console.log("applying styles")
    console.log(styles.length)
    for(let i = 0 ; i <styles.length; i++){
        console.log(styles[i])
        worksheet.cell(styles[i].from.row, styles[i].from.col, styles[i].to.row, styles[i].to.col).style(styles[i].style);
    }
}
// Set value of cell A1 to 100 as a number type styled with paramaters of style
// ws.cell(1, 1)
//   .number(100)
//   .style(style);
 
// // Set value of cell B1 to 200 as a number type styled with paramaters of style
// ws.cell(1, 2)
//   .number(200)
//   .style(style);
 
// // Set value of cell C1 to a formula styled with paramaters of style
// ws.cell(1, 3)
//   .formula('SUM(A1:B1)')
//   .style(style);
 
// // Set value of cell A2 to 'string' styled with paramaters of style
// ws.cell(2, 1)
//   .string('string')
//   .style(style);
 
// // Set value of cell A3 to true as a boolean type styled with paramaters of style but with an adjustment to the font size.
// ws.cell(3, 1)
//   .bool(true)
//   .style(style)
//   .style({font: {size: 14}});
 
applyCell(ws, input.cells);
applyStyle(ws, input.styles);
// wb.write('Excel.xlsx');
wb.write('ExcelFile.xlsx', function(err, stats) {
    if (err) {
      console.error(err);
    } else {
      console.log(stats); // Prints out an instance of a node.js fs.Stats object
    }
  });



