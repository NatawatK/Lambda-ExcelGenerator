

module.exports.excel = async (event, context, callback) => {
    var EX = require('testExcel4Node');
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
 
    
}