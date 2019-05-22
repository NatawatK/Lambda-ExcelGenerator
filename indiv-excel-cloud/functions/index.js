const functions = require('firebase-functions');
const express = require('express')
var Excel = require('./ExcelGenerator')

const csv_stringify = require('csv-stringify')
var app = express()

app.post('/csv2excel', async (req, res) => {
    let csv = req.body.csv;
    let merge = req.body.merge || [];
    let style = req.body.style;
    console.log("csv ", csv)
    console.log("merge ", merge)
    console.log("style", style)
    if(!csv){
        res.send(400, "invalid query : CSV")
        return
    }
    
    let wb = await Excel.generate(csv, merge, style)
    // wb.write("output.xlsx")
    let filename = "output.xlsx"
    wb.write(filename, (err, stats) => {
        if (err) {
          console.error(err);
          res.status(500).send("write file error")
        } else {
            console.log("write file complete")
            res.download(filename)
            // console.log(stats); // Prints out an instance of a node.js fs.Stats object
        }
    });
})

app.get('/sample_csv', async (req, res) => {
    let csv = await Excel.read_csv_file('sheet.csv')
    res.send(csv)
})

exports.api = functions.https.onRequest(app)