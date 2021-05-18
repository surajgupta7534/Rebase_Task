const fs = require('fs');

const express = require('express')

const app = express();

const port =3000;

const xl = require('excel4node');

let noTrade = 0;

var rawData = (fs.readFileSync('strategy-builder.json'));

var Data = JSON.parse(rawData);

var imp = (Data.data);
var x;
imp.forEach(element => {
    x = element.legs;
   noTrade++;
});

const wb = new xl.Workbook();

const ws = wb.addWorksheet("Sheet 1");

var t = "Trade No";

const ColNames = ["Trade No","Lots", "Legs","Entry Date","Strike","B/S","Options","Entry Price","Exit Date","Exit Price","Days","Profit","Total Profits"];

let ColIndex = 1;


ColNames.forEach(heading=>{
    ws.cell(1,ColIndex++).string(heading);
})

let rowIndex = 2;

let j=1;
imp.forEach(element => {
    let colInd = 1;
    ws.cell(rowIndex,colInd).number(j);
    let i=0;
    let pro=0;
    var c =0;
    while(i<(element.legs).length)
    {
        var subCol=2;
        ws.cell(rowIndex,subCol++).number(element.legs[i].lots)
        ws.cell(rowIndex,subCol++).string(element.legs[i].legName)
        ws.cell(rowIndex,subCol++).string(element.legs[i].entryDate)
        ws.cell(rowIndex,subCol++).number(element.legs[i].strikePrice)
        ws.cell(rowIndex,subCol++).string(element.legs[i].buyOrSell)
        ws.cell(rowIndex,subCol++).string(element.legs[i].futuresOrOptions)
        ws.cell(rowIndex,subCol++).number(element.legs[i].entryValue)
        ws.cell(rowIndex,subCol++).string(element.legs[i].exitDate)
        ws.cell(rowIndex,subCol++).number(element.legs[i].exitValue)
        ws.cell(rowIndex,subCol++).number(7)
        ws.cell(rowIndex,subCol++).number((element.legs[i].exitValue-x[i].entryValue)*75)
        pro = pro + (element.legs[i].exitValue-element.legs[i].entryValue);
        i++;
        rowIndex++;
        c = subCol;
    }      
    ws.cell(rowIndex--,c--).number(pro*75);
    rowIndex++;
    j++;
});
wb.write('Data.xlsx');

app.get('/', (req, res) => {
    const file ='Data.xlsx';
    res.download(file)
    console.log('Download Finish.')
})
app.listen(port, () => {
    console.log(`Example app listening at http://localhost:${port}`)
})


