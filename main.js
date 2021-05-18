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

function calDay(s1,s2) {
    var d1 = parseInt(s1.substring(8,10), 10)
    var d2 = parseInt(s2.substring(8,10), 10)

    var m1 = parseInt(s1.substring(5,7))
    var m2 = parseInt(s2.substring(5,7))

    var y1 =  parseInt(s1.substring(0,4))
    var y2 =  parseInt(s1.substring(0,4))

    var td;

    if(y1>y2 || m1>m2)
    {
        td = (30-d2)+d1;
    }
    else
    {
        td = d1-d2;
    }
    return td;
}

const wb = new xl.Workbook();


var t = "Trade No";

const ColNames = ["Trade No","Lots", "Legs","Entry Date","Strike","B/S","Options","Entry Price","Exit Date","Exit Price","Days","Profit","Total Profits"];

let ColIndex = 1;

const aliStyle = wb.createStyle({
    alignment: {
        wrapText: true,
        horizontal: 'center',
      },
      border: {
		left: {
			style: 'thin',
			color: 'black',
		},
		right: {
			style: 'thin',
			color: 'black',
		},
		top: {
			style: 'thin',
			color: 'black',
		},
		bottom: {
			style: 'thin',
			color: 'black',
		},
		outline: false,
	},
})

const ws = wb.addWorksheet("Sheet 1");

ColNames.forEach(heading=>{
    ws.cell(1,ColIndex++).string(heading);
})

let rowIndex = 2;



const bgStyle = wb.createStyle({
    fill: {
        type: "pattern",
        patternType: "solid",
       bgColor: '#33FF35',
       fgColor: '#33FF35'
    },
    alignment: {
        wrapText: true,
        horizontal: 'center',
      },
      border: {
		left: {
			style: 'thin',
			color: 'black',
		},
		right: {
			style: 'thin',
			color: 'black',
		},
		top: {
			style: 'thin',
			color: 'black',
		},
		bottom: {
			style: 'thin',
			color: 'black',
		},
		outline: false,
	},
  });

const bgStyle2 = wb.createStyle({
    fill: {
        type: "pattern",
        patternType: "solid",
        bgColor: '#edc266',
        fgColor: '#edc266'
    },
    alignment: {
        wrapText: true,
        horizontal: 'center',
    },
    border: {
		left: {
			style: 'thin',
			color: 'black',
		},
		right: {
			style: 'thin',
			color: 'black',
		},
		top: {
			style: 'thin',
			color: 'black',
		},
		bottom: {
			style: 'thin',
			color: 'black',
		},
		outline: false,
	},
      
});

for(var i=1;i<ColNames.length;i++)
{
    ws.column(i).setWidth(15)
    ws.column(i).a
}


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
        ws.cell(rowIndex,subCol++).number(element.legs[i].lots).style(aliStyle)
        ws.cell(rowIndex,subCol++).string(element.legs[i].legName).style(aliStyle)
        ws.cell(rowIndex,subCol++).string(element.legs[i].entryDate).style(aliStyle)
        ws.cell(rowIndex,subCol++).number(element.legs[i].strikePrice).style(aliStyle)
        ws.cell(rowIndex,subCol++).string(element.legs[i].buyOrSell).style(aliStyle)
        ws.cell(rowIndex,subCol++).string(element.legs[i].futuresOrOptions).style(aliStyle)
        ws.cell(rowIndex,subCol++).number(element.legs[i].entryValue).style(bgStyle);
        ws.cell(rowIndex,subCol++).string(element.legs[i].exitDate).style(bgStyle);
        ws.cell(rowIndex,subCol++).number(element.legs[i].exitValue).style(aliStyle)
        var day = calDay(element.legs[i].exitDate,element.legs[i].entryDate)
        ws.cell(rowIndex,subCol++).number(day).style(bgStyle);
        ws.cell(rowIndex,subCol++).number((element.legs[i].exitValue-x[i].entryValue)*75).style(bgStyle2);
        pro = pro + (element.legs[i].exitValue-element.legs[i].entryValue);
        i++;
        rowIndex++;
        c = subCol;
    }      
    ws.cell(rowIndex--,c--).number(pro*75).style(bgStyle2);
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


