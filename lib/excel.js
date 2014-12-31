/**
 * User: ngnono
 * Date: 14-12-27
 * Time: 下午6:22
 * To change this template use File | Settings | File Templates.
 */

'use strict';

var excel = require('excel4node');
var _ = require('lodash');


module.exports = Excel;

function Excel(opts) {
    opts = opts || {};

    if (!opts.dataMap) {

    }

    this.dataMap = opts.dataMap;


}

var borderStyle =          {
    left:{
        color:'#000000',
        style:'thin'

    },
    right:{
        color:'#000000',
        style:'thin'
    },
    top:{
        color:'#000000',
        style:'thin'
    },
    bottom:{
        color:'#000000',
        style:'thin'
    }


};

Excel.prototype.save = function (opts, callback) {
    opts = opts || {};

    var wb = new excel.WorkBook();


    console.log(JSON.stringify(opts));

    var style4Title = wb.Style();
    //style4Title.Font.Bold();
    style4Title.Font.Size(13);
    style4Title.Font.Color('#222222');
    style4Title.Font.Family('微软雅黑');
    style4Title.Fill.Pattern('solid');
    style4Title.Fill.Color('#c8c5c5');
    style4Title.Border(
        borderStyle
    );

    var styleString = wb.Style();

    styleString.Font.Alignment.Vertical('top');
    styleString.Font.Alignment.Horizontal('left');
    styleString.Font.WrapText(true);
    styleString.Font.Color('#222222');
    styleString.Font.Size(12);
    styleString.Font.Family('微软雅黑');
    styleString.Border(borderStyle);
    //styleString.Font.Bold();

    var styleNum = wb.Style();

    styleNum.Font.Size(12);
    styleNum.Font.Alignment.Horizontal('right');
    styleNum.Font.Color('#222222');
    styleNum.Border(borderStyle);
    styleNum.Font.Family('Arial');




    var ws = wb.WorkSheet(opts.sheetName);

    var t = wb.Style();
    t.Border(borderStyle);
    ws.Cell(1,1,1,6).Style(t);


    var fileName = opts.fileName;

    var datas = opts.datas || [];
    var cellTitle = opts.cellTitle;


    //title
    var currentRow = 1;
    var currentHeight = 1;

    var tableLength;
    if (this.dataMap) {
        tableLength = this.dataMap.length;
    } else {
        if (!datas || !datas[0]) {
            console.log('datas is empty or null.')
        } else {
            var keys = _.keys(datas[0]);
            tableLength = keys.length;

        }
    }

    if (cellTitle) {
        ws.Cell(currentRow, currentHeight, currentRow, tableLength, true).String(cellTitle);
        currentRow += 1;
    }

    var self = this;

    ws.Cell(currentRow, 1, currentRow, tableLength).Style(style4Title);


    if (self.dataMap) {
        self.dataMap.forEach(function (title) {
            ws.Cell(currentRow, currentHeight).String(title.desc);

            currentHeight += 1;
        });
    } else {
        var keys = _.keys(datas[0]);
        keys.forEach(function (title) {
            ws.Cell(currentRow, currentHeight).String(title);

            currentHeight += 1;
        });
    }

//data
    datas.forEach(function (data) {
        currentHeight = 1;
        currentRow += 1;

        if (self.dataMap) {
            self.dataMap.forEach(function (title) {
                var cell = ws.Cell(currentRow, currentHeight);

                var dataMapKey = title.dataMapKey;
                var dataType = title.type || '';
                var val = data[dataMapKey] || '';


                switch (dataType) {
                    case 'num':
                        cell.Style(styleNum).Number(val);
                        break;
                    default :
                        cell.Style(styleString).String(val);
                        break;
                }

                currentHeight += 1;

            });
        } else {

            var keys = _.keys(data);
            _.forEach(keys, function (k) {
                var cell = ws.Cell(currentRow, currentHeight);
                cell.Style(styleString).String(data[k].toString());
                currentHeight += 1;
            });

        }


    });


// Synchronously write file
    wb.write(fileName, function (err) {
        if (err) {
            return callback(err);
        }

        callback(null, 'file saved');
    });

};

