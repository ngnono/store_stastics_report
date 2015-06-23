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
    this.dataMap = opts.dataMap;
}


var borderStyle = {
    left: {
        color: '#D6DCE0',
        style: 'thin'

    },
    right: {
        color: '#D6DCE0',
        style: 'thin'
    },
    top: {
        color: '#D6DCE0',
        style: 'thin'
    },
    bottom: {
        color: '#D6DCE0',
        style: 'thin'
    }


};


var report_templat = {
    title: {

    },
    content4ColorChange: {
        single: '#FFFFFF',
        double: '#EDF3F3'
    },
    data:{
        num:{
            single:{},
            double:{}
        },
        str:{}

    }
};


Excel.prototype.save = function (opts, callback) {
    opts = opts || {};

    var wb = new excel.WorkBook();
    //console.log(JSON.stringify(opts));

    var style4Title = wb.Style();
    style4Title.Font.Bold();
    style4Title.Font.Size(10);
    style4Title.Font.Color('#ffffff');
    style4Title.Font.Family('宋体');
    style4Title.Fill.Pattern('solid');
    style4Title.Fill.Color('#9DBEC3');
    //style4Title.Font.Alignment.Horizontal('center');
    style4Title.Font.Alignment.Vertical('center');
    style4Title.Border(borderStyle);

//    var styleString = wb.Style();
//
//    //styleString.Font.Alignment.Vertical('top');
//    styleString.Font.Alignment.Horizontal('left');
//    styleString.Font.WrapText(true);
//    styleString.Font.Color('#4B7279');
//    styleString.Font.Size(10);
//    //styleString.Font.Family('Palatino Linotype');
//
//
//    styleString.Border(borderStyle);
    //styleString.Font.Bold();


    var ws = wb.WorkSheet(opts.sheetName);

    //var t = wb.Style();
    // t.Border(borderStyle);
    //ws.Cell(1, 1, 1, 6).Style(t);
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

    ws.Row(currentRow).Height(25);

//data
//    var contentRow = {
//        start: currentRow + 1,
//        end: datas.length + 1
//    };
    var contentColumn = {
        start: 1
    };

    if (self.dataMap) {

        contentColumn.end = self.dataMap.length;
    }
    else {
        contentColumn.end = keys.length;
    }

//    var strDoubleStyle = wb.Style();
//
//    //styleString.Font.Alignment.Vertical('top');
//    s.Font.Alignment.Horizontal('left');
//    s.Font.Alignment.Vertical('center');
//    s.Font.WrapText(true);
//    s.Font.Color('#4B7279');
//    s.Font.Size(10);
//    s.Font.Family('宋体');
//    s.Border(borderStyle);
//
//    s.Fill.Pattern('solid');
//    s.Fill.Color(report_templat.content4ColorChange.double);

    var numStyleSingle= wb.Style();

    numStyleSingle.Font.Size(10);
    numStyleSingle.Font.Alignment.Horizontal('right');
    numStyleSingle.Font.Alignment.Vertical('center');
    numStyleSingle.Font.Color('#4B7279');
    numStyleSingle.Border(borderStyle);
    numStyleSingle.Font.Family('宋体');
    numStyleSingle.Fill.Pattern('solid');
    numStyleSingle.Fill.Color(report_templat.content4ColorChange.single);


    var numStyleDouble = wb.Style();
    numStyleDouble.Font.Size(10);
    numStyleDouble.Font.Alignment.Horizontal('right');
    numStyleDouble.Font.Alignment.Vertical('center');
    numStyleDouble.Font.Color('#4B7279');
    numStyleDouble.Border(borderStyle);
    numStyleDouble.Font.Family('宋体');
    numStyleDouble.Fill.Pattern('solid');
    numStyleDouble.Fill.Color(report_templat.content4ColorChange.double);

    var strStyleSingle = wb.Style();

    //styleString.Font.Alignment.Vertical('top');
    strStyleSingle.Font.Alignment.Horizontal('left');
    strStyleSingle.Font.Alignment.Vertical('center');
    strStyleSingle.Font.WrapText(true);
    strStyleSingle.Font.Color('#4B7279');
    strStyleSingle.Font.Size(10);
    strStyleSingle.Font.Family('宋体');
    strStyleSingle.Border(borderStyle);
    strStyleSingle.Fill.Pattern('solid');
    strStyleSingle.Fill.Color(report_templat.content4ColorChange.single);

    var strStyleDouble = wb.Style();
    strStyleDouble.Font.Alignment.Horizontal('left');
    strStyleDouble.Font.Alignment.Vertical('center');
    strStyleDouble.Font.WrapText(true);
    strStyleDouble.Font.Color('#4B7279');
    strStyleDouble.Font.Size(10);
    strStyleDouble.Font.Family('宋体');
    strStyleDouble.Border(borderStyle);
    strStyleDouble.Fill.Pattern('solid');

    strStyleDouble.Fill.Color(report_templat.content4ColorChange.double);

    //var rouNum = 1;
    //get first data
    var xiaoji = datas.shift();
    //edit storename
    if(xiaoji){
        if(xiaoji['Store']){
            xiaoji['Store'] = ' ';
        }

    }


    var dataHandle = function(datas){
        datas.forEach(function (data) {
            currentHeight = 1;
            currentRow += 1;
            ws.Row(currentRow).Height(25);
            //console.log(rouNum);
            if (self.dataMap) {
                self.dataMap.forEach(function (title) {
                    var cell = ws.Cell(currentRow, currentHeight);
                    var dataMapKey = title.dataMapKey;
                    var dataType = title.type || '';
                    var val = data[dataMapKey] || '';

                    var isNum = false;
                    switch (dataType) {
                        case 'num':
                            isNum = true;
                            cell.Number(val || 0);
                            break;
                        default :

                            cell.String(val);
                            break;
                    }

                    var s = {};
                    if (currentRow % 2 === 0) {
                        if(isNum){
                            s = numStyleDouble;
                        }else{
                            s = strStyleDouble;
                        }

                    } else {
                        if(isNum){
                            s = numStyleSingle;
                        }else{
                            s = strStyleSingle;
                        }

                    }

                    cell.Style(s);

                    currentHeight += 1;
                });
            }
            else {

                var keys = _.keys(data);
                _.forEach(keys, function (k) {
                    var cell = ws.Cell(currentRow, currentHeight);


                    var s = {};
                    if (currentRow % 2 === 0) {

                        s = strStyleDouble;

                    } else {
                        s = strStyleSingle;

                    }

                    cell.Style(s).String(data[k].toString());

                    currentHeight += 1;
                });
            }

        });
    };

    dataHandle(datas);
    dataHandle([xiaoji]);


// Synchronously write file
    wb.write(fileName, function (err) {
        if (err) {
            return callback(err);
        }

        callback(null, 'file saved');
    });
};