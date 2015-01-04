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


    //var rouNum = 1;
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

                var s;// =  wb.Style();

                switch (dataType) {
                    case 'num':

                        s = wb.Style();

                        s.Font.Size(10);
                        s.Font.Alignment.Horizontal('right');
                        s.Font.Alignment.Vertical('center');
                        s.Font.Color('#4B7279');
                        s.Border(borderStyle);
                        s.Font.Family('宋体');



                        cell.Number(val || 0);
                        break;
                    default :
                        s = wb.Style();

                        //styleString.Font.Alignment.Vertical('top');
                        s.Font.Alignment.Horizontal('left');
                        s.Font.Alignment.Vertical('center');
                        s.Font.WrapText(true);
                        s.Font.Color('#4B7279');
                        s.Font.Size(10);
                        s.Font.Family('宋体');
                        s.Border(borderStyle);
                        //s = styleString.clone();
                        cell.String(val);
                        break;
                }

                //var fill = ;
                s.Fill.Pattern('solid');
                if (currentRow % 2 === 0) {
                    s.Fill.Color(report_templat.content4ColorChange.double);
                } else {
                    s.Fill.Color(report_templat.content4ColorChange.single);
                }

                cell.Style(s);

                currentHeight += 1;
            });
        }
        else {

            var keys = _.keys(data);
            _.forEach(keys, function (k) {
                var cell = ws.Cell(currentRow, currentHeight);
                //cell
                var s = wb.Style();

                //styleString.Font.Alignment.Vertical('top');
                s.Font.Alignment.Horizontal('left');
                s.Font.Alignment.Vertical('center');
                s.Font.WrapText(true);
                s.Font.Color('#4B7279');
                s.Font.Size(8);
                s.Font.Family('宋体');
                s.Border(borderStyle);


                s.Fill.Pattern('solid');
                if (currentRow % 2 === 0) {

                    //fff
                    s.Fill.Color(report_templat.content4ColorChange.double);
                } else {
                    s.Fill.Color(report_templat.content4ColorChange.single);
                }

                cell.Style(s).String(data[k].toString());//.String(val);

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