/**
 * User: ngnono
 * Date: 14-12-29
 * Time: 下午3:41
 * To change this template use File | Settings | File Templates.
 */

'use strict';


var Excel = require('../lib/excel');
var config = require('config');



var excelCfg = config.get('EXCEL');

var excel = new Excel(excelCfg);

var d = [
    { Id: 3,
        DateKey: '2014-12-26',
        Store: '中文OK',
        ApplyCount: 111,
        ActiveCount: 1111,
        DailyOrderCount: 2222,
        DailyOrderAmount: 33333,
        WeeklyOrderCount: 55555,
        WeeklyOrderAmount: 77777,
        MonthlyOrderCount: 8888,
        MonthlyOrderAmount: 1122.22 },
    { Id: 2,
        DateKey: '2014-12-26',
        Store: '中文门店',
        ApplyCount: 12,
        ActiveCount: 12,
        DailyOrderCount: 12,
        DailyOrderAmount: 12,
        WeeklyOrderCount: 12,
        WeeklyOrderAmount: 12,
        MonthlyOrderCount: 12,
        MonthlyOrderAmount: 12 } ,
    { Id: 2,
        DateKey: '2014-12-26',
        Store: '中文门店',
        ApplyCount: 12,
        ActiveCount: 12,
        DailyOrderCount: 12,
        DailyOrderAmount: 12,
        WeeklyOrderCount: 12,
        WeeklyOrderAmount: 12,
        MonthlyOrderCount: 12,
        MonthlyOrderAmount: 12 }
    ,
    { Id: 2,
        DateKey: '2014-12-26',
        Store: '中文门店',
        ApplyCount: 12,
        ActiveCount: 12,
        DailyOrderCount: 12,
        DailyOrderAmount: 12,
        WeeklyOrderCount: 12,
        WeeklyOrderAmount: 12,
        MonthlyOrderCount: 12,
        MonthlyOrderAmount: 12 }
    ,
    { Id: 2,
        DateKey: '2014-12-26',
        Store: '中文门店',
        ApplyCount: 12,
        ActiveCount: 12,
        DailyOrderCount: 12,
        DailyOrderAmount: 12,
        WeeklyOrderCount: 12,
        WeeklyOrderAmount: 12,
        MonthlyOrderCount: 12,
        MonthlyOrderAmount: 12 }
    ,
    { Id: 2,
        DateKey: '2014-12-26',
        Store: '中文门店',
        ApplyCount: 12,
        ActiveCount: 12,
        DailyOrderCount: 12,
        DailyOrderAmount: 12,
        WeeklyOrderCount: 12,
        WeeklyOrderAmount: 12,
        MonthlyOrderCount: 12,
        MonthlyOrderAmount: 12 }
    ,
    { Id: 2,
        DateKey: '2014-12-26',
        Store: '中文门店',
        ApplyCount: 12,
        ActiveCount: 12,
        DailyOrderCount: 12,
        DailyOrderAmount: 12,
        WeeklyOrderCount: 12,
        WeeklyOrderAmount: 12,
        MonthlyOrderCount: 12,
        MonthlyOrderAmount: 12 }
    ,
    { Id: 2,
        DateKey: '2014-12-26',
        Store: '中文门店',
        ApplyCount: 12,
        ActiveCount: 12,
        DailyOrderCount: 12,
        DailyOrderAmount: 12,
        WeeklyOrderCount: 12,
        WeeklyOrderAmount: 12,
        MonthlyOrderCount: 12,
        MonthlyOrderAmount: 12 }
    ,
    { Id: 2,
        DateKey: '2014-12-26',
        Store: '中文门店',
        ApplyCount: 12,
        ActiveCount: 12,
        DailyOrderCount: 12,
        DailyOrderAmount: 12,
        WeeklyOrderCount: 12,
        WeeklyOrderAmount: 12,
        MonthlyOrderCount: 12,
        MonthlyOrderAmount: 12 }
    ,
    { Id: 2,
        DateKey: '2014-12-26',
        Store: '中文门店',
        ApplyCount: 12,
        ActiveCount: 12,
        DailyOrderCount: 12,
        DailyOrderAmount: 12,
        WeeklyOrderCount: 12,
        WeeklyOrderAmount: 12,
        MonthlyOrderCount: 12,
        MonthlyOrderAmount: 12 }
];

var opts = {
    sheetName: '中文的sheet',
    fileName: __dirname + '/23.xlsx',
    cellTitle: null,
    datas: d
};

excel.save(opts, function (err, rst) {
    if (err) {
        return console.log('err:' + JSON.stringify(err));
    }

    console.log('ok:' + JSON.stringify(rst));

});


var opts2 = {
    sheetName: '无的sheet',
    fileName: __dirname + '/报名名称1.xlsx',
    cellTitle: null,
    datas: d
};


var excel2 = new Excel();

excel2.save(opts2, function (err, rst) {
    if (err) {
        return console.log('err:' + JSON.stringify(err));
    }

    console.log('ok:' + JSON.stringify(rst));

});


