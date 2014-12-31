/**
 * User: ngnono
 * Date: 14-12-29
 * Time: 下午4:26
 * To change this template use File | Settings | File Templates.
 */

'use strict';

var Email = require('../lib/email');
var config = require('config');
var fs = require('fs');
var Table = require('cli-table');
var _ = require('lodash');

var emailCfg = config.get('EMAIL');


var cfg={};
_.merge(cfg,emailCfg);
console.log(cfg);
var email = new Email(cfg);

var filePath = __dirname + '/报名名称1.xlsx';


var datas = [
    { Id: 3,
        DateKey: '2014-12-26',
        Store: '中文OK',
        ApplyCount: 111,
        ActiveCount: 1111
        //,
//        DailyOrderCount: 2222,
//        DailyOrderAmount: 33333,
//        WeeklyOrderCount: 55555,
//        WeeklyOrderAmount: 77777,
//        MonthlyOrderCount: 8888,
//        MonthlyOrderAmount: 1122.22
    },
    { Id: 2,
        DateKey: '2014-12-26',
        Store: '中文门店',
        ApplyCount: 12,
        ActiveCount: 12
        //,
//        DailyOrderCount: 12,
//        DailyOrderAmount: 12,
//        WeeklyOrderCount: 12,
//        WeeklyOrderAmount: 12,
//        MonthlyOrderCount: 12,
//        MonthlyOrderAmount: 12
    }
];


var keys = _.keys(datas[0]);
var chars = {
    'mid': '',
    'left-mid': '',
    'mid-mid': '',
    'right-mid': '' ,
    'top-mid':'',
    'bottom-mid':'',
    'left':'',
    'right':'',
    'middle':'',
    'top':'',
    'top-left':'',
    'top-right':'',
    'bottom-left':'',
    'bottom-right':'',
    'bottom':''
};

console.log(keys);
var table = new Table({
    chars :chars,
    head: keys//,
    //colWidths:[20,20]

      ,style:{'padding-left':5,'padding-right':5}
});

// table.push(keys);
_.forEach(datas, function (d) {
    var t = [];
    _.forEach(keys, function (k) {
        var v = d[k];
        t.push(v);
    });
    console.log(t);
    table.push(t);
});


console.log(table.toString());


fs.stat(filePath, function (err, info) {

    console.log(err);
    console.log(JSON.stringify(info));

});

//中文名 有问题
var sendInfo = {
    sender: emailCfg.storeReportSender,
    file: {
        name: '中文名OK.xlsx',
        path: filePath
    }//,
    // text: table.toString()


};


_.merge(sendInfo.sender, emailCfg.storeReportSender);
sendInfo.sender.text = table.toString();


email.send(sendInfo, function (err, rst) {
    if (err) {
        return console.log(err);
    }


    console.log(JSON.stringify(rst));

});

