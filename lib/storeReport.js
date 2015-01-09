/**
 * User: ngnono
 * Date: 14-12-27
 * Time: 下午4:18
 * To change this template use File | Settings | File Templates.
 */

'use strict';


var moment = require('moment');
var _ = require('lodash');
var path = require('path');
var DB = require('./db');
var config = require('config');
var util = require('util');
var events = require('events');
var Excel = require('./excel');
var EmailSender = require('./email');
var Datas = require('./storeReportData');
var TABLE = require('cli-table');


var chars = {
    'mid': '',
    'left-mid': '',
    'mid-mid': '',
    'right-mid': '',
    'top-mid': '',
    'bottom-mid': '',
    'left': '',
    'right': '',
    'middle': '',
    'top': '',
    'top-left': '',
    'top-right': '',
    'bottom-left': '',
    'bottom-right': '',
    'bottom': ''
};


/**
 * 获取邮件主体
 * @param opts
 * @returns {*}
 */
function getTextData(opts) {

    var keys = opts.keys;
    var datas = opts.datas;

    var table = new Table({
        chars: chars,
        head: keys,
        //colWidths:[20,20]

        style: {'padding-left': 5, 'padding-right': 5}
    });


    _.forEach(datas, function (d) {
        var t = [];
        _.forEach(keys, function (k) {
            var v = d[k];
            t.push(v);
        });
        console.log(t);
        table.push(t);
    });


    return table.toString();
}


/**
 * 门店 统计
 * @constructor
 */
function Report(opts) {
    opts = opts || {};
    var cfgs = opts.config || config;
    this.db = new DB(cfgs.get('DB'));
    this.excel = new Excel(cfgs.get('EXCEL'));

    var emailCfg = {};
    _.merge(emailCfg, cfgs.get('EMAIL'));

    this.emailSender = new EmailSender(emailCfg);
    this.getStoreReportDatas = Datas.gerStoreReportData;

    events.EventEmitter.call(this);

    this._init();
}


util.inherits(Report, events.EventEmitter);


function ok(obj) {
    console.log(JSON.stringify(obj));
    console.log('ok');
}


function error(err) {
    console.log(err);
}


/**
 * 初始化监听器
 */
Report.prototype._init = function () {

    /**
     * 数据处理
     */
    this.on('data', this.data2ExcelHandler);

    /**
     * 发送
     */
    this.on('send', this.sendEmailHandler);

    this.on('end', ok);

    this.on('error', error);
};


/**
 * 写文件
 * @param datas
 * @param callback
 */
Report.prototype.data2ExcelHandler = function (opts) {
    opts = opts || {};
    if (!opts.datas) {
        return this.emit('error', 'opts.datas is null');
    }

    var fileName = opts.fileName || 'store_' + opts.reportDate.format('YYYY-M-D') + '.xlsx';

    opts.file = {
        name: opts.reportDate.format('YYYY-M-D') + '.xlsx',
        path: path.normalize(__dirname + '/../files/' + fileName)
    };

    //console.log(JSON.stringify(opts.file));
    var self = this;
    this.excel.save({
        datas: opts.datas,
        sheetName: opts.reportDate.format('YYYY-MM-DD'),
        fileName: opts.file.path,
        cellTitle: opts.cellTitle || null
    }, function (err) {
        if (err) {
            return self.emit('error', err);
        }

        self.emit('send', opts);
    });
};


/**
 * 发送EMAL
 * @param opts
 * @param callback
 */
Report.prototype.sendEmailHandler = function (opts) {
    opts = opts || {};

    opts.sender = opts.sender || {};

    _.merge(opts.sender, config.get('EMAIL').storeReportSender);

    //console.log(JSON.stringify(opts.sender));
    if (!opts.sender.subject) {
        opts.sender.subject = '迷你银店铺数据' + opts.reportDate.format('YYYY-M-D');
    }

    //取DATA TOP 10
    var self = this;
    this.emailSender.send(opts, function (err, info) {
        if (err) {
            return self.emit('error', err);
        }
        opts.sendResult = info;

        self.emit('end', opts);
    });
};


/**
 * 报表程序
 * @param opts
 */
Report.prototype.report = function (opts) {
    opts = opts || {};

    //报表日期未当天的前一天（昨日）
    if (opts.reportDate && _.isString(opts.reportDate)) {
        opts.reportDate = moment(opts.reportDate);
    }

    opts.reportDate = opts.reportDate || moment().subtract(1, 'day');

    console.log('reportDate:'+opts.reportDate.toString('YYYY-MM-DD'));
    var self = this;
    this.getStoreReportDatas({
        DB: config.get('DB'),
        reportDate: opts.reportDate
    }, function (err, datas) {
        if (err) {
            return self.emit('error', err);
        }

        if (!datas || datas.length == 0) {
            return self.emit('error', '报表数据为空,reportDate' + opts.reportDate);
        }

        opts.datas = datas;

        self.emit('data', opts);
    });
};


module.exports = Report;