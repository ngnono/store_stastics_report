/**
 * User: ngnono
 * Date: 14-12-29
 * Time: 下午2:27
 * To change this template use File | Settings | File Templates.
 */

'use strict';

var DB = require('./db');
var moment = require('moment');

var db;
exports.gerStoreReportData = getStoreReportData;


var sql = "SELECT * FROM [dbo].[IMS_StoreStastics] WITH(nolock) WHERE [DateKey] = @DateKey ORDER BY [DailyOrderAmount] DESC";

/**
 * GET 门店 报表 数据
 * @param opts   opts.reportDate,opts.DB
 * @param callback
 */
function getStoreReportData(opts, callback) {
    if (!db) {
        db = new DB(opts.DB);
    }

    var date = opts.reportDate;

    //2014-1-1
    var DateKey = moment(date).format('YYYY-M-D');

    var params = {
        DateKey: DateKey
    };

    var s = sql.replace(/[@](\w+)+/g, function (w, w2) {
        return '\'' + params[w2] + '\'' || w;
    });

    //console.log('sql:' + s);

    db.query(s, function (err, datas) {
        if (err) {
            return callback(err);
        }

        callback(null, datas);
    });
}