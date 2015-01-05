/**
 * User: ngnono
 * Date: 14-12-27
 * Time: 下午4:33
 * To change this template use File | Settings | File Templates.
 */

'use strict';


var mssql = require('mssql');
var _ = require('lodash');

/**
 * DB
 * @param opts
 * @constructor
 */
function DB(opts) {
    this.options = opts || {};
}

/**
 * 查询
 * @param sql
 * @param opts
 * @param callback
 */
DB.prototype.query = function (sql, opts, callback) {

    if (_.isFunction(opts)) {
        callback = opts;
        opts = null;
    }

    mssql.connect(this.options, function (err) {
        if (err) {
            return callback(err);
        }

        var db = new mssql.Request();

        db.query(sql, function (err, recordSet) {
            if (err) {
                mssql.close();
                return callback(err);
            }

            mssql.close();
            callback(null, recordSet);
        });
    });
};


module.exports = DB;
