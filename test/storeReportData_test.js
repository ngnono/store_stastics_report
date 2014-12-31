/**
 * User: ngnono
 * Date: 14-12-29
 * Time: 下午2:34
 * To change this template use File | Settings | File Templates.
 */

'use strict';


var config = require('config');
var _ = require('lodash');

var getData = require('../lib/storeReportData');

//console.log(getData);

var datas;
var dbConfig = config;//.get('DB');
console.log(dbConfig);


dbConfig.reportDate='2014-12-26';


getData.gerStoreReportData(dbConfig, function (err, recordSet) {
    if(err){
        return console.log('err:'+JSON.stringify(err))
    }
    console.log(recordSet);

    datas = recordSet;

    _.forEach(datas, function (d) {
        console.log(d);
    });

});
