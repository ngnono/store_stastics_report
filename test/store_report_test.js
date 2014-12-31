/**
 * User: ngnono
 * Date: 14-12-29
 * Time: 下午5:22
 * To change this template use File | Settings | File Templates.
 */

'use strict';

var StoreReport = require('../lib/store_report');


var sr = new StoreReport();

console.log(sr.init);
sr.init();


var opts = {
    reportDate: '2014-12-26'

};

sr.report(opts);

