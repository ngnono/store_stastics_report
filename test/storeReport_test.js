/**
 * User: ngnono
 * Date: 14-12-29
 * Time: 下午5:22
 * To change this template use File | Settings | File Templates.
 */

'use strict';

var StoreReport = require('../lib/storeReport');


var sr = new StoreReport();

var opts = {
    reportDate: '2015-1-26'

};

sr.report(opts);

