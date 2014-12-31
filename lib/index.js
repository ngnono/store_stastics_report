/*
 * sales_report_task
 * user/repo
 *
 * Copyright (c) 2014 
 * Licensed under the MIT license.
 */

'use strict';

var SR = require('./storeReport');

var sr = new SR();

exports.storeReport = function (args) {
    sr.report(args);
};
