/*
 * sales_report_task
 * user/repo
 *
 * Copyright (c) 2014 
 * Licensed under the MIT license.
 */

'use strict';

var SR = require('./store_report');


exports.awesome = function () {
    return 'awesome';
};


var sr = new SR();
sr.init();

exports.store_report = function (args) {
    sr.report(args);
};
