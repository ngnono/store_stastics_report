#! /usr/bin/env node

/**
 * User: ngnono
 * Date: 14-12-31
 * Time: 下午2:03
 * To change this template use File | Settings | File Templates.
 */

'use strict';

var argv = require('optimist').argv;
var app = require('./lib/index');

if (argv.reportdate && !argv.reportDate) {
    argv.reportDate = argv.reportdate;
    delete argv.reportdate;
}

//console.log(argv);

app.storeReport(argv);


//app.storeReport();