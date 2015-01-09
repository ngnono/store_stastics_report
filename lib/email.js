/**
 * User: ngnono
 * Date: 14-12-27
 * Time: 下午7:50
 * To change this template use File | Settings | File Templates.
 */

'use strict';

var emailer = require('nodemailer');
var DirectTransport = require('nodemailer-smtp-transport');
var fs = require('fs');
var _ = require('lodash');
var mime = require('mime');


/**
 * email
 * @param config,config.server {host,port,auth{user,pass}}
 * @constructor
 */
function Email(config) {
    config = config || {};
    var direct = DirectTransport(config.server);
    this.sender = emailer.createTransport(direct);
}


/**
 * sender
 * @param opts opts.sender{from,to,subject,text,html} ,opts.file {name,path}
 * @param callback
 */
Email.prototype.send = function (opts, callback) {
    opts = opts || {};
    var senderInfo = opts.sender;
    var to = '';
    var cc = '';

    //console.log(JSON.stringify(senderInfo));

    if (_.isArray(senderInfo.to)) {
        _.forEach(senderInfo.to, function (t) {
            to += ',' + t;
        });
    } else {
        to = senderInfo.to;
    }

    if (_.isArray(senderInfo.cc)) {
        _.forEach(senderInfo.cc, function (t) {
            cc += ',' + t;
        });
    } else {
        cc = senderInfo.cc;
    }

    var mailInfo = {};
    _.merge(mailInfo, senderInfo);
    mailInfo.to = to;
    mailInfo.cc = cc;

    var getMime = function (name) {
        name = name || '';

        return mime.lookup(name);
    };

    //file
    if (opts.file) {
        if (_.isArray(opts.file)) {
            mailInfo.attachments = mailInfo.attachments || [];
            _.forEach(opts.file, function (f) {
                var tmp2 = {};

                if (f.name) {
                    tmp2.filename = f.name;
                    //
                }


                tmp2.path = f.path;
                tmp2.contentType = tmp2.filename ? getMime(tmp2.filename) : getMime(tmp2.path);


                mailInfo.attachments.push(tmp2);
            });
        } else {
            var tmp = {};

            if (opts.file.name) {
                tmp.filename = opts.file.name;
            }

            tmp.path = opts.file.path;
            tmp.contentType = tmp.filename ? getMime(tmp.filename) : getMime(tmp.path);
            mailInfo.attachments = tmp;
        }
    }

    this.sender.sendMail(mailInfo, function (err, info) {
        if (err) {
            return callback(err);
        }

        callback(null, info);
    });
};


module.exports = Email;