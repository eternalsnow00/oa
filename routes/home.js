const router = require('koa-router')()
const util = require('co-util');
const path = require('path');
const lib = require('../lib');


const pages = [
    {title: 'test', name: 'index'},
];

lib.router.page('index', pages, router);


module.exports = router
