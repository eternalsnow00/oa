const Koa = require('koa')
const app = new Koa()
const router = new require('koa-router')();
const config = require('./config');
const views = require('koa-views')
const routes = require('./routes');
require('co-log')(config.log);

var indexRoute=require('./lib/excel')(router);
app.use(require('koa-static')(__dirname + '/public'));
app.use(views(__dirname + '/views', {
    map: {html: 'ejs'},
    extension: 'ejs'
}));

app.use(indexRoute.routes())

for (let key in routes) {
    console.log(key)
    if (key == 'home')
        router.use('', routes[key].routes())
    else
        router.use('/' + key, routes[key].routes())
}

// 加载路由中间件
app.use(router.routes()).use(router.allowedMethods())


module.exports = app;
