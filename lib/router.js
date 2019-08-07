/**
 * Created by Administrator on 2018/7/5.
 */
"use strict";
const util = require('co-util');

exports.page = (name, pages, router) => {
    for (let page of pages) {
        router.get(util.format('/%s.html', page.name), async (ctx, next) => {
            let data = {};
            if (page.cb) data = await page.cb(ctx, page);
            data.title = data.title || page.title;
            await ctx.render(name + "/" + page.name, data);
        })
    }
}
