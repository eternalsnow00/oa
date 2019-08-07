"use strict";
let models = [
    'router',
    'excel'
];
models.forEach(function (name) {
    module.exports[name] = require('./' + name);
});


