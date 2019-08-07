"use strict";
let models = [
    'home',
];
models.forEach(function (name) {
    exports[name]=require('./' + name);
});


