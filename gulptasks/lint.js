//******************************************************************************
//* lint.js
//*
//* Defines a custom gulp task for ensuring that all source code in
//* this repository follows recommended TypeScript practices. 
//*
//* Rule violations are output automatically to the console.
//******************************************************************************

var gulp = require("gulp"),
    tslint = require("gulp-tslint"),
    config = require('./@configuration.js');

gulp.task("lint", function () {
    return gulp.src(config.paths.sourceGlob)
        .pipe(tslint({ formatter: "prose" }))
        .pipe(tslint.report({ emitError: false }));
});