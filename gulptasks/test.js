//******************************************************************************
//* test.js
//*
//* Defines a custom gulp task for executing the unit tests (with mocha) and
//* also reporting on code coverage (with istanbul).
//******************************************************************************

var gulp = require("gulp"),
    mocha = require("gulp-mocha"),
    istanbul = require("gulp-istanbul"),
    tsc = require("gulp-typescript"),
    yargs = require('yargs').argv,
    config = require('./@configuration.js'),
    istanbul = require("gulp-istanbul");


gulp.task("_istanbul:hook", ["build:testing"], () => {

    return gulp.src(config.testing.testingSrcDestGlob)
        .pipe(istanbul())
        .pipe(istanbul.hookRequire());
});

gulp.task("test", ["_istanbul:hook"], () => {

    // when using single, grab only that test.js file - otherwise use the entire test.js glob
    let path = yargs.single ? './testing/tests/{path}.test.js'.replace('{path}', yargs.single) : config.testing.testingTestsDestGlob;

    // determine if we show the full coverage table
    let reports = yargs["coverage-details"] ? ['text', 'text-summary'] : ['text-summary'];

    // easiest way for tests to have settings available
    global.settings = config.settings;

        return gulp.src(path)
        .pipe(mocha({ ui: 'bdd', reporter: 'dot', timeout: 10000 }))
        .pipe(istanbul.writeReports({
            reporters: reports
        }));
});