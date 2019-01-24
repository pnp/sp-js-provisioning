//******************************************************************************
//* watch.js
//* 
//******************************************************************************

var gulp = require("gulp"),
    gulpWatch = require("gulp-watch"),
    runSequence = require("run-sequence"),
    config = require('./@configuration.js');

gulp.task("watch", () => {
    gulpWatch(config.paths.sourceGlob).on("change", () => {
        runSequence("build");
    });
});