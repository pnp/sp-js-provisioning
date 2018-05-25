//******************************************************************************
//* build.js
//* 
//* Defines a custom gulp task for compiling TypeScript source code into
//* js files.  It outputs the details as to what it generated to the console.
//******************************************************************************

var gulp = require("gulp"),
    tsc = require("gulp-typescript"),
    config = require('./@configuration.js'),
    merge = require("merge2"),
    sourcemaps = require('gulp-sourcemaps'),
    pkg = require("../package.json");

gulp.task("build:lib", () => {

    var project = tsc.createProject("tsconfig.json", { declaration: true });

    var built = gulp.src(config.paths.sourceGlob)
        .pipe(project());

    return merge([
        built.dts.pipe(gulp.dest(config.paths.lib)),
        built.js.pipe(gulp.dest(config.paths.lib))
    ]);
});

gulp.task("build:testing", ["clean"], () => {

    var projectSrc = tsc.createProject("tsconfig.json");
    var projectTests = tsc.createProject("tsconfig.json");

    return merge([
        gulp.src(config.testing.testsSourceGlob)
            .pipe(projectTests({
                compilerOptions: {
                    types: [
                        "chai",
                        "chai-as-promised",
                        "node",
                        "mocha"
                    ]
                }
            }))
            .pipe(gulp.dest(config.testing.testingTestsDest)),
        gulp.src(config.paths.sourceGlob)
            .pipe(projectSrc())
            .pipe(gulp.dest(config.testing.testingSrcDest))
    ]);
});

gulp.task("build:debug", ["clean"], () => {

    var srcProject = tsc.createProject("tsconfig.json");
    var debugProject = tsc.createProject("tsconfig.json");
    var schemaProject = tsc.createProject("tsconfig.json");

    let sourceMapSettings = {
        includeContent: false,
        sourceRoot: (file) => {
            return "..\\.." + file.base.replace(file.cwd, "");
        }
    };

    return merge([
        gulp.src(config.paths.sourceGlob)
            .pipe(sourcemaps.init())
            .pipe(srcProject())
            .pipe(sourcemaps.write(".", sourceMapSettings))
            .pipe(gulp.dest(config.debug.outputSrc)),
        gulp.src(config.debug.debugSourceGlob)
            .pipe(sourcemaps.init())
            .pipe(debugProject())
            .pipe(sourcemaps.write(".", sourceMapSettings))
            .pipe(gulp.dest(config.debug.outputDebug)),
        gulp.src(config.debug.schemasSourceGlob)
            .pipe(sourcemaps.init())
            .pipe(schemaProject())
            .pipe(sourcemaps.write(".", sourceMapSettings))
            .pipe(gulp.dest(config.debug.schemasOutput))
    ]);
});

// run the build chain for lib
gulp.task("build", ["clean", "lint", "build:lib"]);
