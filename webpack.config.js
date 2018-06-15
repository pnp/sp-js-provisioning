var path = require("path"),
    webpack = require('webpack'),
    config = require('./gulptasks/@configuration.js');

module.exports = [{
    // core file + sourcemap -> dist
    cache: true,
    entry: './lib/webprovisioner.js',
    output: {
        path: path.join(__dirname, "dist"),
        publicPath: "/assets/",
        filename: "pnp-provisioning.js",
        libraryTarget: "umd",
        library: "$pnp"
    },
    devtool: "source-map",
    resolve: {
        extensions: ['.js']
    },
    plugins: [
        new webpack.BannerPlugin(config.header)
    ],
    module: {
        loaders: [
            { test: /\.js$/, loader: 'babel-loader?presets[]=es2015' }
        ]
    }
},
{
    // minified core file + sourcemap -> dist
    cache: true,
    entry: './lib/webprovisioner.js',
    output: {
        path: path.join(__dirname, "dist"),
        publicPath: "/assets/",
        filename: "pnp-provisioning.min.js",
        libraryTarget: "umd",
        library: "$pnp"
    },
    devtool: "source-map",
    resolve: {
        extensions: ['.js']
    },
    plugins: [
        new webpack.BannerPlugin(config.header),
        new webpack.DefinePlugin({
            "process.env": {
                "NODE_ENV": JSON.stringify("production")
            }
        }),
        new webpack.optimize.UglifyJsPlugin()
    ],
    module: {
        loaders: [
            { test: /\.js$/, loader: 'babel-loader?presets[]=es2015' },
        ]
    }
}];
