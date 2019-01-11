"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const composedlook_1 = require("./composedlook");
const customactions_1 = require("./customactions");
const features_1 = require("./features");
const websettings_1 = require("./websettings");
const navigation_1 = require("./navigation");
const lists_1 = require("./lists");
const files_1 = require("./files");
const propertybagentries_1 = require("./propertybagentries");
exports.DefaultHandlerMap = (config) => ({
    ComposedLook: new composedlook_1.ComposedLook(config),
    CustomActions: new customactions_1.CustomActions(config),
    Features: new features_1.Features(config),
    Files: new files_1.Files(config),
    Lists: new lists_1.Lists(config),
    Navigation: new navigation_1.Navigation(config),
    PropertyBagEntries: new propertybagentries_1.PropertyBagEntries(config),
    WebSettings: new websettings_1.WebSettings(config),
});
exports.DefaultHandlerSort = {
    ComposedLook: 6,
    CustomActions: 5,
    Features: 2,
    Files: 4,
    Lists: 3,
    Navigation: 7,
    PropertyBagEntries: 8,
    WebSettings: 1,
};
