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
exports.DefaultHandlerMap = {
    ComposedLook: new composedlook_1.ComposedLook(),
    CustomActions: new customactions_1.CustomActions(),
    Features: new features_1.Features(),
    Files: new files_1.Files(),
    Lists: new lists_1.Lists(),
    Navigation: new navigation_1.Navigation(),
    PropertyBagEntries: new propertybagentries_1.PropertyBagEntries(),
    WebSettings: new websettings_1.WebSettings(),
};
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
