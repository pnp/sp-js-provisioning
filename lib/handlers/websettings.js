"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const handlerbase_1 = require("./handlerbase");
const omit = require("object.omit");
const util_1 = require("../util");
/**
 * Describes the WebSettings Object Handler
 */
class WebSettings extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the WebSettings class
     */
    constructor() {
        super("WebSettings");
    }
    /**
     * Provisioning WebSettings
     *
     * @param {Web} web The web
     * @param {IWebSettings} settings The settings
     */
    ProvisionObjects(web, settings) {
        super.scope_started();
        return new Promise((resolve, reject) => {
            Object.keys(settings)
                .filter(key => typeof (settings[key]) === "string")
                .forEach(key => {
                let value = settings[key];
                settings[key] = util_1.replaceUrlTokens(value);
            });
            Promise.all([
                web.rootFolder.update({ WelcomePage: settings.WelcomePage }),
                web.update(omit(settings, "WelcomePage")),
            ]).then(_ => {
                super.scope_ended();
                resolve();
            }).catch(e => reject(e));
        });
    }
}
exports.WebSettings = WebSettings;
