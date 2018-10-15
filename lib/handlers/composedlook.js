"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const handlerbase_1 = require("./handlerbase");
const util_1 = require("../util");
/**
 * Describes the Composed Look Object Handler
 */
class ComposedLook extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the ObjectComposedLook class
     */
    constructor() {
        super("ComposedLook");
    }
    /**
     * Provisioning Composed Look
     *
     * @param {Web} web The web
     * @param {IComposedLook} object The Composed Look to provision
     */
    ProvisionObjects(web, composedLook) {
        super.scope_started();
        return new Promise((resolve, reject) => {
            web.applyTheme(util_1.makeUrlRelative(util_1.replaceUrlTokens(composedLook.ColorPaletteUrl)), util_1.makeUrlRelative(util_1.replaceUrlTokens(composedLook.FontSchemeUrl)), composedLook.BackgroundImageUrl ? util_1.makeUrlRelative(util_1.replaceUrlTokens(composedLook.BackgroundImageUrl)) : null, false).then(_ => {
                super.scope_ended();
                resolve();
            }).catch(e => {
                super.scope_ended();
                reject(e);
            });
        });
    }
}
exports.ComposedLook = ComposedLook;
