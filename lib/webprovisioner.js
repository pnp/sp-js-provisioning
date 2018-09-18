"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const sp_pnp_js_1 = require("sp-pnp-js");
const exports_1 = require("./handlers/exports");
const provisioningcontext_1 = require("./provisioningcontext");
/**
 * Root class of Provisioning
 */
class WebProvisioner {
    /**
     * Creates a new instance of the Provisioner class
     *
     * @param {Web} web The Web instance to which we want to apply templates
     * @param {TypedHash<HandlerBase>} handlermap A set of handlers we want to apply. The keys of the map need to match the property names in the template
     */
    constructor(web, context = new provisioningcontext_1.ProvisioningContext(), handlerMap = exports_1.DefaultHandlerMap, handlerSort = exports_1.DefaultHandlerSort) {
        this.web = web;
        this.context = context;
        this.handlerMap = handlerMap;
        this.handlerSort = handlerSort;
    }
    /**
     * Applies the supplied template to the web used to create this Provisioner instance
     *
     * @param {Schema} template The template to apply
     * @param {Function} progressCallback Callback for progress updates
     */
    applyTemplate(template, progressCallback) {
        sp_pnp_js_1.Logger.write(`Beginning processing of web [${this.web.toUrl()}]`, sp_pnp_js_1.LogLevel.Info);
        let operations = Object.getOwnPropertyNames(template).sort((name1, name2) => {
            let sort1 = this.handlerSort.hasOwnProperty(name1) ? this.handlerSort[name1] : 99;
            let sort2 = this.handlerSort.hasOwnProperty(name2) ? this.handlerSort[name2] : 99;
            return sort1 - sort2;
        });
        return operations.reduce((chain, name) => {
            let handler = this.handlerMap[name];
            return chain.then(_ => {
                if (progressCallback) {
                    progressCallback(name);
                }
                return handler.ProvisionObjects(this.web, template[name], this.context);
            });
        }, Promise.resolve()).then(_ => {
            sp_pnp_js_1.Logger.write(`Done processing of web [${this.web.toUrl()}]`, sp_pnp_js_1.LogLevel.Info);
        });
    }
}
exports.WebProvisioner = WebProvisioner;
