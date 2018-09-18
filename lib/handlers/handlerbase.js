"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const sp_pnp_js_1 = require("sp-pnp-js");
/**
 * Describes the Object Handler Base
 */
class HandlerBase {
    /**
     * Creates a new instance of the ObjectHandlerBase class
     */
    constructor(name) {
        this.name = name;
    }
    /**
     * Provisioning objects
     */
    ProvisionObjects(web, templatePart, _context) {
        sp_pnp_js_1.Logger.log({ data: templatePart, level: sp_pnp_js_1.LogLevel.Warning, message: `Handler ${this.name} for web [${web.toUrl()}] does not override ProvisionObjects.` });
        return Promise.resolve();
    }
    /**
     * Writes to Logger when scope has started
     */
    scope_started() {
        sp_pnp_js_1.Logger.write(`${this.name}: Code execution scope started`);
    }
    /**
     * Writes to Logger when scope has stopped
     */
    scope_ended() {
        sp_pnp_js_1.Logger.write(`${this.name}: Code execution scope stopped`);
    }
}
exports.HandlerBase = HandlerBase;
