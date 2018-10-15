"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const logging_1 = require("@pnp/logging");
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
        logging_1.Logger.log({ data: templatePart, level: 2 /* Warning */, message: `Handler ${this.name} for web [${web.toUrl()}] does not override ProvisionObjects.` });
        return Promise.resolve();
    }
    /**
     * Writes to Logger when scope has started
     */
    scope_started() {
        logging_1.Logger.write(`${this.name}: Code execution scope started`);
    }
    /**
     * Writes to Logger when scope has stopped
     */
    scope_ended() {
        logging_1.Logger.write(`${this.name}: Code execution scope stopped`);
    }
}
exports.HandlerBase = HandlerBase;
