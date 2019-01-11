"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const logging_1 = require("@pnp/logging");
/**
 * Describes the Object Handler Base
 */
class HandlerBase {
    /**
     * Creates a new instance of the ObjectHandlerBase class
     *
     * @param {string} name Name
     * @param {IProvisioningConfig} config Config
     */
    constructor(name, config) {
        this.config = {};
        this.name = name;
        this.config = config;
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
        this.log_info("ProvisionObjects", "Code execution scope started");
    }
    /**
     * Writes to Logger when scope has stopped
     */
    scope_ended() {
        this.log_info("ProvisionObjects", "Code execution scope ended");
    }
    /**
     * Writes to Logger
     *
     * @param {string} scope Scope
     * @param {string} message Message
     */
    log_info(scope, message) {
        let prefix = (this.config.logging && this.config.logging.prefix) ? `(${this.config.logging.prefix}) ` : "";
        logging_1.Logger.write(`${prefix}${this.name}: ${scope}: ${message}`);
    }
}
exports.HandlerBase = HandlerBase;
