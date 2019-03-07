import { Logger } from "@pnp/logging";
/**
 * Describes the Object Handler Base
 */
var HandlerBase = (function () {
    /**
     * Creates a new instance of the ObjectHandlerBase class
     *
     * @param {string} name Name
     * @param {IProvisioningConfig} config Config
     */
    function HandlerBase(name, config) {
        if (config === void 0) { config = {}; }
        this.config = {};
        this.name = name;
        this.config = config;
    }
    /**
     * Provisioning objects
     */
    HandlerBase.prototype.ProvisionObjects = function (web, templatePart, _context) {
        Logger.log({ data: templatePart, level: 2 /* Warning */, message: "Handler " + this.name + " for web [" + web.toUrl() + "] does not override ProvisionObjects." });
        return Promise.resolve();
    };
    /**
     * Writes to Logger when scope has started
     */
    HandlerBase.prototype.scope_started = function () {
        this.log_info("ProvisionObjects", "Code execution scope started");
    };
    /**
     * Writes to Logger when scope has stopped
     */
    HandlerBase.prototype.scope_ended = function () {
        this.log_info("ProvisionObjects", "Code execution scope ended");
    };
    /**
     * Writes to Logger
     *
     * @param {string} scope Scope
     * @param {string} message Message
     * @param {any} data Data
     */
    HandlerBase.prototype.log_info = function (scope, message, data) {
        var prefix = (this.config.logging && this.config.logging.prefix) ? this.config.logging.prefix + " " : "";
        Logger.log({ message: prefix + "(" + this.name + "): " + scope + ": " + message, data: data, level: 1 /* Info */ });
    };
    /**
     * Writes to Logger
     *
     * @param {string} scope Scope
     * @param {string} message Message
     * @param {any} data Data
     */
    HandlerBase.prototype.log_warn = function (scope, message, data) {
        var prefix = (this.config.logging && this.config.logging.prefix) ? this.config.logging.prefix + " " : "";
        Logger.log({ message: prefix + "(" + this.name + "): " + scope + ": " + message, data: data, level: 2 /* Warning */ });
    };
    return HandlerBase;
}());
export { HandlerBase };
