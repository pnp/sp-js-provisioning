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
     */
    HandlerBase.prototype.log_info = function (scope, message) {
        var prefix = (this.config.logging && this.config.logging.prefix) ? "(" + this.config.logging.prefix + ") " : "";
        Logger.write("" + prefix + this.name + ": " + scope + ": " + message);
    };
    return HandlerBase;
}());
export { HandlerBase };
