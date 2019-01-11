import { Logger, ConsoleListener } from "@pnp/logging";
import { DefaultHandlerMap, DefaultHandlerSort } from "./handlers/exports";
import { ProvisioningContext } from "./provisioningcontext";
/**
 * Root class of Provisioning
 */
var WebProvisioner = (function () {
    /**
     * Creates a new instance of the Provisioner class
     *
     * @param {Web} web The Web instance to which we want to apply templates
     * @param {TypedHash<HandlerBase>} handlermap A set of handlers we want to apply. The keys of the map need to match the property names in the template
     */
    function WebProvisioner(web, config, context, handlerSort) {
        if (context === void 0) { context = new ProvisioningContext(); }
        if (handlerSort === void 0) { handlerSort = DefaultHandlerSort; }
        this.web = web;
        this.config = config;
        this.context = context;
        this.handlerSort = handlerSort;
        this.setup(config);
        this.handlerMap = DefaultHandlerMap(this.config);
    }
    /**
     * Applies the supplied template to the web used to create this Provisioner instance
     *
     * @param {Schema} template The template to apply
     * @param {Function} progressCallback Callback for progress updates
     */
    WebProvisioner.prototype.applyTemplate = function (template, progressCallback) {
        var _this = this;
        Logger.write("Beginning processing of web [" + this.web.toUrl() + "]", 1 /* Info */);
        var operations = Object.getOwnPropertyNames(template).sort(function (name1, name2) {
            var sort1 = _this.handlerSort.hasOwnProperty(name1) ? _this.handlerSort[name1] : 99;
            var sort2 = _this.handlerSort.hasOwnProperty(name2) ? _this.handlerSort[name2] : 99;
            return sort1 - sort2;
        });
        return operations.reduce(function (chain, name) {
            var handler = _this.handlerMap[name];
            return chain.then(function () {
                if (progressCallback) {
                    progressCallback(name);
                }
                return handler.ProvisionObjects(_this.web, template[name], _this.context);
            });
        }, Promise.resolve())
            .then(function (_) {
            Logger.write("Done processing of web [" + _this.web.toUrl() + "]", 1 /* Info */);
        })
            .catch(function (_) {
            Logger.write("Processing of web [" + _this.web.toUrl() + "] failed", 3 /* Error */);
        });
    };
    /**
    * Sets up the web provisioner
    *
    * @param {IProvisioningConfig} config Provisioning config
    */
    WebProvisioner.prototype.setup = function (config) {
        this.config = config;
        if (this.config && this.config.logging) {
            Logger.subscribe(new ConsoleListener());
            Logger.activeLogLevel = this.config.logging.activeLogLevel;
        }
    };
    return WebProvisioner;
}());
export { WebProvisioner };
