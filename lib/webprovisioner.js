var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = y[op[0] & 2 ? "return" : op[0] ? "throw" : "next"]) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [0, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
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
    }
    /**
     * Applies the supplied template to the web used to create this Provisioner instance
     *
     * @param {Schema} template The template to apply
     * @param {Function} progressCallback Callback for progress updates
     */
    WebProvisioner.prototype.applyTemplate = function (template, progressCallback) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            var operations, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        Logger.write("Beginning processing of web [" + this.web.toUrl() + "]", 1 /* Info */);
                        operations = Object.getOwnPropertyNames(template).sort(function (name1, name2) {
                            var sort1 = _this.handlerSort.hasOwnProperty(name1) ? _this.handlerSort[name1] : 99;
                            var sort2 = _this.handlerSort.hasOwnProperty(name2) ? _this.handlerSort[name2] : 99;
                            return sort1 - sort2;
                        });
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, operations.reduce(function (chain, name) {
                                var handler = _this.handlerMap[name];
                                return chain.then(function () {
                                    if (progressCallback) {
                                        progressCallback(name);
                                    }
                                    return handler.ProvisionObjects(_this.web, template[name], _this.context);
                                });
                            }, Promise.resolve())];
                    case 2:
                        _a.sent();
                        Logger.write("Done processing of web [" + this.web.toUrl() + "]", 1 /* Info */);
                        return [3 /*break*/, 4];
                    case 3:
                        error_1 = _a.sent();
                        Logger.write("Processing of web [" + this.web.toUrl() + "] failed", 3 /* Error */);
                        throw error_1;
                    case 4: return [2 /*return*/];
                }
            });
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
        this.handlerMap = DefaultHandlerMap(this.config);
    };
    return WebProvisioner;
}());
export { WebProvisioner };
