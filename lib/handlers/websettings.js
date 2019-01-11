"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
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
        const _super = Object.create(null, {
            scope_started: { get: () => super.scope_started },
            scope_ended: { get: () => super.scope_ended }
        });
        return __awaiter(this, void 0, void 0, function* () {
            _super.scope_started.call(this);
            Object.keys(settings)
                .filter(key => typeof (settings[key]) === "string")
                .forEach(key => {
                let value = settings[key];
                settings[key] = util_1.replaceUrlTokens(value);
            });
            try {
                yield Promise.all([
                    web.rootFolder.update({ WelcomePage: settings.WelcomePage }),
                    web.update(omit(settings, "WelcomePage")),
                ]);
                _super.scope_ended.call(this);
            }
            catch (err) {
                _super.scope_ended.call(this);
                throw err;
            }
        });
    }
}
exports.WebSettings = WebSettings;
