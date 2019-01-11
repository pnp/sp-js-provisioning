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
        const _super = Object.create(null, {
            scope_started: { get: () => super.scope_started },
            scope_ended: { get: () => super.scope_ended }
        });
        return __awaiter(this, void 0, void 0, function* () {
            _super.scope_started.call(this);
            try {
                yield web.applyTheme(util_1.makeUrlRelative(util_1.replaceUrlTokens(composedLook.ColorPaletteUrl)), util_1.makeUrlRelative(util_1.replaceUrlTokens(composedLook.FontSchemeUrl)), composedLook.BackgroundImageUrl ? util_1.makeUrlRelative(util_1.replaceUrlTokens(composedLook.BackgroundImageUrl)) : null, false);
                _super.scope_ended.call(this);
            }
            catch (err) {
                _super.scope_ended.call(this);
                throw err;
            }
        });
    }
}
exports.ComposedLook = ComposedLook;
