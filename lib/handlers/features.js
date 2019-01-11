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
/**
 * Describes the Features Object Handler
 */
class Features extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the ObjectFeatures class
     */
    constructor() {
        super("Features");
    }
    /**
     * Provisioning features
     *
     * @param {Web} web The web
     * @param {Array<IFeature>} features The features to provision
     */
    ProvisionObjects(web, features) {
        const _super = Object.create(null, {
            scope_started: { get: () => super.scope_started },
            scope_ended: { get: () => super.scope_ended }
        });
        return __awaiter(this, void 0, void 0, function* () {
            _super.scope_started.call(this);
            try {
                yield features.reduce((chain, feature) => {
                    if (feature.deactivate) {
                        return chain.then(() => web.features.remove(feature.id, feature.force));
                    }
                    else {
                        return chain.then(() => web.features.add(feature.id, feature.force));
                    }
                }, Promise.resolve({}));
                _super.scope_ended.call(this);
            }
            catch (err) {
                _super.scope_ended.call(this);
                throw err;
            }
        });
    }
}
exports.Features = Features;
