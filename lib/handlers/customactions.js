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
 * Describes the Custom Actions Object Handler
 */
class CustomActions extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the ObjectCustomActions class
     *
     * @param {IProvisioningConfig} config Provisioning config
     */
    constructor(config) {
        super("CustomActions", config);
    }
    /**
     * Provisioning Custom Actions
     *
     * @param {Web} web The web
     * @param {Array<ICustomAction>} customactions The Custom Actions to provision
     */
    ProvisionObjects(web, customActions) {
        const _super = Object.create(null, {
            scope_started: { get: () => super.scope_started },
            scope_ended: { get: () => super.scope_ended }
        });
        return __awaiter(this, void 0, void 0, function* () {
            _super.scope_started.call(this);
            try {
                const existingActions = yield web.userCustomActions.select("Title").get();
                let batch = web.createBatch();
                customActions
                    .filter(action => {
                    return !existingActions.some(existingAction => existingAction.Title === action.Title);
                })
                    .map(action => {
                    web.userCustomActions.inBatch(batch).add(action);
                });
                yield batch.execute();
                _super.scope_ended.call(this);
            }
            catch (err) {
                _super.scope_ended.call(this);
                throw err;
            }
        });
    }
}
exports.CustomActions = CustomActions;
