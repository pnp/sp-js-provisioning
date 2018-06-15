"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const handlerbase_1 = require("./handlerbase");
/**
 * Describes the Custom Actions Object Handler
 */
class CustomActions extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the ObjectCustomActions class
     */
    constructor() {
        super("CustomActions");
    }
    /**
     * Provisioning Custom Actions
     *
     * @param {Web} web The web
     * @param {Array<ICustomAction>} customactions The Custom Actions to provision
     */
    ProvisionObjects(web, customActions) {
        super.scope_started();
        return new Promise((resolve, reject) => {
            web.userCustomActions.select("Title").getAs().then(existingActions => {
                let batch = web.createBatch();
                customActions.filter(action => {
                    return !existingActions.some(existingAction => existingAction.Title === action.Title);
                }).map(action => {
                    web.userCustomActions.inBatch(batch).add(action);
                });
                batch.execute().then(_ => {
                    super.scope_ended();
                    resolve();
                });
            }).catch(e => {
                super.scope_ended();
                reject(e);
            });
        });
    }
}
exports.CustomActions = CustomActions;
