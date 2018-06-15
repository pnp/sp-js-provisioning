import { HandlerBase } from "./handlerbase";
import { ICustomAction } from "../schema";
import { Web } from "sp-pnp-js";

/**
 * Describes the Custom Actions Object Handler
 */
export class CustomActions extends HandlerBase {
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
    public ProvisionObjects(web: Web, customActions: ICustomAction[]): Promise<void> {
        super.scope_started();

        return new Promise<void>((resolve, reject) => {

            web.userCustomActions.select("Title").getAs<{ Title: string }[]>().then(existingActions => {

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
