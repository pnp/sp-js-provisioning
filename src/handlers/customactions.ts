import { HandlerBase } from "./handlerbase";
import { ICustomAction } from "../schema";
import { Web } from "@pnp/sp";

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
    public async ProvisionObjects(web: Web, customActions: ICustomAction[]): Promise<void> {
        super.scope_started();
        try {
            const existingActions = await web.userCustomActions.select("Title").get<{ Title: string }[]>();

            let batch = web.createBatch();

            customActions
                .filter(action => {
                    return !existingActions.some(existingAction => existingAction.Title === action.Title);
                })
                .map(action => {
                    web.userCustomActions.inBatch(batch).add(action);
                });

            await batch.execute();
            super.scope_ended();
        } catch (err) {
            super.scope_ended();
            throw err;
        }
    }
}

