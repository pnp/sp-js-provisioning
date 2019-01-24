import { IClientSidePage } from "../schema";
import { HandlerBase } from "./handlerbase";
import { Web } from "@pnp/sp";
import { ProvisioningContext } from "../provisioningcontext";
import { IProvisioningConfig } from "../provisioningconfig";
/**
 * Describes the Composed Look Object Handler
 */
export declare class ClientSidePages extends HandlerBase {
    private tokenHelper;
    /**
     * Creates a new instance of the ObjectClientSidePages class
     */
    constructor(config: IProvisioningConfig);
    /**
     * Provisioning Client Side Pages
     *
     * @param {Web} web The web
     * @param {IClientSidePage[]} clientSidePages The client side pages to provision
     * @param {ProvisioningContext} context Provisioning context
     */
    ProvisionObjects(web: Web, clientSidePages: IClientSidePage[], context: ProvisioningContext): Promise<void>;
    /**
     * Provision a client side page
     *
     * @param {Web} web The web
     * @param {IClientSidePage} clientSidePage Cient side page
     */
    private processClientSidePage(web, clientSidePage);
}
