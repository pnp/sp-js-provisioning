import { HandlerBase } from "./handlerbase";
import { Web } from "@pnp/sp";
import { ProvisioningContext } from "../provisioningcontext";
import { IProvisioningConfig } from "../provisioningconfig";
/**
 * Describes the Site Fields Object Handler
 */
export declare class SiteFields extends HandlerBase {
    private context;
    private tokenHelper;
    /**
     * Creates a new instance of the ObjectSiteFields class
     */
    constructor(config: IProvisioningConfig);
    /**
     * Provisioning Client Side Pages
     *
     * @param {Web} web The web
     * @param {string[]} siteFields The site fields
     * @param {ProvisioningContext} context Provisioning context
     */
    ProvisionObjects(web: Web, siteFields: string[], context: ProvisioningContext): Promise<void>;
    /**
     * Provision a site field
     *
     * @param {Web} web The web
     * @param {IClientSidePage} clientSidePage Cient side page
     */
    private processSiteField(web, schemaXml);
}
