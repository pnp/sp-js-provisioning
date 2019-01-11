import * as xmljs from 'xml-js';
import { HandlerBase } from "./handlerbase";
import { Web, FieldAddResult } from "@pnp/sp";
import { ProvisioningContext } from "../provisioningcontext";
import { IProvisioningConfig } from "../provisioningconfig";
import { TokenHelper } from '../util/tokenhelper';

/**
 * Describes the Site Fields Object Handler
 */
export class SiteFields extends HandlerBase {
    private tokenHelper: TokenHelper;

    /**
     * Creates a new instance of the ObjectSiteFields class
     */
    constructor(config: IProvisioningConfig) {
        super("SiteFields", config);
    }

    /**
     * Provisioning Client Side Pages
     *
     * @param {Web} web The web
     * @param {string[]} siteFields The site fields
     * @param {ProvisioningContext} context Provisioning context
     */
    public async ProvisionObjects(web: Web, siteFields: string[], context: ProvisioningContext): Promise<void> {
        this.tokenHelper = new TokenHelper(context, this.config);
        super.scope_started();
        try {
            await siteFields.reduce((chain: any, schemaXml) => chain.then(() => this.processSiteField(web, schemaXml)), Promise.resolve());
        } catch (err) {
            super.scope_ended();
            throw err;
        }
    }

    /**
     * Provision a site field
     *
     * @param {Web} web The web
     * @param {IClientSidePage} clientSidePage Cient side page
     */
    private async processSiteField(web: Web, schemaXml: string): Promise<FieldAddResult> {
        try {
            schemaXml = this.tokenHelper.replaceTokens(schemaXml);
            const schemaXmlJson = JSON.parse(xmljs.xml2json(schemaXml));
            const fieldAttributes = schemaXmlJson.elements[0].attributes;
            super.log_info("processSiteField", `Processing site field ${fieldAttributes.DisplayName}`, { schemaXmlJson, schemaXml });
            return await web.fields.createFieldAsXml(schemaXml);
        } catch (err) {
            throw err;
        }
    }
}
