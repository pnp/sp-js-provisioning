import { HandlerBase } from "./handlerbase";
import { Web } from "@pnp/sp";
import { ProvisioningContext } from "../provisioningcontext";
import { IProvisioningConfig } from "../provisioningconfig";
import { IContentType } from '../schema';
/**
 * Describes the Content Types Object Handler
 */
export declare class ContentTypes extends HandlerBase {
    private jsomContext;
    private context;
    /**
     * Creates a new instance of the ObjectSiteFields class
     */
    constructor(config: IProvisioningConfig);
    /**
     * Provisioning Content Types
     *
     * @param {Web} web The web
     * @param {IContentType[]} contentTypes The content types
     * @param {ProvisioningContext} context Provisioning context
     */
    ProvisionObjects(web: Web, contentTypes: IContentType[], context: ProvisioningContext): Promise<void>;
    /**
     * Provision a content type
     *
     * @param {Web} web The web
     * @param {IContentType} contentType Content type
     */
    private processContentType(web, contentType);
    /**
     * Add a content type
     *
     * @param {Web} web The web
     * @param {IContentType} contentType Content type
     */
    private addContentType(web, contentType);
    /**
 * Adding content type field refs
 *
 * @param {IContentType} contentType Content type
 * @param {SP.ContentType} spContentType Content type
 */
    private processContentTypeFieldRefs(contentType, spContentType);
}
