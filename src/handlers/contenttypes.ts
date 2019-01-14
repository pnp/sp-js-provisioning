import initSpfxJsom, { ExecuteJsomQuery, JsomContext } from "spfx-jsom";
import { HandlerBase } from "./handlerbase";
import { Web, ContentTypeAddResult } from "@pnp/sp";
import { ProvisioningContext } from "../provisioningcontext";
import { IProvisioningConfig } from "../provisioningconfig";
import { TokenHelper } from '../util/tokenhelper';
import { IContentType } from '../schema';

/**
 * Describes the Content Types Object Handler
 */
export class ContentTypes extends HandlerBase {
    private jsomContext: JsomContext;
    private context: ProvisioningContext;

    /**
     * Creates a new instance of the ObjectSiteFields class
     */
    constructor(config: IProvisioningConfig) {
        super("ContentTypes", config);
    }

    /**
     * Provisioning Content Types
     *
     * @param {Web} web The web
     * @param {IContentType[]} contentTypes The content types
     * @param {ProvisioningContext} context Provisioning context
     */
    public async ProvisionObjects(web: Web, contentTypes: IContentType[], context: ProvisioningContext): Promise<void> {
        this.jsomContext = await initSpfxJsom(context.web.ServerRelativeUrl);
        this.context = context;
        super.scope_started();
        try {
            this.context.contentTypes = (await web.contentTypes.select('Id', 'Name', 'FieldLinks').expand('FieldLinks').get<Array<any>>()).reduce((obj, contentType) => {
                obj[contentType.Name] = {
                    ID: contentType.Id.StringValue,
                    Name: contentType.Name,
                    FieldRefs: contentType.FieldLinks.map(fl => ({
                        ID: fl.Id,
                        Name: fl.Name,
                        Required: fl.Required,
                        Hidden: fl.Hidden,
                    })),
                };
                return obj;
            }, {});
            await contentTypes
                .sort((a, b) => {
                    if (a.ID < b.ID) { return -1; }
                    if (a.ID > b.ID) { return 1; }
                    return 0;
                })
                .reduce((chain: any, contentType) => chain.then(() => this.processContentType(web, contentType)), Promise.resolve());
        } catch (err) {
            super.scope_ended();
            throw err;
        }
    }

    /**
     * Provision a content type
     *
     * @param {Web} web The web
     * @param {IContentType} contentType Content type
     */
    private async processContentType(web: Web, contentType: IContentType): Promise<void> {
        try {
            let contentTypeId = this.context.contentTypes[contentType.Name].ID;
            if (!contentTypeId) {
                const contentTypeAddResult = await this.addContentType(web, contentType);
                contentTypeId = contentTypeAddResult.data.Id;
            }
            super.log_info("processContentType", `Processing content type [${contentType.Name}] (${contentTypeId})`);
            const spContentType = this.jsomContext.web.get_contentTypes().getById(contentTypeId);
            if (contentType.Description) {
                spContentType.set_description(contentType.Description);
            }
            if (contentType.Group) {
                spContentType.set_group(contentType.Group);
            }
            spContentType.update(true);
            await ExecuteJsomQuery(this.jsomContext);
            if (contentType.FieldRefs) {
                await this.processContentTypeFieldRefs(contentType, spContentType);
            }
        } catch (err) {
            throw err;
        }
    }

    /**
     * Add a content type
     *
     * @param {Web} web The web
     * @param {IContentType} contentType Content type
     */
    private async addContentType(web: Web, contentType: IContentType): Promise<ContentTypeAddResult> {
        try {
            super.log_info("addContentType", `Adding content type [${contentType.Name}] (${contentType.ID})`);
            return await web.contentTypes.add(contentType.ID, contentType.Name, contentType.Description, contentType.Group);
        } catch (err) {
            throw err;
        }
    }

    /**
 * Adding content type field refs
 *
 * @param {IContentType} contentType Content type
 * @param {SP.ContentType} spContentType Content type
 */
    private async processContentTypeFieldRefs(contentType: IContentType, spContentType: SP.ContentType): Promise<void> {
        try {
            for (let i = 0; i < contentType.FieldRefs.length; i++) {
                const fieldRef = contentType.FieldRefs[i];
                let [existingFieldLink] = this.context.contentTypes[contentType.Name].FieldRefs.filter(fr => fr.Name === fieldRef.Name);
                let fieldLink: SP.FieldLink;
                if (existingFieldLink) {
                    fieldLink = spContentType.get_fieldLinks().getById(new SP.Guid(existingFieldLink.ID));
                } else {
                    super.log_info("processContentTypeFieldRefs", `Adding field ref ${fieldRef.Name} to content type [${contentType.Name}] (${contentType.ID})`);
                    const siteField = this.jsomContext.web.get_fields().getByInternalNameOrTitle(fieldRef.Name);
                    const fieldLinkCreationInformation = new SP.FieldLinkCreationInformation();
                    fieldLinkCreationInformation.set_field(siteField);
                    fieldLink = spContentType.get_fieldLinks().add(fieldLinkCreationInformation);
                }
                if (contentType.FieldRefs[i].hasOwnProperty("Required")) {
                    fieldLink.set_required(contentType.FieldRefs[i].Required);
                }
                if (contentType.FieldRefs[i].hasOwnProperty("Hidden")) {
                    fieldLink.set_hidden(contentType.FieldRefs[i].Hidden);
                }
            }
            spContentType.update(true);
            await ExecuteJsomQuery(this.jsomContext);
            super.log_info("processContentTypeFieldRefs", `Successfully processed field refs for content type [${contentType.Name}] (${contentType.ID})`);
        } catch (error) {
            super.log_info("processContentTypeFieldRefs", `Failed to process field refs for content type [${contentType.Name}] (${contentType.ID})`, { error: error.args && error.args.get_message() });
        }
    }
}
