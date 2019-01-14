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
            this.context.contentTypes = (await web.contentTypes.select('Id', 'Name').get<Array<{ Id: { StringValue: string }, Name: string }>>()).reduce((obj, l) => {
                obj[l.Name] = l.Id.StringValue;
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
            super.log_info("processContentType", `Processing content type ${contentType.Name}`);
            let contentTypeId = this.context.contentTypes[contentType.Name];
            if (!contentTypeId) {
                const contentTypeAddResult = await this.addContentType(web, contentType);
                contentTypeId = contentTypeAddResult.data.Id;
            }
            if (contentType.FieldRefs) {
                await this.processContentTypeFieldRefs(contentType, contentTypeId);
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
            super.log_info("addContentType", `Adding content type ${contentType.Name} (${contentType.ID})`);
            return await web.contentTypes.add(contentType.ID, contentType.Name, contentType.Description, contentType.Group);
        } catch (err) {
            throw err;
        }
    }

    /**
     * Adding content type field refs
     *
     * @param {IContentType} contentType Content type
     * @param {string} contentTypeId Content type id
     */
    private async processContentTypeFieldRefs(contentType: IContentType, contentTypeId: string): Promise<void> {
        try {
            const _contentType = this.jsomContext.web.get_contentTypes().getById(contentTypeId);
            _contentType.set_description(contentType.Description);
            _contentType.set_group(contentType.Group);
            const fieldRefNames = contentType.FieldRefs.map(fieldRef => fieldRef.Name);
            super.log_info("processContentTypeFieldRefs", `Processing field refs for content type ${contentType.Name} (${contentType.ID})`, { fieldRefs: fieldRefNames });
            const fieldLinks: SP.FieldLink[] = fieldRefNames.map((fieldRefName, i) => {
                const siteField = this.jsomContext.web.get_fields().getByInternalNameOrTitle(fieldRefName);
                const fieldLinkCreationInformation = new SP.FieldLinkCreationInformation();
                fieldLinkCreationInformation.set_field(siteField);
                const fieldLink = _contentType.get_fieldLinks().add(fieldLinkCreationInformation);
                if(contentType.FieldRefs[i].hasOwnProperty("Required")) {
                    fieldLink.set_required(contentType.FieldRefs[i].Required);
                }
                if(contentType.FieldRefs[i].hasOwnProperty("Hidden")) {
                    fieldLink.set_hidden(contentType.FieldRefs[i].Hidden);                    
                }
                return fieldLink;
            });
            _contentType.update(false);
            await ExecuteJsomQuery(this.jsomContext, fieldLinks.map(fieldLink => ({ clientObject: fieldLink })));
        } catch ({ sender, args }) {
            super.log_info("processContentTypeFieldRefs", `Failed to process field refs for content type ${contentType.Name} (${contentType.ID})`, { error: args.get_message() });
        }
    }
}
