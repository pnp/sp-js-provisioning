import { List, Web } from '@pnp/sp';
import initSpfxJsom, { ExecuteJsomQuery, JsomContext } from "spfx-jsom";
import * as xmljs from 'xml-js';
import { IProvisioningConfig } from '../provisioningconfig';
import { ProvisioningContext } from '../provisioningcontext';
import { IContentTypeBinding, IList, IListInstanceFieldRef, IListView } from '../schema';
import { TokenHelper } from '../util/tokenhelper';
import { HandlerBase } from './handlerbase';

export interface ISPField {
    Id: string;
    InternalName: string;
    SchemaXml: string;
}

/**
 * Describes the Lists Object Handler
 */
export class Lists extends HandlerBase {
    public tokenHelper: TokenHelper;
    public jsomContext: JsomContext;
    public context: ProvisioningContext;

    /**
     * Creates a new instance of the Lists class
     *
     * @param {IProvisioningConfig} config Provisioning config
     */
    constructor(config: IProvisioningConfig) {
        super('Lists', config);
    }

    /**
     * Provisioning lists
     *
     * @param {Web} web The web
     * @param {Array<IList>} lists The lists to provision
     */
    public async ProvisionObjects(web: Web, lists: IList[], context: ProvisioningContext): Promise<void> {
        this.jsomContext = (await initSpfxJsom(context.web.ServerRelativeUrl)).jsomContext;
        this.context = context;
        this.tokenHelper = new TokenHelper(this.context, this.config);
        super.scope_started();
        try {
            this.context.lists = (await web.lists.select('Id', 'Title').get<Array<{ Id: String, Title: string }>>()).reduce((obj, l) => {
                obj[l.Title] = l.Id;
                return obj;
            }, {});
            await lists.reduce((chain: any, list) => chain.then(() => this.processList(web, list)), Promise.resolve());
            await lists.reduce((chain: any, list) => chain.then(() => this.processListFields(web, list)), Promise.resolve());
            await lists.reduce((chain: any, list) => chain.then(() => this.processListFieldRefs(web, list)), Promise.resolve());
            await lists.reduce((chain: any, list) => chain.then(() => this.processListViews(web, list)), Promise.resolve());
            this.context.lists = (await web.lists.select('Id', 'Title').get<Array<{ Id: String, Title: string }>>()).reduce((obj, l) => {
                obj[l.Title] = l.Id;
                return obj;
            }, {});
            super.scope_ended();
        } catch (err) {
            super.scope_ended();
            throw err;
        }
    }

    /**
     * Processes a list
     *
     * @param {Web} web The web
     * @param {IList} lc The list
     */
    private async processList(web: Web, lc: IList): Promise<void> {
        super.log_info('processList', `Processing list ${lc.Title}`);
        let list: List;
        if (this.context.lists[lc.Title]) {
            super.log_info('processList', `List ${lc.Title} already exists. Ensuring...`);
            const listEnsure = await web.lists.ensure(lc.Title, lc.Description, lc.Template, lc.ContentTypesEnabled, lc.AdditionalSettings);
            list = listEnsure.list;
        } else {
            super.log_info('processList', `List ${lc.Title} doesn't exist. Creating...`);
            let listAdd = await web.lists.add(lc.Title, lc.Description, lc.Template, lc.ContentTypesEnabled, lc.AdditionalSettings);
            list = listAdd.list;
            this.context.lists[listAdd.data.Title] = listAdd.data.Id;
        }
        if (lc.ContentTypeBindings) {
            await this.processContentTypeBindings(lc, list, lc.ContentTypeBindings, lc.RemoveExistingContentTypes);
        }
    }

    /**
     * Processes content type bindings for a list
     *
     * @param {IList} lc The list configuration
     * @param {List} list The pnp list
     * @param {Array<IContentTypeBinding>} contentTypeBindings Content type bindings
     * @param {boolean} removeExisting Remove existing content type bindings
     */
    private async processContentTypeBindings(lc: IList, list: List, contentTypeBindings: IContentTypeBinding[], removeExisting: boolean): Promise<any> {
        super.log_info('processContentTypeBindings', `Processing content types for list ${lc.Title}.`);
        await contentTypeBindings.reduce((chain, ct) => chain.then(() => this.processContentTypeBinding(lc, list, ct.ContentTypeID)), Promise.resolve());
        if (removeExisting) {
            let promises = [];
            const contentTypes = await list.contentTypes.get();
            contentTypes.forEach(({ Id: { StringValue: ContentTypeId } }) => {
                let shouldRemove = (contentTypeBindings.filter(ctb => ContentTypeId.indexOf(ctb.ContentTypeID) !== -1).length === 0)
                    && (ContentTypeId.indexOf('0x0120') === -1);
                if (shouldRemove) {
                    super.log_info('processContentTypeBindings', `Removing content type ${ContentTypeId} from list ${lc.Title}`);
                    promises.push(list.contentTypes.getById(ContentTypeId).delete());
                } else {
                    super.log_info('processContentTypeBindings', `Skipping removal of content type ${ContentTypeId} from list ${lc.Title}`);
                }
            });
            await Promise.all(promises);
        }
    }

    /**
     * Processes a content type binding for a list
     *
     * @param {IList} lc The list configuration
     * @param {List} list The pnp list
     * @param {string} contentTypeID The Content Type ID
     */
    private async processContentTypeBinding(lc: IList, list: List, contentTypeID: string): Promise<any> {
        try {
            super.log_info('processContentTypeBinding', `Adding content Type ${contentTypeID} to list ${lc.Title}.`);
            await list.contentTypes.addAvailableContentType(contentTypeID);
            super.log_info('processContentTypeBinding', `Content Type ${contentTypeID} added successfully to list ${lc.Title}.`);
        } catch (err) {
            super.log_info('processContentTypeBinding', `Failed to add Content Type ${contentTypeID} to list ${lc.Title}.`);
        }
    }


    /**
     * Processes fields for a list
     *
     * @param {Web} web The web
     * @param {IList} list The pnp list
     */
    private async processListFields(web: Web, list: IList): Promise<any> {
        if (list.Fields) {
            await list.Fields.reduce((chain, field) => chain.then(() => this.processField(web, list, field)), Promise.resolve());
        }
    }

    /**
     * Processes a field for a lit
     *
     * @param {Web} web The web
     * @param {IList} lc The list configuration
     * @param {string} fieldXml Field xml
     */
    private async processField(web: Web, lc: IList, fieldXml: string): Promise<any> {
        const list = web.lists.getByTitle(lc.Title);
        const fXmlJson = JSON.parse(xmljs.xml2json(fieldXml));
        const fieldAttr = fXmlJson.elements[0].attributes;
        const fieldName = fieldAttr.Name;
        const fieldDisplayName = fieldAttr.DisplayName;

        super.log_info('processField', `Processing field ${fieldName} (${fieldDisplayName}) for list ${lc.Title}.`);
        fXmlJson.elements[0].attributes.DisplayName = fieldName;
        fieldXml = xmljs.json2xml(fXmlJson);

        // Looks like e.g. lookup fields can't be updated, so we'll need to re-create the field
        try {
            await list.fields.getById(fieldAttr.ID).delete();
            super.log_info('processField', `Field ${fieldName} (${fieldDisplayName}) successfully deleted from list ${lc.Title}.`);
        } catch (err) {
            super.log_info('processField', `Field ${fieldName} (${fieldDisplayName}) does not exist in list ${lc.Title}.`);
        }

        // Looks like e.g. lookup fields can't be updated, so we'll need to re-create the field
        try {
            let fieldAddResult = await list.fields.createFieldAsXml(this.tokenHelper.replaceTokens(fieldXml));
            await fieldAddResult.field.update({ Title: fieldDisplayName });
            super.log_info('processField', `Field '${fieldDisplayName}' added successfully to list ${lc.Title}.`);
        } catch (err) {
            super.log_info('processField', `Failed to add field '${fieldDisplayName}' to list ${lc.Title}.`);
        }
    }

    /**
   * Processes field refs for a list
   *
   * @param {Web} web The web
   * @param {IList} lc The list configuration
   */
    private async processListFieldRefs(web: Web, lc: IList): Promise<any> {
        if (lc.FieldRefs) {
            super.log_info('processListFieldRefs', `Retrieving fields for list ${lc.Title} and web.`);
            let list = web.lists.getByTitle(lc.Title);
            let [listFields, webFields] = await Promise.all([
                list.fields.select('Id', 'InternalName', 'SchemaXml').usingCaching().get<ISPField[]>(),
                web.fields.select('Id', 'InternalName', 'SchemaXml').usingCaching().get<ISPField[]>(),
            ]);
            super.log_info('processListFieldRefs', `Fields for list ${lc.Title} and web retrieved. Processing field refs.`);
            await
                await lc.FieldRefs.reduce((chain: any, fieldRef) => {
                    return chain.then(() => this.processFieldRef(list, lc, fieldRef, listFields, webFields));
                }, Promise.resolve());
        }
    }

    /**
     *
     * Processes a field ref for a list
     *
     * @param {List} list The list
     * @param {IList} lc The list configuration
     * @param {IListInstanceFieldRef} fieldRef The list field ref
     * @param {any[]} listFields The list fields
     * @param {any[]} webFields The web fields
     */
    private async processFieldRef(list: List, lc: IList, fieldRef: IListInstanceFieldRef, listFields: ISPField[], webFields: ISPField[]): Promise<void> {
        let [listFld] = listFields.filter(f => f.Id == fieldRef.ID);
        let [webFld] = webFields.filter(f => f.Id == fieldRef.ID);
        super.log_info('processFieldRef', `Processing field ref '${fieldRef.ID}' for list ${lc.Title}.`);
        if (listFld) {
            await list.fields.getById(fieldRef.ID).update({ Hidden: fieldRef.Hidden, Required: fieldRef.Required, Title: fieldRef.DisplayName });
            super.log_info('processFieldRef', `Field '${fieldRef.ID}' updated for list ${lc.Title}.`);
        } else if (webFld) {
            let fieldAddResult = await list.fields.createFieldAsXml(webFld.SchemaXml);
            fieldAddResult.field.update({ Title: fieldRef.DisplayName, Required: fieldRef.Required, Hidden: fieldRef.Hidden });
            super.log_info('processFieldRef', `Field '${fieldRef.ID}' added from web.`);
        }
    }

    /**
     * Processes views for a list
     *
     * @param {Web} web The web
     * @param {IList} lc The list configuration
     */
    private async processListViews(web: Web, lc: IList): Promise<any> {
        if (lc.Views) {
            await lc.Views.reduce((chain: any, view) => chain.then(() => this.processView(web, lc, view)), Promise.resolve());
        }
        this.context.listViews = (await web.lists.getByTitle(lc.Title).views.select('Id', 'Title').get<Array<{ Id: string, Title: string }>>()).reduce((obj, view) => {
            obj[`${lc.Title}|${view.Title}`] = view.Id;
            return obj;
        }, this.context.listViews);
    }

    /**
     * Processes a view for a list
     *
     * @param {Web} web The web
     * @param {IList} lc The list configuration
     * @param {IListView} lvc The view configuration
     */
    private async processView(web: Web, lc: IList, lvc: IListView): Promise<void> {
        super.log_info('processView', `Processing view ${lvc.Title} for list ${lc.Title}.`);
        let existingView = web.lists.getByTitle(lc.Title).views.getByTitle(lvc.Title);
        let viewExists = false;
        try {
            await existingView.get();
            viewExists = true;
        } catch (err) { }
        try {
            if (viewExists) {
                super.log_info('processView', `View ${lvc.Title} for list ${lc.Title} already exists, updating.`);
                await existingView.update(lvc.AdditionalSettings);
                super.log_info('processView', `View ${lvc.Title} successfully updated for list ${lc.Title}.`);
                await this.processViewFields(existingView, lvc);
            } else {
                super.log_info('processView', `View ${lvc.Title} for list ${lc.Title} doesn't exists, creating.`);
                const result = await web.lists.getByTitle(lc.Title).views.add(lvc.Title, lvc.PersonalView, lvc.AdditionalSettings);
                super.log_info('processView', `View ${lvc.Title} added successfully to list ${lc.Title}.`);
                await this.processViewFields(result.view, lvc);
            }
        } catch (err) {
            super.log_info('processViewFields', `Failed to process view for view ${lvc.Title}.`);
        }
    }

    /**
     * Processes view fields for a view
     *
     * @param {any} view The pnp view
     * @param {IListView} lvc The view configuration
     */
    private async processViewFields(view: any, lvc: IListView): Promise<void> {
        try {
            super.log_info('processViewFields', `Processing view fields for view ${lvc.Title}.`);
            await view.fields.removeAll();
            await lvc.ViewFields.reduce((chain, viewField) => chain.then(() => view.fields.add(viewField)), Promise.resolve());
            super.log_info('processViewFields', `View fields successfully processed for view ${lvc.Title}.`);
        } catch (err) {
            super.log_info('processViewFields', `Failed to process view fields for view ${lvc.Title}.`);
        }
    }
}
