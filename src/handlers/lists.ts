import * as xmljs from "xml-js";
import { HandlerBase } from "./handlerbase";
import { IContentTypeBinding, IList, IListInstanceFieldRef, IListView } from "../schema";
import { Web, List, Logger, LogLevel } from "sp-pnp-js";
import { ProvisioningContext } from "../provisioningcontext";

/**
 * Describes the Lists Object Handler
 */
export class Lists extends HandlerBase {
    private context: ProvisioningContext;

    /**
     * Creates a new instance of the Lists class
     */
    constructor() {
        super("Lists");
    }

    /**
     * Provisioning lists
     *
     * @param {Web} web The web
     * @param {Array<IList>} lists The lists to provision
     */
    public async ProvisionObjects(web: Web, lists: IList[], context: ProvisioningContext): Promise<void> {
        this.context = context;
        super.scope_started();
        try {
            await lists.reduce((chain, list) => chain.then(_ => this.processList(web, list)), Promise.resolve());
            await lists.reduce((chain, list) => chain.then(_ => this.processFields(web, list)), Promise.resolve());
            await lists.reduce((chain, list) => chain.then(_ => this.processFieldRefs(web, list)), Promise.resolve());
            await lists.reduce((chain, list) => chain.then(_ => this.processViews(web, list)), Promise.resolve());
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
        const { created, list, data } = await web.lists.ensure(lc.Title, lc.Description, lc.Template, lc.ContentTypesEnabled, lc.AdditionalSettings);
        this.context.lists.push(data);
        if (created) {
            Logger.log({ data: list, level: LogLevel.Info, message: `List ${lc.Title} created successfully.` });
        }
        await this.processContentTypeBindings(lc, list, lc.ContentTypeBindings, lc.RemoveExistingContentTypes);
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
        if (contentTypeBindings) {
            await contentTypeBindings.reduce((chain, ct) => chain.then(_ => this.processContentTypeBinding(lc, list, ct.ContentTypeID)), Promise.resolve());
            if (removeExisting) {
                let promises = [];
                const contentTypes = await list.contentTypes.get();
                contentTypes.forEach(({ Id: { StringValue: ContentTypeId } }) => {
                    let shouldRemove = (contentTypeBindings.filter(ctb => ContentTypeId.indexOf(ctb.ContentTypeID) !== -1).length === 0)
                        && (ContentTypeId.indexOf("0x0120") === -1);
                    if (shouldRemove) {
                        Logger.write(`Removing content type ${ContentTypeId} from list ${lc.Title}`, LogLevel.Info);
                        promises.push(list.contentTypes.getById(ContentTypeId).delete());
                    }
                });
                await Promise.all(promises);
            }
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
            await list.contentTypes.addAvailableContentType(contentTypeID);
            Logger.log({ message: `Content Type ${contentTypeID} added successfully to list ${lc.Title}.`, level: LogLevel.Info });
        } catch (err) {
            Logger.log({ message: `Failed to add Content Type ${contentTypeID} to list ${lc.Title}.`, level: LogLevel.Warning });
        }
    }


    /**
     * Processes fields for a list
     *
     * @param {Web} web The web
     * @param {IList} list The pnp list
     */
    private async processFields(web: Web, list: IList): Promise<any> {
        if (list.Fields) {
            await list.Fields.reduce((chain, field) => chain.then(_ => this.processField(web, list, field)), Promise.resolve());
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

        Logger.log({ message: `Processing field ${fieldName} (${fieldDisplayName}) for list ${lc.Title}.`, level: LogLevel.Info, data: fieldAttr });
        fXmlJson.elements[0].attributes.DisplayName = fieldName;
        fieldXml = xmljs.json2xml(fXmlJson);

        // Looks like e.g. lookup fields can't be updated, so we'll need to re-create the field
        try {
            let field = await list.fields.getById(fieldAttr.ID);
            await field.delete();
            Logger.log({ message: `Field ${fieldName} (${fieldDisplayName}) successfully deleted from list ${lc.Title}.`, level: LogLevel.Info });
        } catch (err) {
            Logger.log({ message: `Field ${fieldName} (${fieldDisplayName}) does not exist in list ${lc.Title}.`, level: LogLevel.Info });
        }

        // Looks like e.g. lookup fields can't be updated, so we'll need to re-create the field
        try {
            let fieldAddResult = await list.fields.createFieldAsXml(this.context.replaceTokens(fieldXml));
            await fieldAddResult.field.update({ Title: fieldDisplayName });
            Logger.log({ message: `Field '${fieldDisplayName}' added successfully to list ${lc.Title}.`, level: LogLevel.Info });
        } catch (err) {
            Logger.log({ message: `Failed to add field '${fieldDisplayName}' to list ${lc.Title}.`, level: LogLevel.Warning });
        }
    }

    /**
   * Processes field refs for a list
   *
   * @param {Web} web The web
   * @param {IList} list The pnp list
   */
    private async processFieldRefs(web: Web, list: IList): Promise<any> {
        if (list.FieldRefs) {
            await list.FieldRefs.reduce((chain, fieldRef) => chain.then(_ => this.processFieldRef(web, list, fieldRef)), Promise.resolve());
        }
    }

    /**
     *
     * Processes a field ref for a list
     *
     * @param {Web} web The web
     * @param {IList} lc The list configuration
     * @param {IListInstanceFieldRef} fieldRef The list field ref
     */
    private async processFieldRef(web: Web, lc: IList, fieldRef: IListInstanceFieldRef): Promise<void> {
        const list = web.lists.getByTitle(lc.Title);

        try {
            await list.fields.getById(fieldRef.ID).update({ Hidden: fieldRef.Hidden, Required: fieldRef.Required, Title: fieldRef.DisplayName });
            Logger.log({ data: fieldRef, level: LogLevel.Info, message: `Field '${fieldRef.ID}' updated for list ${lc.Title}.` });
        } catch (err) {
            Logger.log({ message: `Failed to update field '${fieldRef.ID}' for list ${lc.Title}.`, data: fieldRef, level: LogLevel.Warning });
        }
    }

    /**
     * Processes views for a list
     *
     * @param web The web
     * @param lc The list configuration
     */
    private async processViews(web: Web, lc: IList): Promise<any> {
        if (lc.Views) {
            await lc.Views.reduce((chain, view) => chain.then(_ => this.processView(web, lc, view)), Promise.resolve());
        }
    }

    /**
     * Processes a view for a list
     *
     * @param {Web} web The web
     * @param {IList} lc The list configuration
     * @param {IListView} lvc The view configuration
     */
    private async processView(web: Web, lc: IList, lvc: IListView): Promise<void> {
        Logger.log({ message: `Processing view ${lvc.Title} for list ${lc.Title}.`, level: LogLevel.Info });
        let view = web.lists.getByTitle(lc.Title).views.getByTitle(lvc.Title);
        try {
            await view.get();
            await view.update(lvc.AdditionalSettings);
            await this.processViewFields(view, lvc);
        } catch (err) {
            const result = await web.lists.getByTitle(lc.Title).views.add(lvc.Title, lvc.PersonalView, lvc.AdditionalSettings);
            Logger.log({ message: `View ${lvc.Title} added successfully to list ${lc.Title}.`, level: LogLevel.Info });
            await this.processViewFields(result.view, lvc);
        }
    }

    /**
     * Processes view fields for a view
     *
     * @param {any} view The pnp view
     * @param {IListView} lvc The view configuration
     */
    private async processViewFields(view, lvc: IListView): Promise<void> {
        try {
            Logger.log({ message: `Processing view fields for view ${lvc.Title}.`, data: { viewFields: lvc.ViewFields }, level: LogLevel.Info });
            await view.fields.removeAll();
            await lvc.ViewFields.reduce((chain, viewField) => chain.then(_ => view.fields.add(viewField)), Promise.resolve());
            Logger.log({ message: `View fields successfully processed for view ${lvc.Title}.`, level: LogLevel.Info });
        } catch (err) {
            Logger.log({ message: `Failed to process view fields for view ${lvc.Title}.`, level: LogLevel.Info });
        }
    }
}
