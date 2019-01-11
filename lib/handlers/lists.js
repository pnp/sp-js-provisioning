"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const xmljs = require("xml-js");
const handlerbase_1 = require("./handlerbase");
/**
 * Describes the Lists Object Handler
 */
class Lists extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the Lists class
     *
     * @param {IProvisioningConfig} config Provisioning config
     */
    constructor(config) {
        super("Lists", config);
    }
    /**
     * Provisioning lists
     *
     * @param {Web} web The web
     * @param {Array<IList>} lists The lists to provision
     */
    ProvisionObjects(web, lists, context) {
        const _super = Object.create(null, {
            scope_started: { get: () => super.scope_started },
            scope_ended: { get: () => super.scope_ended }
        });
        return __awaiter(this, void 0, void 0, function* () {
            this.context = context;
            _super.scope_started.call(this);
            try {
                yield lists.reduce((chain, list) => chain.then(_ => this.processList(web, list)), Promise.resolve());
                yield lists.reduce((chain, list) => chain.then(_ => this.processListFields(web, list)), Promise.resolve());
                yield lists.reduce((chain, list) => chain.then(_ => this.processListFieldRefs(web, list)), Promise.resolve());
                yield lists.reduce((chain, list) => chain.then(_ => this.processListViews(web, list)), Promise.resolve());
                _super.scope_ended.call(this);
            }
            catch (err) {
                _super.scope_ended.call(this);
                throw err;
            }
        });
    }
    /**
     * Processes a list
     *
     * @param {Web} web The web
     * @param {IList} lc The list
     */
    processList(web, lc) {
        const _super = Object.create(null, {
            log_info: { get: () => super.log_info }
        });
        return __awaiter(this, void 0, void 0, function* () {
            _super.log_info.call(this, "processList", `Processing list ${lc.Title}`);
            const { list, data } = yield web.lists.add(lc.Title, lc.Description, lc.Template, lc.ContentTypesEnabled, lc.AdditionalSettings);
            this.context.lists.push(data);
            if (lc.ContentTypeBindings) {
                yield this.processContentTypeBindings(lc, list, lc.ContentTypeBindings, lc.RemoveExistingContentTypes);
            }
        });
    }
    /**
     * Processes content type bindings for a list
     *
     * @param {IList} lc The list configuration
     * @param {List} list The pnp list
     * @param {Array<IContentTypeBinding>} contentTypeBindings Content type bindings
     * @param {boolean} removeExisting Remove existing content type bindings
     */
    processContentTypeBindings(lc, list, contentTypeBindings, removeExisting) {
        const _super = Object.create(null, {
            log_info: { get: () => super.log_info }
        });
        return __awaiter(this, void 0, void 0, function* () {
            yield contentTypeBindings.reduce((chain, ct) => chain.then(_ => this.processContentTypeBinding(lc, list, ct.ContentTypeID)), Promise.resolve());
            if (removeExisting) {
                let promises = [];
                const contentTypes = yield list.contentTypes.get();
                contentTypes.forEach(({ Id: { StringValue: ContentTypeId } }) => {
                    let shouldRemove = (contentTypeBindings.filter(ctb => ContentTypeId.indexOf(ctb.ContentTypeID) !== -1).length === 0)
                        && (ContentTypeId.indexOf("0x0120") === -1);
                    if (shouldRemove) {
                        _super.log_info.call(this, "processContentTypeBindings", `Removing content type ${ContentTypeId} from list ${lc.Title}`);
                        promises.push(list.contentTypes.getById(ContentTypeId).delete());
                    }
                });
                yield Promise.all(promises);
            }
        });
    }
    /**
     * Processes a content type binding for a list
     *
     * @param {IList} lc The list configuration
     * @param {List} list The pnp list
     * @param {string} contentTypeID The Content Type ID
     */
    processContentTypeBinding(lc, list, contentTypeID) {
        const _super = Object.create(null, {
            log_info: { get: () => super.log_info }
        });
        return __awaiter(this, void 0, void 0, function* () {
            try {
                yield list.contentTypes.addAvailableContentType(contentTypeID);
                _super.log_info.call(this, "processContentTypeBinding", `Content Type ${contentTypeID} added successfully to list ${lc.Title}.`);
            }
            catch (err) {
                _super.log_info.call(this, "processContentTypeBinding", `Failed to add Content Type ${contentTypeID} to list ${lc.Title}.`);
            }
        });
    }
    /**
     * Processes fields for a list
     *
     * @param {Web} web The web
     * @param {IList} list The pnp list
     */
    processListFields(web, list) {
        return __awaiter(this, void 0, void 0, function* () {
            if (list.Fields) {
                yield list.Fields.reduce((chain, field) => chain.then(_ => this.processField(web, list, field)), Promise.resolve());
            }
        });
    }
    /**
     * Processes a field for a lit
     *
     * @param {Web} web The web
     * @param {IList} lc The list configuration
     * @param {string} fieldXml Field xml
     */
    processField(web, lc, fieldXml) {
        const _super = Object.create(null, {
            log_info: { get: () => super.log_info }
        });
        return __awaiter(this, void 0, void 0, function* () {
            const list = web.lists.getByTitle(lc.Title);
            const fXmlJson = JSON.parse(xmljs.xml2json(fieldXml));
            const fieldAttr = fXmlJson.elements[0].attributes;
            const fieldName = fieldAttr.Name;
            const fieldDisplayName = fieldAttr.DisplayName;
            _super.log_info.call(this, "processField", `Processing field ${fieldName} (${fieldDisplayName}) for list ${lc.Title}.`);
            fXmlJson.elements[0].attributes.DisplayName = fieldName;
            fieldXml = xmljs.json2xml(fXmlJson);
            // Looks like e.g. lookup fields can't be updated, so we'll need to re-create the field
            try {
                let field = yield list.fields.getById(fieldAttr.ID);
                yield field.delete();
                _super.log_info.call(this, "processField", `Field ${fieldName} (${fieldDisplayName}) successfully deleted from list ${lc.Title}.`);
            }
            catch (err) {
                _super.log_info.call(this, "processField", `Field ${fieldName} (${fieldDisplayName}) does not exist in list ${lc.Title}.`);
            }
            // Looks like e.g. lookup fields can't be updated, so we'll need to re-create the field
            try {
                let fieldAddResult = yield list.fields.createFieldAsXml(this.context.replaceTokens(fieldXml));
                yield fieldAddResult.field.update({ Title: fieldDisplayName });
                _super.log_info.call(this, "processField", `Field '${fieldDisplayName}' added successfully to list ${lc.Title}.`);
            }
            catch (err) {
                _super.log_info.call(this, "processField", `Failed to add field '${fieldDisplayName}' to list ${lc.Title}.`);
            }
        });
    }
    /**
   * Processes field refs for a list
   *
   * @param {Web} web The web
   * @param {IList} list The pnp list
   */
    processListFieldRefs(web, list) {
        return __awaiter(this, void 0, void 0, function* () {
            if (list.FieldRefs) {
                yield list.FieldRefs.reduce((chain, fieldRef) => chain.then(_ => this.processFieldRef(web, list, fieldRef)), Promise.resolve());
            }
        });
    }
    /**
     *
     * Processes a field ref for a list
     *
     * @param {Web} web The web
     * @param {IList} lc The list configuration
     * @param {IListInstanceFieldRef} fieldRef The list field ref
     */
    processFieldRef(web, lc, fieldRef) {
        const _super = Object.create(null, {
            log_info: { get: () => super.log_info }
        });
        return __awaiter(this, void 0, void 0, function* () {
            const list = web.lists.getByTitle(lc.Title);
            try {
                yield list.fields.getById(fieldRef.ID).update({ Hidden: fieldRef.Hidden, Required: fieldRef.Required, Title: fieldRef.DisplayName });
                _super.log_info.call(this, "processFieldRef", `Field '${fieldRef.ID}' updated for list ${lc.Title}.`);
            }
            catch (err) {
                _super.log_info.call(this, "processFieldRef", `Failed to update field '${fieldRef.ID}' for list ${lc.Title}.`);
            }
        });
    }
    /**
     * Processes views for a list
     *
     * @param web The web
     * @param lc The list configuration
     */
    processListViews(web, lc) {
        return __awaiter(this, void 0, void 0, function* () {
            if (lc.Views) {
                yield lc.Views.reduce((chain, view) => chain.then(_ => this.processView(web, lc, view)), Promise.resolve());
            }
        });
    }
    /**
     * Processes a view for a list
     *
     * @param {Web} web The web
     * @param {IList} lc The list configuration
     * @param {IListView} lvc The view configuration
     */
    processView(web, lc, lvc) {
        const _super = Object.create(null, {
            log_info: { get: () => super.log_info }
        });
        return __awaiter(this, void 0, void 0, function* () {
            _super.log_info.call(this, "processView", `Processing view ${lvc.Title} for list ${lc.Title}.`);
            let view = web.lists.getByTitle(lc.Title).views.getByTitle(lvc.Title);
            try {
                yield view.get();
                yield view.update(lvc.AdditionalSettings);
                yield this.processViewFields(view, lvc);
            }
            catch (err) {
                const result = yield web.lists.getByTitle(lc.Title).views.add(lvc.Title, lvc.PersonalView, lvc.AdditionalSettings);
                _super.log_info.call(this, "processView", `View ${lvc.Title} added successfully to list ${lc.Title}.`);
                yield this.processViewFields(result.view, lvc);
            }
        });
    }
    /**
     * Processes view fields for a view
     *
     * @param {any} view The pnp view
     * @param {IListView} lvc The view configuration
     */
    processViewFields(view, lvc) {
        const _super = Object.create(null, {
            log_info: { get: () => super.log_info }
        });
        return __awaiter(this, void 0, void 0, function* () {
            try {
                _super.log_info.call(this, "processViewFields", `Processing view fields for view ${lvc.Title}.`);
                yield view.fields.removeAll();
                yield lvc.ViewFields.reduce((chain, viewField) => chain.then(_ => view.fields.add(viewField)), Promise.resolve());
                _super.log_info.call(this, "processViewFields", `View fields successfully processed for view ${lvc.Title}.`);
            }
            catch (err) {
                _super.log_info.call(this, "processViewFields", `Failed to process view fields for view ${lvc.Title}.`);
            }
        });
    }
}
exports.Lists = Lists;
