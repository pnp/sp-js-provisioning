import { HandlerBase } from "./handlerbase";
import { IList } from "../schema";
import { Web } from "sp-pnp-js";
import { ProvisioningContext } from "../provisioningcontext";
/**
 * Describes the Lists Object Handler
 */
export declare class Lists extends HandlerBase {
    private context;
    /**
     * Creates a new instance of the Lists class
     */
    constructor();
    /**
     * Provisioning lists
     *
     * @param {Web} web The web
     * @param {Array<IList>} lists The lists to provision
     */
    ProvisionObjects(web: Web, lists: IList[], context: ProvisioningContext): Promise<void>;
    /**
     * Processes a list
     *
     * @param {Web} web The web
     * @param {IList} lc The list
     */
    private processList(web, lc);
    /**
     * Processes content type bindings for a list
     *
     * @param {IList} lc The list configuration
     * @param {List} list The pnp list
     * @param {Array<IContentTypeBinding>} contentTypeBindings Content type bindings
     * @param {boolean} removeExisting Remove existing content type bindings
     */
    private processContentTypeBindings(lc, list, contentTypeBindings, removeExisting);
    /**
     * Processes a content type binding for a list
     *
     * @param {IList} lc The list configuration
     * @param {List} list The pnp list
     * @param {string} contentTypeID The Content Type ID
     */
    private processContentTypeBinding(lc, list, contentTypeID);
    /**
     * Processes fields for a list
     *
     * @param {Web} web The web
     * @param {IList} list The pnp list
     */
    private processFields(web, list);
    /**
     * Processes a field for a lit
     *
     * @param {Web} web The web
     * @param {IList} lc The list configuration
     * @param {string} fieldXml Field xml
     */
    private processField(web, lc, fieldXml);
    /**
   * Processes field refs for a list
   *
   * @param {Web} web The web
   * @param {IList} list The pnp list
   */
    private processFieldRefs(web, list);
    /**
     *
     * Processes a field ref for a list
     *
     * @param {Web} web The web
     * @param {IList} lc The list configuration
     * @param {IListInstanceFieldRef} fieldRef The list field ref
     */
    private processFieldRef(web, lc, fieldRef);
    /**
     * Processes views for a list
     *
     * @param web The web
     * @param lc The list configuration
     */
    private processViews(web, lc);
    /**
     * Processes a view for a list
     *
     * @param {Web} web The web
     * @param {IList} lc The list configuration
     * @param {IListView} lvc The view configuration
     */
    private processView(web, lc, lvc);
    /**
     * Processes view fields for a view
     *
     * @param {any} view The pnp view
     * @param {IListView} lvc The view configuration
     */
    private processViewFields(view, lvc);
}
