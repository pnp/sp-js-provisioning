"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const handlerbase_1 = require("./handlerbase");
const Util = require("../util");
const sp_pnp_js_1 = require("sp-pnp-js");
/**
 * Describes the PropertyBagEntries Object Handler
 */
class PropertyBagEntries extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the PropertyBagEntries class
     */
    constructor() {
        super("PropertyBagEntries");
    }
    /**
     * Provisioning property bag entries
     *
     * @param {Web} web The web
     * @param {Array<IPropertyBagEntry>} entries The property bag entries to provision
     */
    ProvisionObjects(web, entries) {
        super.scope_started();
        return new Promise((resolve, reject) => {
            if (Util.isNode()) {
                sp_pnp_js_1.Logger.write("PropertyBagEntries Handler not supported in Node.", sp_pnp_js_1.LogLevel.Error);
                reject();
            }
            else {
                web.get().then(({ ServerRelativeUrl }) => {
                    let ctx = new SP.ClientContext(ServerRelativeUrl), spWeb = ctx.get_web(), propBag = spWeb.get_allProperties(), idxProps = [];
                    entries.filter(entry => entry.Overwrite).forEach(entry => {
                        propBag.set_item(entry.Key, entry.Value);
                        if (entry.Indexed) {
                            idxProps.push(Util.base64EncodeString(entry.Key));
                        }
                    });
                    spWeb.update();
                    ctx.load(propBag);
                    ctx.executeQueryAsync(() => {
                        if (idxProps.length > 0) {
                            propBag.set_item("vti_indexedpropertykeys", idxProps.join("|"));
                            spWeb.update();
                            ctx.executeQueryAsync(() => {
                                super.scope_ended();
                                resolve();
                            }, () => {
                                super.scope_ended();
                                reject();
                            });
                        }
                        else {
                            super.scope_ended();
                            resolve();
                        }
                    }, () => {
                        super.scope_ended();
                        reject();
                    });
                });
            }
        });
    }
}
exports.PropertyBagEntries = PropertyBagEntries;
