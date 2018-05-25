import { HandlerBase } from "./handlerbase";
import { IPropertyBagEntry } from "../schema";
import * as Util from "../util";
import { Web, Logger, LogLevel } from "sp-pnp-js";

/**
 * Describes the PropertyBagEntries Object Handler
 */
export class PropertyBagEntries extends HandlerBase {
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
    public ProvisionObjects(web: Web, entries: IPropertyBagEntry[]): Promise<void> {
        super.scope_started();
        return new Promise<any>((resolve, reject) => {
            if (Util.isNode()) {
                Logger.write("PropertyBagEntries Handler not supported in Node.", LogLevel.Error);
                reject();
            } else {
                web.get().then(({ ServerRelativeUrl }) => {
                    let ctx = new SP.ClientContext(ServerRelativeUrl),
                        spWeb = ctx.get_web(),
                        propBag = spWeb.get_allProperties(),
                        idxProps = [];
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
                        } else {
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
