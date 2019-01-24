var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import { HandlerBase } from "./handlerbase";
import * as Util from "../util";
import { Logger } from "@pnp/logging";
/**
 * Describes the PropertyBagEntries Object Handler
 */
var PropertyBagEntries = (function (_super) {
    __extends(PropertyBagEntries, _super);
    /**
     * Creates a new instance of the PropertyBagEntries class
     *
     * @param {IProvisioningConfig} config Provisioning config
     */
    function PropertyBagEntries(config) {
        return _super.call(this, "PropertyBagEntries", config) || this;
    }
    /**
     * Provisioning property bag entries
     *
     * @param {Web} web The web
     * @param {Array<IPropertyBagEntry>} entries The property bag entries to provision
     */
    PropertyBagEntries.prototype.ProvisionObjects = function (web, entries) {
        var _this = this;
        _super.prototype.scope_started.call(this);
        return new Promise(function (resolve, reject) {
            if (Util.isNode()) {
                Logger.write("PropertyBagEntries Handler not supported in Node.", 3 /* Error */);
                reject();
            }
            else if (_this.config.spfxContext) {
                Logger.write("PropertyBagEntries Handler not supported in SPFx.", 3 /* Error */);
                reject();
            }
            else {
                web.get().then(function (_a) {
                    var ServerRelativeUrl = _a.ServerRelativeUrl;
                    var ctx = new SP.ClientContext(ServerRelativeUrl), spWeb = ctx.get_web(), propBag = spWeb.get_allProperties(), idxProps = [];
                    entries.filter(function (entry) { return entry.Overwrite; }).forEach(function (entry) {
                        propBag.set_item(entry.Key, entry.Value);
                        if (entry.Indexed) {
                            idxProps.push(Util.base64EncodeString(entry.Key));
                        }
                    });
                    spWeb.update();
                    ctx.load(propBag);
                    ctx.executeQueryAsync(function () {
                        if (idxProps.length > 0) {
                            propBag.set_item("vti_indexedpropertykeys", idxProps.join("|"));
                            spWeb.update();
                            ctx.executeQueryAsync(function () {
                                _super.prototype.scope_ended.call(_this);
                                resolve();
                            }, function () {
                                _super.prototype.scope_ended.call(_this);
                                reject();
                            });
                        }
                        else {
                            _super.prototype.scope_ended.call(_this);
                            resolve();
                        }
                    }, function () {
                        _super.prototype.scope_ended.call(_this);
                        reject();
                    });
                });
            }
        });
    };
    return PropertyBagEntries;
}(HandlerBase));
export { PropertyBagEntries };
