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
const handlerbase_1 = require("./handlerbase");
const sp_1 = require("@pnp/sp");
/**
 * Describes the Composed Look Object Handler
 */
class ClientSidePages extends handlerbase_1.HandlerBase {
    /**
     * Creates a new instance of the ObjectClientSidePages class
     */
    constructor(config) {
        super("ClientSidePages", config);
    }
    /**
     * Provisioning Client Side Pages
     *
     * @param {Web} web The web
     * @param {IClientSidePage[]} clientSidePages The client side pages to provision
     * @param {ProvisioningContext} context Provisioning context
     */
    ProvisionObjects(web, clientSidePages, context) {
        const _super = Object.create(null, {
            scope_started: { get: () => super.scope_started },
            scope_ended: { get: () => super.scope_ended }
        });
        return __awaiter(this, void 0, void 0, function* () {
            this.context = context;
            _super.scope_started.call(this);
            try {
                yield clientSidePages.reduce((chain, clientSidePage) => chain.then(_ => this.processClientSidePage(web, clientSidePage)), Promise.resolve());
            }
            catch (err) {
                _super.scope_ended.call(this);
                throw err;
            }
        });
    }
    /**
     * Provision a client side page
     *
     * @param {Web} web The web
     * @param {IClientSidePage} clientSidePage Cient side page
     */
    processClientSidePage(web, clientSidePage) {
        const _super = Object.create(null, {
            log_info: { get: () => super.log_info }
        });
        return __awaiter(this, void 0, void 0, function* () {
            _super.log_info.call(this, "processClientSidePage", `Processing client side page ${clientSidePage.Name}.`);
            const page = yield web.addClientSidePage(clientSidePage.Name, clientSidePage.Title, clientSidePage.LibraryTitle);
            if (clientSidePage.Sections) {
                clientSidePage.Sections.forEach(s => {
                    const section = page.addSection();
                    s.Columns.forEach(col => {
                        const column = section.addColumn(col.Factor);
                        col.Controls.forEach(control => {
                            let controlJsonString = this.context.replaceTokens(JSON.stringify(control));
                            control = JSON.parse(controlJsonString);
                            column.addControl(new sp_1.ClientSideWebpart(control.Title, control.Description, control.ClientSideComponentProperties, control.ClientSideComponentId));
                        });
                    });
                });
                yield page.save();
            }
            yield page.publish();
            if (clientSidePage.CommentsDisabled) {
                yield page.disableComments();
            }
            if (clientSidePage.PageLayoutType) {
                const pageItem = yield page.getItem();
                yield pageItem.update({ PageLayoutType: clientSidePage.PageLayoutType });
            }
        });
    }
}
exports.ClientSidePages = ClientSidePages;
