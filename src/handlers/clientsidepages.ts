import { IClientSidePage } from "../schema";
import { HandlerBase } from "./handlerbase";
import { Web, ClientSideWebpart } from "@pnp/sp";
import { IProvisioningConfig } from "../provisioningconfig";

/**
 * Describes the Composed Look Object Handler
 */
export class ClientSidePages extends HandlerBase {
    /**
     * Creates a new instance of the ObjectClientSidePages class
     */
    constructor(config: IProvisioningConfig) {
        super("ClientSidePages", config);
    }

    /**
     * Provisioning Client Side Pages
     *
     * @param {Web} web The web
     * @param {IClientSidePage[]} clientSidePages The client side pages to provision
     */
    public async ProvisionObjects(web: Web, clientSidePages: IClientSidePage[]): Promise<void> {
        super.scope_started();
        try {
            await clientSidePages.reduce((chain, clientSidePage) => chain.then(_ => this.processClientSidePage(web, clientSidePage)), Promise.resolve());
        } catch (err) {
            super.scope_ended();
            throw err;
        }
    }

    /**
     * Provision a client side page
     *
     * @param {Web} web The web
     * @param {IClientSidePage} clientSidePage Cient side page
     */
    private async processClientSidePage(web: Web, clientSidePage: IClientSidePage) {
        const page = await web.addClientSidePage(clientSidePage.Name, clientSidePage.Title, clientSidePage.LibraryTitle);
        if (clientSidePage.Sections) {
            clientSidePage.Sections.forEach(s => {
                const section = page.addSection();
                s.Columns.forEach(col => {
                    const column = section.addColumn(col.Factor);
                    col.Controls.forEach(control => {
                        column.addControl(new ClientSideWebpart(control.Title, control.Description, control.ClientSideComponentProperties, control.ClientSideComponentId));
                    });
                });
            });
            await page.save();
        }
        if (clientSidePage.CommentsDisabled) {
            await page.disableComments();
        }
        if (clientSidePage.PageLayoutType) {
            const pageItem = await page.getItem();
            await pageItem.update({ PageLayoutType: clientSidePage.PageLayoutType });
        }
    }
}
