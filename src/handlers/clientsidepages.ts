import { IClientSidePage } from "../schema";
import { HandlerBase } from "./handlerbase";
import { Web, ClientSideWebpart } from "@pnp/sp";
import { ProvisioningContext } from "../provisioningcontext";
import { IProvisioningConfig } from "../provisioningconfig";

/**
 * Describes the Composed Look Object Handler
 */
export class ClientSidePages extends HandlerBase {
    private context: ProvisioningContext;

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
     * @param {ProvisioningContext} context Provisioning context
     */
    public async ProvisionObjects(web: Web, clientSidePages: IClientSidePage[], context: ProvisioningContext): Promise<void> {
        this.context = context;
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
        super.log_info("processClientSidePage", `Processing client side page ${clientSidePage.Name}.`);
        const page = await web.addClientSidePage(clientSidePage.Name, clientSidePage.Title, clientSidePage.LibraryTitle);
        if (clientSidePage.Sections) {
            clientSidePage.Sections.forEach(s => {
                const section = page.addSection();
                s.Columns.forEach(col => {
                    const column = section.addColumn(col.Factor);
                    col.Controls.forEach(control => {
                        let controlJsonString = this.context.replaceTokens(JSON.stringify(control));
                        control = JSON.parse(controlJsonString);
                        column.addControl(new ClientSideWebpart(control.Title, control.Description, control.ClientSideComponentProperties, control.ClientSideComponentId));
                    });
                });
            });
            await page.save();
        }
        await page.publish();
        if (clientSidePage.CommentsDisabled) {
            await page.disableComments();
        }
        if (clientSidePage.PageLayoutType) {
            const pageItem = await page.getItem();
            await pageItem.update({ PageLayoutType: clientSidePage.PageLayoutType });
        }
    }
}
