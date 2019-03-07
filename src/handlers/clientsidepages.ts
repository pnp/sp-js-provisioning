import { IClientSidePage } from "../schema";
import { HandlerBase } from "./handlerbase";
import { Web, ClientSideWebpart, LayoutType } from "@pnp/sp";
import { ProvisioningContext } from "../provisioningcontext";
import { IProvisioningConfig } from "../provisioningconfig";
import { TokenHelper } from '../util/tokenhelper';

/**
 * Describes the Composed Look Object Handler
 */
export class ClientSidePages extends HandlerBase {
    private tokenHelper: TokenHelper;

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
        this.tokenHelper = new TokenHelper(context, this.config);
        super.scope_started();
        try {
            await clientSidePages.reduce((chain: any, clientSidePage) => chain.then(() => this.processClientSidePage(web, clientSidePage)), Promise.resolve());
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
        super.log_info("processClientSidePage", `Processing client side page ${clientSidePage.Name}`);
        const page = await web.addClientSidePage(clientSidePage.Name, clientSidePage.Title);
        if (clientSidePage.Sections) {
            clientSidePage.Sections.forEach(s => {
                const section = page.addSection();
                s.Columns.forEach(col => {
                    const column = section.addColumn(col.Factor);
                    col.Controls.forEach(control => {
                        let controlJsonString = this.tokenHelper.replaceTokens(JSON.stringify(control));
                        control = JSON.parse(controlJsonString);
                        super.log_info("processClientSidePage", `Adding ${control.webPartData.title} to client side page ${clientSidePage.Name}`);
                        column.addControl(new ClientSideWebpart(control));
                    });
                });
            });
            await page.save();
        }
        if (clientSidePage.CommentsDisabled) {
            super.log_info("processClientSidePage", `Disabling comments for client side page ${clientSidePage.Name}`);
            await page.disableComments();
        }
    }
}
