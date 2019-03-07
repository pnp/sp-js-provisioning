import { IClientSidePage } from "../schema";
import { HandlerBase } from "./handlerbase";
import { Web, ClientSideWebpart, ClientSidePageComponent, ClientSideWebpartPropertyTypes } from "@pnp/sp";
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
            const partDefinitions = await web.getClientSideWebParts();
            await clientSidePages.reduce((chain: any, clientSidePage) => chain.then(() => this.processClientSidePage(web, clientSidePage, partDefinitions)), Promise.resolve());
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
     * @param {ClientSidePageComponent[]} partDefinitions Cient side web parts
     */
    private async processClientSidePage(web: Web, { Name, Title, Sections, CommentsDisabled }: IClientSidePage, partDefinitions: ClientSidePageComponent[]) {
        super.log_info("processClientSidePage", `Processing client side page ${Name}`);
        const page = await web.addClientSidePage(Name, Title);
        if (Sections) {
            Sections.forEach(s => {
                const section = page.addSection();
                s.Columns.forEach(col => {
                    const column = section.addColumn(col.Factor);
                    col.Controls.forEach(control => {
                        const [partDef] = partDefinitions.filter(c => c.Id.toLowerCase().indexOf(control.Id.toLowerCase()) !== -1);
                        if (partDef) {
                            try {
                                const part = ClientSideWebpart
                                    .fromComponentDef(partDef)
                                    .setProperties<any>(JSON.parse(this.tokenHelper.replaceTokens(JSON.stringify(control.Properties))));
                                super.log_info("processClientSidePage", `Adding ${partDef.Name} to client side page ${Name}`);
                                column.addControl(part);
                            } catch (error) {
                                console.log(error);
                                super.log_info("processClientSidePage", `Failed adding part ${partDef.Name} to client side page ${Name}`);
                            }
                        } else {
                            super.log_warn("processClientSidePage", `Client side web part with definition id ${control.Id} not found.`);
                        }
                    });
                });
            });
            await page.save();
        }
        if (CommentsDisabled) {
            super.log_info("processClientSidePage", `Disabling comments for client side page ${Name}`);
            await page.disableComments();
        }
    }
}
