import { IComposedLook } from "../schema";
import { HandlerBase } from "./handlerbase";
import { Web } from "@pnp/sp";
import { replaceUrlTokens, makeUrlRelative } from "../util";
import { IProvisioningConfig} from "../provisioningconfig";

/**
 * Describes the Composed Look Object Handler
 */
export class ComposedLook extends HandlerBase {
    /**
     * Creates a new instance of the ObjectComposedLook class
     */
    constructor(config: IProvisioningConfig) {
        super("ComposedLook", config);
    }

    /**
     * Provisioning Composed Look
     *
     * @param {Web} web The web
     * @param {IComposedLook} object The Composed Look to provision
     */
    public async ProvisionObjects(web: Web, composedLook: IComposedLook): Promise<void> {
        super.scope_started();
        try {
            await web.applyTheme(
                makeUrlRelative(replaceUrlTokens(composedLook.ColorPaletteUrl, this.config)),
                makeUrlRelative(replaceUrlTokens(composedLook.FontSchemeUrl, this.config)),
                composedLook.BackgroundImageUrl ? makeUrlRelative(replaceUrlTokens(composedLook.BackgroundImageUrl, this.config)) : null,
                false);
            super.scope_ended();
        } catch (err) {
            super.scope_ended();
            throw err;
        }
    }
}
