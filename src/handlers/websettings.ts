import { HandlerBase } from "./handlerbase";
import { IWebSettings } from "../schema";
import { Web } from "sp-pnp-js";
import * as omit from "object.omit";
import { replaceUrlTokens } from "../util";

/**
 * Describes the WebSettings Object Handler
 */
export class WebSettings extends HandlerBase {
    /**
     * Creates a new instance of the WebSettings class
     */
    constructor() {
        super("WebSettings");
    }

    /**
     * Provisioning WebSettings
     *
     * @param {Web} web The web
     * @param {IWebSettings} settings The settings
     */
    public ProvisionObjects(web: Web, settings: IWebSettings): Promise<void> {
        super.scope_started();
        return new Promise<void>((resolve, reject) => {
            Object.keys(settings)
                .filter(key => typeof (settings[key]) === "string")
                .forEach(key => {
                    let value: string = <any>settings[key];
                    settings[key] = replaceUrlTokens(value);
                });
            Promise.all([
                web.rootFolder.update({ WelcomePage: settings.WelcomePage }),
                web.update(omit(settings, "WelcomePage")),
            ]).then(_ => {
                super.scope_ended();
                resolve();
            }).catch(e => reject(e));
        });
    }
}
