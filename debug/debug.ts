declare var require: any;
import {
    setup,
    Logger,
    LogLevel,
    ConsoleListener,
    NodeFetchClient,
    default as pnp, Web
} from "sp-pnp-js";

// setup the connection to SharePoint using the settings file, you can
// override any of the values as you want here, just be sure not to commit
// your account details :)
// if you don't have a settings file defined this will error
// you can comment it out and put the values here directly, or better yet
// create a settings file using settings.example.js as a template
let settings = require("../../settings.js");

// configure your node options
setup({
    fetchClientFactory: () => {
        return new NodeFetchClient(settings.testing.siteUrl, settings.testing.clientId, settings.testing.clientSecret);
    }
});

// setup console logger
Logger.subscribe(new ConsoleListener());
Logger.activeLogLevel = LogLevel.Verbose;

import { Example } from "./example";

cleanUpAllSubsites().then(_ => {

    // importing the example debug scenario and running it
    // adding your debugging to other files and importing them will keep them out of git
    // PRs updating the debug.ts or example.ts will not be accepted unless they are fixing bugs
    // add your debugging imports here and prior to submitting a PR git checkout debug/debug.ts
    // will allow you to keep all your debugging files locally
    // comment out the example
    Example();
});


// you can also set break points inside the src folder to examine how things are working
// within the library while debugging!

// this can be used to clean up lots of test sub webs :)
function cleanUpAllSubsites(): Promise<void> {

    return new Promise<void>((resolve, reject) => {

        pnp.sp.site.rootWeb.webs.select("Title").get().then((w) => {
            return Promise.all(w.map((element: any) => {
                console.log(element["odata.id"]);
                let web = new Web(element["odata.id"], "");
                return web.webs.select("Title").get().then((sw: any[]) => {
                    return Promise.all(sw.map((value) => {
                        console.log(value["odata.id"]);
                        let web2 = new Web(value["odata.id"], "");
                        return web2.delete().catch(e => {});
                    }));
                }).then(() => { web.delete(); }).catch(e => {});
            }));

        }).catch(e => {
            reject(e);
        }).then(_ => {
            resolve();
        });
    });
}
