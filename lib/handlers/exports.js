import { ComposedLook } from "./composedlook";
import { CustomActions } from "./customactions";
import { Features } from "./features";
import { WebSettings } from "./websettings";
import { Navigation } from "./navigation";
import { Lists } from "./lists";
import { Files } from "./files";
import { ClientSidePages } from "./clientsidepages";
import { PropertyBagEntries } from "./propertybagentries";
import { SiteFields } from "./sitefields";
export var DefaultHandlerMap = function (config) { return ({
    ClientSidePages: new ClientSidePages(config),
    ComposedLook: new ComposedLook(config),
    CustomActions: new CustomActions(config),
    Features: new Features(config),
    Files: new Files(config),
    Lists: new Lists(config),
    Navigation: new Navigation(config),
    PropertyBagEntries: new PropertyBagEntries(config),
    WebSettings: new WebSettings(config),
    SiteFields: new SiteFields(config),
}); };
export var DefaultHandlerSort = {
    ClientSidePages: 7,
    ComposedLook: 6,
    CustomActions: 5,
    Features: 2,
    Files: 4,
    Lists: 3,
    Navigation: 9,
    PropertyBagEntries: 8,
    WebSettings: 10,
    SiteFields: 0,
};
