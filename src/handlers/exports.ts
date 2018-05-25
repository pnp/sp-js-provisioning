import { TypedHash } from "sp-pnp-js";
import { HandlerBase } from "./handlerbase";
import { ComposedLook } from "./composedlook";
import { CustomActions } from "./customactions";
import { Features } from "./features";
import { WebSettings } from "./websettings";
import { Navigation } from "./navigation";
import { Lists } from "./lists";
import { Files } from "./files";
import { PropertyBagEntries } from "./propertybagentries";

export const DefaultHandlerMap: TypedHash<HandlerBase> = {
    ComposedLook: new ComposedLook(),
    CustomActions: new CustomActions(),
    Features: new Features(),
    Files: new Files(),
    Lists: new Lists(),
    Navigation: new Navigation(),
    PropertyBagEntries: new PropertyBagEntries(),
    WebSettings: new WebSettings(),
};

export const DefaultHandlerSort: TypedHash<number> = {
    ComposedLook: 6,
    CustomActions: 5,
    Features: 2,
    Files: 4,
    Lists: 3,
    Navigation: 7,
    PropertyBagEntries: 8,
    WebSettings: 1,
};

