"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
function replaceUrlTokens(str, config) {
    let siteAbsoluteUrl = null;
    if (config.spfxContext) {
        siteAbsoluteUrl = config.spfxContext.pageContext.site.absoluteUrl;
    }
    else if (window.hasOwnProperty("_spPageContextInfo")) {
        siteAbsoluteUrl = _spPageContextInfo.siteAbsoluteUrl;
    }
    return str
        .replace(/{sitecollection}/g, siteAbsoluteUrl)
        .replace(/{wpgallery}/g, `${siteAbsoluteUrl}/_catalogs/wp`)
        .replace(/{hosturl}/g, `${window.location.protocol}//${window.location.host}:${window.location.port}`)
        .replace(/{themegallery}/g, `${siteAbsoluteUrl}/_catalogs/theme/15`);
}
exports.replaceUrlTokens = replaceUrlTokens;
function makeUrlRelative(absUrl) {
    return absUrl.replace(`${document.location.protocol}//${document.location.hostname}`, "");
}
exports.makeUrlRelative = makeUrlRelative;
function base64EncodeString(str) {
    let bytes = [];
    for (let i = 0; i < str.length; ++i) {
        bytes.push(str.charCodeAt(i));
        bytes.push(0);
    }
    let b64encoded = window.btoa(String.fromCharCode.apply(null, bytes));
    return b64encoded;
}
exports.base64EncodeString = base64EncodeString;
function isNode() {
    return typeof window === "undefined";
}
exports.isNode = isNode;
