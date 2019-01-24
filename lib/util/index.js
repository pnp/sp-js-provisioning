export function replaceUrlTokens(str, config) {
    var siteAbsoluteUrl = null;
    if (config.spfxContext) {
        siteAbsoluteUrl = config.spfxContext.pageContext.site.absoluteUrl;
    }
    else if (window.hasOwnProperty("_spPageContextInfo")) {
        siteAbsoluteUrl = _spPageContextInfo.siteAbsoluteUrl;
    }
    return str
        .replace(/{sitecollection}/g, siteAbsoluteUrl)
        .replace(/{wpgallery}/g, siteAbsoluteUrl + "/_catalogs/wp")
        .replace(/{hosturl}/g, window.location.protocol + "//" + window.location.host + ":" + window.location.port)
        .replace(/{themegallery}/g, siteAbsoluteUrl + "/_catalogs/theme/15");
}
export function makeUrlRelative(absUrl) {
    return absUrl.replace(document.location.protocol + "//" + document.location.hostname, "");
}
export function base64EncodeString(str) {
    var bytes = [];
    for (var i = 0; i < str.length; ++i) {
        bytes.push(str.charCodeAt(i));
        bytes.push(0);
    }
    var b64encoded = window.btoa(String.fromCharCode.apply(null, bytes));
    return b64encoded;
}
export function isNode() {
    return typeof window === "undefined";
}
