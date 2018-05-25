"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
function ReplaceTokens(str) {
    return str.replace(/{sitecollection}/g, _spPageContextInfo.siteAbsoluteUrl)
        .replace(/{wpgallery}/g, `${_spPageContextInfo.siteAbsoluteUrl}/_catalogs/wp`)
        .replace(/{themegallery}/g, `${_spPageContextInfo.siteAbsoluteUrl}/_catalogs/theme/15`);
}
exports.ReplaceTokens = ReplaceTokens;
function MakeUrlRelative(absUrl) {
    return absUrl.replace(`${document.location.protocol}//${document.location.hostname}`, "");
}
exports.MakeUrlRelative = MakeUrlRelative;
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
