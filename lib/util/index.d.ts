/// <reference types="sharepoint" />
export declare function replaceTokens(str: string): string;
export declare function makeUrlRelative(absUrl: string): string;
export declare function base64EncodeString(str: string): string;
export declare function isNode(): boolean;
export declare function executeQuery(ctx: SP.ClientContext): Promise<{}>;
