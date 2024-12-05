"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.downloadFile = exports.statExt = void 0;
const tslib_1 = require("tslib");
const client_1 = require("./client");
const core_1 = require("@actions/core");
const promises_1 = require("node:fs/promises");
const Path = tslib_1.__importStar(require("node:path"));
const groupsMap = new Map();
const decomposeSharePointURI = async (uri, client) => {
    const [groupName, ...pathParts] = uri.replace("sharepoint://", "").split("/");
    // sharepoint://Team/Builds/Automation/alpha/v2.1.4/docs/library.pdf
    // [___________][___][______________________________________________]
    //       |        |                         |
    //   Pseudo-    Group      Path to file on Team's Shared Documents
    //   protocol   name                 (relative to root)
    if (!groupName) {
        throw new Error(`No group name could be found in URI '${uri}'`);
    }
    const groupNameLC = groupName.toLowerCase();
    if (!groupsMap.has(groupNameLC)) {
        const res = await client.getGroup(groupNameLC);
        if (res) {
            groupsMap.set(groupNameLC, res);
        }
    }
    const group = groupsMap.get(groupNameLC);
    if (!group) {
        throw new Error(`Could not find SharePoint group ${groupName}`);
    }
    return { group, pathParts };
};
const statExt = async (path, options) => {
    let s;
    try {
        s = await (0, promises_1.stat)(path, options);
        return Object.assign(s, { exists: true });
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
    }
    catch (ex) {
        // no-op.
    }
    return { ...{}, exists: false };
};
exports.statExt = statExt;
const downloadFile = async (uri, to, config) => {
    if (!uri.startsWith("sharepoint:")) {
        throw new Error(`Invalid URI: ${uri}`);
    }
    const client = (0, client_1.getClient)(config);
    const { group, pathParts } = await decomposeSharePointURI(uri, client);
    const fileName = pathParts[pathParts.length - 1];
    if (!fileName) {
        throw new Error(`Invalid URI: ${uri}`);
    }
    (0, core_1.info)(`Downloading '${uri}' to '${to}'...`);
    const itemPath = pathParts.join("/");
    // Try once to create the directory if it doesn't exist yet. No mkdirp.
    const stat = await (0, exports.statExt)(to);
    if (!stat.exists) {
        await (0, promises_1.mkdir)(to);
    }
    try {
        // This fetch may yield a 404 error and is cheaper than a full download.
        await client.getDriveItem(group, itemPath);
        await client.downloadItem(group, itemPath, Path.join(to, fileName));
    }
    catch (ex) {
        throw new Error(`Could not download '${itemPath}' with error ${ex.message}`);
    }
    (0, core_1.info)("done!");
};
exports.downloadFile = downloadFile;
(async function () {
    const config = {
        clientId: (0, core_1.getInput)("azure-client-id"),
        clientSecret: (0, core_1.getInput)("azure-client-secret"),
        tenantId: (0, core_1.getInput)("azure-tenant-id"),
    };
    await (0, exports.downloadFile)((0, core_1.getInput)("uri"), (0, core_1.getInput)("target"), config);
})();
