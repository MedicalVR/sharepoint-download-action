"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.getClient = exports.SharePointClient = void 0;
const tslib_1 = require("tslib");
require("isomorphic-fetch");
const identity_1 = require("@azure/identity");
const microsoft_graph_client_1 = require("@microsoft/microsoft-graph-client");
const node_fs_1 = require("node:fs");
const Path = tslib_1.__importStar(require("node:path"));
const node_util_1 = require("node:util");
const promises_1 = require("node:fs/promises");
const Stream = tslib_1.__importStar(require("node:stream"));
const azureTokenCredentials_1 = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");
const Finished = (0, node_util_1.promisify)(Stream.finished);
const TEMP_FILE_CONTENTS = "Temporary contents...";
class SharePointClient {
    constructor(config) {
        this.config = config;
        this.clientSecretCredential = new identity_1.ClientSecretCredential(this.config.tenantId, this.config.clientId, this.config.clientSecret);
    }
    getGraph() {
        if (this.graph) {
            return this.graph;
        }
        const authProvider = new azureTokenCredentials_1.TokenCredentialAuthenticationProvider(this.clientSecretCredential, {
            scopes: ["https://graph.microsoft.com/.default"],
        });
        this.graph = microsoft_graph_client_1.Client.initWithMiddleware({ authProvider });
        return this.graph;
    }
    async getGroup(name) {
        const groups = await this.getGraph().api("/groups").select(["displayName", "id"]).get();
        const nameLC = name.toLowerCase();
        for (const group of groups.value) {
            if (group.displayName.toLowerCase() === nameLC) {
                return {
                    id: group.id,
                    name: group.displayName,
                };
            }
        }
        return undefined;
    }
    async getDriveItem(group, path) {
        const res = await this.getGraph()
            .api(`/groups/${group.id}/drive/root:/${path}`)
            .select(["id"])
            .get();
        if (res) {
            return {
                id: res.id,
                path,
            };
        }
        return undefined;
    }
    async downloadItem(group, path, targetPath) {
        const stream = await this.getGraph()
            .api(`/groups/${group.id}/drive/root:/${path}:/content`)
            .responseType(microsoft_graph_client_1.ResponseType.BLOB)
            .get();
        if (stream) {
            const writer = (0, node_fs_1.createWriteStream)(targetPath);
            Stream.Readable.fromWeb(stream instanceof Blob ? stream.stream() : stream).pipe(writer);
            await Finished(writer);
        }
    }
    async uploadItem(group, path, targetPath, callback) {
        const pathDetails = Path.parse(targetPath);
        let item;
        try {
            item = await this.getDriveItem(group, targetPath);
            // eslint-disable-next-line @typescript-eslint/no-unused-vars
        }
        catch (ex) {
            // Do nothing.
        }
        const graph = this.getGraph();
        if (!item) {
            // First we need to create a new (placeholder) item that's located at the same path.
            const parent = await this.getDriveItem(group, pathDetails.dir);
            if (!parent) {
                throw new Error(`Directory not found on SharePoint of group ${group.name}: ${pathDetails.dir}`);
            }
            const res = await graph
                .api(`/groups/${group.id}/drive/items/${parent.id}:/${pathDetails.base}:/content`)
                .put(TEMP_FILE_CONTENTS);
            item = {
                id: res.id,
                path: targetPath,
            };
        }
        if (!item) {
            throw new Error(`Drive item at path ${targetPath} could not be created.`);
        }
        const stats = await (0, promises_1.stat)(path);
        const totalSize = stats.size;
        const progress = (range) => {
            if (!range) {
                callback(100);
                return;
            }
            const current = range.maxValue;
            const newer = Math.min(totalSize, current);
            const old = Math.max(totalSize, current);
            const frac = Math.abs(newer - old) / old;
            callback(100 - Math.floor(frac * 100));
        };
        const uploadEventHandlers = {
            progress,
        };
        const options = {
            rangeSize: 1024 * 1024,
            uploadEventHandlers,
        };
        // Create upload session for SharePoint upload.
        const payload = {
            item: {
                "@microsoft.graph.conflictBehavior": "replace",
            },
        };
        const reader = (0, node_fs_1.createReadStream)(path);
        const uploadSession = await microsoft_graph_client_1.LargeFileUploadTask.createUploadSession(graph, `https://graph.microsoft.com/v1.0/groups/${group.id}/drive/items/${item.id}/createuploadsession`, payload);
        const fileObject = new microsoft_graph_client_1.StreamUpload(reader, pathDetails.base, totalSize);
        const task = new microsoft_graph_client_1.LargeFileUploadTask(graph, fileObject, uploadSession, options);
        await task.upload();
        return item;
    }
}
exports.SharePointClient = SharePointClient;
let gClient = undefined;
const getClient = (config) => {
    if (gClient) {
        return gClient;
    }
    gClient = new SharePointClient(config);
    return gClient;
};
exports.getClient = getClient;
