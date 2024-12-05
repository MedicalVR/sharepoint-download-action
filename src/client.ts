import "isomorphic-fetch";
import { ClientSecretCredential } from "@azure/identity";
import {
  Client,
  LargeFileUploadSession,
  LargeFileUploadTask,
  LargeFileUploadTaskOptions,
  Range,
  ResponseType,
  StreamUpload,
  UploadEventHandlers,
} from "@microsoft/microsoft-graph-client";
import { createWriteStream, createReadStream } from "node:fs";
import * as Path from "node:path";
import { promisify } from "node:util";
import { stat } from "node:fs/promises";
import * as Stream from "node:stream";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";

const Finished = promisify(Stream.finished);

const TEMP_FILE_CONTENTS = "Temporary contents...";

export type DriveItem = {
  id: string;
  path: string;
};

export type Group = {
  id: string;
  name: string;
};

export type GroupDrive = {
  id: string;
  groupId: string;
};

export type SharePointClientConfig = {
  tenantId: string;
  clientId: string;
  clientSecret: string;
};

export class SharePointClient {
  private clientSecretCredential: ClientSecretCredential;
  private graph!: Client;

  constructor(private readonly config: SharePointClientConfig) {
    this.clientSecretCredential = new ClientSecretCredential(
      this.config.tenantId,
      this.config.clientId,
      this.config.clientSecret,
    );
  }

  getGraph(): Client {
    if (this.graph) {
      return this.graph;
    }

    const authProvider = new TokenCredentialAuthenticationProvider(this.clientSecretCredential, {
      scopes: ["https://graph.microsoft.com/.default"],
    });

    this.graph = Client.initWithMiddleware({ authProvider });
    return this.graph;
  }

  async getGroup(name: string): Promise<Group | undefined> {
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

  async getDriveItem(group: Group, path: string): Promise<DriveItem | undefined> {
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

  async downloadItem(group: Group, path: string, targetPath: string) {
    const stream = await this.getGraph()
      .api(`/groups/${group.id}/drive/root:/${path}:/content`)
      .responseType(ResponseType.BLOB)
      .get();
    if (stream) {
      const writer = createWriteStream(targetPath);
      Stream.Readable.fromWeb(stream instanceof Blob ? stream.stream() : stream).pipe(writer);
      await Finished(writer);
    }
  }

  async uploadItem(
    group: Group,
    path: string,
    targetPath: string,
    callback: (progress: number) => void,
  ): Promise<DriveItem> {
    const pathDetails = Path.parse(targetPath);
    let item: DriveItem | undefined;
    try {
      item = await this.getDriveItem(group, targetPath);
      // eslint-disable-next-line @typescript-eslint/no-unused-vars
    } catch (ex: any) {
      // Do nothing.
    }

    const graph = this.getGraph();
    if (!item) {
      // First we need to create a new (placeholder) item that's located at the same path.
      const parent = await this.getDriveItem(group, pathDetails.dir);
      if (!parent) {
        throw new Error(
          `Directory not found on SharePoint of group ${group.name}: ${pathDetails.dir}`,
        );
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

    const stats = await stat(path);
    const totalSize = stats.size;

    const progress = (range?: Range) => {
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

    const uploadEventHandlers: UploadEventHandlers = {
      progress,
    };

    const options: LargeFileUploadTaskOptions = {
      rangeSize: 1024 * 1024,
      uploadEventHandlers,
    };

    // Create upload session for SharePoint upload.
    const payload = {
      item: {
        "@microsoft.graph.conflictBehavior": "replace",
      },
    };

    const reader = createReadStream(path);
    const uploadSession: LargeFileUploadSession = await LargeFileUploadTask.createUploadSession(
      graph,
      `https://graph.microsoft.com/v1.0/groups/${group.id}/drive/items/${item.id}/createuploadsession`,
      payload,
    );
    const fileObject = new StreamUpload(reader, pathDetails.base, totalSize);
    const task = new LargeFileUploadTask(graph, fileObject, uploadSession, options);
    await task.upload();

    return item;
  }
}

let gClient: SharePointClient | undefined = undefined;

export const getClient = (config: SharePointClientConfig): SharePointClient => {
  if (gClient) {
    return gClient;
  }

  gClient = new SharePointClient(config);
  return gClient;
};
