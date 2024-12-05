import type { BigIntStats, Stats } from "node:fs";
import { getClient, Group, SharePointClient, SharePointClientConfig } from "./client";
import { getInput, info } from "@actions/core";
import { mkdir, stat } from "node:fs/promises";
import * as Path from "node:path";

type DecomposedSharePointURI = {
  group: Group;
  pathParts: string[];
};

export type ExtStats = (Stats | BigIntStats) & {
  exists: boolean;
};

export type StatOptions = {
  bigint: boolean;
};

const groupsMap: Map<string, Group> = new Map();

const decomposeSharePointURI = async (
  uri: string,
  client: SharePointClient,
): Promise<DecomposedSharePointURI> => {
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

export const statExt = async (path: string, options?: StatOptions): Promise<ExtStats> => {
  let s: Stats | BigIntStats;
  try {
    s = await stat(path, options);
    return Object.assign(s, { exists: true });
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
  } catch (ex: any) {
    // no-op.
  }

  return { ...({} as Stats), exists: false };
};

export const downloadFile = async (uri: string, to: string, config: SharePointClientConfig) => {
  if (!uri.startsWith("sharepoint:")) {
    throw new Error(`Invalid URI: ${uri}`);
  }

  const client = getClient(config);
  const { group, pathParts } = await decomposeSharePointURI(uri, client);
  const fileName = pathParts[pathParts.length - 1];
  if (!fileName) {
    throw new Error(`Invalid URI: ${uri}`);
  }

  info(`Downloading '${uri}' to '${to}'...`);
  const itemPath = pathParts.join("/");
  // Try once to create the directory if it doesn't exist yet. No mkdirp.
  const stat = await statExt(to);
  if (!stat.exists) {
    await mkdir(to);
  }

  try {
    // This fetch may yield a 404 error and is cheaper than a full download.
    await client.getDriveItem(group, itemPath);
    await client.downloadItem(group, itemPath, Path.join(to, fileName));
  } catch (ex: any) {
    throw new Error(`Could not download '${itemPath}' with error ${ex.message}`);
  }

  info("done!");
};

(async function () {
  const config: SharePointClientConfig = {
    clientId: getInput("azure-client-id"),
    clientSecret: getInput("azure-client-secret"),
    tenantId: getInput("azure-tenant-id"),
  };

  await downloadFile(getInput("uri"), getInput("target"), config);
})();
