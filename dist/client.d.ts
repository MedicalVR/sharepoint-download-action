import "isomorphic-fetch";
import { Client } from "@microsoft/microsoft-graph-client";
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
export declare class SharePointClient {
    private readonly config;
    private clientSecretCredential;
    private graph;
    constructor(config: SharePointClientConfig);
    getGraph(): Client;
    getGroup(name: string): Promise<Group | undefined>;
    getDriveItem(group: Group, path: string): Promise<DriveItem | undefined>;
    downloadItem(group: Group, path: string, targetPath: string): Promise<void>;
    uploadItem(group: Group, path: string, targetPath: string, callback: (progress: number) => void): Promise<DriveItem>;
}
export declare const getClient: (config: SharePointClientConfig) => SharePointClient;
