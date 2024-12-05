import type { BigIntStats, Stats } from "node:fs";
import { SharePointClientConfig } from "./client";
export type ExtStats = (Stats | BigIntStats) & {
    exists: boolean;
};
export type StatOptions = {
    bigint: boolean;
};
export declare const statExt: (path: string, options?: StatOptions) => Promise<ExtStats>;
export declare const downloadFile: (uri: string, to: string, config: SharePointClientConfig) => Promise<void>;
