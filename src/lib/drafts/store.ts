import { promises as fs } from "node:fs";
import path from "node:path";

import { getEnv } from "@/lib/env";
import { SharePointGraphClient } from "@/lib/sharepoint/graph-client";
import type { DraftManifest, DraftRevisionRecord, EditorPresence } from "@/lib/drafts/types";
import type { NormalizedReportSnapshot } from "@/lib/workbook/types";

interface JsonFileStore {
  readJson<T>(filePath: string, fallback: T): Promise<T>;
  writeJson(filePath: string, value: unknown): Promise<void>;
  listChildren(dirPath: string): Promise<string[]>;
}

class LocalJsonFileStore implements JsonFileStore {
  private readonly rootDir = path.resolve(process.cwd(), getEnv().LOCAL_STORAGE_DIR, "report-drafts");

  private resolve(filePath: string): string {
    return path.join(this.rootDir, filePath);
  }

  async readJson<T>(filePath: string, fallback: T): Promise<T> {
    try {
      const raw = await fs.readFile(this.resolve(filePath), "utf8");
      return JSON.parse(raw) as T;
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === "ENOENT") {
        return fallback;
      }

      throw error;
    }
  }

  async writeJson(filePath: string, value: unknown): Promise<void> {
    const target = this.resolve(filePath);
    const tempTarget = `${target}.${process.pid}.${Date.now()}.${Math.random().toString(36).slice(2)}.tmp`;
    await fs.mkdir(path.dirname(target), { recursive: true });
    await fs.writeFile(tempTarget, `${JSON.stringify(value, null, 2)}\n`, "utf8");
    await fs.rename(tempTarget, target);
  }

  async listChildren(dirPath: string): Promise<string[]> {
    const target = this.resolve(dirPath);
    try {
      return await fs.readdir(target);
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === "ENOENT") {
        return [];
      }

      throw error;
    }
  }
}

class SharePointJsonFileStore implements JsonFileStore {
  private readonly client = new SharePointGraphClient();

  async readJson<T>(filePath: string, fallback: T): Promise<T> {
    try {
      return await this.client.getJson<T>(filePath);
    } catch (error) {
      if (error instanceof Error && error.message.includes(" 404")) {
        return fallback;
      }

      throw error;
    }
  }

  async writeJson(filePath: string, value: unknown): Promise<void> {
    await this.client.putJson(filePath, value);
  }

  async listChildren(dirPath: string): Promise<string[]> {
    return this.client.listChildren(dirPath);
  }
}

function getFileStore(): JsonFileStore {
  return getEnv().STORAGE_MODE === "sharepoint" ? new SharePointJsonFileStore() : new LocalJsonFileStore();
}

function getIndexPath(): string {
  return "index.json";
}

function getManifestPath(reportId: string): string {
  return path.posix.join(reportId, "manifest.json");
}

function getCurrentDraftPath(reportId: string): string {
  return path.posix.join(reportId, "current", "draft.json");
}

function getRevisionIndexPath(reportId: string): string {
  return path.posix.join(reportId, "revisions", "index.json");
}

function getRevisionSnapshotPath(reportId: string, revisionId: string): string {
  return path.posix.join(reportId, "revisions", `${revisionId}.json`);
}

function getPresencePath(reportId: string, reportingMonth: string, userId: string): string {
  return path.posix.join(reportId, "presence", reportingMonth, `${userId}.json`);
}

export async function listDraftManifests(): Promise<DraftManifest[]> {
  const store = getFileStore();
  return store.readJson<DraftManifest[]>(getIndexPath(), []);
}

async function saveDraftManifestIndex(manifests: DraftManifest[]): Promise<void> {
  const store = getFileStore();
  await store.writeJson(getIndexPath(), manifests.sort((left, right) => right.updatedAt.localeCompare(left.updatedAt)));
}

export async function getDraftManifest(reportId: string): Promise<DraftManifest | null> {
  const manifests = await listDraftManifests();
  return manifests.find((manifest) => manifest.reportId === reportId) ?? null;
}

export async function upsertDraftManifest(manifest: DraftManifest): Promise<void> {
  const store = getFileStore();
  const manifests = await listDraftManifests();
  const next = [manifest, ...manifests.filter((entry) => entry.reportId !== manifest.reportId)];
  await Promise.all([
    store.writeJson(getManifestPath(manifest.reportId), manifest),
    saveDraftManifestIndex(next),
  ]);
}

export async function getCurrentDraftSnapshot(reportId: string): Promise<NormalizedReportSnapshot | null> {
  const store = getFileStore();
  return store.readJson<NormalizedReportSnapshot | null>(getCurrentDraftPath(reportId), null);
}

export async function writeCurrentDraftSnapshot(reportId: string, snapshot: NormalizedReportSnapshot): Promise<void> {
  const store = getFileStore();
  await store.writeJson(getCurrentDraftPath(reportId), snapshot);
}

export async function listDraftRevisions(reportId: string): Promise<DraftRevisionRecord["meta"][]> {
  const store = getFileStore();
  const revisions = await store.readJson<DraftRevisionRecord["meta"][]>(getRevisionIndexPath(reportId), []);
  return revisions.sort((left, right) => right.revisionNumber - left.revisionNumber);
}

export async function saveDraftRevision(reportId: string, record: DraftRevisionRecord): Promise<void> {
  const store = getFileStore();
  const revisions = await listDraftRevisions(reportId);
  const nextRevisions = [record.meta, ...revisions.filter((entry) => entry.revisionId !== record.meta.revisionId)];

  await Promise.all([
    store.writeJson(getRevisionIndexPath(reportId), nextRevisions),
    store.writeJson(getRevisionSnapshotPath(reportId, record.meta.revisionId), record),
  ]);
}

export async function getDraftRevisionRecord(reportId: string, revisionId: string): Promise<DraftRevisionRecord | null> {
  const store = getFileStore();
  return store.readJson<DraftRevisionRecord | null>(getRevisionSnapshotPath(reportId, revisionId), null);
}

export async function upsertDraftPresence(presence: EditorPresence): Promise<void> {
  const store = getFileStore();
  await store.writeJson(getPresencePath(presence.reportId, presence.reportingMonth, presence.user.id), presence);
}

export async function listDraftPresence(reportId: string, reportingMonth: string): Promise<EditorPresence[]> {
  const store = getFileStore();
  const names = await store.listChildren(path.posix.join(reportId, "presence", reportingMonth));
  const records = await Promise.all(
    names
      .filter((name) => name.endsWith(".json"))
      .map((name) =>
        store.readJson<EditorPresence | null>(path.posix.join(reportId, "presence", reportingMonth, name), null),
      ),
  );

  return records.filter((record): record is EditorPresence => record !== null);
}
