import { cloneManifest } from "@/lib/platform/types";
import { syncPublishedManifestToGit, type GitSyncResult } from "@/lib/platform/git-sync";
import type { PlatformManifest } from "@/lib/platform/types";

export interface PublishArtifacts extends GitSyncResult {
  manifest: PlatformManifest;
}

export async function buildPublishedArtifacts(input: {
  manifest: PlatformManifest;
  tenantSlug: string;
  environmentSlug: string;
  versionNumber: number;
  versionId: string;
}): Promise<PublishArtifacts> {
  const manifest = cloneManifest(input.manifest);
  manifest.metadata = {
    ...manifest.metadata,
    publishedAt: new Date().toISOString(),
    publishedVersionId: input.versionId,
  };

  const gitResult = await syncPublishedManifestToGit({
    manifest,
    tenantSlug: input.tenantSlug,
    environmentSlug: input.environmentSlug,
    versionNumber: input.versionNumber,
  });

  return {
    manifest,
    manifestPath: gitResult.manifestPath,
    gitCommitSha: gitResult.gitCommitSha,
  };
}

