import { promises as fs } from "node:fs";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

import { getEnv } from "@/lib/env";
import type { PlatformManifest } from "@/lib/platform/types";

const execFileAsync = promisify(execFile);

export interface GitSyncResult {
  manifestPath: string;
  gitCommitSha: string | null;
}

async function writeManifestFile(filePath: string, manifest: PlatformManifest): Promise<void> {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, `${JSON.stringify(manifest, null, 2)}\n`, "utf8");
}

export async function syncPublishedManifestToGit(input: {
  manifest: PlatformManifest;
  tenantSlug: string;
  environmentSlug: string;
  versionNumber: number;
}): Promise<GitSyncResult> {
  const env = getEnv();
  const versionDir = path.resolve(
    process.cwd(),
    env.PLATFORM_GIT_OUTPUT_DIR,
    input.tenantSlug,
    input.environmentSlug,
    `v${input.versionNumber}`,
  );
  const currentDir = path.resolve(process.cwd(), env.PLATFORM_GIT_OUTPUT_DIR, input.tenantSlug, input.environmentSlug, "current");
  const manifestPath = path.join(versionDir, "manifest.json");
  const currentManifestPath = path.join(currentDir, "manifest.json");

  await writeManifestFile(manifestPath, input.manifest);
  await writeManifestFile(currentManifestPath, input.manifest);

  if (!env.PLATFORM_GIT_AUTO_COMMIT) {
    return {
      manifestPath: path.relative(process.cwd(), manifestPath),
      gitCommitSha: null,
    };
  }

  try {
    const relativeOutputDir = path.relative(process.cwd(), path.resolve(process.cwd(), env.PLATFORM_GIT_OUTPUT_DIR));
    await execFileAsync("git", ["add", relativeOutputDir], {
      cwd: process.cwd(),
      env: {
        ...process.env,
        GIT_AUTHOR_NAME: env.PLATFORM_GIT_AUTHOR_NAME,
        GIT_AUTHOR_EMAIL: env.PLATFORM_GIT_AUTHOR_EMAIL,
        GIT_COMMITTER_NAME: env.PLATFORM_GIT_AUTHOR_NAME,
        GIT_COMMITTER_EMAIL: env.PLATFORM_GIT_AUTHOR_EMAIL,
      },
    });

    await execFileAsync(
      "git",
      [
        "commit",
        "-m",
        `platform(${input.tenantSlug}/${input.environmentSlug}): publish v${input.versionNumber}`,
      ],
      {
        cwd: process.cwd(),
        env: {
          ...process.env,
          GIT_AUTHOR_NAME: env.PLATFORM_GIT_AUTHOR_NAME,
          GIT_AUTHOR_EMAIL: env.PLATFORM_GIT_AUTHOR_EMAIL,
          GIT_COMMITTER_NAME: env.PLATFORM_GIT_AUTHOR_NAME,
          GIT_COMMITTER_EMAIL: env.PLATFORM_GIT_AUTHOR_EMAIL,
        },
      },
    );

    const { stdout } = await execFileAsync("git", ["rev-parse", "HEAD"], { cwd: process.cwd() });
    return {
      manifestPath: path.relative(process.cwd(), manifestPath),
      gitCommitSha: stdout.trim(),
    };
  } catch {
    return {
      manifestPath: path.relative(process.cwd(), manifestPath),
      gitCommitSha: null,
    };
  }
}

