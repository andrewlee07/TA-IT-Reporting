import { headers } from "next/headers";
import Link from "next/link";
import { redirect } from "next/navigation";

import { PublishedRuntime } from "@/components/platform/published-runtime";
import styles from "@/components/platform/platform-shell.module.css";
import { getDraftPreviewManifest } from "@/lib/platform/service";

export const dynamic = "force-dynamic";

interface PreviewPageProps {
  params: Promise<{ tenantSlug: string; segments?: string[] }>;
}

export default async function PlatformDraftPreviewPage({ params }: PreviewPageProps) {
  const { tenantSlug, segments } = await params;
  const requestHeaders = await headers();
  const requestedRoute = segments?.[0];

  let manifest;
  try {
    manifest = await getDraftPreviewManifest({
      tenantSlug,
      route: requestedRoute,
      request: requestHeaders,
    });
  } catch (error) {
    const status = typeof error === "object" && error !== null && "status" in error ? Number(error.status) : 500;
    if (status === 401 || status === 403) {
      redirect(`/platform/login?tenant=${encodeURIComponent(tenantSlug)}`);
    }
    throw error;
  }

  return (
    <div className={styles.previewShell}>
      <div className={styles.previewBanner}>
        <div>
          <p className={styles.cardEyebrow}>Draft preview</p>
          <h2>Builder-only view of the current draft manifest</h2>
        </div>
        <div className={styles.inlineList}>
          <Link className={styles.secondaryLink} href={`/platform/${tenantSlug}`}>
            Back to studio
          </Link>
          <Link className={styles.secondaryLink} href={`/platform/runtime/${tenantSlug}/${requestedRoute ?? ""}`}>
            Open published runtime
          </Link>
        </div>
      </div>
      <PublishedRuntime manifest={manifest} requestedRoute={requestedRoute} tenantSlug={tenantSlug} mode="draft-preview" />
    </div>
  );
}
