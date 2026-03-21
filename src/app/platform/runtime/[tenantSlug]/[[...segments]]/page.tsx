import { headers } from "next/headers";
import Link from "next/link";

import { PublishedRuntime } from "@/components/platform/published-runtime";
import styles from "@/components/platform/platform-shell.module.css";
import { getRuntimeManifest } from "@/lib/platform/service";

export const dynamic = "force-dynamic";

interface RuntimePageProps {
  params: Promise<{ tenantSlug: string; segments?: string[] }>;
}

export default async function PlatformRuntimePage({ params }: RuntimePageProps) {
  const { tenantSlug, segments } = await params;
  const requestHeaders = await headers();
  const manifest = await getRuntimeManifest({ tenantSlug, request: requestHeaders });
  const requestedRoute = segments?.[0];

  if (!manifest) {
    return (
      <div className={styles.runtimeShell}>
        <aside className={styles.runtimeSidebar}>
          <div className={styles.brandBlock}>
            <div className={styles.brandRow}>
              <div className={styles.brandMark}>TA</div>
              <div className={styles.brandCopy}>
                <p className={styles.sidebarTitle}>TeacherActive</p>
                <p className={styles.sidebarSub}>Published runtime</p>
              </div>
            </div>
            <div className={styles.sidebarModePill}>Runtime shell · awaiting publish</div>
          </div>
          <div className={styles.sidebarSection}>
            <p className={styles.sidebarLabel}>Status</p>
            <div className={styles.sidebarMeta}>
              <span>Tenant</span>
              <strong>{tenantSlug}</strong>
            </div>
            <div className={styles.sidebarMeta}>
              <span>Runtime</span>
              <strong>Draft only</strong>
            </div>
          </div>
          <div className={styles.sidebarSection}>
            <p className={styles.sidebarLabel}>Next step</p>
            <div className={styles.sidebarPanel}>
              Publish the current draft from the studio to activate a versioned runtime manifest for this tenant.
            </div>
          </div>
          <div className={styles.sidebarSection}>
            <Link className={styles.secondaryLink} href={`/platform/${tenantSlug}`}>
              Open platform studio
            </Link>
          </div>
          <div className={styles.sidebarFooter}>
            Internal runtime preview.
            <br />
            Publish is required before the generated app becomes active.
          </div>
        </aside>

        <main className={styles.runtimeMain}>
          <header className={styles.runtimeHeader}>
            <div className={styles.headerLead}>
              <p className={styles.eyebrow}>Platform Runtime</p>
              <h1>No published runtime is active</h1>
              <p className={styles.headerCopy}>
                Publish the draft metadata for <strong>{tenantSlug}</strong> from the builder before using the generated runtime.
              </p>
            </div>
            <div className={styles.headerRail}>
              <div className={styles.headerNote}>
                <span>Current state</span>
                <strong>Draft only</strong>
                <p>The runtime route is ready, but it will render the published manifest only after activation.</p>
              </div>
            </div>
          </header>
        </main>
      </div>
    );
  }

  return <PublishedRuntime manifest={manifest} requestedRoute={requestedRoute} tenantSlug={tenantSlug} />;
}
