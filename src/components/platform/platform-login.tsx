"use client";

import Link from "next/link";
import { useMemo, useState, useTransition } from "react";

import type { PlatformSessionSummary } from "@/lib/platform/types";

import styles from "./platform-shell.module.css";

interface PlatformLoginProps {
  initialSession: PlatformSessionSummary;
  inviteToken?: string;
  requestedTenantSlug?: string;
}

async function fetchJson<T>(input: RequestInfo, init?: RequestInit): Promise<T> {
  const response = await fetch(input, init);
  const payload = (await response.json()) as T & { error?: string };

  if (!response.ok) {
    throw new Error(payload.error ?? "Request failed.");
  }

  return payload;
}

export function PlatformLogin({ initialSession, inviteToken = "", requestedTenantSlug }: PlatformLoginProps) {
  const [session, setSession] = useState(initialSession);
  const [token, setToken] = useState(inviteToken);
  const [email, setEmail] = useState(initialSession.actor?.email ?? "");
  const [name, setName] = useState(initialSession.actor?.name ?? "");
  const [message, setMessage] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isPending, startTransition] = useTransition();

  const orderedMemberships = useMemo(
    () => [...session.memberships].sort((left, right) => left.tenantName.localeCompare(right.tenantName)),
    [session.memberships],
  );

  async function handleAcceptInvite(): Promise<void> {
    const payload = await fetchJson<{ session: PlatformSessionSummary; tenantSlug: string }>("/api/platform/auth/session", {
      method: "POST",
      headers: {
        "content-type": "application/json",
      },
      body: JSON.stringify({
        token,
        email,
        name,
      }),
    });

    setSession(payload.session);
    setMessage("Invite accepted. Opening the tenant workspace.");
    window.location.href = `/platform/${requestedTenantSlug ?? payload.tenantSlug}`;
  }

  async function handleLogout(): Promise<void> {
    await fetchJson("/api/platform/auth/session", {
      method: "DELETE",
    });
    setSession({
      actor: null,
      memberships: [],
      source: "none",
    });
    setMessage("Signed out.");
  }

  return (
    <div className={styles.runtimeShell}>
      <aside className={styles.runtimeSidebar}>
        <div className={styles.brandBlock}>
          <div className={styles.brandRow}>
            <div className={styles.brandMark}>TA</div>
            <div className={styles.brandCopy}>
              <p className={styles.sidebarTitle}>TeacherActive</p>
              <p className={styles.sidebarSub}>Platform access</p>
            </div>
          </div>
          <div className={styles.sidebarModePill}>Invite-based beta access</div>
        </div>
        <div className={styles.sidebarSection}>
          <p className={styles.sidebarLabel}>Session</p>
          <div className={styles.sidebarMeta}>
            <span>Source</span>
            <strong>{session.source}</strong>
          </div>
          <div className={styles.sidebarMeta}>
            <span>Memberships</span>
            <strong>{session.memberships.length}</strong>
          </div>
        </div>
        <div className={styles.sidebarSection}>
          <p className={styles.sidebarLabel}>Beta model</p>
          <div className={styles.sidebarPanel}>
            Manual tenant provisioning plus invite-only builder access. Published runtime stays isolated from draft preview.
          </div>
        </div>
      </aside>

      <main className={styles.runtimeMain}>
        <header className={styles.runtimeHeader}>
          <div className={styles.headerLead}>
            <p className={styles.eyebrow}>Adaptive platform beta</p>
            <h1>Builder access</h1>
            <p className={styles.headerCopy}>
              Accept a tenant invite to create a signed builder session, or pick one of your current tenant workspaces.
            </p>
          </div>
        </header>

        {message ? <div className={styles.successBanner}>{message}</div> : null}
        {error ? <div className={styles.errorBanner}>{error}</div> : null}

        <div className={styles.workspaceGrid}>
          <section className={styles.panel}>
            <div className={styles.sectionHeader}>
              <div>
                <p className={styles.cardEyebrow}>Invite</p>
                <h2>Accept access</h2>
              </div>
            </div>
            <div className={styles.formGrid}>
              <label className={styles.formFieldSpan}>
                <span>Invite token</span>
                <input className={styles.input} onChange={(event) => setToken(event.target.value)} value={token} />
              </label>
              <label className={styles.formField}>
                <span>Email</span>
                <input className={styles.input} onChange={(event) => setEmail(event.target.value)} value={email} />
              </label>
              <label className={styles.formField}>
                <span>Name</span>
                <input className={styles.input} onChange={(event) => setName(event.target.value)} value={name} />
              </label>
            </div>
            <div className={styles.actionsRow}>
              <button
                className={styles.primaryButton}
                disabled={isPending || !token.trim() || !email.trim() || !name.trim()}
                onClick={() =>
                  startTransition(() => {
                    void handleAcceptInvite().catch((caughtError) => {
                      setError(caughtError instanceof Error ? caughtError.message : "Failed to accept invite.");
                    });
                  })
                }
                type="button"
              >
                {isPending ? "Opening..." : "Accept invite"}
              </button>
            </div>
          </section>

          <section className={styles.panelWide}>
            <div className={styles.sectionHeader}>
              <div>
                <p className={styles.cardEyebrow}>Your tenants</p>
                <h2>Select a workspace</h2>
              </div>
              {session.actor ? (
                <button
                  className={styles.secondaryButton}
                  onClick={() =>
                    startTransition(() => {
                      void handleLogout().catch((caughtError) => {
                        setError(caughtError instanceof Error ? caughtError.message : "Failed to end session.");
                      });
                    })
                  }
                  type="button"
                >
                  Sign out
                </button>
              ) : null}
            </div>

            {orderedMemberships.length === 0 ? (
              <div className={styles.emptyState}>
                No tenant memberships yet. Create an invite from a super-admin tenant, then return here to accept it.
              </div>
            ) : (
              <div className={styles.tileGrid}>
                {orderedMemberships.map((membership) => (
                  <article className={styles.metricCard} key={`${membership.tenantId}-${membership.role}`}>
                    <span>{membership.role}</span>
                    <strong>{membership.tenantName}</strong>
                    <p className={styles.metricMeta}>{membership.tenantSlug}</p>
                    <Link className={styles.secondaryLink} href={`/platform/${membership.tenantSlug}`}>
                      Open studio
                    </Link>
                  </article>
                ))}
              </div>
            )}
          </section>
        </div>
      </main>
    </div>
  );
}
