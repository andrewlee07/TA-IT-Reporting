import { headers } from "next/headers";
import { redirect } from "next/navigation";

import { PlatformStudio } from "@/components/platform/platform-studio";
import { getEnv } from "@/lib/env";
import { getPlatformBootstrap } from "@/lib/platform/service";

export const dynamic = "force-dynamic";

interface PlatformTenantPageProps {
  params: Promise<{ tenantSlug: string }>;
  searchParams: Promise<{ workspace?: string | string[] }>;
}

function getSingleValue(value: string | string[] | undefined): string | undefined {
  return Array.isArray(value) ? value[0] : value;
}

export default async function PlatformTenantPage({ params, searchParams }: PlatformTenantPageProps) {
  const env = getEnv();
  if (!env.PLATFORM_WORKSPACE_ENABLED) {
    redirect("/");
  }

  const [{ tenantSlug }, query] = await Promise.all([params, searchParams]);
  const workspace = getSingleValue(query.workspace) ?? "data-model";
  const requestHeaders = await headers();
  let bootstrap;

  try {
    bootstrap = await getPlatformBootstrap({ tenantSlug, request: requestHeaders });
  } catch (error) {
    const status = typeof error === "object" && error !== null && "status" in error ? Number(error.status) : 500;
    if (status === 401 || status === 403) {
      redirect(`/platform/login?tenant=${encodeURIComponent(tenantSlug)}`);
    }
    throw error;
  }

  return <PlatformStudio initialBootstrap={bootstrap} initialWorkspace={workspace} />;
}
