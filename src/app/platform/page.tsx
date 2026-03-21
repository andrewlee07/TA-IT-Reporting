import { headers } from "next/headers";
import { redirect } from "next/navigation";

import { getEnv } from "@/lib/env";
import { getCurrentPlatformSession } from "@/lib/platform/service";

export const dynamic = "force-dynamic";

export default async function PlatformIndexPage() {
  const env = getEnv();

  if (!env.PLATFORM_WORKSPACE_ENABLED) {
    redirect("/");
  }

  const requestHeaders = await headers();
  const session = await getCurrentPlatformSession({ request: requestHeaders });
  const firstMembership = session.memberships[0];

  if (firstMembership) {
    redirect(`/platform/${firstMembership.tenantSlug}`);
  }

  if (env.PLATFORM_LOCAL_DEV_MODE) {
    redirect(`/platform/${env.PLATFORM_DEFAULT_TENANT_SLUG}`);
  }

  redirect("/platform/login");
}
