import { headers } from "next/headers";

import { PlatformLogin } from "@/components/platform/platform-login";
import { getEnv } from "@/lib/env";
import { getCurrentPlatformSession } from "@/lib/platform/service";

export const dynamic = "force-dynamic";

interface PlatformLoginPageProps {
  searchParams: Promise<{ invite?: string | string[]; tenant?: string | string[] }>;
}

function singleValue(value: string | string[] | undefined): string | undefined {
  return Array.isArray(value) ? value[0] : value;
}

export default async function PlatformLoginPage({ searchParams }: PlatformLoginPageProps) {
  if (!getEnv().PLATFORM_WORKSPACE_ENABLED) {
    return null;
  }

  const requestHeaders = await headers();
  const session = await getCurrentPlatformSession({ request: requestHeaders });
  const query = await searchParams;

  return (
    <PlatformLogin
      initialSession={session}
      inviteToken={singleValue(query.invite)}
      requestedTenantSlug={singleValue(query.tenant)}
    />
  );
}
