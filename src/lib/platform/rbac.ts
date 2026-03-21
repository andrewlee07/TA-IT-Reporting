import type { PlatformActor, PlatformRole } from "@/lib/platform/types";

const roleWeight: Record<PlatformRole, number> = {
  USER: 1,
  BUILDER_ADMIN: 2,
  SUPER_ADMIN: 3,
};

export function hasRole(actor: PlatformActor, minimumRole: PlatformRole): boolean {
  return roleWeight[actor.role] >= roleWeight[minimumRole];
}

export function assertRole(actor: PlatformActor, minimumRole: PlatformRole, message?: string): void {
  if (!hasRole(actor, minimumRole)) {
    throw new Error(message ?? `Requires ${minimumRole} privileges.`);
  }
}

