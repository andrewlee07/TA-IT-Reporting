import { Pool } from "pg";
import { PrismaPg } from "@prisma/adapter-pg";

import { PrismaClient } from "@/generated/prisma/client";
import { requireDatabaseUrl } from "@/lib/env";

const globalForPrisma = globalThis as unknown as {
  prisma?: PrismaClient;
};

function hasRequiredDelegates(client: PrismaClient): boolean {
  const reportAnnotationStateDelegate = (client as PrismaClient & {
    reportAnnotationState?: { findUnique?: unknown; upsert?: unknown };
  }).reportAnnotationState;

  return (
    typeof client.report?.findMany === "function" &&
    typeof reportAnnotationStateDelegate?.findUnique === "function" &&
    typeof reportAnnotationStateDelegate?.upsert === "function"
  );
}

export function getPrisma(): PrismaClient {
  if (globalForPrisma.prisma) {
    if (hasRequiredDelegates(globalForPrisma.prisma)) {
      return globalForPrisma.prisma;
    }

    void globalForPrisma.prisma.$disconnect().catch(() => undefined);
    delete globalForPrisma.prisma;
  }

  const adapter = new PrismaPg(
    new Pool({
      connectionString: requireDatabaseUrl(),
    }),
  );

  const prisma = new PrismaClient({
    adapter,
    log: process.env.NODE_ENV === "development" ? ["warn", "error"] : ["error"],
  });

  if (process.env.NODE_ENV !== "production") {
    globalForPrisma.prisma = prisma;
  }

  return prisma;
}
