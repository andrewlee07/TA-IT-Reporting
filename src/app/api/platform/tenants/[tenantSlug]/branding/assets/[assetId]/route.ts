import { NextResponse } from "next/server";

import { getObjectStorage } from "@/lib/storage";
import { getPublicRuntimeManifest } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string; assetId: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(_: Request, { params }: RouteProps) {
  const { tenantSlug, assetId } = await params;
  const manifest = await getPublicRuntimeManifest({ tenantSlug });
  const asset = manifest?.branding.assets.find((candidate) => candidate.id === assetId);

  if (!asset) {
    return NextResponse.json({ error: "Asset not found." }, { status: 404 });
  }

  const buffer = await getObjectStorage().getBuffer(asset.storageKey);
  return new NextResponse(new Uint8Array(buffer), {
    status: 200,
    headers: {
      "content-type": asset.contentType,
      "cache-control": "public, max-age=300",
    },
  });
}
