import { NextResponse } from "next/server";

import { jsonError } from "@/lib/platform/http";
import { uploadBrandAsset } from "@/lib/platform/service";

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const formData = await request.formData();
    const file = formData.get("file");
    const kind = String(formData.get("kind") ?? "reference") as "logo" | "icon" | "brand_book" | "reference";
    const label = String(formData.get("label") ?? "Brand asset");

    if (!(file instanceof File)) {
      return NextResponse.json({ error: "A file is required." }, { status: 400 });
    }

    const asset = await uploadBrandAsset({
      tenantSlug,
      kind,
      label,
      fileName: file.name,
      contentType: file.type || "application/octet-stream",
      buffer: Buffer.from(await file.arrayBuffer()),
      request,
    });

    return NextResponse.json({ asset });
  } catch (error) {
    return jsonError(error, "Failed to upload branding asset.", 400);
  }
}
