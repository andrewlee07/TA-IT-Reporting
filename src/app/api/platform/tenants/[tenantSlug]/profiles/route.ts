import { NextResponse } from "next/server";
import { z } from "zod";

import { jsonError } from "@/lib/platform/http";
import { saveProfileConfiguration } from "@/lib/platform/service";

const profileConfigSchema = z.object({
  pageTitle: z.string().trim().min(1),
  visibleFieldKeys: z.array(z.string().trim().min(1)),
  profilePageKey: z.string().trim().min(1).optional(),
  settingsPageKey: z.string().trim().min(1).optional(),
});

interface RouteProps {
  params: Promise<{ tenantSlug: string }>;
}

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(request: Request, { params }: RouteProps) {
  try {
    const { tenantSlug } = await params;
    const payload = profileConfigSchema.parse(await request.json());
    const profiles = await saveProfileConfiguration({
      tenantSlug,
      ...payload,
      request,
    });

    return NextResponse.json({ profiles });
  } catch (error) {
    return jsonError(error, "Failed to save profile configuration.", 400);
  }
}
