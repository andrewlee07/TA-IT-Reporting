import { headers } from "next/headers";
import { notFound } from "next/navigation";

import { PlatformPublicForm } from "@/components/platform/platform-public-form";
import { getDraftPreviewManifest } from "@/lib/platform/service";

export const dynamic = "force-dynamic";

interface PreviewFormPageProps {
  params: Promise<{ tenantSlug: string; formKey: string }>;
}

export default async function PlatformPreviewFormPage({ params }: PreviewFormPageProps) {
  const { tenantSlug, formKey } = await params;
  const requestHeaders = await headers();
  const manifest = await getDraftPreviewManifest({
    tenantSlug,
    request: requestHeaders,
    route: formKey,
  });
  const form = manifest.forms.find((candidate) => candidate.key === formKey || candidate.route === formKey);

  if (!form) {
    notFound();
  }

  return <PlatformPublicForm form={form} manifest={manifest} mode="draft-preview" tenantSlug={tenantSlug} />;
}
