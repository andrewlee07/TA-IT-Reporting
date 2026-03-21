import { notFound } from "next/navigation";

import { PlatformPublicForm } from "@/components/platform/platform-public-form";
import { getPublicRuntimeManifest } from "@/lib/platform/service";

export const dynamic = "force-dynamic";

interface FormPageProps {
  params: Promise<{ tenantSlug: string; formKey: string }>;
}

export default async function PlatformPublicFormPage({ params }: FormPageProps) {
  const { tenantSlug, formKey } = await params;
  const manifest = await getPublicRuntimeManifest({ tenantSlug });
  const form = manifest?.forms.find((candidate) => candidate.key === formKey || candidate.route === formKey);

  if (!manifest || !form) {
    notFound();
  }

  return <PlatformPublicForm form={form} manifest={manifest} tenantSlug={tenantSlug} />;
}
