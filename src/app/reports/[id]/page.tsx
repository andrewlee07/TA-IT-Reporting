import { redirect } from "next/navigation";

interface CompatibilityReportRedirectProps {
  params: Promise<{ id: string }>;
  searchParams: Promise<{ month?: string | string[]; page?: string | string[] }>;
}

function getSingleValue(value: string | string[] | undefined): string | undefined {
  return Array.isArray(value) ? value[0] : value;
}

export default async function CompatibilityReportRedirect({ params, searchParams }: CompatibilityReportRedirectProps) {
  const { id } = await params;
  const query = await searchParams;
  const url = new URLSearchParams();

  url.set("report", id);

  const month = getSingleValue(query.month);
  const page = getSingleValue(query.page);

  if (month) {
    url.set("month", month);
  }

  if (page) {
    url.set("page", page);
  }

  redirect(`/?${url.toString()}`);
}
