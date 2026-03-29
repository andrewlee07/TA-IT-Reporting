import { redirect } from "next/navigation";

import { requireAppUser } from "@/lib/auth/app-user";
import { EDITOR_SECTIONS } from "@/lib/drafts/types";
import { getStoredReport } from "@/lib/reports/service";

interface ReportEditorPageProps {
  params: Promise<{ id: string }>;
  searchParams: Promise<{ month?: string | string[]; section?: string | string[] }>;
}

function getSingleValue(value: string | string[] | undefined): string | undefined {
  return Array.isArray(value) ? value[0] : value;
}

export default async function ReportEditorPage({ params, searchParams }: ReportEditorPageProps) {
  await requireAppUser();
  const { id } = await params;
  const query = await searchParams;
  const report = await getStoredReport(id);

  if (!report) {
    redirect("/");
  }

  const month = getSingleValue(query.month);
  const selectedMonth = month && report.availableMonths.includes(month) ? month : report.currentMonth;
  const requestedSection = getSingleValue(query.section);
  const selectedSection = EDITOR_SECTIONS.includes(requestedSection as (typeof EDITOR_SECTIONS)[number])
    ? (requestedSection as (typeof EDITOR_SECTIONS)[number])
    : EDITOR_SECTIONS[0];
  const paramsToRedirect = new URLSearchParams({
    report: id,
    month: selectedMonth,
    page: "p-data",
    tab: selectedSection,
  });

  redirect(`/?${paramsToRedirect.toString()}`);
}
