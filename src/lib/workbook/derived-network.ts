import type {
  DerivedNetworkMetricRow,
  NormalizedReportSnapshot,
  OfficeNetworkAvailabilityRow,
  ServiceAvailabilityRow,
} from "@/lib/workbook/types";

export const NETWORK_SERVICE_NAME = "Network";
export const NETWORK_TARGET_PCT = 99.9;

function roundTo(value: number, digits: number): number {
  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
}

export function deriveNetworkMetrics(snapshot: Pick<NormalizedReportSnapshot, "periods" | "officeLocations" | "officeNetworkAvailability">): DerivedNetworkMetricRow[] {
  const inScopeOffices = snapshot.officeLocations.filter((office) => office.inScope);
  const officeNames = new Set(inScopeOffices.map((office) => office.officeName));

  return snapshot.periods.map((period) => {
    const monthRows = snapshot.officeNetworkAvailability.filter(
      (row) => row.reportingMonth === period.reportingMonth && officeNames.has(row.officeName),
    );
    const byOffice = new Map<string, OfficeNetworkAvailabilityRow[]>();

    for (const row of monthRows) {
      const rows = byOffice.get(row.officeName) ?? [];
      rows.push(row);
      byOffice.set(row.officeName, rows);
    }

    const rows = inScopeOffices
      .map((office) => byOffice.get(office.officeName)?.[0])
      .filter((row): row is OfficeNetworkAvailabilityRow => Boolean(row));

    if (rows.length === 0) {
      return {
        reportingMonth: period.reportingMonth,
        availabilityPct: 0,
        outageMinutes: 0,
        majorIncidents: 0,
        perfectOffices: 0,
        below99_9Offices: 0,
        below99Offices: 0,
        worstOffice: null,
        worstAvailabilityPct: null,
      };
    }

    const availabilityPct = roundTo(rows.reduce((total, row) => total + row.availabilityPct, 0) / rows.length, 2);
    const outageMinutes = rows.reduce((total, row) => total + row.outageMinutes, 0);
    const majorIncidents = rows.reduce((total, row) => total + row.majorIncidents, 0);
    const perfectOffices = rows.filter((row) => row.availabilityPct === 100).length;
    const below99_9Offices = rows.filter((row) => row.availabilityPct < 99.9).length;
    const below99Offices = rows.filter((row) => row.availabilityPct < 99).length;
    const worst = rows.reduce((lowest, row) => (lowest.availabilityPct <= row.availabilityPct ? lowest : row), rows[0]);

    return {
      reportingMonth: period.reportingMonth,
      availabilityPct,
      outageMinutes,
      majorIncidents,
      perfectOffices,
      below99_9Offices,
      below99Offices,
      worstOffice: worst.officeName,
      worstAvailabilityPct: worst.availabilityPct,
    };
  });
}

export function buildDerivedNetworkServiceRows(metrics: DerivedNetworkMetricRow[]): ServiceAvailabilityRow[] {
  return metrics.map((metric) => ({
    reportingMonth: metric.reportingMonth,
    serviceName: NETWORK_SERVICE_NAME,
    serviceType: NETWORK_SERVICE_NAME,
    availabilityPct: metric.availabilityPct,
    targetPct: NETWORK_TARGET_PCT,
    outageMinutes: metric.outageMinutes,
    majorIncidents: metric.majorIncidents,
    backupSuccessPct: null,
    restoreTestStatus: "",
    commentary: metric.worstOffice
      ? `${metric.perfectOffices} offices at 100%. Worst office: ${metric.worstOffice} (${metric.worstAvailabilityPct?.toFixed(2)}%).`
      : "No in-scope office network data available.",
  }));
}
