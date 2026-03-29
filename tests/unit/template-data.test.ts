import { describe, expect, it } from "vitest";

import { buildTemplateData } from "@/lib/report/template-data";
import { createBlankSnapshot } from "@/lib/workbook/blank-snapshot";

describe("buildTemplateData", () => {
  it("backfills runtime-critical rows for the active month when collections are empty", () => {
    const snapshot = createBlankSnapshot("2026-06", "Blank Report");

    snapshot.supportOperations = [];
    snapshot.securityPatching = [];
    snapshot.assetsLifecycle = [];
    snapshot.changeRelease = [];
    snapshot.devDelivery = [];
    snapshot.derivedNetworkMetrics = [];

    const data = buildTemplateData(snapshot, "2026-06");

    expect(data.support.find((row) => row.Month === "2026-06")).toMatchObject({
      ResolutionSLA: "95.0%",
    });
    expect(data.security.find((row) => row.Month === "2026-06")).toMatchObject({
      CritVulns: 0,
    });
    expect(data.change.find((row) => row.Month === "2026-06")).toMatchObject({
      SuccessRate: "89.0%",
    });
    expect(data.dev.find((row) => row.Month === "2026-06")).toMatchObject({
      BacklogEnd: 0,
    });
    expect(
      data.assets
        .filter((row) => row.Month === "2026-06")
        .map((row) => row.AssetType)
        .sort(),
    ).toEqual(["Laptop", "Mobile", "Monitor"]);
    expect(data.derivedNetwork.find((row) => row.Month === "2026-06")).toMatchObject({
      WorstOffice: "",
      WorstAvailability: "0.0%",
    });
  });
});
