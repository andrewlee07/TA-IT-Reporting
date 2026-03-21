import { describe, expect, it } from "vitest";

import { createStarterManifest } from "@/lib/platform/defaults";
import { validateRecordInput } from "@/lib/platform/records";

describe("validateRecordInput", () => {
  it("calculates computed fields for valid input", () => {
    const manifest = createStarterManifest("teacheractive");
    const objectDefinition = manifest.objects.find((object) => object.key === "booking_request");
    expect(objectDefinition).toBeDefined();

    const result = validateRecordInput({
      objectDefinition: objectDefinition!,
      manifest,
      rawData: {
        request_title: "North Leeds cover",
        school_name: "North Leeds Primary",
        contact_email: "ops@example.com",
        vacancies: 2,
        daily_rate: 180,
        days_requested: 5,
        status: "Open",
      },
      existingRecords: [],
    });

    expect(result.errors).toHaveLength(0);
    expect(result.data.total_value).toBe(1800);
  });

  it("enforces required and unique fields", () => {
    const manifest = createStarterManifest("teacheractive");
    const objectDefinition = manifest.objects.find((object) => object.key === "booking_request");
    expect(objectDefinition).toBeDefined();

    const result = validateRecordInput({
      objectDefinition: objectDefinition!,
      manifest,
      rawData: {
        school_name: "North Leeds Primary",
      },
      existingRecords: [
        {
          id: "rec-1",
          objectKey: "booking_request",
          data: {
            request_title: "Duplicated title",
          },
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
        },
      ],
      currentRecordId: "rec-2",
    });

    expect(result.errors.some((error) => error.includes("required"))).toBe(true);

    const duplicate = validateRecordInput({
      objectDefinition: objectDefinition!,
      manifest,
      rawData: {
        request_title: "Duplicated title",
        school_name: "North Leeds Primary",
        contact_email: "ops@example.com",
        vacancies: 1,
        daily_rate: 180,
        days_requested: 3,
        status: "Open",
      },
      existingRecords: [
        {
          id: "rec-1",
          objectKey: "booking_request",
          data: {
            request_title: "Duplicated title",
          },
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
        },
      ],
    });

    expect(duplicate.errors.some((error) => error.includes("unique"))).toBe(true);
  });
});

