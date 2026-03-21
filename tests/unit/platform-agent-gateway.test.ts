import { describe, expect, it } from "vitest";

import { prepareAgentInvocation } from "@/lib/platform/agent-gateway";
import { createStarterManifest } from "@/lib/platform/defaults";

describe("prepareAgentInvocation", () => {
  it("masks protected fields before handing records to an agent", () => {
    const manifest = createStarterManifest("teacheractive");

    const invocation = prepareAgentInvocation({
      manifest,
      agentId: "booking_analyst",
      objectKey: "booking_request",
      records: [
        {
          id: "rec-1",
          objectKey: "booking_request",
          data: {
            request_title: "North Leeds cover",
            school_name: "North Leeds Primary",
            contact_email: "ops@example.com",
            vacancies: 2,
          },
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
        },
      ],
    });

    expect(invocation.metadata.masked).toBe(true);
    expect(invocation.inputRecords[0]?.contact_email).not.toBe("ops@example.com");
    expect(String(invocation.inputRecords[0]?.school_name)).toContain("*");
  });

  it("rejects providers without zero-retention support when the agent requires it", () => {
    const manifest = createStarterManifest("teacheractive");
    manifest.modelProviders[0] = {
      ...manifest.modelProviders[0],
      supportsZeroRetention: false,
    };

    expect(() =>
      prepareAgentInvocation({
        manifest,
        agentId: "booking_analyst",
        objectKey: "booking_request",
        records: [],
      }),
    ).toThrow(/zero retention/i);
  });
});

