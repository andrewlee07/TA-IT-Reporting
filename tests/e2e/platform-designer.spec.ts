import { expect, test } from "@playwright/test";

test("platform pages workspace exposes the designer canvas and draft preview", async ({ page, context }) => {
  await page.goto("/platform/teacheractive?workspace=pages");

  await expect(page.getByRole("heading", { level: 1, name: "Pages" })).toBeVisible();
  await expect(page.getByText("Hybrid designer")).toBeVisible();
  await expect(page.getByRole("button", { name: "Desktop" })).toBeVisible();
  await expect(page.getByRole("button", { name: "Undo" })).toBeVisible();
  await expect(page.getByText("Page tree")).toBeVisible();
  await expect(page.getByText("Inspector")).toBeVisible();

  const [previewPage] = await Promise.all([
    context.waitForEvent("page"),
    page.getByRole("link", { name: "Open draft preview" }).click(),
  ]);

  await previewPage.waitForLoadState("domcontentloaded");
  await expect(previewPage.getByRole("heading", { name: "Builder-only view of the current draft manifest" })).toBeVisible();
  await expect(previewPage.getByText("Internal draft preview.")).toBeVisible();
});
