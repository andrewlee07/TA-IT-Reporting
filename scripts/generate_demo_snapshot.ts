import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { parseWorkbookBuffer } from "@/lib/workbook/parser";

async function main() {
  const fixturePath = path.resolve(process.cwd(), "fixtures", "IT_Exec_Reporting_Ingestion_Template_v2_dummy_data.xlsx");
  const outputPath = path.resolve(process.cwd(), "fixtures", "demo-snapshot.json");
  const workbook = await readFile(fixturePath);
  const { snapshot } = await parseWorkbookBuffer(workbook, path.basename(fixturePath));

  await mkdir(path.dirname(outputPath), { recursive: true });
  await writeFile(outputPath, `${JSON.stringify(snapshot, null, 2)}\n`, "utf8");

  console.log(`Wrote ${outputPath}`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
