import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

import { buildWorkbookSheetExports } from "@/lib/workbook/export-records";
import type { NormalizedReportSnapshot } from "@/lib/workbook/types";

const execFileAsync = promisify(execFile);

async function getTemplateWorkbookPath(): Promise<string> {
  const master = path.resolve(process.cwd(), "public", "templates", "IT_Exec_Reporting_Ingestion_Template_master.xlsx");
  const fallback = path.resolve(process.cwd(), "public", "templates", "IT_Exec_Reporting_Ingestion_Template_v4_dummy_data.xlsx");
  try {
    await fs.access(master);
    return master;
  } catch {
    await fs.access(fallback);
    return fallback;
  }
}

async function resolvePythonBinary(): Promise<string> {
  const localVenv = path.resolve(process.cwd(), ".venv", "bin", "python");
  try {
    await fs.access(localVenv);
    return localVenv;
  } catch {
    return "python3";
  }
}

export async function renderWorkbookFromSnapshot(snapshot: NormalizedReportSnapshot): Promise<Buffer> {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "ta-it-reporting-"));
  const recordsPath = path.join(tempDir, "sheet-records.json");
  const outputPath = path.join(tempDir, "report.xlsx");
  const templatePath = await getTemplateWorkbookPath();
  const pythonBinary = await resolvePythonBinary();
  const scriptPath = path.resolve(process.cwd(), "scripts", "render_workbook_from_records.py");

  try {
    await fs.writeFile(recordsPath, `${JSON.stringify(buildWorkbookSheetExports(snapshot), null, 2)}\n`, "utf8");
    await execFileAsync(pythonBinary, [scriptPath, templatePath, recordsPath, outputPath], {
      cwd: process.cwd(),
    });
    return fs.readFile(outputPath);
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}
