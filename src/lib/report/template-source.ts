import { promises as fs } from "node:fs";
import path from "node:path";

const templatePath = path.resolve(process.cwd(), "src/lib/report/template.html");

let cachedTemplate: string | null = null;
let cachedTemplateStyles: string | null = null;
let cachedTemplateBody: string | null = null;

function shouldBypassCache(): boolean {
  return process.env.NODE_ENV !== "production";
}

export async function loadTemplateSource(): Promise<string> {
  if (shouldBypassCache()) {
    return fs.readFile(templatePath, "utf8");
  }

  if (!cachedTemplate) {
    cachedTemplate = await fs.readFile(templatePath, "utf8");
  }

  return cachedTemplate;
}

export async function loadTemplateStyles(): Promise<string> {
  if (shouldBypassCache()) {
    const template = await loadTemplateSource();
    const match = template.match(/<style>([\s\S]*?)<\/style>/);

    if (!match) {
      throw new Error("Unable to extract report styles from template.");
    }

    return match[1].trim();
  }

  if (!cachedTemplateStyles) {
    const template = await loadTemplateSource();
    const match = template.match(/<style>([\s\S]*?)<\/style>/);

    if (!match) {
      throw new Error("Unable to extract report styles from template.");
    }

    cachedTemplateStyles = match[1].trim();
  }

  return cachedTemplateStyles;
}

export async function loadTemplateBodyMarkup(): Promise<string> {
  if (shouldBypassCache()) {
    const template = await loadTemplateSource();
    const match = template.match(/<body>([\s\S]*?)<script>\s*const D = __REPORT_DATA__/);

    if (!match) {
      throw new Error("Unable to extract report body markup from template.");
    }

    return match[1].trim();
  }

  if (!cachedTemplateBody) {
    const template = await loadTemplateSource();
    const match = template.match(/<body>([\s\S]*?)<script>\s*const D = __REPORT_DATA__/);

    if (!match) {
      throw new Error("Unable to extract report body markup from template.");
    }

    cachedTemplateBody = match[1].trim();
  }

  return cachedTemplateBody;
}
