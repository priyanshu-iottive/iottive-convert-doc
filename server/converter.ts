import fs from "fs";
import os from "os";
import path from "path";
import { fileURLToPath } from "url";
import mammoth from "mammoth";
import { PDFParse, VerbosityLevel } from "pdf-parse";

// ---- Types for parsed content ----
export interface ParsedElement {
  type: "heading" | "paragraph" | "bullet" | "table";
  level?: number; // heading level 1-3
  text?: string;
  bold?: boolean;
  runs?: { text: string; bold?: boolean; italic?: boolean }[];
  rows?: { cells: string[] }[];
}

export interface ParsedDocument {
  elements: ParsedElement[];
}

// ---- Parse uploaded DOCX via mammoth ----
export async function parseDocx(buffer: Buffer): Promise<ParsedDocument> {
  const result = await mammoth.convertToHtml({ buffer });
  const html = result.value;
  return parseHtmlToElements(html);
}

// ---- Parse uploaded PDF ----
export async function parsePdf(buffer: Buffer): Promise<ParsedDocument> {
  const parser = new PDFParse({
    data: buffer,
    verbosity: VerbosityLevel.ERRORS
  });
  const result = await parser.getText();
  const text = result.text;
  const lines = text.split("\n").filter((l) => l.trim());
  const elements: ParsedElement[] = [];

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;

    // Heuristic: lines that are ALL CAPS or short + no period = heading
    if (
      trimmed.length < 80 &&
      (trimmed === trimmed.toUpperCase() || /^[A-Z][A-Za-z\s&\-–:]+$/.test(trimmed)) &&
      !trimmed.endsWith(".")
    ) {
      elements.push({
        type: "heading",
        level: trimmed === trimmed.toUpperCase() ? 1 : 2,
        text: trimmed,
        runs: [{ text: trimmed, bold: true }],
      });
    } else if (/^[\•\-\*\◦\▪\●]\s/.test(trimmed) || /^\d+[\.\)]\s/.test(trimmed)) {
      const bulletText = trimmed.replace(/^[\•\-\*\◦\▪\●\d\.\)]+\s*/, "");
      elements.push({
        type: "bullet",
        text: bulletText,
        runs: [{ text: bulletText }],
      });
    } else {
      elements.push({
        type: "paragraph",
        text: trimmed,
        runs: [{ text: trimmed }],
      });
    }
  }

  return { elements };
}

// ---- Parse HTML (from mammoth) into structured elements ----
function parseHtmlToElements(html: string): ParsedDocument {
  const elements: ParsedElement[] = [];

  // Simple regex-based HTML parser for mammoth output
  // mammoth produces clean HTML: <h1>, <h2>, <h3>, <p>, <ul>/<ol>/<li>, <table>
  const tagRegex =
    /<(h[1-6]|p|li|tr|table|\/table|\/tr|td|th)[^>]*>([\s\S]*?)(?=<(?:h[1-6]|p|li|tr|table|\/table|\/tr|td|th)[^>]*>|$)/gi;

  let inTable = false;
  let currentRows: { cells: string[] }[] = [];
  let currentCells: string[] = [];

  // Split HTML by major tags
  const parts = html.split(/(<\/?(?:h[1-6]|p|li|table|tr|td|th|ul|ol)[^>]*>)/gi);

  let currentTag = "";
  let buffer = "";

  for (const part of parts) {
    const tagMatch = part.match(
      /^<(\/?)(?:h([1-6])|p|li|table|tr|td|th|ul|ol)([^>]*)>$/i
    );
    if (tagMatch) {
      const isClosing = tagMatch[1] === "/";
      const tagName = part
        .replace(/<\/?/, "")
        .replace(/[^a-zA-Z0-9].*/, "")
        .toLowerCase();

      if (isClosing) {
        if (tagName === "table") {
          if (currentRows.length > 0) {
            elements.push({ type: "table", rows: currentRows });
          }
          currentRows = [];
          inTable = false;
        } else if (tagName === "tr") {
          if (currentCells.length > 0) {
            currentRows.push({ cells: [...currentCells] });
          }
          currentCells = [];
        } else if (tagName === "td" || tagName === "th") {
          currentCells.push(stripHtml(buffer).trim());
          buffer = "";
        } else if (
          tagName.startsWith("h") ||
          tagName === "p" ||
          tagName === "li"
        ) {
          // Inside a table, don't consume the buffer on </p> — let </td> handle it
          if (!inTable) {
            const text = stripHtml(buffer).trim();
            if (text) {
              if (tagName.startsWith("h")) {
                const level = parseInt(tagName[1]) || 2;
                elements.push({
                  type: "heading",
                  level: Math.min(level, 3),
                  text,
                  runs: parseRuns(buffer),
                });
              } else if (tagName === "li") {
                elements.push({
                  type: "bullet",
                  text,
                  runs: parseRuns(buffer),
                });
              } else {
                elements.push({
                  type: "paragraph",
                  text,
                  runs: parseRuns(buffer),
                });
              }
            }
            buffer = "";
          }
          // When inside table, buffer keeps accumulating for the parent <td>
        }
        currentTag = "";
      } else {
        // Opening tag
        if (tagName === "table") {
          inTable = true;
          currentRows = [];
        } else if (tagName === "tr") {
          currentCells = [];
        }
        currentTag = tagName;
        // Don't reset buffer inside table cells — <p> inside <td> should not clear accumulated text
        if (!inTable) {
          buffer = "";
        }
      }
    } else {
      buffer += part;
    }
  }

  return { elements };
}

function stripHtml(html: string): string {
  return html
    .replace(/<[^>]+>/g, "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, " ");
}

function parseRuns(
  html: string
): { text: string; bold?: boolean; italic?: boolean }[] {
  const runs: { text: string; bold?: boolean; italic?: boolean }[] = [];

  // Extract text with bold/italic markers
  let remaining = html;
  const partRegex =
    /(<(?:strong|b|em|i)>)([\s\S]*?)(<\/(?:strong|b|em|i)>)|([^<]+)/gi;
  let match;

  while ((match = partRegex.exec(remaining)) !== null) {
    if (match[2]) {
      const tag = match[1].toLowerCase();
      const text = stripHtml(match[2]);
      if (text.trim()) {
        runs.push({
          text,
          bold: tag.includes("strong") || tag.includes("<b>"),
          italic: tag.includes("em") || tag.includes("<i>"),
        });
      }
    } else if (match[4]) {
      const text = stripHtml(match[4]);
      if (text.trim()) {
        runs.push({ text });
      }
    }
  }

  if (runs.length === 0) {
    const text = stripHtml(html);
    if (text.trim()) {
      runs.push({ text });
    }
  }

  return runs;
}

// ---- Build branded DOCX using python-docx via subprocess ----
export async function buildBrandedDocx(
  parsed: ParsedDocument,
  clientName: string,
  contactName: string,
  contactEmail: string,
  projectTitle: string
): Promise<Buffer> {
  const { execSync } = await import("child_process");

  // Write parsed data to temp JSON
  const tmpDir = path.join(os.tmpdir(), `docconv_${Date.now()}`);
  fs.mkdirSync(tmpDir, { recursive: true });

  const dataPath = path.join(tmpDir, "parsed.json");
  fs.writeFileSync(
    dataPath,
    JSON.stringify({
      elements: parsed.elements,
      clientName,
      contactName,
      contactEmail,
      projectTitle,
    })
  );

  const outputPath = path.join(tmpDir, "output.docx");
  // In dev (ESM): import.meta.url works. In prod (CJS): use __dirname.
  let scriptDir: string;
  try {
    scriptDir = path.dirname(fileURLToPath(import.meta.url));
  } catch {
    // @ts-ignore - __dirname available in CJS
    scriptDir = typeof __dirname !== 'undefined' ? __dirname : process.cwd();
  }
  
  // Look for brand-template.docx in server/ dir or next to the built file
  let templatePath = path.join(scriptDir, "brand-template.docx");
  if (!fs.existsSync(templatePath)) {
    templatePath = path.join(process.cwd(), "server", "brand-template.docx");
  }
  let pythonScript = path.join(scriptDir, "build_docx.py");
  if (!fs.existsSync(pythonScript)) {
    pythonScript = path.join(process.cwd(), "server", "build_docx.py");
  }

  // Try python3 first, then fallback to python
  try {
    execSync(
      `python3 "${pythonScript}" "${dataPath}" "${templatePath}" "${outputPath}"`,
      { timeout: 30000, stdio: 'ignore' }
    );
  } catch (e) {
    execSync(
      `python "${pythonScript}" "${dataPath}" "${templatePath}" "${outputPath}"`,
      { timeout: 30000 }
    );
  }

  const result = fs.readFileSync(outputPath);

  // Cleanup
  fs.rmSync(tmpDir, { recursive: true, force: true });

  return result;
}
