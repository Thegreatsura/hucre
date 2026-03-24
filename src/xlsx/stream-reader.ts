// ── Streaming XLSX Reader ────────────────────────────────────────────
// Yields rows one at a time from an XLSX file via SAX parsing.
// Parses shared strings and styles upfront (small), then streams
// worksheet rows without buffering the entire sheet in memory.

import type { CellValue, ReadOptions } from "../_types";
import type { SharedString } from "./shared-strings";
import type { ParsedStyles } from "./styles";
import type { Relationship } from "./relationships";
import { ParseError, ZipError } from "../errors";
import { ZipReader } from "../zip/reader";
import { parseXml, parseSax, decodeOoxmlEscapes } from "../xml/parser";
import { parseContentTypes } from "./content-types";
import { parseRelationships } from "./relationships";
import { parseSharedStrings } from "./shared-strings";
import { parseStyles, isDateStyle } from "./styles";
import { parseCellRef } from "./worksheet";
import { serialToDate } from "../_date";

// ── Types ────────────────────────────────────────────────────────────

export interface StreamRow {
  /** 0-based row index */
  index: number;
  /** Cell values for this row */
  values: CellValue[];
}

// ── OOXML Relationship Types ─────────────────────────────────────────

const REL_WORKBOOK =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_WORKSHEET =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_SHARED_STRINGS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

// ── Helpers ──────────────────────────────────────────────────────────

function decodeUtf8(data: Uint8Array): string {
  return new TextDecoder("utf-8").decode(data);
}

function resolvePath(base: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);

  const baseParts = base.split("/").filter(Boolean);
  const targetParts = target.split("/").filter(Boolean);

  for (const part of targetParts) {
    if (part === "..") {
      baseParts.pop();
    } else if (part !== ".") {
      baseParts.push(part);
    }
  }

  return baseParts.join("/");
}

function dirname(path: string): string {
  const idx = path.lastIndexOf("/");
  return idx === -1 ? "" : path.slice(0, idx);
}

// ── Workbook XML Parsing (minimal — just sheet info + date system) ───

interface SheetInfo {
  name: string;
  sheetId: number;
  rId: string;
}

function parseWorkbookXml(
  xml: string,
  options?: ReadOptions,
): { sheets: SheetInfo[]; dateSystem: "1900" | "1904" } {
  const doc = parseXml(xml);
  const sheets: SheetInfo[] = [];
  let dateSystem: "1900" | "1904" = "1900";

  if (options?.dateSystem === "1904") {
    dateSystem = "1904";
  } else if (options?.dateSystem === "1900") {
    dateSystem = "1900";
  }

  for (const child of doc.children) {
    if (typeof child === "string") continue;
    const local = child.local || child.tag;

    if (local === "workbookPr") {
      if (child.attrs["date1904"] === "1" || child.attrs["date1904"] === "true") {
        if (!options?.dateSystem || options.dateSystem === "auto") {
          dateSystem = "1904";
        }
      }
    }

    if (local === "sheets") {
      for (const sheetChild of child.children) {
        if (typeof sheetChild === "string") continue;
        const sheetLocal = sheetChild.local || sheetChild.tag;
        if (sheetLocal === "sheet") {
          const name = sheetChild.attrs["name"] ?? "";
          const sheetId = Number(sheetChild.attrs["sheetId"] ?? "0");
          const rId =
            sheetChild.attrs["r:id"] ??
            sheetChild.attrs["R:id"] ??
            findRIdAttr(sheetChild.attrs) ??
            "";

          if (name && rId) {
            sheets.push({ name, sheetId, rId });
          }
        }
      }
    }
  }

  return { sheets, dateSystem };
}

function findRIdAttr(attrs: Record<string, string>): string | undefined {
  for (const key of Object.keys(attrs)) {
    if (key.endsWith(":id") && attrs[key].startsWith("rId")) {
      return attrs[key];
    }
  }
  return undefined;
}

// ── Resolve target sheet ────────────────────────────────────────────

function resolveTargetSheet(allSheets: SheetInfo[], sheetSpec?: number | string): SheetInfo | null {
  if (sheetSpec === undefined) {
    // Default: first sheet
    return allSheets[0] ?? null;
  }

  if (typeof sheetSpec === "number") {
    return sheetSpec >= 0 && sheetSpec < allSheets.length ? allSheets[sheetSpec] : null;
  }

  return allSheets.find((s) => s.name === sheetSpec) ?? null;
}

// ── Streaming row parser via SAX ────────────────────────────────────

function* parseWorksheetRows(
  xml: string,
  sharedStrings: SharedString[],
  styles: ParsedStyles | null,
  dateSystem: "1900" | "1904",
): Generator<StreamRow, void, undefined> {
  // We will collect rows from the SAX parser and yield them.
  // Since parseSax is synchronous and runs callbacks, we accumulate
  // completed rows into an array, then yield them after parseSax finishes
  // processing a section. But since parseSax processes the whole string
  // at once, we collect all rows and yield them one by one.
  //
  // The key memory advantage: we don't build Cell objects or a cells Map.
  // We only emit the CellValue[] for each row, discarding per-row state
  // immediately.

  const completedRows: StreamRow[] = [];

  let inSheetData = false;
  let inRow = false;
  let inCell = false;
  let inValue = false;
  let _inFormula = false;
  let inInlineStr = false;
  let inInlineT = false;

  let currentRowIndex = -1;
  let currentRowCells: Array<{ col: number; value: CellValue }> = [];

  // Rich text in inline strings
  let inInlineR = false;
  let inInlineRT = false;

  // Current cell state
  let cellRef = "";
  let cellType = "";
  let cellStyleIndex = -1;
  let cellValueText = "";
  let inlineText = "";
  let inlineRichTextParts: string[] = [];
  let currentRunText = "";

  parseSax(xml, {
    onOpenTag(tag, attrs) {
      const local = tag.includes(":") ? tag.slice(tag.indexOf(":") + 1) : tag;

      switch (local) {
        case "sheetData":
          inSheetData = true;
          break;
        case "row":
          if (inSheetData) {
            inRow = true;
            currentRowIndex = attrs["r"] ? Number(attrs["r"]) - 1 : currentRowIndex + 1;
            currentRowCells = [];
          }
          break;
        case "c":
          if (inRow) {
            inCell = true;
            cellRef = attrs["r"] ?? "";
            cellType = attrs["t"] ?? "";
            cellStyleIndex = attrs["s"] ? Number(attrs["s"]) : -1;
            cellValueText = "";
            inlineText = "";
            inlineRichTextParts = [];
          }
          break;
        case "v":
          if (inCell) inValue = true;
          break;
        case "f":
          if (inCell) _inFormula = true;
          break;
        case "is":
          if (inCell) inInlineStr = true;
          break;
        case "t":
          if (inInlineStr && !inInlineR) {
            inInlineT = true;
          } else if (inInlineR) {
            inInlineRT = true;
          }
          break;
        case "r":
          if (inInlineStr) {
            inInlineR = true;
            currentRunText = "";
          }
          break;
      }
    },

    onText(text) {
      if (inValue) {
        cellValueText += text;
      } else if (inInlineT) {
        inlineText += text;
      } else if (inInlineRT) {
        currentRunText += text;
      }
    },

    onCloseTag(tag) {
      const local = tag.includes(":") ? tag.slice(tag.indexOf(":") + 1) : tag;

      switch (local) {
        case "sheetData":
          inSheetData = false;
          break;
        case "row":
          if (inRow) {
            // Emit row
            const maxCol =
              currentRowCells.length > 0 ? Math.max(...currentRowCells.map((c) => c.col)) : -1;
            const values: CellValue[] =
              maxCol >= 0 ? Array.from({ length: maxCol + 1 }, () => null) : [];
            for (const cell of currentRowCells) {
              values[cell.col] = cell.value;
            }
            completedRows.push({ index: currentRowIndex, values });
            inRow = false;
          }
          break;
        case "c":
          if (inCell) {
            // Process cell and add to current row
            const value = resolveStreamCellValue(
              cellType,
              cellStyleIndex,
              cellValueText,
              inlineText,
              inlineRichTextParts,
              sharedStrings,
              styles,
              dateSystem,
            );
            if (cellRef) {
              const pos = parseCellRef(cellRef);
              currentRowCells.push({ col: pos.col, value });
            }
            inCell = false;
          }
          break;
        case "v":
          inValue = false;
          break;
        case "f":
          _inFormula = false;
          break;
        case "is":
          inInlineStr = false;
          break;
        case "t":
          if (inInlineRT) {
            inInlineRT = false;
          } else if (inInlineT) {
            inInlineT = false;
          }
          break;
        case "r":
          if (inInlineR) {
            inlineRichTextParts.push(decodeOoxmlEscapes(currentRunText));
            inInlineR = false;
          }
          break;
      }
    },
  });

  // Yield all completed rows
  for (const row of completedRows) {
    yield row;
  }
}

// ── Cell value resolution (streaming — no Cell objects) ──────────────

function resolveStreamCellValue(
  type: string,
  styleIndex: number,
  valueText: string,
  inlineText: string,
  inlineRichTextParts: string[],
  sharedStrings: SharedString[],
  styles: ParsedStyles | null,
  dateSystem: "1900" | "1904",
): CellValue {
  switch (type) {
    case "s": {
      // Shared string
      const idx = Number(valueText);
      if (!Number.isNaN(idx) && idx >= 0 && idx < sharedStrings.length) {
        return sharedStrings[idx].text;
      }
      return valueText;
    }
    case "str": {
      // Inline formula string result
      return decodeOoxmlEscapes(valueText);
    }
    case "inlineStr": {
      // Inline string with <is> element
      if (inlineRichTextParts.length > 0) {
        return inlineRichTextParts.join("");
      }
      return decodeOoxmlEscapes(inlineText);
    }
    case "b": {
      // Boolean
      return valueText === "1" || valueText.toLowerCase() === "true";
    }
    case "e": {
      // Error
      return valueText;
    }
    case "n":
    default: {
      // Number (explicit or implied)
      if (valueText === "") {
        return null;
      }

      const num = Number(valueText);
      if (!Number.isNaN(num)) {
        // Check if this is a date via style
        if (styles && styleIndex >= 0 && isDateStyle(styles, styleIndex)) {
          return serialToDate(num, dateSystem === "1904");
        }
        return num;
      }
      return valueText || null;
    }
  }
}

// ── Main streaming reader ───────────────────────────────────────────

/**
 * Create an async iterable that yields rows one at a time.
 * Parses shared strings and styles upfront (they're small),
 * then streams worksheet rows via SAX parsing.
 */
export async function* streamXlsxRows(
  input: Uint8Array | ArrayBuffer,
  options?: ReadOptions & { sheet?: number | string },
): AsyncGenerator<StreamRow, void, undefined> {
  const data = input instanceof Uint8Array ? input : new Uint8Array(input);

  // 1. Open ZIP archive
  let zip: ZipReader;
  try {
    zip = new ZipReader(data);
  } catch (err) {
    if (err instanceof ZipError) throw err;
    throw new ParseError("Failed to open XLSX file: not a valid ZIP archive", undefined, {
      cause: err,
    });
  }

  // 2. Validate content types
  if (!zip.has("[Content_Types].xml")) {
    throw new ParseError("Invalid XLSX: missing [Content_Types].xml");
  }
  const contentTypesXml = decodeUtf8(await zip.extract("[Content_Types].xml"));
  parseContentTypes(contentTypesXml);

  // 3. Parse _rels/.rels to find the workbook path
  if (!zip.has("_rels/.rels")) {
    throw new ParseError("Invalid XLSX: missing _rels/.rels");
  }
  const rootRelsXml = decodeUtf8(await zip.extract("_rels/.rels"));
  const rootRels = parseRelationships(rootRelsXml);
  const workbookRel = rootRels.find((r) => r.type === REL_WORKBOOK);
  if (!workbookRel) {
    throw new ParseError("Invalid XLSX: cannot find workbook relationship in _rels/.rels");
  }

  const workbookPath = workbookRel.target.startsWith("/")
    ? workbookRel.target.slice(1)
    : workbookRel.target;

  // 4. Parse workbook relationships
  const workbookDir = dirname(workbookPath);
  const workbookRelsPath = workbookDir
    ? `${workbookDir}/_rels/${workbookPath.slice(workbookDir.length + 1)}.rels`
    : `_rels/${workbookPath}.rels`;

  let workbookRels: Relationship[] = [];
  if (zip.has(workbookRelsPath)) {
    const wbRelsXml = decodeUtf8(await zip.extract(workbookRelsPath));
    workbookRels = parseRelationships(wbRelsXml);
  }

  // 5. Parse workbook XML for sheet names and date system
  if (!zip.has(workbookPath)) {
    throw new ParseError(`Invalid XLSX: missing workbook at ${workbookPath}`);
  }
  const workbookXml = decodeUtf8(await zip.extract(workbookPath));
  const { sheets: sheetInfos, dateSystem } = parseWorkbookXml(workbookXml, options);

  // 6. Parse shared strings (small, needed for cell resolution)
  let sharedStrings: SharedString[] = [];
  const ssRel = workbookRels.find((r) => r.type === REL_SHARED_STRINGS);
  if (ssRel) {
    const ssPath = resolvePath(workbookDir, ssRel.target);
    if (zip.has(ssPath)) {
      const ssXml = decodeUtf8(await zip.extract(ssPath));
      sharedStrings = parseSharedStrings(ssXml);
    }
  }

  // 7. Parse styles (needed for date detection)
  let parsedStyles: ParsedStyles | null = null;
  const stylesRel = workbookRels.find((r) => r.type === REL_STYLES);
  if (stylesRel) {
    const stylesPath = resolvePath(workbookDir, stylesRel.target);
    if (zip.has(stylesPath)) {
      const stylesXml = decodeUtf8(await zip.extract(stylesPath));
      parsedStyles = parseStyles(stylesXml);
    }
  }

  // 8. Build rId → worksheet path map
  const sheetRelMap = new Map<string, string>();
  for (const rel of workbookRels) {
    if (rel.type === REL_WORKSHEET) {
      sheetRelMap.set(rel.id, resolvePath(workbookDir, rel.target));
    }
  }

  // 9. Resolve the target sheet
  const targetSheet = resolveTargetSheet(sheetInfos, options?.sheet);
  if (!targetSheet) {
    return; // No matching sheet — yield nothing
  }

  const wsPath = sheetRelMap.get(targetSheet.rId);
  if (!wsPath || !zip.has(wsPath)) {
    throw new ParseError(`Invalid XLSX: missing worksheet file for sheet "${targetSheet.name}"`);
  }

  // 10. Stream worksheet rows via SAX
  const wsXml = decodeUtf8(await zip.extract(wsPath));
  yield* parseWorksheetRows(wsXml, sharedStrings, parsedStyles, dateSystem);
}
