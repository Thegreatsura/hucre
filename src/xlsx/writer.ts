// ── XLSX Writer ──────────────────────────────────────────────────────
// Generates valid Office Open XML spreadsheet files (XLSX).

import type { WriteOptions, WriteOutput } from "../_types";
import { ZipWriter } from "../zip/writer";
import { writeContentTypes } from "./content-types-writer";
import type { ContentTypesOptions } from "./content-types-writer";
import { writeRootRels, writeWorkbookXml, writeWorkbookRels } from "./workbook-writer";
import { createStylesCollector } from "./styles-writer";
import { createSharedStrings, writeSharedStringsXml, writeWorksheetXml } from "./worksheet-writer";
import type { WorksheetResult } from "./worksheet-writer";
import { writeDrawing } from "./drawing-writer";
import type { DrawingResult } from "./drawing-writer";
import { writeComments } from "./comments-writer";
import type { CommentsResult } from "./comments-writer";
import { xmlDocument, xmlSelfClose } from "../xml/writer";

const encoder = /* @__PURE__ */ new TextEncoder();

const NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL_HYPERLINK =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const REL_VML_DRAWING =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";

/**
 * Write a Workbook to XLSX format.
 * Returns a Uint8Array containing the ZIP archive.
 */
export async function writeXlsx(options: WriteOptions): Promise<WriteOutput> {
  const { sheets, defaultFont, dateSystem } = options;

  // Create shared collectors
  const styles = createStylesCollector(defaultFont);
  const sharedStrings = createSharedStrings();

  // Generate worksheet XMLs (also populates styles and shared strings)
  const worksheetResults: WorksheetResult[] = [];
  for (const sheet of sheets) {
    const result = writeWorksheetXml(sheet, styles, sharedStrings, dateSystem);
    worksheetResults.push(result);
  }

  const hasSharedStrings = sharedStrings.count() > 0;

  // Generate drawing data for sheets that have images
  const drawingResults: Array<DrawingResult | null> = [];
  const drawingIndices: number[] = [];
  const imageExtensions = new Set<string>();
  let globalImageIndex = 1;

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    if (sheet.images && sheet.images.length > 0) {
      const result = writeDrawing(sheet.images, globalImageIndex);
      drawingResults.push(result);
      drawingIndices.push(i + 1); // 1-based drawing index matches sheet index

      // Track image extensions and advance global counter
      for (const img of result.images) {
        const ext = img.path.split(".").pop();
        if (ext) imageExtensions.add(ext);
      }
      globalImageIndex += sheet.images.length;
    } else {
      drawingResults.push(null);
    }
  }

  // Generate comments data for sheets that have comments
  const commentsResults: Array<CommentsResult | null> = [];
  const commentIndices: number[] = [];

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    if (sheet.cells) {
      const result = writeComments(sheet.cells, i);
      if (result) {
        commentsResults.push(result);
        commentIndices.push(i + 1);
      } else {
        commentsResults.push(null);
      }
    } else {
      commentsResults.push(null);
    }
  }

  // Build ZIP archive
  const zip = new ZipWriter();

  // [Content_Types].xml
  const ctOpts: ContentTypesOptions = {
    sheetCount: sheets.length,
    hasSharedStrings,
    drawingIndices: drawingIndices.length > 0 ? drawingIndices : undefined,
    imageExtensions: imageExtensions.size > 0 ? imageExtensions : undefined,
    commentIndices: commentIndices.length > 0 ? commentIndices : undefined,
  };
  zip.add("[Content_Types].xml", encoder.encode(writeContentTypes(ctOpts)));

  // _rels/.rels
  zip.add("_rels/.rels", encoder.encode(writeRootRels()));

  // xl/workbook.xml
  zip.add("xl/workbook.xml", encoder.encode(writeWorkbookXml(sheets)));

  // xl/_rels/workbook.xml.rels
  zip.add(
    "xl/_rels/workbook.xml.rels",
    encoder.encode(writeWorkbookRels(sheets.length, hasSharedStrings)),
  );

  // xl/styles.xml
  zip.add("xl/styles.xml", encoder.encode(styles.toXml()));

  // xl/sharedStrings.xml (if any strings)
  if (hasSharedStrings) {
    zip.add("xl/sharedStrings.xml", encoder.encode(writeSharedStringsXml(sharedStrings)));
  }

  // xl/worksheets/sheetN.xml + optional xl/worksheets/_rels/sheetN.xml.rels
  for (let i = 0; i < worksheetResults.length; i++) {
    const result = worksheetResults[i];
    const drawing = drawingResults[i];
    const comments = commentsResults[i];

    zip.add(`xl/worksheets/sheet${i + 1}.xml`, encoder.encode(result.xml));

    // Generate worksheet .rels if there are hyperlinks, a drawing, or comments
    const hasHyperlinks = result.hyperlinkRelationships.length > 0;
    const hasDrawing = drawing !== null && result.drawingRId !== null;
    const hasComments = comments !== null && result.legacyDrawingRId !== null;

    if (hasHyperlinks || hasDrawing || hasComments) {
      const relElements: string[] = [];

      // Hyperlink relationships
      for (const rel of result.hyperlinkRelationships) {
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: rel.id,
            Type: REL_HYPERLINK,
            Target: rel.target,
            TargetMode: "External",
          }),
        );
      }

      // Drawing relationship
      if (hasDrawing && result.drawingRId) {
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: result.drawingRId,
            Type: REL_DRAWING,
            Target: `../drawings/drawing${i + 1}.xml`,
          }),
        );
      }

      // Comments relationships (VML drawing + comments file)
      if (hasComments && result.legacyDrawingRId) {
        // Legacy drawing (VML) relationship
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: result.legacyDrawingRId,
            Type: REL_VML_DRAWING,
            Target: `../drawings/vmlDrawing${i + 1}.vml`,
          }),
        );

        // Comments file relationship — use the next rId after legacyDrawingRId
        const legacyRIdNum = parseInt(result.legacyDrawingRId.replace("rId", ""), 10);
        const commentsRId = `rId${legacyRIdNum + 1}`;
        relElements.push(
          xmlSelfClose("Relationship", {
            Id: commentsRId,
            Type: REL_COMMENTS,
            Target: `../comments${i + 1}.xml`,
          }),
        );
      }

      const relsXml = xmlDocument("Relationships", { xmlns: NS_RELATIONSHIPS }, relElements);
      zip.add(`xl/worksheets/_rels/sheet${i + 1}.xml.rels`, encoder.encode(relsXml));
    }

    // Add drawing files
    if (drawing) {
      zip.add(`xl/drawings/drawing${i + 1}.xml`, encoder.encode(drawing.drawingXml));
      zip.add(`xl/drawings/_rels/drawing${i + 1}.xml.rels`, encoder.encode(drawing.drawingRels));

      // Add image files to ZIP (store, don't compress — images are already compressed)
      for (const img of drawing.images) {
        zip.add(img.path, img.data, { compress: false });
      }
    }

    // Add comments and VML drawing files
    if (comments) {
      zip.add(`xl/comments${i + 1}.xml`, encoder.encode(comments.commentsXml));
      zip.add(`xl/drawings/vmlDrawing${i + 1}.vml`, encoder.encode(comments.vmlXml));
    }
  }

  return zip.build();
}
