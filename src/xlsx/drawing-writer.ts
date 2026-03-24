// ── Drawing Writer ────────────────────────────────────────────────────
// Generates xl/drawings/drawingN.xml and related relationship files
// for embedding images in XLSX worksheets.

import type { SheetImage } from "../_types";
import { xmlDocument, xmlElement, xmlSelfClose } from "../xml/writer";

// ── Namespaces ───────────────────────────────────────────────────────

const NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

// ── Constants ────────────────────────────────────────────────────────

/** Default image width in EMU (English Metric Units). 1 inch = 914400 EMU */
const DEFAULT_WIDTH_EMU = 3000000; // ~3.28 inches
/** Default image height in EMU */
const DEFAULT_HEIGHT_EMU = 2000000; // ~2.19 inches
/** Pixels to EMU conversion factor (96 DPI assumption) */
const EMU_PER_PIXEL = 9525;

// ── Types ────────────────────────────────────────────────────────────

export interface DrawingImage {
  /** Path inside the ZIP: xl/media/imageN.ext */
  path: string;
  /** Raw image bytes */
  data: Uint8Array;
  /** MIME content type */
  contentType: string;
}

export interface DrawingResult {
  /** The drawing XML content (xl/drawings/drawingN.xml) */
  drawingXml: string;
  /** The drawing relationships XML (xl/drawings/_rels/drawingN.xml.rels) */
  drawingRels: string;
  /** Image files to add to the ZIP */
  images: DrawingImage[];
}

// ── Extension / Content Type Mapping ─────────────────────────────────

const IMAGE_EXTENSIONS: Record<string, string> = {
  png: "png",
  jpeg: "jpeg",
  gif: "gif",
};

const IMAGE_CONTENT_TYPES: Record<string, string> = {
  png: "image/png",
  jpeg: "image/jpeg",
  gif: "image/gif",
};

// ── Writer ───────────────────────────────────────────────────────────

/**
 * Generate drawing XML, drawing relationships, and image entries
 * for a worksheet's images.
 *
 * @param images - Array of SheetImage objects to embed
 * @param imageStartIndex - Global image counter start (for unique image file names across sheets)
 * @returns DrawingResult with XML and image file data
 */
export function writeDrawing(images: SheetImage[], imageStartIndex: number): DrawingResult {
  const drawingImages: DrawingImage[] = [];
  const relElements: string[] = [];
  const anchorElements: string[] = [];

  for (let i = 0; i < images.length; i++) {
    const img = images[i];
    const imageIndex = imageStartIndex + i;
    const rId = `rId${i + 1}`;
    const ext = IMAGE_EXTENSIONS[img.type] ?? "png";
    const contentType = IMAGE_CONTENT_TYPES[img.type] ?? "image/png";
    const mediaPath = `xl/media/image${imageIndex}.${ext}`;
    const relTarget = `../media/image${imageIndex}.${ext}`;

    // Add image file entry
    drawingImages.push({
      path: mediaPath,
      data: img.data,
      contentType,
    });

    // Add relationship entry
    relElements.push(
      xmlSelfClose("Relationship", {
        Id: rId,
        Type: REL_IMAGE,
        Target: relTarget,
      }),
    );

    // Calculate dimensions in EMU
    const widthEmu = img.width ? img.width * EMU_PER_PIXEL : DEFAULT_WIDTH_EMU;
    const heightEmu = img.height ? img.height * EMU_PER_PIXEL : DEFAULT_HEIGHT_EMU;

    // Build twoCellAnchor element
    const fromCol = img.anchor.from.col;
    const fromRow = img.anchor.from.row;
    const toCol = img.anchor.to?.col ?? fromCol + 3;
    const toRow = img.anchor.to?.row ?? fromRow + 5;

    const fromElement = xmlElement("xdr:from", undefined, [
      xmlElement("xdr:col", undefined, String(fromCol)),
      xmlElement("xdr:colOff", undefined, "0"),
      xmlElement("xdr:row", undefined, String(fromRow)),
      xmlElement("xdr:rowOff", undefined, "0"),
    ]);

    const toElement = xmlElement("xdr:to", undefined, [
      xmlElement("xdr:col", undefined, String(toCol)),
      xmlElement("xdr:colOff", undefined, "0"),
      xmlElement("xdr:row", undefined, String(toRow)),
      xmlElement("xdr:rowOff", undefined, "0"),
    ]);

    const nvPicPr = xmlElement("xdr:nvPicPr", undefined, [
      xmlSelfClose("xdr:cNvPr", { id: i + 2, name: `Picture ${i + 1}` }),
      xmlElement("xdr:cNvPicPr", undefined, [xmlSelfClose("a:picLocks", { noChangeAspect: 1 })]),
    ]);

    const blipFill = xmlElement("xdr:blipFill", undefined, [
      xmlSelfClose("a:blip", { "r:embed": rId }),
      xmlElement("a:stretch", undefined, [xmlSelfClose("a:fillRect")]),
    ]);

    const spPr = xmlElement("xdr:spPr", undefined, [
      xmlElement("a:xfrm", undefined, [
        xmlSelfClose("a:off", { x: 0, y: 0 }),
        xmlSelfClose("a:ext", { cx: widthEmu, cy: heightEmu }),
      ]),
      xmlElement("a:prstGeom", { prst: "rect" }, [xmlSelfClose("a:avLst")]),
    ]);

    const pic = xmlElement("xdr:pic", undefined, [nvPicPr, blipFill, spPr]);

    const anchor = xmlElement("xdr:twoCellAnchor", undefined, [
      fromElement,
      toElement,
      pic,
      xmlSelfClose("xdr:clientData"),
    ]);

    anchorElements.push(anchor);
  }

  // Build drawing XML
  const drawingXml = xmlDocument(
    "xdr:wsDr",
    {
      "xmlns:xdr": NS_XDR,
      "xmlns:a": NS_A,
      "xmlns:r": NS_R,
    },
    anchorElements,
  );

  // Build drawing relationships XML
  const drawingRels = xmlDocument("Relationships", { xmlns: NS_RELATIONSHIPS }, relElements);

  return {
    drawingXml,
    drawingRels,
    images: drawingImages,
  };
}
