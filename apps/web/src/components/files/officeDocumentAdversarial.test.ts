import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";
import * as XLSX from "xlsx";

import { parseOfficeDocument } from "./officeDocumentParser";

const presentationNamespace = "http://schemas.openxmlformats.org/presentationml/2006/main";
const drawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
const relationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function oneSlide(shape: string, additions: Record<string, string> = {}) {
  return zipSync({
    "_rels/.rels": strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="officeDocument" Type="${relationshipNamespace}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
    ),
    "ppt/presentation.xml": strToU8(
      `<p:presentation xmlns:p="${presentationNamespace}" xmlns:r="${relationshipNamespace}"><p:sldIdLst><p:sldId id="256" r:id="slide"/></p:sldIdLst></p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": strToU8(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="slide" Type="${relationshipNamespace}/slide" Target="slides/slide1.xml"/></Relationships>`,
    ),
    "ppt/slides/slide1.xml": strToU8(
      `<p:sld xmlns:p="${presentationNamespace}" xmlns:a="${drawingNamespace}"><p:cSld><p:spTree>${shape}</p:spTree></p:cSld></p:sld>`,
    ),
    ...Object.fromEntries(Object.entries(additions).map(([path, xml]) => [path, strToU8(xml)])),
  });
}

const placedShape = (attributes: string, text: string) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title" ${attributes}/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="500000" y="500000"/><a:ext cx="8000000" cy="1000000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2800"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;

describe("Office document adversarial independent review", () => {
  it("rejects unreferenced FAT heads that share a physical tail before workbook decoding", () => {
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet([["Ordinary workbook"]]),
      "Sheet",
    );
    const original = new Uint8Array(XLSX.write(workbook, { bookType: "biff8", type: "array" }));
    const physicalSectors = original.length / 512 - 1;
    if (physicalSectors + 3 >= 128) throw new Error("Expected a small one-FAT-sector fixture.");
    const bytes = new Uint8Array(original.length + 3 * 512);
    bytes.set(original);
    const view = new DataView(bytes.buffer);
    const fatSector = view.getUint32(76, true);
    const setSuccessor = (sector: number, successor: number) =>
      view.setUint32((fatSector + 1) * 512 + sector * 4, successor, true);
    setSuccessor(physicalSectors, 0xfffffffe);
    setSuccessor(physicalSectors + 1, physicalSectors);
    setSuccessor(physicalSectors + 2, physicalSectors);
    expect(() => parseOfficeDocument(bytes, "xls")).toThrow(
      /allocation|chain|sector|overlap|shared/i,
    );
  });

  it("matches a layout placeholder to the master by type when indices differ", () => {
    const placeholder = (shape: string, type: string, index: number) =>
      shape.replace("<p:nvPr/>", `<p:nvPr><p:ph type="${type}" idx="${index}"/></p:nvPr>`);
    const slide = placeholder(placedShape("", "Body content"), "body", 2).replace(
      /<p:spPr>.*?<\/p:spPr>/,
      "<p:spPr/>",
    );
    const body = placeholder(placedShape("", ""), "body", 1);
    const footer = placeholder(placedShape("", ""), "ftr", 2).replace('y="500000"', 'y="4500000"');
    const layout = placeholder(placedShape("", ""), "body", 2).replace(
      /<p:spPr>.*?<\/p:spPr>/,
      "<p:spPr/>",
    );
    const result = parseOfficeDocument(
      oneSlide(slide, {
        "ppt/slides/_rels/slide1.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="layout" Type="${relationshipNamespace}/slideLayout" Target="../slideLayouts/layout.xml"/></Relationships>`,
        "ppt/slideLayouts/layout.xml": `<p:sldLayout xmlns:p="${presentationNamespace}" xmlns:a="${drawingNamespace}"><p:cSld><p:spTree>${layout}</p:spTree></p:cSld></p:sldLayout>`,
        "ppt/slideLayouts/_rels/layout.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="master" Type="${relationshipNamespace}/slideMaster" Target="../slideMasters/master.xml"/></Relationships>`,
        "ppt/slideMasters/master.xml": `<p:sldMaster xmlns:p="${presentationNamespace}" xmlns:a="${drawingNamespace}"><p:cSld><p:spTree>${body}${footer}</p:spTree></p:cSld></p:sldMaster>`,
      }),
      "pptx",
    );
    if (result.kind !== "presentation") throw new Error("Expected presentation");
    expect(result.slides[0]?.elements).toContainEqual(
      expect.objectContaining({ kind: "text", text: "Body content", y: 500000 }),
    );
  });

  it.each(["1", "true"])("does not paint a PowerPoint shape marked hidden=%s", (hidden) => {
    const result = parseOfficeDocument(
      oneSlide(placedShape(`hidden="${hidden}"`, "Invisible selection-pane item")),
      "pptx",
    );
    if (result.kind !== "presentation") throw new Error("Expected presentation");
    expect(result.slides[0]?.elements).toEqual([]);
    expect(result.slides[0]?.text).toEqual([]);
  });

  it("does not extract text from a hidden parent group", () => {
    const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="8" name="Hidden group" hidden="true"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${placedShape("", "Hidden group child")}</p:grpSp>`;
    const result = parseOfficeDocument(oneSlide(group), "pptx");
    if (result.kind !== "presentation") throw new Error("Expected presentation");
    expect(result.slides[0]?.elements).toEqual([]);
    expect(result.slides[0]?.text).toEqual([]);
  });

  it("retains a shape explicitly marked visible", () => {
    const result = parseOfficeDocument(
      oneSlide(placedShape('hidden="0"', "Visible content")),
      "pptx",
    );
    if (result.kind !== "presentation") throw new Error("Expected presentation");
    expect(result.slides[0]?.elements).toContainEqual(
      expect.objectContaining({ kind: "text", text: "Visible content" }),
    );
    expect(result.slides[0]?.text).toEqual(["Visible content"]);
  });

  it("rejects an actual inflater expansion that disagrees with both ZIP size declarations", () => {
    const bytes = oneSlide(placedShape("", "Value"));
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const decoder = new TextDecoder();
    for (let offset = 0; offset + 46 <= bytes.length; offset++) {
      if (view.getUint32(offset, true) !== 0x02014b50) continue;
      const nameLength = view.getUint16(offset + 28, true);
      const name = decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength));
      if (name === "ppt/slides/slide1.xml") {
        const local = view.getUint32(offset + 42, true);
        view.setUint32(offset + 24, 1, true);
        view.setUint32(local + 22, 1, true);
        break;
      }
    }
    expect(() => parseOfficeDocument(bytes, "pptx")).toThrow(
      /expanded document exceeds preview limits/i,
    );
  });
});
