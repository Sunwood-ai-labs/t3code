import { strFromU8, strToU8, unzipSync, Zip, ZipDeflate, zipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";
import * as XLSX from "xlsx";

import { OFFICE_INPUT_MAX_BYTES, parseOfficeDocument } from "./officeDocumentParser";

function workbookBytes(bookType: "xlsx" | "biff8", sheet?: XLSX.WorkSheet): Uint8Array {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    sheet ??
      XLSX.utils.aoa_to_sheet([
        ["品名", "価格"],
        ["りんご", 125],
      ]),
    "売上",
  );
  return new Uint8Array(XLSX.write(workbook, { type: "array", bookType }));
}

const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL = "http://schemas.openxmlformats.org/package/2006/relationships";

function presentationParts() {
  return {
    "_rels/.rels": strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="presentation" Type="${R}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
    ),
    "[Content_Types].xml": strToU8(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
    ),
    "ppt/presentation.xml": strToU8(
      `<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst><p:sldId id="256" r:id="second"/><p:sldId id="257" r:id="first"/></p:sldIdLst></p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="first" Type="${R}/slide" Target="slides/slide1.xml"/><Relationship Id="second" Type="${R}/slide" Target="slides/slide2.xml"/></Relationships>`,
    ),
    "ppt/slides/slide1.xml": strToU8(
      `<p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Second &amp; final</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`,
    ),
    "ppt/slides/slide2.xml": strToU8(
      `<p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>最初</a:t></a:r><a:br/><a:fld><a:t>42</a:t></a:fld></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`,
    ),
  };
}

function centralOffset(bytes: Uint8Array): number {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let index = 0; index + 4 <= bytes.length; index++) {
    if (view.getUint32(index, true) === 0x02014b50) return index;
  }
  throw new Error("Fixture has no ZIP central directory.");
}

function visualPresentationParts(): Record<string, Uint8Array> {
  const parts: Record<string, Uint8Array> = presentationParts();
  parts["[Content_Types].xml"] = strToU8(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/></Types>',
  );
  const placeholder = (idx: number, text: string, props = "", runs = "") =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${idx}" name="Shape ${idx}"/><p:nvPr><p:ph idx="${idx}"/></p:nvPr></p:nvSpPr><p:spPr>${props}</p:spPr><p:txBody><a:p><a:r>${runs}<a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
  parts["ppt/slides/slide2.xml"] =
    strToU8(`<p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:cSld><p:spTree>${placeholder(7, "日本語の見出し", "", '<a:rPr sz="3600" b="1"/>')}${placeholder(8, "本文はマスターから継承")}
    <p:sp><p:spPr><a:xfrm rot="5400000"><a:off x="100000" y="100000"/><a:ext cx="500000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill></p:spPr></p:sp>
    <p:pic><p:nvPicPr><p:cNvPr id="10" descr="テスト画像"/></p:nvPicPr><p:blipFill><a:blip r:embed="image"/></p:blipFill><p:spPr><a:xfrm><a:off x="4000000" y="3000000"/><a:ext cx="1000000" cy="1000000"/></a:xfrm></p:spPr></p:pic>
    </p:spTree></p:cSld></p:sld>`);
  parts["ppt/slides/_rels/slide2.xml.rels"] = strToU8(
    `<Relationships xmlns="${REL}"><Relationship Id="layout" Type="${R}/slideLayout" Target="../slideLayouts/layout.xml"/><Relationship Id="image" Type="${R}/image" Target="../media/picture.png"/></Relationships>`,
  );
  parts["ppt/slideLayouts/layout.xml"] = strToU8(
    `<p:sldLayout xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph idx="7" type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="500000" y="300000"/><a:ext cx="8000000" cy="1000000"/></a:xfrm></p:spPr></p:sp><p:sp><p:nvSpPr><p:nvPr><p:ph idx="8" type="body"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp></p:spTree></p:cSld></p:sldLayout>`,
  );
  parts["ppt/slideLayouts/_rels/layout.xml.rels"] = strToU8(
    `<Relationships xmlns="${REL}"><Relationship Id="master" Type="${R}/slideMaster" Target="../slideMasters/master.xml"/></Relationships>`,
  );
  parts["ppt/slideMasters/master.xml"] = strToU8(
    `<p:sldMaster xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph idx="8" type="body"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="600000" y="2000000"/><a:ext cx="7000000" cy="2000000"/></a:xfrm></p:spPr><p:txBody><a:p><a:r><a:rPr sz="2400"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:rPr><a:t/></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sldMaster>`,
  );
  parts["ppt/media/picture.png"] = Uint8Array.from(
    atob(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6xTgAAAAASUVORK5CYII=",
    ),
    (char) => char.charCodeAt(0),
  );
  return parts;
}

function presentationWithBitmap(name: string, bytes: Uint8Array, mime: string, override = true) {
  const parts = visualPresentationParts();
  delete parts["ppt/media/picture.png"];
  parts[`ppt/media/${name}`] = bytes;
  const relationships = "ppt/slides/_rels/slide2.xml.rels";
  parts[relationships] = strToU8(
    strFromU8(parts[relationships]!).replace("../media/picture.png", `../media/${name}`),
  );
  parts["[Content_Types].xml"] = strToU8(
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${
      override
        ? `<Override PartName="/ppt/media/${name}" ContentType="${mime}"/>`
        : `<Default Extension="${name.slice(name.lastIndexOf(".") + 1)}" ContentType="${mime}"/>`
    }</Types>`,
  );
  return parts;
}

const xmlEncodings = [
  { label: "UTF-16LE with BOM", littleEndian: true, bom: true, declaration: "UTF-16" },
  { label: "UTF-16BE with BOM", littleEndian: false, bom: true, declaration: "UTF-16" },
  { label: "UTF-16LE without BOM", littleEndian: true, bom: false, declaration: "UTF-16LE" },
  { label: "UTF-16BE without BOM", littleEndian: false, bom: false, declaration: "UTF-16BE" },
] as const;

function transformXmlParts(
  parts: Record<string, Uint8Array>,
  transform: (xml: string) => Uint8Array,
): Record<string, Uint8Array> {
  return Object.fromEntries(
    Object.entries(parts).map(([name, bytes]) => [
      name,
      /\.(?:xml|rels)$/.test(name) ? transform(strFromU8(bytes)) : bytes,
    ]),
  );
}

function utf16Parts(parts: Record<string, Uint8Array>, encoding: (typeof xmlEncodings)[number]) {
  return transformXmlParts(parts, (xml) => {
    const text = `<?xml version="1.0" encoding="${encoding.declaration}"?>${xml.replace(/^<\?xml[^?]*\?>/, "")}`;
    const bytes = new Uint8Array(text.length * 2 + (encoding.bom ? 2 : 0));
    const view = new DataView(bytes.buffer);
    if (encoding.bom) view.setUint16(0, 0xfeff, encoding.littleEndian);
    for (let index = 0; index < text.length; index++) {
      view.setUint16(
        index * 2 + (encoding.bom ? 2 : 0),
        text.charCodeAt(index),
        encoding.littleEndian,
      );
    }
    return bytes;
  });
}

function strictPresentationParts() {
  return transformXmlParts(visualPresentationParts(), (xml) =>
    strToU8(
      xml
        .replaceAll(P, "http://purl.oclc.org/ooxml/presentationml/main")
        .replaceAll(A, "http://purl.oclc.org/ooxml/drawingml/main")
        .replaceAll(R, "http://purl.oclc.org/ooxml/officeDocument/relationships"),
    ),
  );
}

describe("normal OOXML namespace and encoding compatibility", () => {
  it("displays a valid input workbook when its normalized ZIP exceeds the input upload limit", () => {
    let seed = 0x1937cafe;
    const rows = Array.from({ length: 90 }, () => {
      const characters: string[] = [];
      for (let index = 0; index < 30000; index++) {
        seed ^= seed << 13;
        seed ^= seed >>> 17;
        seed ^= seed << 5;
        characters.push(String.fromCharCode(0x4e00 + ((seed >>> 0) % 20992)));
      }
      return [characters.join("")];
    });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), "Unicode notes");
    const parts = unzipSync(
      new Uint8Array(XLSX.write(workbook, { type: "array", bookType: "xlsx", bookSST: true })),
    );
    const encoded = utf16Parts(parts, xmlEncodings[0]);
    const input = zipSync(encoded, { level: 9 });
    // Ordinary Excel text cells fit below 32,767 characters. UTF-16 compresses
    // these CJK notes better than UTF-8, without approaching the XML budgets.
    expect(rows.every(([text]) => text!.length <= 32767)).toBe(true);
    expect(input.length).toBeLessThanOrEqual(OFFICE_INPUT_MAX_BYTES);
    expect(zipSync(parts, { level: 1 }).length).toBeGreaterThan(OFFICE_INPUT_MAX_BYTES);
    for (const contents of [Object.values(encoded), Object.values(parts)]) {
      expect(Math.max(...contents.map((bytes) => bytes.length))).toBeLessThanOrEqual(
        8 * 1024 * 1024,
      );
      expect(contents.reduce((sum, bytes) => sum + bytes.length, 0)).toBeLessThanOrEqual(
        32 * 1024 * 1024,
      );
    }
    const document = parseOfficeDocument(input, "xlsx");
    if (document.kind !== "spreadsheet") throw new Error("Expected spreadsheet");
    expect(document.sheets).toHaveLength(1);
    expect(document.sheets[0]).toMatchObject({
      name: "Unicode notes",
      totalRows: 90,
      totalColumns: 1,
      truncated: true,
    });
    const visibleRows = document.sheets[0]?.rows ?? [];
    expect(visibleRows).toHaveLength(90);
    for (const [index, row] of visibleRows.entries()) {
      expect(row).toHaveLength(1);
      // Compare compact facts first so a failure cannot dump 90 long cells.
      expect(row[0]?.length).toBe(2000);
      expect(row[0]?.slice(0, 32)).toBe(rows[index]?.[0]?.slice(0, 32));
      expect(row[0] === `${rows[index]?.[0]?.slice(0, 1999)}…`).toBe(true);
    }
  });

  it("preserves Strict slide relationship order, inherited geometry, formatted text and embedded images", () => {
    const document = parseOfficeDocument(zipSync(strictPresentationParts()), "pptx");
    expect(document).toEqual(parseOfficeDocument(zipSync(visualPresentationParts()), "pptx"));
    if (document.kind !== "presentation") throw new Error("Expected presentation");
    expect(document.slides.map((slide) => slide.text)).toEqual([
      ["日本語の見出し", "本文はマスターから継承"],
      ["Second & final"],
    ]);
    expect(document.slides[0]?.elements).toMatchObject([
      { kind: "text", text: "日本語の見出し", x: 500000, y: 300000, fontSize: 36, bold: true },
      { kind: "text", text: "本文はマスターから継承", x: 600000, y: 2000000, color: "#123456" },
      { kind: "shape", shape: "rect", rotation: 90 },
      { kind: "image", alt: "テスト画像", x: 4000000, y: 3000000 },
    ]);
    const image = document.slides[0]?.elements[3];
    expect(image?.kind === "image" && image.src.startsWith("data:image/png;base64,")).toBe(true);
  });

  it.each(xmlEncodings)(
    "reads spreadsheet names and Unicode cell values from $label parts",
    (encoding) => {
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.aoa_to_sheet([
          ["品名", "価格"],
          ["りんご 🐈", 125],
        ]),
        "売上",
      );
      const original = new Uint8Array(
        XLSX.write(workbook, { type: "array", bookType: "xlsx", bookSST: true }),
      );
      const parts = unzipSync(original);
      expect(parts["xl/sharedStrings.xml"]).toBeDefined();
      const document = parseOfficeDocument(zipSync(utf16Parts(parts, encoding)), "xlsx");
      expect(document).toEqual(parseOfficeDocument(original, "xlsx"));
      if (document.kind !== "spreadsheet") throw new Error("Expected spreadsheet");
      expect(document.sheets[0]).toMatchObject({
        name: "売上",
        rows: [
          ["品名", "価格"],
          ["りんご 🐈", "125"],
        ],
      });
    },
  );

  it("preserves namespace URI strings written as ordinary Strict slide text", () => {
    const parts = strictPresentationParts();
    const literal =
      "http://purl.oclc.org/ooxml/presentationml/main http://purl.oclc.org/ooxml/drawingml/main http://purl.oclc.org/ooxml/officeDocument/relationships";
    parts["ppt/slides/slide1.xml"] = strToU8(
      strFromU8(parts["ppt/slides/slide1.xml"]!).replace("Second &amp; final", literal),
    );
    const document = parseOfficeDocument(zipSync(parts), "pptx");
    if (document.kind !== "presentation") throw new Error("Expected presentation");
    expect(document.slides[1]?.text).toEqual([literal]);
  });

  it.each(xmlEncodings)(
    "preserves presentation text, relationships and visuals from $label parts",
    (encoding) => {
      expect(
        parseOfficeDocument(zipSync(utf16Parts(visualPresentationParts(), encoding)), "pptx"),
      ).toEqual(parseOfficeDocument(zipSync(visualPresentationParts()), "pptx"));
    },
  );

  it("reads Strict namespace aliases together with UTF-16BE relationships and slide parts", () => {
    expect(
      parseOfficeDocument(zipSync(utf16Parts(strictPresentationParts(), xmlEncodings[1])), "pptx"),
    ).toEqual(parseOfficeDocument(zipSync(visualPresentationParts()), "pptx"));
  });

  it("retains UTF-8 BOM support for spreadsheet and presentation parts", () => {
    const withBom = (xml: string) => strToU8(`\uFEFF${xml}`);
    const workbook = workbookBytes("xlsx");
    expect(
      parseOfficeDocument(zipSync(transformXmlParts(unzipSync(workbook), withBom)), "xlsx"),
    ).toEqual(parseOfficeDocument(workbook, "xlsx"));
    expect(
      parseOfficeDocument(zipSync(transformXmlParts(visualPresentationParts(), withBom)), "pptx"),
    ).toEqual(parseOfficeDocument(zipSync(visualPresentationParts()), "pptx"));
  });

  it("rejects an encoding declaration that contradicts the actual XML bytes", () => {
    const parts = presentationParts();
    parts["ppt/slides/slide2.xml"] = strToU8(
      `<?xml version="1.0" encoding="UTF-16"?>${strFromU8(parts["ppt/slides/slide2.xml"])}`,
    );
    expect(() => parseOfficeDocument(zipSync(parts), "pptx")).toThrow();
  });
});

describe("parseOfficeDocument spreadsheets", () => {
  it("preserves a canonical workbook declared through a Default content type", () => {
    const original = workbookBytes("xlsx");
    const parts = unzipSync(original);
    let contentTypes = strFromU8(parts["[Content_Types].xml"]!)
      .replace(/<Override\b[^>]*PartName="\/xl\/workbook\.xml"[^>]*\/>/, "")
      .replace(
        /<Default\b[^>]*Extension="xml"[^>]*\/>/,
        '<Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
      );
    for (const name of Object.keys(parts)) {
      if (name === "[Content_Types].xml" || name === "xl/workbook.xml" || !name.endsWith(".xml"))
        continue;
      if (!contentTypes.includes(`PartName="/${name}"`)) {
        contentTypes = contentTypes.replace(
          "</Types>",
          `<Override PartName="/${name}" ContentType="application/xml"/></Types>`,
        );
      }
    }
    parts["[Content_Types].xml"] = strToU8(contentTypes);
    const document = parseOfficeDocument(zipSync(parts), "xlsx");
    expect(document).toEqual(parseOfficeDocument(original, "xlsx"));
    if (document.kind !== "spreadsheet") throw new Error("Expected spreadsheet");
    expect(document.sheets[0]).toMatchObject({
      name: "売上",
      rows: [
        ["品名", "価格"],
        ["りんご", "125"],
      ],
    });
  });

  it.each([
    { prefix: "book/", location: "book/main.xml", moveTree: true },
    { prefix: "", location: "rootmain.xml", moveTree: true },
    { prefix: "book/", location: "book/main.xml", moveTree: false },
    { prefix: "", location: "rootmain.xml", moveTree: false },
    { prefix: "book/", location: "book/main.xml", moveTree: false, strictUtf16: true },
  ])(
    "reads relocated workbook $location (move subtree: $moveTree, Strict UTF-16: $strictUtf16) through the relationship graph",
    ({
      prefix,
      location,
      moveTree,
      strictUtf16,
    }: {
      prefix: string;
      location: string;
      moveTree: boolean;
      strictUtf16?: boolean;
    }) => {
      const workbook = XLSX.utils.book_new();
      const visible = XLSX.utils.aoa_to_sheet([
        ["品名 🐈", "数量", "単価", "合計"],
        ["りんご café", 12, 150, 1800],
      ]);
      visible.D2 = { t: "n", f: "B2*C2", v: 1800 };
      XLSX.utils.book_append_sheet(workbook, visible, "売上 日本語");
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([["secret"]]), "Hidden");
      XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.aoa_to_sheet([["very secret"]]),
        "Very hidden",
      );
      XLSX.utils.book_set_sheet_visibility(workbook, "Hidden", 1);
      XLSX.utils.book_set_sheet_visibility(workbook, "Very hidden", 2);
      const original = new Uint8Array(
        XLSX.write(workbook, { type: "array", bookType: "xlsx", bookSST: true }),
      );
      const originalParts = unzipSync(original);
      const baseName = location.slice(location.lastIndexOf("/") + 1);
      const parts = Object.fromEntries(
        Object.entries(originalParts).map(([name, bytes]) => [
          name === "xl/workbook.xml"
            ? location
            : name === "xl/_rels/workbook.xml.rels"
              ? `${prefix}_rels/${baseName}.rels`
              : moveTree && name.startsWith("xl/")
                ? prefix + name.slice(3)
                : name,
          bytes,
        ]),
      );
      parts["_rels/.rels"] = strToU8(
        strFromU8(parts["_rels/.rels"]!).replace(
          'Target="xl/workbook.xml"',
          `Target="/${location}"`,
        ),
      );
      let contentTypes = strFromU8(parts["[Content_Types].xml"]!);
      if (moveTree)
        contentTypes = contentTypes.replaceAll('PartName="/xl/', `PartName="/${prefix}`);
      parts["[Content_Types].xml"] = strToU8(
        contentTypes.replace(
          `PartName="/${moveTree ? prefix : "xl/"}workbook.xml"`,
          `PartName="/${location}"`,
        ),
      );
      if (!moveTree) {
        const sidecar = `${prefix}_rels/${baseName}.rels`;
        parts[sidecar] = strToU8(
          strFromU8(parts[sidecar]!).replace(
            /Target="([^"]+)"/g,
            (_attribute: string, target: string) =>
              `Target="${target.startsWith("/") ? target : `${prefix ? "../" : ""}xl/${target}`}"`,
          ),
        );
      }
      const encoded = strictUtf16
        ? utf16Parts(
            transformXmlParts(parts, (xml) =>
              strToU8(
                xml
                  .replaceAll(
                    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                    "http://purl.oclc.org/ooxml/spreadsheetml/main",
                  )
                  .replaceAll(R, "http://purl.oclc.org/ooxml/officeDocument/relationships"),
              ),
            ),
            xmlEncodings[1],
          )
        : parts;
      const document = parseOfficeDocument(zipSync(encoded), "xlsx");
      expect(document).toEqual(parseOfficeDocument(original, "xlsx"));
      if (document.kind !== "spreadsheet") throw new Error("Expected spreadsheet");
      expect(document.sheets.map((sheet) => sheet.name)).toEqual(["売上 日本語"]);
      expect(document.sheets[0]?.rows).toEqual([
        ["品名 🐈", "数量", "単価", "合計"],
        ["りんご café", "12", "150", "1800"],
      ]);
    },
  );

  it.each([
    { label: "missing", entries: "" },
    {
      label: "external",
      entries: `<Relationship Id="workbook" Type="${R}/officeDocument" Target="https://example.com/workbook.xml" TargetMode="External"/>`,
    },
    {
      label: "ambiguous",
      entries: `<Relationship Id="one" Type="${R}/officeDocument" Target="xl/workbook.xml"/><Relationship Id="two" Type="${R}/officeDocument" Target="xl/workbook.xml"/>`,
    },
  ])(
    "rejects a $label workbook root relationship instead of guessing its conventional path",
    ({ entries }: { entries: string }) => {
      const parts = unzipSync(workbookBytes("xlsx"));
      parts["_rels/.rels"] = strToU8(`<Relationships xmlns="${REL}">${entries}</Relationships>`);
      expect(() => parseOfficeDocument(zipSync(parts), "xlsx")).toThrow(
        /office document|relationship/,
      );
    },
  );

  it.each(["xlsx", "xls"] as const)(
    "previews visible sheets after twenty hidden sheets in %s",
    (format) => {
      const workbook = XLSX.utils.book_new();
      for (let index = 0; index < 20; index++) {
        const name = `Hidden ${index + 1}`;
        XLSX.utils.book_append_sheet(
          workbook,
          XLSX.utils.aoa_to_sheet([[`Hidden value ${index + 1}`]]),
          name,
        );
        XLSX.utils.book_set_sheet_visibility(workbook, name, index % 2 ? 2 : 1);
      }
      XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.aoa_to_sheet([["Public revenue"]]),
        "Visible Revenue",
      );
      XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.aoa_to_sheet([["Public costs"]]),
        "Visible Costs",
      );
      const document = parseOfficeDocument(
        new Uint8Array(
          XLSX.write(workbook, { type: "array", bookType: format === "xls" ? "biff8" : "xlsx" }),
        ),
        format,
      );
      if (document.kind !== "spreadsheet") throw new Error("Expected spreadsheet");
      expect(document.sheets.map((sheet) => ({ name: sheet.name, rows: sheet.rows }))).toEqual([
        { name: "Visible Revenue", rows: [["Public revenue"]] },
        { name: "Visible Costs", rows: [["Public costs"]] },
      ]);
      expect(document.warnings.some((warning) => warning.includes("first 20"))).toBe(false);
    },
  );

  it.each(["xlsx", "xls"] as const)(
    "applies the twenty-sheet cap to visible sheets in %s",
    (format) => {
      const workbook = XLSX.utils.book_new();
      for (let index = 0; index < 21; index++) {
        const hiddenName = `Hidden ${index + 1}`;
        XLSX.utils.book_append_sheet(
          workbook,
          XLSX.utils.aoa_to_sheet([["Hidden value"]]),
          hiddenName,
        );
        XLSX.utils.book_set_sheet_visibility(workbook, hiddenName, 1);
        XLSX.utils.book_append_sheet(
          workbook,
          XLSX.utils.aoa_to_sheet([[`Public value ${index + 1}`]]),
          `Visible ${String(index + 1).padStart(2, "0")}`,
        );
      }
      const document = parseOfficeDocument(
        new Uint8Array(
          XLSX.write(workbook, { type: "array", bookType: format === "xls" ? "biff8" : "xlsx" }),
        ),
        format,
      );
      if (document.kind !== "spreadsheet") throw new Error("Expected spreadsheet");
      expect(document.sheets).toHaveLength(20);
      expect(document.sheets[0]).toMatchObject({ name: "Visible 01", rows: [["Public value 1"]] });
      expect(document.sheets[19]).toMatchObject({
        name: "Visible 20",
        rows: [["Public value 20"]],
      });
      expect(document.sheets.some((sheet) => sheet.name === "Visible 21")).toBe(false);
      expect(document.warnings).toContain("Only the first 20 visible worksheets are shown.");
    },
  );
  it.each([
    { rows: 100, truncated: false },
    { rows: 101, truncated: true },
  ])(
    "detects the row boundary without optional worksheet dimensions: $rows rows",
    ({ rows, truncated }) => {
      const parts = unzipSync(
        workbookBytes(
          "xlsx",
          XLSX.utils.aoa_to_sheet(Array.from({ length: rows }, (_, index) => [`Row ${index + 1}`])),
        ),
      );
      const worksheet = "xl/worksheets/sheet1.xml";
      const xml = strFromU8(parts[worksheet]!);
      // Worksheet dimension metadata is optional; ordinary writers can omit it.
      parts[worksheet] = strToU8(xml.replace(/<dimension\b[^>]*\/>/, ""));
      expect(strFromU8(parts[worksheet])).not.toContain("<dimension");
      const document = parseOfficeDocument(zipSync(parts), "xlsx");
      if (document.kind !== "spreadsheet") throw new Error("Expected spreadsheet");
      expect(document.sheets[0]?.rows).toHaveLength(100);
      expect(document.sheets[0]?.rows[99]).toEqual(["Row 100"]);
      expect(document.sheets[0]?.truncated).toBe(truncated);
      expect(document.sheets[0]?.totalRows).toBe(rows);
    },
  );
  it("decodes Japanese from a pre-Unicode BIFF5 Shift-JIS record", () => {
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([["ABCD"]]), "Legacy");
    const bytes = new Uint8Array(XLSX.write(workbook, { type: "array", bookType: "biff5" }));
    // The BIFF5 writer replaces non-ASCII text. Populate its fixed-width LABEL
    // record with CP932 bytes, then update the workbook CodePage record.
    let labels = 0;
    let codepages = 0;
    for (let offset = 0; offset + 6 <= bytes.length; offset++) {
      if (
        bytes[offset] === 65 &&
        bytes[offset + 1] === 66 &&
        bytes[offset + 2] === 67 &&
        bytes[offset + 3] === 68
      ) {
        bytes.set([0x93, 0xfa, 0x96, 0x7b], offset);
        labels++;
      }
      if (
        bytes[offset] === 0x42 &&
        bytes[offset + 1] === 0 &&
        bytes[offset + 2] === 2 &&
        bytes[offset + 3] === 0 &&
        bytes[offset + 4] === 0xe4 &&
        bytes[offset + 5] === 4
      ) {
        bytes[offset + 4] = 0xa4;
        bytes[offset + 5] = 3;
        codepages++;
      }
    }
    expect(labels).toBe(1);
    expect(codepages).toBe(1);
    const document = parseOfficeDocument(bytes, "xls");
    if (document.kind !== "spreadsheet") throw new Error("Expected spreadsheet");
    expect(document.sheets[0]).toMatchObject({ name: "Legacy", rows: [["日本"]] });
  });
  it.each(["xlsx", "xls"] as const)(
    "reads a genuine %s workbook with Unicode and cached values",
    (format) => {
      const sheet = XLSX.utils.aoa_to_sheet([
        ["品名", "金額"],
        ["りんご", 1234.5],
      ]);
      sheet.B2 = { t: "n", v: 1234.5, z: "#,##0.00", f: "1000+234.5" };
      const document = parseOfficeDocument(
        workbookBytes(format === "xls" ? "biff8" : "xlsx", sheet),
        format,
      );
      expect(document.kind).toBe("spreadsheet");
      if (document.kind !== "spreadsheet") return;
      expect(document.sheets[0]).toMatchObject({ name: "売上", totalRows: 2, totalColumns: 2 });
      expect(document.sheets[0]?.rows).toEqual([
        ["品名", "金額"],
        ["りんご", "1,234.50"],
      ]);
    },
  );

  it("bounds visible rows and columns while retaining full dimensions", () => {
    const rows = Array.from({ length: 105 }, (_, row) =>
      Array.from({ length: 35 }, (_, column) => `${row}:${column}`),
    );
    const document = parseOfficeDocument(
      workbookBytes("xlsx", XLSX.utils.aoa_to_sheet(rows)),
      "xlsx",
    );
    if (document.kind !== "spreadsheet") throw new Error("Expected spreadsheet");
    expect(document.sheets[0]).toMatchObject({ totalRows: 105, totalColumns: 35, truncated: true });
    expect(document.sheets[0]?.rows).toHaveLength(100);
    expect(document.sheets[0]?.rows[99]).toHaveLength(30);
    expect(document.sheets[0]?.rows[99]?.[29]).toBe("99:29");
  });

  it("keeps leading sparse coordinates and signals content beyond the window", () => {
    const sheet = {
      C9000: { t: "s", v: "outside preview" },
      "!ref": "C9000:C9000",
    } satisfies XLSX.WorkSheet;
    const document = parseOfficeDocument(workbookBytes("xlsx", sheet), "xlsx");
    if (document.kind !== "spreadsheet") throw new Error("Expected spreadsheet");
    expect(document.sheets[0]).toMatchObject({ totalRows: 9000, totalColumns: 3, truncated: true });
    expect(document.sheets[0]?.rows).toHaveLength(100);
    expect(document.sheets[0]?.rows[0]).toEqual(["", "", ""]);
  });

  it("caps each cell and aggregate text", () => {
    let seed = 17;
    const value = () =>
      Array.from({ length: 2100 }, () => {
        seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
        return String.fromCharCode(33 + ((seed >>> 16) % 80));
      }).join("");
    const rows = Array.from({ length: 17 }, () => Array.from({ length: 10 }, value));
    const bytes = workbookBytes("xlsx", XLSX.utils.aoa_to_sheet(rows));
    const document = parseOfficeDocument(bytes, "xlsx");
    if (document.kind !== "spreadsheet") throw new Error("Expected spreadsheet");
    const cells = document.sheets.flatMap((sheet) => sheet.rows.flat());
    expect(Math.max(...cells.map((cell) => cell.length))).toBe(2000);
    expect(cells.reduce((total, cell) => total + cell.length * 2, 0)).toBeLessThanOrEqual(
      512 * 1024,
    );
    expect(document.sheets[0]?.truncated).toBe(true);
  });

  it.each(["xlsx", "biff8"] as const)("preserves emoji at %s cell text boundaries", (format) => {
    const values = [
      "a".repeat(1998) + "🐈tail",
      "a".repeat(1997) + "🐈tail",
      "a".repeat(1998) + "🐈",
    ];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet(values.map((value) => [value])),
      "Emoji",
    );
    const document = parseOfficeDocument(
      new Uint8Array(XLSX.write(workbook, { type: "array", bookType: format, bookSST: true })),
      format === "biff8" ? "xls" : format,
    );
    if (document.kind !== "spreadsheet") throw new Error("Expected spreadsheet");
    expect(document.sheets[0]?.rows).toEqual([
      ["a".repeat(1998) + "…"],
      ["a".repeat(1997) + "🐈…"],
      [values[2]],
    ]);
  });

  it("preserves emoji when the aggregate cell text budget is exhausted", () => {
    const values = [...Array<string>(131).fill("a".repeat(2000)), "a".repeat(142) + "🐈tail", "🐈"];
    const rows = Array.from({ length: Math.ceil(values.length / 10) }, (_, row) =>
      values.slice(row * 10, row * 10 + 10),
    );
    const document = parseOfficeDocument(
      workbookBytes("xlsx", XLSX.utils.aoa_to_sheet(rows)),
      "xlsx",
    );
    if (document.kind !== "spreadsheet") throw new Error("Expected spreadsheet");
    const cells = document.sheets[0]!.rows.flat();
    expect(cells.slice(131, 133)).toEqual(["a".repeat(142) + "…", "…"]);
    expect(cells.slice(133).every((cell) => cell === "")).toBe(true);
    expect(cells.reduce((length, cell) => length + cell.length, 0)).toBe((512 * 1024) / 2);
  });

  it("uses a formula expression when no cached result exists", () => {
    const document = parseOfficeDocument(
      workbookBytes("xlsx", { A1: { t: "n", f: "1+1" }, "!ref": "A1" }),
      "xlsx",
    );
    if (document.kind !== "spreadsheet") throw new Error("Expected spreadsheet");
    expect(document.sheets[0]?.rows[0]?.[0]).toBe("=1+1");
  });

  it("does not expose hidden worksheets", () => {
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([["secret"]]), "hidden");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([["public"]]), "visible");
    XLSX.utils.book_set_sheet_visibility(workbook, "hidden", 1);
    const document = parseOfficeDocument(
      new Uint8Array(XLSX.write(workbook, { type: "array", bookType: "xlsx" })),
      "xlsx",
    );
    if (document.kind !== "spreadsheet") throw new Error("Expected spreadsheet");
    expect(document.sheets.map((sheet) => sheet.name)).toEqual(["visible"]);
  });
});

describe("parseOfficeDocument presentation outlines", () => {
  it.each([
    { name: "hero.bin", override: true },
    { name: "hero.data", override: false },
    { name: "hero.jpg", override: true },
  ])(
    "preserves PNG bytes and the whole slide model when named $name (override: $override)",
    ({ name, override }: { name: string; override: boolean }) => {
      const original = visualPresentationParts();
      const bytes = original["ppt/media/picture.png"]!;
      const document = parseOfficeDocument(
        zipSync(presentationWithBitmap(name, bytes, "image/png", override)),
        "pptx",
      );
      expect(document).toEqual(parseOfficeDocument(zipSync(original), "pptx"));
      if (document.kind !== "presentation") throw new Error("Expected presentation");
      const image = document.slides[0]?.elements.find((element) => element.kind === "image");
      if (image?.kind !== "image") throw new Error("Expected slide bitmap");
      expect(image).toMatchObject({
        x: 4000000,
        y: 3000000,
        width: 1000000,
        height: 1000000,
        alt: "テスト画像",
      });
      expect(image.src.startsWith("data:image/png;base64,")).toBe(true);
      expect(
        Uint8Array.from(atob(image.src.split(",")[1]!), (character) => character.charCodeAt(0)),
      ).toEqual(bytes);
    },
  );

  it("keeps actual JPEG bytes and MIME when a declared JPEG image uses an arbitrary part name", () => {
    // A normal Pillow RGB 2x2 JPEG, independently verified before embedding this fixture.
    const bytes = Uint8Array.from(
      atob(
        "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAUDBAQEAwUEBAQFBQUGBwwIBwcHBw8LCwkMEQ8SEhEPERETFhwXExQaFRERGCEYGh0dHx8fExciJCIeJBweHx7/2wBDAQUFBQcGBw4ICA4eFBEUHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh7/wAARCAACAAIDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDmqKKK/VT80P/Z",
      ),
      (character) => character.charCodeAt(0),
    );
    const expected = parseOfficeDocument(
      zipSync(presentationWithBitmap("picture.jpg", bytes, "image/jpeg")),
      "pptx",
    );
    const document = parseOfficeDocument(
      zipSync(presentationWithBitmap("hero.bin", bytes, "image/jpeg")),
      "pptx",
    );
    expect(document).toEqual(expected);
    if (document.kind !== "presentation") throw new Error("Expected presentation");
    const image = document.slides[0]?.elements.find((element) => element.kind === "image");
    if (image?.kind !== "image") throw new Error("Expected JPEG bitmap");
    expect(image.src).toBe(`data:image/jpeg;base64,${btoa(String.fromCharCode(...bytes))}`);
  });

  it("omits non-bitmap bytes even when their part declares a supported image MIME", () => {
    const document = parseOfficeDocument(
      zipSync(presentationWithBitmap("hero.bin", strToU8("ordinary text content"), "image/png")),
      "pptx",
    );
    if (document.kind !== "presentation") throw new Error("Expected presentation");
    expect(document.slides[0]?.elements.some((element) => element.kind === "image")).toBe(false);
    expect(document.slides[0]?.text).toEqual(["日本語の見出し", "本文はマスターから継承"]);
    expect(document.warnings.length).toBeGreaterThan(0);
  });

  it("still omits an ordinary animated GIF when its declared GIF part has a png filename", () => {
    const bytes = Uint8Array.from(
      atob(
        "R0lGODlhAgACAIEAACh4yAAAAAAAAAAAACH/C05FVFNDQVBFMi4wAwEAAAAh+QQACgAAACwAAAAAAgACAAAIBgABCAQQEAAh+QQBCgABACwAAAAAAgACAIHIeCgAAAAAAAAAAAAIBgABCAQQEAA7",
      ),
      (character) => character.charCodeAt(0),
    );
    const document = parseOfficeDocument(
      zipSync(presentationWithBitmap("hero.png", bytes, "image/gif")),
      "pptx",
    );
    const expected = parseOfficeDocument(zipSync(visualPresentationParts()), "pptx");
    if (document.kind !== "presentation" || expected.kind !== "presentation")
      throw new Error("Expected presentation");
    expect(document.slides[0]?.elements).toEqual(
      expected.slides[0]?.elements.filter((element) => element.kind !== "image"),
    );
    expect(document.slides.map((slide) => slide.text)).toEqual(
      expected.slides.map((slide) => slide.text),
    );
    expect(document.warnings.join(" ")).toContain("Only static PNG, JPEG, GIF, and WebP images");
  });

  it.each([
    { location: "deck/main.xml", strict: false },
    { location: "main.xml", strict: false },
    { location: "deck/main.xml", strict: true },
    { location: "main.xml", strict: true },
  ])("resolves the $strict Strict package entry at $location", ({ location, strict }) => {
    const source = strict ? strictPresentationParts() : visualPresentationParts();
    const prefix = location.slice(0, location.lastIndexOf("/") + 1);
    const parts = Object.fromEntries(
      Object.entries(source).map(([name, bytes]) => {
        const relocated =
          name === "ppt/presentation.xml"
            ? location
            : name === "ppt/_rels/presentation.xml.rels"
              ? `${prefix}_rels/main.xml.rels`
              : name.startsWith("ppt/")
                ? prefix + name.slice(4)
                : name;
        return [relocated, bytes];
      }),
    );
    const relationshipNamespace = strict
      ? "http://purl.oclc.org/ooxml/officeDocument/relationships"
      : R;
    parts["_rels/.rels"] = strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="presentation" Type="${relationshipNamespace}/officeDocument" Target="${strict ? "/" : ""}${location}"/></Relationships>`,
    );
    const document = parseOfficeDocument(zipSync(parts), "pptx");
    expect(document).toEqual(parseOfficeDocument(zipSync(source), "pptx"));
    if (document.kind !== "presentation") throw new Error("Expected presentation");
    expect(document.slides.map((slide) => slide.text)).toEqual([
      ["日本語の見出し", "本文はマスターから継承"],
      ["Second & final"],
    ]);
    expect(document.slides[0]?.elements).toContainEqual(
      expect.objectContaining({
        kind: "text",
        text: "本文はマスターから継承",
        x: 600000,
        y: 2000000,
      }),
    );
    expect(
      document.slides[0]?.elements.some(
        (element) => element.kind === "image" && element.src.startsWith("data:image/png;base64,"),
      ),
    ).toBe(true);
  });

  it.each([
    { label: "missing", entries: "", error: "exactly one internal office document" },
    {
      label: "multiple",
      entries: `<Relationship Id="first" Type="${R}/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="second" Type="${R}/officeDocument" Target="ppt/presentation.xml"/>`,
      error: "exactly one internal office document",
    },
    {
      label: "duplicate IDs",
      entries: `<Relationship Id="same" Type="${R}/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="same" Type="${R}/metadata" Target="metadata.xml"/>`,
      error: "Duplicate presentation relationship",
    },
    {
      label: "external",
      entries: `<Relationship Id="presentation" Type="${R}/officeDocument" Target="https://example.com/deck.xml" TargetMode="External"/>`,
      error: "must be internal",
    },
    {
      label: "missing part",
      entries: `<Relationship Id="presentation" Type="${R}/officeDocument" Target="missing.xml"/>`,
      error: "missing its slide index",
    },
  ])("rejects a $label root office document relationship", ({ entries, error }) => {
    const parts = presentationParts();
    parts["_rels/.rels"] = strToU8(`<Relationships xmlns="${REL}">${entries}</Relationships>`);
    expect(() => parseOfficeDocument(zipSync(parts), "pptx")).toThrow(error);
  });

  it("omits external images while preserving internal slide relationships", () => {
    const parts = visualPresentationParts();
    const path = "ppt/slides/_rels/slide2.xml.rels";
    parts[path] = strToU8(
      strFromU8(parts[path]!).replace(
        'Target="../media/picture.png"',
        'Target="https://example.com/picture.png" TargetMode="External"',
      ),
    );
    const document = parseOfficeDocument(zipSync(parts), "pptx");
    if (document.kind !== "presentation") throw new Error("Expected presentation");
    expect(document.slides[0]?.text).toEqual(["日本語の見出し", "本文はマスターから継承"]);
    expect(document.slides[0]?.elements.some((element) => element.kind === "image")).toBe(false);
    expect(document.slides[1]?.text).toEqual(["Second & final"]);
  });

  it.each([
    { type: "ctrTitle", y: 100000 },
    { type: "subTitle", y: 2000000 },
  ])("normalizes $type when inheriting master placement", ({ type, y }) => {
    const parts = visualPresentationParts();
    const decoder = new TextDecoder();
    parts["ppt/slideLayouts/layout.xml"] = strToU8(
      decoder
        .decode(parts["ppt/slideLayouts/layout.xml"])
        .replace('type="title"', `type="${type}"`)
        .replace(/<p:spPr>.*?<\/p:spPr>/, "<p:spPr/>"),
    );
    parts["ppt/slideMasters/master.xml"] = strToU8(
      decoder
        .decode(parts["ppt/slideMasters/master.xml"])
        .replace(
          "<p:spTree>",
          '<p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="title" idx="0"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="200000" y="100000"/><a:ext cx="8000000" cy="1000000"/></a:xfrm></p:spPr></p:sp>',
        ),
    );
    const document = parseOfficeDocument(zipSync(parts), "pptx");
    if (document.kind !== "presentation") throw new Error("Expected presentation");
    expect(document.slides[0]?.elements[0]).toMatchObject({
      kind: "text",
      text: "日本語の見出し",
      y,
    });
  });
  it("falls back from placeholder index to matching type through layout and master", () => {
    const parts = visualPresentationParts();
    parts["ppt/slideLayouts/layout.xml"] = strToU8(
      new TextDecoder().decode(parts["ppt/slideLayouts/layout.xml"]).replace('idx="8"', 'idx="99"'),
    );
    const document = parseOfficeDocument(zipSync(parts), "pptx");
    if (document.kind !== "presentation") throw new Error("Expected presentation");
    expect(document.slides[0]?.elements[1]).toMatchObject({ kind: "text", x: 600000, y: 2000000 });
  });

  it("caps slide count and reports the omitted remainder", () => {
    const parts: Record<string, Uint8Array> = presentationParts();
    parts["ppt/presentation.xml"] = strToU8(
      `<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst>${Array.from({ length: 21 }, (_, index) => `<p:sldId id="${index + 256}" r:id="r${index}"/>`).join("")}</p:sldIdLst></p:presentation>`,
    );
    parts["ppt/_rels/presentation.xml.rels"] = strToU8(
      `<Relationships xmlns="${REL}">${Array.from({ length: 21 }, (_, index) => `<Relationship Id="r${index}" Type="${R}/slide" Target="slides/page${index}.xml"/>`).join("")}</Relationships>`,
    );
    for (let index = 0; index < 21; index++)
      parts[`ppt/slides/page${index}.xml`] = parts["ppt/slides/slide1.xml"]!;
    const document = parseOfficeDocument(zipSync(parts), "pptx");
    if (document.kind !== "presentation") throw new Error("Expected presentation");
    expect(document.slides).toHaveLength(20);
    expect(document.truncated).toBe(true);
  });
  it("builds an ordered visual model with layout/master positions, rotation, Japanese text and a bounded image", () => {
    const document = parseOfficeDocument(zipSync(visualPresentationParts()), "pptx");
    if (document.kind !== "presentation") throw new Error("Expected presentation");
    expect(document.slides[0]?.elements).toMatchObject([
      {
        kind: "text",
        text: "日本語の見出し",
        x: 500000,
        y: 300000,
        width: 8000000,
        height: 1000000,
        fontSize: 36,
        bold: true,
      },
      {
        kind: "text",
        text: "本文はマスターから継承",
        x: 600000,
        y: 2000000,
        width: 7000000,
        height: 2000000,
        fontSize: 24,
        color: "#123456",
      },
      { kind: "shape", shape: "rect", fill: "#ff0000", rotation: 90 },
      { kind: "image", alt: "テスト画像", x: 4000000, y: 3000000 },
    ]);
    const image = document.slides[0]?.elements[3];
    expect(image?.kind === "image" && image.src.startsWith("data:image/png;base64,")).toBe(true);
    expect(document.slides[0]?.width).toBe(9144000);
  });

  it("keeps unresolved-placement text in the outline without overlapping the canvas", () => {
    const document = parseOfficeDocument(zipSync(presentationParts()), "pptx");
    if (document.kind !== "presentation") throw new Error("Expected presentation");
    expect(document.slides[0]?.elements).toEqual([]);
    expect(document.slides[0]?.text).toEqual(["最初\n42"]);
    expect(document.warnings.join(" ")).toContain("no supported placement");
  });

  it("omits images beyond decoded-pixel budgets", () => {
    const parts = visualPresentationParts();
    new DataView(parts["ppt/media/picture.png"]!.buffer).setUint32(16, 5000);
    const document = parseOfficeDocument(zipSync(parts), "pptx");
    if (document.kind !== "presentation") throw new Error("Expected presentation");
    expect(document.slides[0]?.elements.some((element) => element.kind === "image")).toBe(false);
    expect(document.warnings.join(" ")).toContain("image budget");
  });
  it("follows relationship order and extracts fields, breaks, Unicode and XML entities", () => {
    const document = parseOfficeDocument(zipSync(presentationParts()), "pptx");
    expect(document).toMatchObject({
      kind: "presentation",
      truncated: false,
      slides: [
        { number: 1, text: ["最初\n42"] },
        { number: 2, text: ["Second & final"] },
      ],
    });
    expect(document.warnings[0]).toContain("Basic slide preview");
  });

  it("fails rather than showing a deck with missing slides", () => {
    const parts = presentationParts();
    const { "ppt/slides/slide2.xml": _removed, ...remaining } = parts;
    expect(() => parseOfficeDocument(zipSync(remaining), "pptx")).toThrow("missing slide");
  });

  it("caps paragraph text and reports truncation", () => {
    const parts = presentationParts();
    parts["ppt/slides/slide2.xml"] = strToU8(
      `<p:sld xmlns:p="${P}" xmlns:a="${A}"><a:p><a:r><a:t>${"a".repeat(3000)}</a:t></a:r></a:p></p:sld>`,
    );
    const document = parseOfficeDocument(zipSync(parts), "pptx");
    if (document.kind !== "presentation") throw new Error("Expected presentation");
    expect(document.slides[0]?.text[0]).toHaveLength(2000);
    expect(document.slides[0]?.truncated).toBe(true);
    expect(document.truncated).toBe(true);
  });

  it("preserves emoji at paragraph boundaries across text runs", () => {
    const parts = presentationParts();
    const values = ["a".repeat(1998) + "🐈", "a".repeat(1997) + "🐈"];
    parts["ppt/slides/slide2.xml"] = strToU8(
      `<p:sld xmlns:p="${P}" xmlns:a="${A}">${values.map((value) => `<a:p><a:r><a:t>${value}</a:t></a:r><a:r><a:t>tail</a:t></a:r></a:p>`).join("")}<a:p><a:r><a:t>${values[0]}</a:t></a:r></a:p></p:sld>`,
    );
    const document = parseOfficeDocument(zipSync(parts), "pptx");
    if (document.kind !== "presentation") throw new Error("Expected presentation");
    expect(document.slides[0]?.text).toEqual([
      "a".repeat(1998) + "…",
      "a".repeat(1997) + "🐈…",
      values[0],
    ]);
    expect(document.slides[0]?.truncated).toBe(true);
  });

  it("rejects entity declarations before processing XML", () => {
    const parts = presentationParts();
    parts["ppt/slides/slide1.xml"] = strToU8('<!DOCTYPE x [<!ENTITY e "expanded">]><x>&e;</x>');
    expect(() => parseOfficeDocument(zipSync(parts), "pptx")).toThrow("document types");
  });
});

describe("Office archive boundaries", () => {
  it("reads legacy version 3 workbook values despite uninitialized high stream sizes", () => {
    const bytes = workbookBytes("biff8");
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    expect(view.getUint16(26, true)).toBe(3);
    const directoryOffset = (view.getUint32(48, true) + 1) * 512;
    for (let offset = directoryOffset; offset < directoryOffset + 512; offset += 128) {
      if (bytes[offset + 66] === 2 || bytes[offset + 66] === 5)
        view.setUint32(offset + 124, 0x12345678, true);
    }
    const document = parseOfficeDocument(bytes, "xls");
    if (document.kind !== "spreadsheet") throw new Error("Expected spreadsheet");
    expect(document.sheets[0]?.name).toBe("売上");
    expect(document.sheets[0]?.rows).toEqual([
      ["品名", "価格"],
      ["りんご", "125"],
    ]);
  });

  it("rejects full 64-bit version 4 stream sizes while reading the bounded baseline", () => {
    const rows = [
      ["品名", "価格"],
      ["りんご", 125],
      ...Array.from({ length: 30 }, (_, row) => [`商品${row}`, "日本語".repeat(30)]),
    ];
    const original = workbookBytes("biff8", XLSX.utils.aoa_to_sheet(rows));
    const cfb = XLSX.CFB.read(original, { type: "array" }) as {
      FileIndex: { name: string; content: Uint8Array }[];
    };
    const workbook = cfb.FileIndex.find((entry) => entry.name === "Workbook")!.content;
    expect(workbook.length).toBeGreaterThanOrEqual(4096);
    const streamSectors = Math.ceil(workbook.length / 4096);
    const bytes = new Uint8Array((3 + streamSectors) * 4096);
    bytes.set(original.subarray(0, 512));
    const view = new DataView(bytes.buffer);
    view.setUint16(26, 4, true);
    view.setUint16(30, 12, true);
    for (const [offset, value] of [
      [40, 1],
      [44, 1],
      [48, 0],
      [60, 0xfffffffe],
      [64, 0],
      [68, 0xfffffffe],
      [72, 0],
    ])
      view.setUint32(offset!, value!, true);
    for (let index = 0; index < 109; index++)
      view.setUint32(76 + index * 4, index ? 0xffffffff : 1, true);
    for (let index = 0; index < 32; index++) {
      const offset = 4096 + index * 128;
      for (const edge of [68, 72, 76]) view.setUint32(offset + edge, 0xffffffff, true);
    }
    for (const [index, name] of ["Root Entry", "Workbook"].entries()) {
      const offset = 4096 + index * 128;
      for (let char = 0; char < name.length; char++)
        view.setUint16(offset + char * 2, name.charCodeAt(char), true);
      view.setUint16(offset + 64, (name.length + 1) * 2, true);
      bytes[offset + 66] = index ? 2 : 5;
      bytes[offset + 67] = 1;
      view.setUint32(offset + 116, index ? 2 : 0xfffffffe, true);
      view.setUint32(offset + 120, index ? workbook.length : 0, true);
    }
    view.setUint32(4096 + 76, 1, true);
    for (let index = 0; index < 1024; index++) view.setUint32(8192 + index * 4, 0xffffffff, true);
    view.setUint32(8192, 0xfffffffe, true);
    view.setUint32(8196, 0xfffffffd, true);
    for (let index = 0; index < streamSectors; index++)
      view.setUint32(8200 + index * 4, index + 1 === streamSectors ? 0xfffffffe : index + 3, true);
    bytes.set(workbook, 3 * 4096);
    const document = parseOfficeDocument(bytes, "xls");
    if (document.kind !== "spreadsheet") throw new Error("Expected spreadsheet");
    expect(document.sheets[0]?.name).toBe("売上");
    expect(document.sheets[0]?.rows.slice(0, 2)).toEqual([
      ["品名", "価格"],
      ["りんご", "125"],
    ]);
    view.setUint32(4096 + 128 + 124, 1, true);
    expect(() => parseOfficeDocument(bytes, "xls")).toThrow("streams exceed preview limits");
  });

  it("rejects unbounded binary workbook stream sizes before invoking SheetJS", () => {
    const bytes = workbookBytes("biff8");
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const directoryOffset = (view.getUint32(48, true) + 1) * 512;
    for (let offset = directoryOffset; offset < directoryOffset + 512; offset += 128) {
      if (bytes[offset + 66] === 2) {
        view.setUint32(offset + 120, bytes.length + 1, true);
        break;
      }
    }
    expect(() => parseOfficeDocument(bytes, "xls")).toThrow("streams exceed preview limits");
  });

  it("rejects cyclic binary workbook directories before invoking SheetJS", () => {
    const bytes = workbookBytes("biff8");
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const sector = view.getUint32(48, true);
    const fat = view.getUint32(76, true);
    view.setUint32((fat + 1) * 512 + sector * 4, sector, true);
    expect(() => parseOfficeDocument(bytes, "xls")).toThrow("allocation chain");
  });
  it("accepts a streamed ZIP using data descriptors after comparing both directories", () => {
    const chunks: Uint8Array[] = [];
    const archive = new Zip((error, data) => {
      if (error) throw error;
      chunks.push(data);
    });
    for (const [name, bytes] of Object.entries(presentationParts())) {
      const entry = new ZipDeflate(name);
      archive.add(entry);
      entry.push(bytes, true);
    }
    archive.end();
    const bytes = new Uint8Array(chunks.reduce((total, chunk) => total + chunk.length, 0));
    let offset = 0;
    for (const chunk of chunks) {
      bytes.set(chunk, offset);
      offset += chunk.length;
    }
    expect(parseOfficeDocument(bytes, "pptx").kind).toBe("presentation");
  });

  it("rejects excessive XML nesting before passing it to a document decoder", () => {
    const parts = presentationParts();
    parts["ppt/slides/slide1.xml"] = strToU8("<a>".repeat(129) + "x" + "</a>".repeat(129));
    expect(() => parseOfficeDocument(zipSync(parts), "pptx")).toThrow("too complex");
  });
  it("rejects oversized input and non-Office bytes", () => {
    expect(() => parseOfficeDocument(new Uint8Array(OFFICE_INPUT_MAX_BYTES + 1), "xlsx")).toThrow(
      "5 MiB",
    );
    expect(() => parseOfficeDocument(strToU8("not an Excel file"), "xls")).toThrow("binary Excel");
    expect(() => parseOfficeDocument(strToU8("not a zip"), "pptx")).toThrow("ZIP");
  });

  it("rejects compressed expansion bombs before reading document XML", () => {
    const bytes = zipSync({ bomb: new Uint8Array(2 * 1024 * 1024) });
    expect(() => parseOfficeDocument(bytes, "xlsx")).toThrow("preview limits");
  });

  it("rejects too many entries", () => {
    const files = Object.fromEntries(
      Array.from({ length: 501 }, (_, index) => [`file${index}`, new Uint8Array(0)]),
    );
    expect(() => parseOfficeDocument(zipSync(files), "xlsx")).toThrow("directory");
  });

  it("rejects duplicate case aliases and path traversal", () => {
    expect(() =>
      parseOfficeDocument(zipSync({ a: strToU8("1"), A: strToU8("2") }), "xlsx"),
    ).toThrow("duplicate");
    expect(() => parseOfficeDocument(zipSync({ "../escape": strToU8("1") }), "xlsx")).toThrow(
      "unsafe",
    );
  });

  it("rejects local/central name and size discrepancies", () => {
    const names = zipSync({ first: strToU8("test") });
    names[30] = "z".charCodeAt(0);
    expect(() => parseOfficeDocument(names, "xlsx")).toThrow("disagree");
    const sizes = zipSync({ first: strToU8("test") });
    new DataView(sizes.buffer).setUint32(centralOffset(sizes) + 24, 1, true);
    expect(() => parseOfficeDocument(sizes, "xlsx")).toThrow("sizes disagree");
  });

  it("checks actual expanded bytes even when both declared sizes lie", () => {
    const bytes = zipSync({ first: strToU8("test") });
    const view = new DataView(bytes.buffer);
    view.setUint32(22, 1, true);
    view.setUint32(centralOffset(bytes) + 24, 1, true);
    expect(() => parseOfficeDocument(bytes, "xlsx")).toThrow("expanded document");
  });

  it("rejects encrypted and symlink entries", () => {
    const encrypted = zipSync({ a: strToU8("test") });
    const encryptedView = new DataView(encrypted.buffer);
    encryptedView.setUint16(centralOffset(encrypted) + 8, 1, true);
    expect(() => parseOfficeDocument(encrypted, "xlsx")).toThrow("unsupported entries");
    const symlink = zipSync({ a: strToU8("test") });
    new DataView(symlink.buffer).setUint32(centralOffset(symlink) + 38, 0xa1ff0000, true);
    expect(() => parseOfficeDocument(symlink, "xlsx")).toThrow("unsupported entries");
  });

  it("verifies checksums rather than trusting the directory", () => {
    const bytes = zipSync({ first: strToU8("test") }, { level: 0 });
    bytes[35] = 65;
    expect(() => parseOfficeDocument(bytes, "xlsx")).toThrow("checksum");
  });
});
