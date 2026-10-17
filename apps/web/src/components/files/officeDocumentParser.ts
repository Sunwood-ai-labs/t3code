import { readImageDimensions } from "@t3tools/shared/imageDimensions";
import { Unzip, UnzipInflate, Zip, ZipDeflate } from "fflate";
import { SaxesParser, type SaxesTagNS } from "saxes";
import * as XLSX from "xlsx";
import * as codepages from "xlsx/dist/cpexcel.full.mjs";

XLSX.set_cptable(codepages);

export type OfficeDocumentFormat = "xlsx" | "xls" | "pptx";

interface SlideGeometry {
  x: number;
  y: number;
  width: number;
  height: number;
  rotation?: number;
}

export type OfficeSlideElement = SlideGeometry &
  (
    | {
        kind: "text";
        text: string;
        fontSize: number;
        color?: string;
        bold?: boolean;
        align?: "left" | "center" | "right";
      }
    | { kind: "image"; src: string; alt: string }
    | { kind: "shape"; shape: "rect" | "ellipse"; fill?: string; stroke?: string }
  );

export type OfficeDocument =
  | {
      kind: "spreadsheet";
      sheets: {
        name: string;
        rows: string[][];
        totalRows: number;
        totalColumns: number;
        truncated: boolean;
      }[];
      warnings: string[];
    }
  | {
      kind: "presentation";
      slides: {
        number: number;
        text: string[];
        truncated: boolean;
        width: number;
        height: number;
        elements: OfficeSlideElement[];
        background?: string;
      }[];
      truncated: boolean;
      warnings: string[];
    };

export interface OfficeWorkerRequest {
  id: number;
  name: string;
  format: OfficeDocumentFormat;
  buffer: ArrayBuffer;
}

export type OfficeWorkerResponse =
  | { readonly type: "ready" }
  | { id: number; ok: true; document: OfficeDocument }
  | { id: number; ok: false; error: string };

export const OFFICE_INPUT_MAX_BYTES = 5 * 1024 * 1024;
const ARCHIVE_MAX_BYTES = 32 * 1024 * 1024;
const ENTRY_MAX_BYTES = 8 * 1024 * 1024;
const MAX_ENTRIES = 500;
const MAX_RATIO = 200;
const RATIO_GRACE_BYTES = 1024 * 1024;
const MAX_XML_ELEMENTS = 200_000;
const MAX_ROWS = 100;
const MAX_COLUMNS = 30;
const MAX_SHEETS = 20;
const MAX_SLIDES = 20;
const MAX_CELL_CHARS = 2000;
const MAX_TEXT_BYTES = 512 * 1024;
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const DOCUMENT_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

function normalizeOfficeUri(uri: string): string {
  if (uri === "http://purl.oclc.org/ooxml/drawingml/main") return DRAWING_NS;
  if (uri === "http://purl.oclc.org/ooxml/presentationml/main") return PRESENTATION_NS;
  if (uri === "http://purl.oclc.org/ooxml/spreadsheetml/main") return SPREADSHEET_NS;
  const strictRelationships = "http://purl.oclc.org/ooxml/officeDocument/relationships";
  return uri === strictRelationships || uri.startsWith(`${strictRelationships}/`)
    ? DOCUMENT_REL_NS + uri.slice(strictRelationships.length)
    : uri;
}

function normalizeOfficeTag(tag: SaxesTagNS): SaxesTagNS {
  return {
    ...tag,
    uri: normalizeOfficeUri(tag.uri),
    attributes: Object.fromEntries(
      Object.entries(tag.attributes).map(([name, value]) => [
        name,
        { ...value, uri: normalizeOfficeUri(value.uri) },
      ]),
    ),
  };
}
const CRC_TABLE = Uint32Array.from({ length: 256 }, (_, value) => {
  let crc = value;
  for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
  return crc >>> 0;
});

function fail(message: string): never {
  throw new Error(message);
}

function containsControlCharacters(value: string): boolean {
  return [...value].some((character) => character.charCodeAt(0) < 32);
}

interface ZipEntry {
  name: string;
  compressedSize: number;
  expandedSize: number;
  offset: number;
  dataEnd: number;
  flags: number;
  crc: number;
}

/** Require one unambiguous ZIP directory before passing its bytes to either decoder. */
function archiveDirectory(bytes: Uint8Array): Map<string, ZipEntry> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u16 = (offset: number) => view.getUint16(offset, true);
  const u32 = (offset: number) => view.getUint32(offset, true);
  let end = -1;
  for (let offset = bytes.length - 22; offset >= Math.max(0, bytes.length - 65_557); offset--) {
    if (u32(offset) === 0x06054b50 && offset + 22 + u16(offset + 20) === bytes.length) {
      end = offset;
      break;
    }
  }
  if (end < 0) fail("The document is not a supported ZIP archive.");
  const count = u16(end + 10);
  const directorySize = u32(end + 12);
  const directoryOffset = u32(end + 16);
  if (
    u16(end + 4) !== 0 ||
    u16(end + 6) !== 0 ||
    u16(end + 8) !== count ||
    count === 0 ||
    count > MAX_ENTRIES ||
    directoryOffset + directorySize !== end
  ) {
    fail("The archive directory exceeds preview limits or uses an unsupported format.");
  }
  const entries = new Map<string, ZipEntry>();
  const aliases = new Set<string>();
  const decoder = new TextDecoder("utf-8", { fatal: true });
  let cursor = directoryOffset;
  let total = 0;
  function extraFields(start: number, length: number) {
    const stop = start + length;
    for (let position = start; position < stop;) {
      if (position + 4 > stop) fail("Invalid ZIP extra field.");
      const type = u16(position);
      const size = u16(position + 2);
      if (type === 0x0001 || type === 0x7075 || position + 4 + size > stop) {
        fail("ZIP64 and alternative ZIP path names are not supported for preview.");
      }
      position += 4 + size;
    }
  }
  for (let index = 0; index < count; index++) {
    if (cursor + 46 > end || u32(cursor) !== 0x02014b50) fail("Invalid ZIP directory entry.");
    const flags = u16(cursor + 8);
    const method = u16(cursor + 10);
    const compressedSize = u32(cursor + 20);
    const expandedSize = u32(cursor + 24);
    const nameLength = u16(cursor + 28);
    const extraLength = u16(cursor + 30);
    const commentLength = u16(cursor + 32);
    const offset = u32(cursor + 42);
    const entryEnd = cursor + 46 + nameLength + extraLength + commentLength;
    if (
      entryEnd > end ||
      nameLength === 0 ||
      (flags & ~0x080e) !== 0 ||
      (method !== 0 && method !== 8) ||
      u16(cursor + 34) !== 0 ||
      ((u32(cursor + 38) >>> 16) & 0xf000) === 0xa000 ||
      expandedSize > ENTRY_MAX_BYTES ||
      (expandedSize > RATIO_GRACE_BYTES && expandedSize > compressedSize * MAX_RATIO)
    ) {
      fail("The archive contains unsupported entries or exceeds preview limits.");
    }
    const nameBytes = bytes.subarray(cursor + 46, cursor + 46 + nameLength);
    if ((flags & 0x0800) === 0 && nameBytes.some((byte) => byte >= 128)) {
      fail("Non-UTF-8 ZIP path names are not supported for preview.");
    }
    const name = decoder.decode(nameBytes);
    const parts = name.replace(/\/$/, "").split("/");
    if (
      /[\\:]/.test(name) ||
      containsControlCharacters(name) ||
      parts.some((part) => part === "" || part === "." || part === "..") ||
      aliases.has(name.toLowerCase())
    ) {
      fail("The archive contains duplicate or unsafe path names.");
    }
    aliases.add(name.toLowerCase());
    extraFields(cursor + 46 + nameLength, extraLength);
    if (offset + 30 > directoryOffset || u32(offset) !== 0x04034b50) {
      fail("Invalid ZIP local header.");
    }
    const localNameLength = u16(offset + 26);
    const localExtraLength = u16(offset + 28);
    const dataStart = offset + 30 + localNameLength + localExtraLength;
    const dataEnd = dataStart + compressedSize;
    if (
      u16(offset + 6) !== flags ||
      u16(offset + 8) !== method ||
      localNameLength !== nameLength ||
      dataEnd > directoryOffset ||
      nameBytes.some((byte, position) => bytes[offset + 30 + position] !== byte)
    ) {
      fail("The ZIP local header and directory disagree.");
    }
    extraFields(offset + 30 + localNameLength, localExtraLength);
    const crc = u32(cursor + 16);
    if ((flags & 8) === 0) {
      if (
        u32(offset + 14) !== crc ||
        u32(offset + 18) !== compressedSize ||
        u32(offset + 22) !== expandedSize
      ) {
        fail("The ZIP local header and directory sizes disagree.");
      }
    } else {
      const descriptor = dataEnd + (u32(dataEnd) === 0x08074b50 ? 4 : 0);
      if (
        descriptor + 12 > directoryOffset ||
        u32(descriptor) !== crc ||
        u32(descriptor + 4) !== compressedSize ||
        u32(descriptor + 8) !== expandedSize
      ) {
        fail("Invalid ZIP data descriptor.");
      }
    }
    total += expandedSize;
    if (total > ARCHIVE_MAX_BYTES) fail("The expanded document exceeds the 32 MiB preview limit.");
    entries.set(name, { name, compressedSize, expandedSize, offset, dataEnd, flags, crc });
    cursor = entryEnd;
  }
  if (cursor !== end) fail("Invalid ZIP directory size.");
  const ordered = [...entries.values()].toSorted((left, right) => left.offset - right.offset);
  let expected = 0;
  for (const entry of ordered) {
    if (entry.offset !== expected) fail("Overlapping or unlisted ZIP records are not supported.");
    expected = entry.dataEnd;
    if ((entry.flags & 8) !== 0) expected += u32(expected) === 0x08074b50 ? 16 : 12;
  }
  if (expected !== directoryOffset) fail("Unlisted ZIP data is not supported.");
  return entries;
}

function expandArchive(bytes: Uint8Array): Map<string, Uint8Array> {
  const directory = archiveDirectory(bytes);
  const result = new Map<string, Uint8Array>();
  const seen = new Set<string>();
  let total = 0;
  const unzip = new Unzip((file) => {
    const entry = directory.get(file.name);
    if (!entry || seen.has(file.name)) fail("The ZIP local records and directory disagree.");
    seen.add(file.name);
    const chunks: Uint8Array[] = [];
    let size = 0;
    file.ondata = (error, chunk, final) => {
      if (error) throw error;
      size += chunk.length;
      total += chunk.length;
      if (
        size > ENTRY_MAX_BYTES ||
        size > entry.expandedSize ||
        total > ARCHIVE_MAX_BYTES ||
        (size > RATIO_GRACE_BYTES && size > entry.compressedSize * MAX_RATIO)
      ) {
        fail("The expanded document exceeds preview limits.");
      }
      chunks.push(chunk);
      if (final) {
        if (size !== entry.expandedSize) fail("The expanded ZIP entry size is invalid.");
        const content = new Uint8Array(size);
        let offset = 0;
        for (const part of chunks) {
          content.set(part, offset);
          offset += part.length;
        }
        let crc = 0xffffffff;
        for (const byte of content) crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ byte) & 255]!;
        if ((crc ^ 0xffffffff) >>> 0 !== entry.crc) fail("The ZIP entry checksum is invalid.");
        result.set(file.name, content);
      }
    };
    file.start();
  });
  unzip.register(UnzipInflate);
  // Small compressed pushes bound work before each expanded-size callback can reject a bomb.
  for (let offset = 0; offset < bytes.length; offset += 512) {
    unzip.push(bytes.subarray(offset, offset + 512), offset + 512 >= bytes.length);
  }
  if (result.size !== directory.size) fail("The ZIP archive is incomplete.");
  return result;
}

type XmlEncoding = "utf-8" | "utf-16le" | "utf-16be";

function xmlEncoding(bytes: Uint8Array): { encoding: XmlEncoding; bom: boolean } {
  if (
    (bytes[0] === 0xff && bytes[1] === 0xfe && bytes[2] === 0 && bytes[3] === 0) ||
    (bytes[0] === 0 && bytes[1] === 0 && bytes[2] === 0xfe && bytes[3] === 0xff)
  ) {
    fail("UTF-32 XML is not supported for preview.");
  }
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return { encoding: "utf-16le", bom: true };
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return { encoding: "utf-16be", bom: true };
  if (bytes[0] === 0x3c && bytes[1] === 0 && bytes[2] === 0x3f && bytes[3] === 0)
    return { encoding: "utf-16le", bom: false };
  if (bytes[0] === 0 && bytes[1] === 0x3c && bytes[2] === 0 && bytes[3] === 0x3f)
    return { encoding: "utf-16be", bom: false };
  return { encoding: "utf-8", bom: bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf };
}

function parseXml(
  bytes: Uint8Array,
  handlers: {
    open?: (tag: SaxesTagNS) => void;
    close?: (tag: SaxesTagNS) => void;
    text?: (text: string) => void;
  } = {},
) {
  if (bytes.length > ENTRY_MAX_BYTES) fail("The document XML exceeds preview limits.");
  const { encoding, bom } = xmlEncoding(bytes);
  const parser = new SaxesParser({ xmlns: true });
  let depth = 0;
  let elements = 0;
  let declaredEncoding: string | undefined;
  parser.on("xmldecl", (declaration) => {
    declaredEncoding = declaration.encoding?.toLowerCase();
    if (declaredEncoding === undefined) return;
    if (declaredEncoding === "utf-16") {
      if (encoding === "utf-8" || !bom) fail("UTF-16 XML requires a matching byte-order mark.");
    } else if (
      declaredEncoding === "utf-8" ||
      declaredEncoding === "utf8" ||
      declaredEncoding === "utf-16le" ||
      declaredEncoding === "utf-16be"
    ) {
      if ((declaredEncoding === "utf8" ? "utf-8" : declaredEncoding) !== encoding)
        fail("The XML encoding declaration disagrees with its bytes.");
    } else fail("The XML encoding is not supported for preview.");
  });
  parser.on("doctype", () => fail("XML document types are not supported for preview."));
  parser.on("error", (error) => {
    throw error;
  });
  parser.on("opentag", (tag) => {
    if (++depth > 128 || ++elements > MAX_XML_ELEMENTS)
      fail("The document XML is too complex to preview.");
    handlers.open?.(normalizeOfficeTag(tag));
  });
  parser.on("closetag", (tag) => {
    handlers.close?.(normalizeOfficeTag(tag));
    depth--;
  });
  parser.on("text", (text) => handlers.text?.(text));
  parser.on("cdata", (text) => handlers.text?.(text));
  const decoder = new TextDecoder(encoding, { fatal: true });
  for (let offset = 0; offset < bytes.length; offset += 4096) {
    parser.write(decoder.decode(bytes.subarray(offset, offset + 4096), { stream: true }));
  }
  parser.write(decoder.decode()).close();
  if (encoding !== "utf-8" && !bom && !declaredEncoding)
    fail("XML without a byte-order mark must declare its UTF-16 byte order.");
  return encoding;
}

/** Normalize validated XML for SheetJS without retaining a second set of parts. */
function normalizedWorkbookArchive(
  parts: Map<string, Uint8Array>,
  encodings: Map<string, XmlEncoding>,
): Uint8Array {
  const chunks: Uint8Array[] = [];
  let size = 0;
  let expanded = 0;
  const zip = new Zip((error, chunk) => {
    if (error) throw error;
    if (size + chunk.length > ARCHIVE_MAX_BYTES)
      fail("The normalized workbook exceeds preview limits.");
    if (chunk.length) chunks.push(chunk);
    size += chunk.length;
  });
  for (const [name, original] of parts) {
    const encoding = encodings.get(name);
    const content = encoding
      ? new TextEncoder().encode(
          new TextDecoder(encoding, { fatal: true })
            .decode(original)
            .replace(/^(<\?xml\s[^?]*?\bencoding\s*=\s*)(["'])[^"']+\2/, '$1"UTF-8"'),
        )
      : original;
    expanded += content.length;
    if (content.length > ENTRY_MAX_BYTES || expanded > ARCHIVE_MAX_BYTES)
      fail("The normalized workbook XML exceeds preview limits.");
    const entry = new ZipDeflate(name, { level: 1 });
    zip.add(entry);
    if (content.length === 0) entry.push(content, true);
    for (let offset = 0; offset < content.length; offset += 512)
      entry.push(content.subarray(offset, offset + 512), offset + 512 >= content.length);
  }
  zip.end();
  const output = new Uint8Array(size);
  let offset = 0;
  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.length;
  }
  return output;
}

function attribute(tag: SaxesTagNS, local: string, uri = ""): string | undefined {
  return Object.values(tag.attributes).find((value) => value.local === local && value.uri === uri)
    ?.value;
}

function resolvePart(base: string, target: string): string {
  if (/[\\?#:]/.test(target) || containsControlCharacters(target))
    fail("Unsupported presentation relationship target.");
  const parts = target.startsWith("/") ? [] : base.split("/").slice(0, -1);
  for (const part of target.split("/")) {
    if (part === "" || part === ".") continue;
    if (part === "..") {
      if (parts.length === 0) fail("A presentation relationship leaves its archive.");
      parts.pop();
    } else parts.push(part);
  }
  return parts.join("/");
}

function textPrefix(text: string, length: number): string {
  let end = Math.min(text.length, length);
  if (
    end > 0 &&
    text.charCodeAt(end - 1) >= 0xd800 &&
    text.charCodeAt(end - 1) <= 0xdbff &&
    text.charCodeAt(end) >= 0xdc00 &&
    text.charCodeAt(end) <= 0xdfff
  )
    end--;
  return text.slice(0, end);
}

function textBudget() {
  let remaining = MAX_TEXT_BYTES / 2;
  let truncated = false;
  return {
    take(text: string) {
      const length = Math.min(MAX_CELL_CHARS, remaining);
      const value =
        text.length > length
          ? textPrefix(text, Math.max(0, length - 1)) + (length ? "…" : "")
          : text;
      remaining -= value.length;
      truncated ||= value !== text;
      return value;
    },
    get truncated() {
      return truncated;
    },
  };
}

const CFB_FREE = 0xffffffff;
const CFB_END = 0xfffffffe;
const CFB_FAT = 0xfffffffd;

/** Validate even unreferenced allocations before the CFB decoder scans physical sectors. */
function validateSectorGraph(table: Uint32Array, reserved: Set<number>): Uint32Array {
  const incoming = new Uint8Array(table.length);
  for (let sector = 0; sector < table.length; sector++) {
    const next = table[sector]!;
    if (reserved.has(sector)) {
      if (next !== CFB_FAT) fail("Invalid workbook reserved sector allocation.");
      continue;
    }
    if (next === CFB_FREE || next === CFB_END) continue;
    if (next >= table.length || reserved.has(next) || table[next] === CFB_FREE) {
      fail("Invalid workbook allocation sector reference.");
    }
    if (++incoming[next]! > 1) fail("Shared workbook allocation chain tails are not supported.");
  }
  const states = new Uint8Array(table.length);
  const lengths = new Uint32Array(table.length);
  for (let start = 0; start < table.length; start++) {
    if (states[start]) continue;
    const path: number[] = [];
    let current = start;
    while (current < table.length && states[current] === 0) {
      states[current] = 1;
      path.push(current);
      current = table[current]!;
    }
    if (current < table.length && states[current] === 1) fail("Cyclic workbook allocation chain.");
    let length = current < table.length ? lengths[current]! : 0;
    for (let index = path.length - 1; index >= 0; index--) {
      const sector = path[index]!;
      lengths[sector] = ++length;
      states[sector] = 2;
    }
  }
  return lengths;
}

function validateCompoundWorkbook(bytes: Uint8Array) {
  const signature = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  if (bytes.length < 512 || signature.some((value, index) => bytes[index] !== value)) {
    fail("This file is not a supported binary Excel workbook.");
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u32 = (offset: number) => view.getUint32(offset, true);
  const sectorShift = view.getUint16(30, true);
  const major = view.getUint16(26, true);
  if (
    view.getUint16(28, true) !== 0xfffe ||
    !((major === 3 && sectorShift === 9) || (major === 4 && sectorShift === 12)) ||
    view.getUint16(32, true) !== 6
  ) {
    fail("Unsupported binary workbook storage format.");
  }
  const sectorSize = 2 ** sectorShift;
  const sectorCount = bytes.length / sectorSize - 1;
  const fatCount = u32(44);
  // A valid file within 5 MiB needs fewer than the header's 109 FAT slots.
  if (
    !Number.isInteger(sectorCount) ||
    sectorCount < 1 ||
    fatCount < 1 ||
    fatCount > 109 ||
    fatCount > sectorCount ||
    u32(72) !== 0 ||
    (u32(68) !== CFB_END && u32(68) !== CFB_FREE) ||
    u32(64) > sectorCount ||
    u32(56) !== 4096
  ) {
    fail("Binary workbook allocation exceeds preview limits.");
  }
  const fatSectors = Array.from({ length: fatCount }, (_, index) => u32(76 + index * 4));
  if (new Set(fatSectors).size !== fatCount || fatSectors.some((sector) => sector >= sectorCount))
    fail("Invalid workbook allocation table.");
  const slots = fatCount * (sectorSize / 4);
  if (slots < sectorCount) fail("Incomplete workbook allocation table.");
  const allocation = new Uint32Array(sectorCount);
  for (let index = 0; index < sectorCount; index++) {
    const value = u32(
      (fatSectors[Math.floor(index / (sectorSize / 4))]! + 1) * sectorSize +
        (index % (sectorSize / 4)) * 4,
    );
    allocation[index] = value;
  }
  const reserved = new Set(fatSectors);
  const lengths = validateSectorGraph(allocation, reserved);
  const directoryStart = u32(48);
  if (directoryStart >= sectorCount || reserved.has(directoryStart))
    fail("Invalid workbook directory sector.");
  // SheetJS can copy a visited suffix again when physical scan order encounters
  // a fragmented chain backwards. Predict those copies without revisiting tails.
  const copied = new Uint8Array(sectorCount);
  let copiedBytes = 0;
  for (let index = 0; index < sectorCount; index++) {
    let sector = (index + directoryStart) % sectorCount;
    if (copied[sector]) continue;
    copiedBytes += lengths[sector]! * sectorSize;
    if (copiedBytes > ARCHIVE_MAX_BYTES) fail("Workbook allocation copy budget exceeds 32 MiB.");
    while (sector < sectorCount && !copied[sector]) {
      copied[sector] = 1;
      sector = allocation[sector]!;
    }
  }
  const owned = new Uint8Array(sectorCount);
  for (const sector of reserved) owned[sector] = 1;
  function claimChain(start: number, count?: number): number[] {
    const chain: number[] = [];
    for (let sector = start; sector !== CFB_END; sector = allocation[sector]!) {
      if (sector >= sectorCount || owned[sector] || allocation[sector] === CFB_FREE)
        fail("Overlapping or invalid workbook sector chains.");
      owned[sector] = 1;
      chain.push(sector);
    }
    if (count !== undefined && chain.length !== count)
      fail("Workbook sector chain length disagrees with its stream size.");
    return chain;
  }
  const directoryChain = claimChain(directoryStart);
  const directory: { kind: number; size: number; start: number; edges: number[] }[] = [];
  let total = 0;
  for (const sector of directoryChain) {
    for (
      let offset = (sector + 1) * sectorSize;
      offset < (sector + 2) * sectorSize;
      offset += 128
    ) {
      const kind = bytes[offset + 66]!;
      const size = u32(offset + 120);
      if (kind !== 0 && kind !== 1 && kind !== 2 && kind !== 5)
        fail("Invalid workbook directory entry.");
      if (kind !== 2 && kind !== 5 && size !== 0)
        fail("Unused workbook directory entries and storage entries must have zero stream size.");
      if (
        kind !== 0 &&
        (directory.length >= MAX_ENTRIES ||
          // Legacy version 3 writers may leave the unused high size DWORD uninitialized.
          (major === 4 && u32(offset + 124) !== 0) ||
          size > bytes.length ||
          (total += size) > ARCHIVE_MAX_BYTES)
      )
        fail("Binary workbook streams exceed preview limits.");
      const edges = [u32(offset + 68), u32(offset + 72), u32(offset + 76)];
      // SheetJS traverses directory pointers even on unused entries.
      if (kind === 0 && edges.some((target) => target !== CFB_FREE))
        fail("Unused workbook directory references are not supported.");
      directory.push({ kind, size, start: u32(offset + 116), edges });
    }
  }
  if (directory[0]?.kind !== 5 || directory.slice(1).some((entry) => entry.kind === 5))
    fail("Invalid workbook root directory.");
  const directoryIncoming = new Uint8Array(directory.length);
  for (let index = 0; index < directory.length; index++) {
    const entry = directory[index]!;
    if (!entry.kind) continue;
    for (const target of entry.edges) {
      if (target === CFB_FREE) continue;
      if (
        target >= directory.length ||
        target === 0 ||
        !directory[target]?.kind ||
        ++directoryIncoming[target]! > 1
      )
        fail("Invalid or shared workbook directory reference.");
    }
  }
  const reached = new Uint8Array(directory.length);
  const queue = [0];
  for (let index = 0; index < queue.length; index++) {
    const current = queue[index]!;
    if (reached[current]) fail("Cyclic workbook directory references.");
    reached[current] = 1;
    for (const target of directory[current]!.edges) if (target !== CFB_FREE) queue.push(target);
  }
  if (directory.some((entry, index) => entry.kind && !reached[index]))
    fail("Unreachable or cyclic workbook directory entries.");
  const miniFatCount = u32(64);
  const miniFatStart = u32(60);
  const miniFatChain = miniFatCount ? claimChain(miniFatStart, miniFatCount) : [];
  if (!miniFatCount && miniFatStart !== CFB_END && miniFatStart !== CFB_FREE)
    fail("Invalid workbook mini allocation table.");
  const root = directory[0]!;
  const rootChain = root.size ? claimChain(root.start, Math.ceil(root.size / sectorSize)) : [];
  const miniCount = Math.ceil(root.size / 64);
  const miniAllocation = new Uint32Array(miniCount);
  if (miniCount > miniFatChain.length * (sectorSize / 4))
    fail("Incomplete workbook mini allocation table.");
  for (let index = 0; index < miniFatChain.length * (sectorSize / 4); index++) {
    const value = u32(
      (miniFatChain[Math.floor(index / (sectorSize / 4))]! + 1) * sectorSize +
        (index % (sectorSize / 4)) * 4,
    );
    if (index < miniCount) miniAllocation[index] = value;
    else if (value !== CFB_FREE && value !== CFB_END)
      fail("Invalid workbook mini allocation padding.");
  }
  validateSectorGraph(miniAllocation, new Set());
  const miniOwned = new Uint8Array(miniCount);
  for (const entry of directory) {
    if (entry.kind !== 2 || !entry.size) continue;
    if (entry.size >= 4096) {
      claimChain(entry.start, Math.ceil(entry.size / sectorSize));
      continue;
    }
    let count = 0;
    for (let sector = entry.start; sector !== CFB_END; sector = miniAllocation[sector]!) {
      if (sector >= miniCount || miniOwned[sector] || miniAllocation[sector] === CFB_FREE)
        fail("Overlapping or invalid workbook mini sector chains.");
      miniOwned[sector] = 1;
      count++;
    }
    if (count !== Math.ceil(entry.size / 64))
      fail("Workbook mini chain length disagrees with its stream size.");
  }
  if (rootChain.length === 0 && miniCount !== 0) fail("Invalid workbook mini stream.");
}

function visibleSheetIndices(bytes: Uint8Array): number[] {
  // Unlike bookSheets, an empty worksheet selection retains visibility metadata.
  const metadata = XLSX.read(bytes, { type: "array", sheets: [], sheetRows: MAX_ROWS + 1 });
  return metadata.SheetNames.flatMap((_, index) =>
    metadata.Workbook?.Sheets?.[index]?.Hidden ? [] : [index],
  );
}

function parseSpreadsheet(bytes: Uint8Array, format: "xlsx" | "xls"): OfficeDocument {
  const selectedIndices = format === "xlsx" ? visibleSheetIndices(bytes) : null;
  if (selectedIndices?.length === 0) fail("The workbook has no visible worksheets to preview.");
  const workbook = XLSX.read(bytes, {
    type: "array",
    sheetRows: MAX_ROWS + 1,
    ...(selectedIndices ? { sheets: selectedIndices.slice(0, MAX_SHEETS) } : {}),
    cellFormula: true,
    cellText: true,
    cellHTML: false,
    cellStyles: false,
    bookVBA: false,
  });
  const budget = textBudget();
  const visibleNames = workbook.SheetNames.filter(
    (name, index) => !workbook.Workbook?.Sheets?.[index]?.Hidden && workbook.Sheets[name],
  );
  const sheets = visibleNames.slice(0, MAX_SHEETS).map((name) => {
    const sheet = workbook.Sheets[name]!;
    const ref = sheet["!fullref"] ?? sheet["!ref"];
    const range = ref ? XLSX.utils.decode_range(ref) : null;
    const totalRows = range ? range.e.r + 1 : 0;
    const totalColumns = range ? range.e.c + 1 : 0;
    if (
      !Number.isSafeInteger(totalRows) ||
      !Number.isSafeInteger(totalColumns) ||
      totalRows < 0 ||
      totalColumns < 0 ||
      totalRows > 1_048_576 ||
      totalColumns > 16_384
    )
      fail("Invalid worksheet dimensions.");
    let truncated = totalRows > MAX_ROWS || totalColumns > MAX_COLUMNS;
    const rows: string[][] = [];
    for (let row = 0; row < Math.min(totalRows, MAX_ROWS); row++) {
      const cells: string[] = [];
      for (let column = 0; column < Math.min(totalColumns, MAX_COLUMNS); column++) {
        const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })] as
          | XLSX.CellObject
          | undefined;
        const value = cell
          ? (cell.w ?? (cell.v != null ? String(cell.v) : cell.f ? `=${cell.f}` : ""))
          : "";
        const bounded = budget.take(value);
        truncated ||= bounded !== value;
        cells.push(bounded);
      }
      rows.push(cells);
    }
    return { name: textPrefix(name, MAX_CELL_CHARS), rows, totalRows, totalColumns, truncated };
  });
  if (sheets.length === 0) fail("The workbook has no visible worksheets to preview.");
  const warnings = ["Values only; styles, charts, and formula recalculation are not included."];
  if ((selectedIndices?.length ?? visibleNames.length) > MAX_SHEETS)
    warnings.push(`Only the first ${MAX_SHEETS} visible worksheets are shown.`);
  if (sheets.some((sheet) => sheet.truncated))
    warnings.push("Tables are limited to the first 100 rows and 30 columns.");
  if (budget.truncated)
    warnings.push("Cell text is limited to 2,000 characters and 512 KiB in total.");
  return { kind: "spreadsheet", sheets, warnings };
}

interface XmlNode {
  tag: SaxesTagNS;
  children: XmlNode[];
  value: string;
}

function xmlTree(bytes: Uint8Array): XmlNode {
  const stack: XmlNode[] = [];
  let root: XmlNode | undefined;
  let count = 0;
  parseXml(bytes, {
    open(tag) {
      if (++count > 20_000) fail("The slide XML is too complex to preview.");
      const node: XmlNode = { tag, children: [], value: "" };
      if (stack.length) stack.at(-1)!.children.push(node);
      else root = node;
      stack.push(node);
    },
    text(value) {
      const current = stack.at(-1);
      if (current) current.value += value;
    },
    close() {
      stack.pop();
    },
  });
  if (!root) fail("Empty presentation XML.");
  return root;
}

function children(node: XmlNode | undefined, local: string, uri = DRAWING_NS): XmlNode[] {
  return node?.children.filter((child) => child.tag.uri === uri && child.tag.local === local) ?? [];
}

function child(node: XmlNode | undefined, local: string, uri = DRAWING_NS): XmlNode | undefined {
  return children(node, local, uri)[0];
}

function descendants(node: XmlNode | undefined, local: string, uri = DRAWING_NS): XmlNode[] {
  if (!node) return [];
  const result: XmlNode[] = [];
  const visit = (current: XmlNode) => {
    if (current.tag.uri === uri && current.tag.local === local) result.push(current);
    for (const nested of current.children) visit(nested);
  };
  visit(node);
  return result;
}

function nodeAttribute(node: XmlNode | undefined, local: string, uri = "") {
  return node ? attribute(node.tag, local, uri) : undefined;
}

function hiddenObject(node: XmlNode): boolean {
  const properties = node.children.find(
    (candidate) => candidate.tag.uri === PRESENTATION_NS && candidate.tag.local.startsWith("nv"),
  );
  const hidden = nodeAttribute(child(properties, "cNvPr", PRESENTATION_NS), "hidden");
  return hidden === "1" || hidden === "true";
}

function visibleParagraphs(node: XmlNode): XmlNode[] {
  const result: XmlNode[] = [];
  const visit = (current: XmlNode) => {
    if (hiddenObject(current)) return;
    if (current.tag.uri === DRAWING_NS && current.tag.local === "p") result.push(current);
    for (const nested of current.children) visit(nested);
  };
  visit(node);
  return result;
}

function solidColor(node: XmlNode | undefined): string | undefined {
  const value = nodeAttribute(child(child(node, "solidFill"), "srgbClr"), "val");
  return value && /^[a-f\d]{6}$/i.test(value) ? `#${value}` : undefined;
}

function placeholder(shape: XmlNode | undefined) {
  return descendants(shape, "ph", PRESENTATION_NS)[0];
}

function inheritedShape(
  shape: XmlNode | undefined,
  container: XmlNode | undefined,
  matchBy: "index" | "type" = "index",
): XmlNode | undefined {
  const ph = placeholder(shape);
  if (!ph) return undefined;
  const idx = nodeAttribute(ph, "idx") ?? "0";
  const rawType = nodeAttribute(ph, "type") ?? "body";
  const type =
    rawType === "title" || rawType === "ctrTitle"
      ? "title"
      : ["dt", "ftr", "hdr", "sldNum"].includes(rawType)
        ? rawType
        : "body";
  const shapes = descendants(container, "sp", PRESENTATION_NS).filter((candidate) =>
    placeholder(candidate),
  );
  const byType = () =>
    shapes.find((candidate) => (nodeAttribute(placeholder(candidate), "type") ?? "body") === type);
  // Slide indices identify layout placeholders; master placeholders inherit a normalized type.
  return matchBy === "type"
    ? byType()
    : shapes.find((candidate) => (nodeAttribute(placeholder(candidate), "idx") ?? "0") === idx);
}

function geometry(
  chain: (XmlNode | undefined)[],
  slideWidth: number,
  slideHeight: number,
): SlideGeometry | undefined {
  const transforms = chain.map((shape) => child(child(shape, "spPr", PRESENTATION_NS), "xfrm"));
  const transform = transforms.find(
    (candidate) => child(candidate, "off") && child(candidate, "ext"),
  );
  if (!transform) return undefined;
  const x = Number(nodeAttribute(child(transform, "off"), "x"));
  const y = Number(nodeAttribute(child(transform, "off"), "y"));
  const width = Number(nodeAttribute(child(transform, "ext"), "cx"));
  const height = Number(nodeAttribute(child(transform, "ext"), "cy"));
  const rotation =
    Number(
      transforms
        .map((candidate) => nodeAttribute(candidate, "rot"))
        .find((value) => value !== undefined) ?? 0,
    ) / 60_000;
  if (
    ![x, y, width, height, rotation].every(Number.isFinite) ||
    width <= 0 ||
    height <= 0 ||
    width > slideWidth * 4 ||
    height > slideHeight * 4 ||
    Math.abs(x) > slideWidth * 4 ||
    Math.abs(y) > slideHeight * 4
  )
    return undefined;
  return { x, y, width, height, rotation };
}

function paragraphText(node: XmlNode): string {
  let text = "";
  const visit = (current: XmlNode) => {
    if (text.length > MAX_CELL_CHARS) return;
    if (current.tag.uri === DRAWING_NS && current.tag.local === "t")
      text += textPrefix(current.value, MAX_CELL_CHARS + 2 - text.length);
    else if (current.tag.uri === DRAWING_NS && current.tag.local === "br") text += "\n";
    else for (const nested of current.children) visit(nested);
  };
  visit(node);
  return text;
}

function imageSource(bytes: Uint8Array): string | undefined {
  const signature = String.fromCharCode(...bytes.subarray(0, 12));
  const mime = signature.startsWith("\x89PNG\r\n\x1a\n")
    ? "image/png"
    : bytes[0] === 0xff && bytes[1] === 0xd8
      ? "image/jpeg"
      : signature.startsWith("GIF87a") || signature.startsWith("GIF89a")
        ? "image/gif"
        : signature.startsWith("RIFF") && signature.slice(8, 12) === "WEBP"
          ? "image/webp"
          : undefined;
  if (!mime) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (mime === "image/png") {
    for (let offset = 8; offset + 12 <= bytes.length;) {
      const length = view.getUint32(offset);
      if (offset + length + 12 > bytes.length) return undefined;
      if (String.fromCharCode(...bytes.subarray(offset + 4, offset + 8)) === "acTL")
        return undefined;
      offset += length + 12;
    }
  } else if (mime === "image/webp") {
    for (let offset = 12; offset + 8 <= bytes.length;) {
      const type = String.fromCharCode(...bytes.subarray(offset, offset + 4));
      const length = view.getUint32(offset + 4, true);
      if (offset + length + 8 > bytes.length || type === "ANIM" || type === "ANMF")
        return undefined;
      offset += length + 8 + (length & 1);
    }
  } else if (mime === "image/gif") {
    let offset = 13 + (bytes[10]! & 128 ? 3 * (1 << ((bytes[10]! & 7) + 1)) : 0);
    let frames = 0;
    const skipBlocks = () => {
      while (offset < bytes.length) {
        const length = bytes[offset++]!;
        if (length === 0) return;
        offset += length;
      }
    };
    while (offset < bytes.length) {
      const marker = bytes[offset++]!;
      if (marker === 0x3b) break;
      if (marker === 0x21) {
        offset++;
        skipBlocks();
      } else if (marker === 0x2c && offset + 9 <= bytes.length) {
        if (++frames > 1) return undefined;
        const flags = bytes[offset + 8]!;
        offset += 9 + (flags & 128 ? 3 * (1 << ((flags & 7) + 1)) : 0) + 1;
        skipBlocks();
      } else return undefined;
    }
    if (frames !== 1) return undefined;
  }
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 8192)
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
  return `data:${mime};base64,${btoa(binary)}`;
}

function partRelationships(path: string, tree: (path: string) => XmlNode | undefined) {
  const split = path.lastIndexOf("/");
  const relPath = `${path.slice(0, split + 1)}_rels/${path.slice(split + 1)}.rels`;
  const relationships = new Map<string, { type: string; path: string; target: string }>();
  const ids = new Set<string>();
  for (const node of children(tree(relPath), "Relationship", PACKAGE_REL_NS)) {
    const id = nodeAttribute(node, "Id");
    const target = nodeAttribute(node, "Target");
    const type = nodeAttribute(node, "Type");
    if (!id || !target || !type) fail("Invalid presentation relationship.");
    if (ids.has(id)) fail("Duplicate presentation relationship.");
    ids.add(id);
    const normalizedType = normalizeOfficeUri(type);
    const mode = nodeAttribute(node, "TargetMode");
    if (mode === "External") {
      if (!path && normalizedType === `${DOCUMENT_REL_NS}/officeDocument`)
        fail("The office document relationship must be internal.");
      continue;
    }
    if (mode && mode !== "Internal") fail("Invalid presentation relationship target mode.");
    relationships.set(id, { type: normalizedType, path: resolvePart(path, target), target });
  }
  return relationships;
}

function officeDocumentPath(parts: Map<string, Uint8Array>): string {
  const relationships = partRelationships("", (path) => {
    const bytes = parts.get(path);
    return bytes ? xmlTree(bytes) : undefined;
  });
  const officeDocuments = [...relationships.values()].filter(
    (rel) => rel.type === `${DOCUMENT_REL_NS}/officeDocument`,
  );
  if (officeDocuments.length !== 1)
    fail("The document must have exactly one internal office document relationship.");
  return officeDocuments[0]!.path;
}

function normalizeWorkbookRelationships(
  parts: Map<string, Uint8Array>,
  path: string,
  workbook: XmlNode,
): string | undefined {
  const tree = (name: string) => {
    const bytes = parts.get(name);
    return bytes ? xmlTree(bytes) : undefined;
  };
  const workbookTypes = new Set([
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    "application/vnd.ms-excel.sheet.binary.macroEnabled.main",
    "application/vnd.ms-excel.addin.macroEnabled.main+xml",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
  ]);
  const candidates = children(
    tree("[Content_Types].xml"),
    "Override",
    "http://schemas.openxmlformats.org/package/2006/content-types",
  ).filter((node) => workbookTypes.has(nodeAttribute(node, "ContentType") ?? ""));
  // SheetJS selects content types first, with a legacy fallback to xl/workbook.xml.
  if (
    candidates.length > 1 ||
    (candidates.length
      ? nodeAttribute(candidates[0], "PartName") !== `/${path}`
      : path !== "xl/workbook.xml")
  )
    fail("The workbook content type disagrees with its office document relationship.");
  const relationships = partRelationships(path, tree);
  const directory = path.slice(0, path.lastIndexOf("/") + 1);
  const archiveNames = new Map([...parts.keys()].map((name) => [name.toLowerCase(), name]));
  const find = (name: string) => archiveNames.get(name.toLowerCase());
  let changed = false;
  for (const sheet of descendants(workbook, "sheet", SPREADSHEET_NS)) {
    const id = nodeAttribute(sheet, "id", DOCUMENT_REL_NS);
    const relationship = id ? relationships.get(id) : undefined;
    if (!relationship || !find(relationship.path))
      fail("The workbook references a missing worksheet.");
    const target = relationship.target;
    const selected = ["xl/" + target.replace(/\/?xl\//, ""), target, directory + target]
      .map(find)
      .find((name) => name !== undefined);
    changed ||= selected !== find(relationship.path);
    const preferred = find("xl/" + relationship.path.replace(/\/?xl\//, ""));
    if (preferred && preferred !== find(relationship.path))
      fail("The workbook worksheet paths are ambiguous for preview.");
  }
  if (!changed) return undefined;
  const escape = (value: string) =>
    value.replace(/[&<"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", '"': "&quot;" })[char]!);
  let xml = `<Relationships xmlns="${PACKAGE_REL_NS}">`;
  for (const [id, rel] of relationships) {
    const entry = `<Relationship Id="${escape(id)}" Type="${escape(rel.type)}" Target="${escape(rel.path)}"/>`;
    if (xml.length + entry.length + "</Relationships>".length > ENTRY_MAX_BYTES)
      fail("The normalized workbook XML exceeds preview limits.");
    xml += entry;
  }
  xml += "</Relationships>";
  const bytes = new TextEncoder().encode(xml);
  if (bytes.length > ENTRY_MAX_BYTES) fail("The normalized workbook XML exceeds preview limits.");
  const relPath = `${directory}_rels/${path.slice(directory.length)}.rels`;
  parts.set(relPath, bytes);
  return relPath;
}

function parsePresentation(parts: Map<string, Uint8Array>): OfficeDocument {
  const trees = new Map<string, XmlNode>();
  const tree = (path: string | undefined) => {
    if (!path) return undefined;
    const bytes = parts.get(path);
    if (!bytes) return undefined;
    let result = trees.get(path);
    if (!result) {
      result = xmlTree(bytes);
      trees.set(path, result);
    }
    return result;
  };
  const relationships = (path: string) => partRelationships(path, tree);
  const related = (path: string, type: string) =>
    [...relationships(path).values()].find((rel) => rel.type === `${DOCUMENT_REL_NS}/${type}`)
      ?.path;
  const presentationPath = officeDocumentPath(parts);
  const presentation = tree(presentationPath);
  if (
    !presentation ||
    presentation.tag.local !== "presentation" ||
    presentation.tag.uri !== PRESENTATION_NS
  )
    fail("The presentation is missing its slide index.");
  const slideSize = child(presentation, "sldSz", PRESENTATION_NS);
  const width = Number(nodeAttribute(slideSize, "cx") ?? 9_144_000);
  const height = Number(nodeAttribute(slideSize, "cy") ?? 5_143_500);
  if (
    ![width, height].every(
      (value) => Number.isFinite(value) && value >= 100_000 && value <= 100_000_000,
    )
  )
    fail("Unsupported slide dimensions.");
  const slideRelationships = relationships(presentationPath);
  const paths = descendants(presentation, "sldId", PRESENTATION_NS).map((node) => {
    const id = nodeAttribute(node, "id", DOCUMENT_REL_NS);
    const rel = id ? slideRelationships.get(id) : undefined;
    if (!rel || rel.type !== `${DOCUMENT_REL_NS}/slide` || !parts.has(rel.path))
      fail("The presentation references a missing slide.");
    return rel.path;
  });
  if (paths.length === 0) fail("The presentation has no slides.");
  const budget = textBudget();
  const warnings = new Set([
    "Basic slide preview; theme styling, charts, grouped shapes, and animations may differ or be omitted. Use the outline or original file for complete content.",
  ]);
  let imagePixels = 0;
  let imageTextBytes = 0;
  const slides = paths.slice(0, MAX_SLIDES).map((path, index) => {
    const slide = tree(path)!;
    const layoutPath = related(path, "slideLayout");
    const layout = tree(layoutPath);
    const masterPath = layoutPath ? related(layoutPath, "slideMaster") : undefined;
    const master = tree(masterPath);
    let background = "#ffffff";
    for (const source of [slide, layout, master]) {
      const color = solidColor(
        child(child(source, "cSld", PRESENTATION_NS), "bg", PRESENTATION_NS)?.children.find(
          (node) => node.tag.local === "bgPr",
        ),
      );
      if (color) {
        background = color;
        break;
      }
    }
    let truncated = false;
    const paragraphValues = new Map<XmlNode, string>();
    const paragraphs = visibleParagraphs(slide);
    const text = paragraphs.slice(0, 100).flatMap((paragraph) => {
      const original = paragraphText(paragraph);
      const bounded = budget.take(original);
      truncated ||= bounded !== original;
      paragraphValues.set(paragraph, bounded);
      return bounded.trim() ? [bounded] : [];
    });
    truncated ||= paragraphs.length > 100;
    const elements: OfficeSlideElement[] = [];
    const shapeTree = child(child(slide, "cSld", PRESENTATION_NS), "spTree", PRESENTATION_NS);
    for (const shape of shapeTree?.children ?? []) {
      if (shape.tag.uri !== PRESENTATION_NS) continue;
      if (hiddenObject(shape)) continue;
      if (shape.tag.local === "nvGrpSpPr" || shape.tag.local === "grpSpPr") continue;
      if (elements.length >= 100) {
        truncated = true;
        break;
      }
      if (shape.tag.local !== "sp" && shape.tag.local !== "pic") {
        warnings.add(
          "Some grouped shapes, charts, or other slide objects are available only in the original file; extracted text remains in the outline.",
        );
        continue;
      }
      const layoutShape = inheritedShape(shape, layout);
      const masterShape = inheritedShape(layoutShape ?? shape, master, "type");
      const chain = [shape, layoutShape, masterShape];
      const box = geometry(chain, width, height);
      if (!box) {
        warnings.add(
          "Some objects have no supported placement and are omitted from the canvas; their text remains in the outline.",
        );
        continue;
      }
      if (shape.tag.local === "pic") {
        const embed = nodeAttribute(descendants(shape, "blip")[0], "embed", DOCUMENT_REL_NS);
        const imageRelationship = embed ? relationships(path).get(embed) : undefined;
        const imagePath =
          imageRelationship?.type === `${DOCUMENT_REL_NS}/image`
            ? imageRelationship.path
            : undefined;
        const imageBytes = imagePath ? parts.get(imagePath) : undefined;
        const dimensions = imageBytes ? readImageDimensions(imageBytes) : null;
        if (
          !imagePath ||
          !imageBytes ||
          !dimensions ||
          dimensions.width > 4096 ||
          dimensions.height > 4096 ||
          imagePixels + dimensions.width * dimensions.height > 4_000_000 ||
          imageTextBytes + Math.ceil(imageBytes.length / 3) * 4 + 64 > 2 * 1024 * 1024
        ) {
          warnings.add(
            "Some images are unsupported or exceed the preview image budget (4 megapixels and 2 MiB in total).",
          );
          continue;
        }
        const src = imageSource(imageBytes);
        if (!src) {
          warnings.add("Only static PNG, JPEG, GIF, and WebP images are shown.");
          continue;
        }
        imagePixels += dimensions.width * dimensions.height;
        imageTextBytes += src.length;
        const alt =
          nodeAttribute(descendants(shape, "cNvPr", PRESENTATION_NS)[0], "descr") ?? "Slide image";
        elements.push({ kind: "image", ...box, src, alt: textPrefix(alt, 200) });
        if (descendants(shape, "srcRect").length)
          warnings.add("Image cropping is simplified in this preview.");
        continue;
      }
      const properties = child(shape, "spPr", PRESENTATION_NS);
      const preset = nodeAttribute(child(properties, "prstGeom"), "prst");
      const fill = solidColor(properties);
      const stroke = solidColor(child(properties, "ln"));
      if ((preset === "rect" || preset === "ellipse") && (fill || stroke))
        elements.push({
          kind: "shape",
          ...box,
          shape: preset,
          ...(fill ? { fill } : {}),
          ...(stroke ? { stroke } : {}),
        });
      else if (preset && preset !== "rect" && preset !== "ellipse")
        warnings.add("Some shape geometry is simplified or omitted.");
      const shapeText = visibleParagraphs(shape)
        .map((paragraph) => paragraphValues.get(paragraph) ?? "")
        .filter(Boolean)
        .join("\n");
      if (shapeText && elements.length >= 100) truncated = true;
      if (shapeText && elements.length < 100) {
        const inheritedRun = chain.flatMap((candidate) => [
          ...descendants(candidate, "rPr"),
          ...descendants(candidate, "defRPr"),
        ]);
        const size = inheritedRun
          .map((run) => nodeAttribute(run, "sz"))
          .find((value) => value !== undefined);
        const placeholderType = chain
          .map((candidate) => nodeAttribute(placeholder(candidate), "type"))
          .find(Boolean);
        const fontSize = size
          ? Number(size) / 100
          : placeholderType === "ctrTitle" || placeholderType === "title"
            ? 32
            : 18;
        const color = inheritedRun.map(solidColor).find(Boolean) ?? "#111111";
        const bold =
          inheritedRun
            .map((run) => nodeAttribute(run, "b"))
            .find((value) => value !== undefined) === "1";
        const alignment = nodeAttribute(descendants(shape, "pPr")[0], "algn");
        elements.push({
          kind: "text",
          ...box,
          text: shapeText,
          fontSize: Number.isFinite(fontSize) ? Math.max(1, Math.min(fontSize, 400)) : 18,
          color,
          bold,
          align: alignment === "ctr" ? "center" : alignment === "r" ? "right" : "left",
        });
      }
    }
    return { number: index + 1, text, truncated, width, height, elements, background };
  });
  const truncated =
    paths.length > MAX_SLIDES || budget.truncated || slides.some((slide) => slide.truncated);
  if (truncated)
    warnings.add(
      "Preview limited to 20 slides, 100 elements and paragraphs per slide, 2,000 characters per paragraph, and 512 KiB of text.",
    );
  return { kind: "presentation", slides, truncated, warnings: [...warnings] };
}

/** Called inside a disposable worker; the caller owns cancellation and its time limit. */
export function parseOfficeDocument(
  bytes: Uint8Array,
  format: OfficeDocumentFormat,
): OfficeDocument {
  if (bytes.length === 0 || bytes.length > OFFICE_INPUT_MAX_BYTES) {
    fail("Office preview is limited to files up to 5 MiB.");
  }
  if (format === "xls") {
    validateCompoundWorkbook(bytes);
    return parseSpreadsheet(bytes, "xls");
  }
  const parts = expandArchive(bytes);
  const encodings = new Map<string, XmlEncoding>();
  for (const [name, content] of parts) {
    if (/\.(?:xml|rels)$/i.test(name)) {
      const encoding = parseXml(content);
      if (encoding !== "utf-8") encodings.set(name, encoding);
    }
  }
  if (format === "pptx") return parsePresentation(parts);
  const workbookPath = officeDocumentPath(parts);
  const workbookBytes = parts.get(workbookPath);
  const workbook = workbookBytes ? xmlTree(workbookBytes) : undefined;
  if (!workbook || workbook.tag.local !== "workbook" || workbook.tag.uri !== SPREADSHEET_NS)
    fail("This file is not a supported Excel workbook.");
  const normalizedRelationships = normalizeWorkbookRelationships(parts, workbookPath, workbook);
  if (normalizedRelationships) encodings.delete(normalizedRelationships);
  return parseSpreadsheet(
    encodings.size || normalizedRelationships ? normalizedWorkbookArchive(parts, encodings) : bytes,
    "xlsx",
  );
}
