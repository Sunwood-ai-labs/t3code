import { describe, expect, it } from "vite-plus/test";

import {
  filePreviewKind,
  decodeFilePreviewText,
  FILE_TEXT_PREVIEW_MAX_BYTES,
  hostPreviewMimeTypeFromExtension,
  isWorkspaceAudioPreviewPath,
  isWorkspaceBrowserPreviewPath,
  isWorkspaceImagePreviewPath,
  isWorkspacePreviewEntryPath,
  isWorkspaceVideoPreviewPath,
  mediaKindFromPath,
  officePreviewFormat,
  officePreviewMimeTypeFromExtension,
} from "./filePreview.ts";

describe("bounded web Office previews", () => {
  it.each([
    {
      name: "report.xlsx",
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      expected: "xlsx",
    },
    { name: "report.xlsx", mimeType: "application/vnd.ms-excel", expected: "xlsx" },
    { name: "report.xls", mimeType: "application/vnd.ms-excel", expected: "xls" },
    {
      name: "report.xls",
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      expected: "xls",
    },
    {
      name: "slides.pptx",
      mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      expected: "pptx",
    },
    { name: "slides.pptx", mimeType: "application/vnd.ms-powerpoint", expected: "pptx" },
    {
      name: "report.xlsx",
      mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      expected: "xlsx",
    },
    { name: "slides.pptx", mimeType: "application/vnd.ms-excel", expected: "pptx" },
    { name: "REPORT.XlSx", mimeType: " Application/Vnd.Ms-Excel ; version=1 ", expected: "xlsx" },
    {
      name: "REPORT.XLS",
      mimeType:
        " APPLICATION/VND.OPENXMLFORMATS-OFFICEDOCUMENT.SPREADSHEETML.SHEET; charset=binary",
      expected: "xls",
    },
    {
      name: "SLIDES.PpTx",
      mimeType: " APPLICATION/VND.MS-POWERPOINT ; version=1",
      expected: "pptx",
    },
  ])(
    "selects the supported filename format for $name with Office MIME $mimeType",
    ({ name, mimeType, expected }: { name: string; mimeType: string; expected: string }) => {
      expect(officePreviewFormat({ name, mimeType })).toBe(expected);
      expect(filePreviewKind({ name, mimeType })).toBe("unsupported");
    },
  );

  it.each([
    { name: "report.xlsx", mimeType: "image/png" },
    { name: "report.xls", mimeType: "application/zip" },
    { name: "slides.pptx", mimeType: "text/html" },
    { name: "report.xlsx", mimeType: " Application/PDF ; version=1" },
  ])(
    "preserves a non-Office content type conflict for $name ($mimeType)",
    ({ name, mimeType }: { name: string; mimeType: string }) => {
      expect(officePreviewFormat({ name, mimeType })).toBeNull();
    },
  );

  it.each(["legacy.ppt", "download"])(
    "keeps legacy PowerPoint MIME unsupported for $name",
    (name: string) => {
      expect(officePreviewFormat({ name, mimeType: "application/vnd.ms-powerpoint" })).toBeNull();
    },
  );

  it.each(["xlsx", "xls", "pptx"] as const)(
    "recognizes %s without changing native fallback",
    (format) => {
      const name = `report.${format.toUpperCase()}`;
      expect(officePreviewFormat({ name })).toBe(format);
      expect(filePreviewKind({ name })).toBe("unsupported");
      expect(isWorkspacePreviewEntryPath(name)).toBe(true);
      const mimeType = officePreviewMimeTypeFromExtension(`.${format}`)!;
      expect(hostPreviewMimeTypeFromExtension(`.${format}`)).toBe(mimeType);
      expect(officePreviewFormat({ name: "download", mimeType: `${mimeType}; version=1` })).toBe(
        format,
      );
      expect(filePreviewKind({ name, mimeType })).toBe("unsupported");
    },
  );

  it.each([
    "legacy.ppt",
    "report.xlsx.exe",
    "report.xlsx?download=1",
    "report%2Exlsx",
    "report.pptm",
  ])("keeps %s on the existing fallback", (name) => {
    expect(officePreviewFormat({ name })).toBeNull();
    expect(isWorkspacePreviewEntryPath(name)).toBe(false);
  });

  it("honors a specific conflicting content type", () => {
    expect(officePreviewFormat({ name: "report.xlsx", mimeType: "image/png" })).toBeNull();
    expect(officePreviewMimeTypeFromExtension(".ppt")).toBeNull();
  });
});

describe("workspace file previews", () => {
  it.each(["report.html", "report.HTM", "document#draft.pdf", "reports?old/document.pdf"])(
    "recognizes browser preview path %s",
    (path) => {
      expect(isWorkspaceBrowserPreviewPath(path)).toBe(true);
      expect(isWorkspacePreviewEntryPath(path)).toBe(true);
    },
  );

  it.each([
    "icon.png",
    "photo.JPEG",
    "animation.gif",
    "vector#mark.svg",
    "photo?edited.JPEG",
    "images#archive/icon.png",
    "texture.webp",
    "image.avif",
  ])("recognizes image preview path %s", (path) => {
    expect(isWorkspaceImagePreviewPath(path)).toBe(true);
    expect(isWorkspacePreviewEntryPath(path)).toBe(true);
  });

  it.each([
    "README.md",
    "src/index.ts",
    "image.png.ts",
    "png",
    "image.png#notes.txt",
    "image.svg?notes.txt",
    "document.pdf?download=1",
    "report.html#notes.txt",
    "image%2Epng",
  ])("rejects non-preview path %s", (path) => {
    expect(isWorkspacePreviewEntryPath(path)).toBe(false);
  });

  it("serves audio in place from the host like video and browser documents", () => {
    expect(isWorkspaceAudioPreviewPath("notes/recording.WAV")).toBe(true);
    expect(isWorkspaceAudioPreviewPath("recording.wav.ts")).toBe(false);
    expect(hostPreviewMimeTypeFromExtension(".m4a")).toBe("audio/mp4");
    expect(hostPreviewMimeTypeFromExtension(".mp4")).toBe("video/mp4");
    expect(hostPreviewMimeTypeFromExtension(".txt")).toBeNull();
  });
});

describe("media path parsing", () => {
  it.each([
    ["https://cdn.example/clip.webm?download=1#t=2", "video"],
    ["https://example.com/download?name=recording.mp4", null],
    ["https://example.png", null],
    ["images%2Fresult%2Epng", "image"],
    ["images/result%23v2.png", "image"],
    ["images/result.png%23secret.txt", null],
    ["images/result.png%3Fsecret.txt", null],
    ["/tmp/100%.png", "image"],
  ])("classifies the decoded pathname of %s", (source, kind) => {
    expect(mediaKindFromPath(source)).toBe(kind);
  });

  it.each([
    ["recording.mp4#t=2", "video", false],
    ["recording%2Emp4", "video", false],
    ["recording#take2.mp4", null, true],
    ["recording?take2.mp4", null, true],
  ])("distinguishes authored URLs from literal filenames in %s", (source, kind, literalVideo) => {
    expect(mediaKindFromPath(source)).toBe(kind);
    expect(isWorkspaceVideoPreviewPath(source)).toBe(literalVideo);
  });
});

describe("attachment preview classification", () => {
  it.each([
    ["example.json", "application/octet-stream", "text"],
    ["README.md", "text/plain", "markdown"],
    ["component.tsx", "", "text"],
    ["report.pdf", "application/pdf", "pdf"],
    ["page.HTML", "", "html"],
    ["recording.mp3", "", "audio"],
    ["payload", "application/problem+json", "text"],
    ["archive.zip", "application/zip", "unsupported"],
    ["misleading.json", "application/pdf", "pdf"],
    ["misleading.pdf", "application/zip", "unsupported"],
  ])("classifies %s (%s) as %s", (name, mimeType, expected) => {
    expect(filePreviewKind({ name, mimeType })).toBe(expected);
  });
  it("rejects binary and invalid UTF-8 data", () => {
    expect(() => decodeFilePreviewText(new Uint8Array([65, 0, 66]))).toThrow("binary");
    expect(() => decodeFilePreviewText(new Uint8Array([255]))).toThrow("UTF-8");
  });
  it("does not corrupt a multi-byte character at the preview boundary", () => {
    const bytes = new Uint8Array(FILE_TEXT_PREVIEW_MAX_BYTES + 1).fill(97);
    bytes[FILE_TEXT_PREVIEW_MAX_BYTES - 1] = 0xe2;
    bytes[FILE_TEXT_PREVIEW_MAX_BYTES] = 0x82;
    const preview = decodeFilePreviewText(bytes);
    expect(preview.truncated).toBe(true);
    expect(preview.text.endsWith("�")).toBe(false);
    expect(preview.text.length).toBe(FILE_TEXT_PREVIEW_MAX_BYTES - 1);
  });
});
