// @vitest-environment jsdom

import * as NodeBuffer from "node:buffer";
import * as NodeURL from "node:url";
import { EnvironmentId, ThreadId } from "@t3tools/contracts";
import { act, type ReactNode } from "react";
import { createRoot, type Root } from "react-dom/client";
import * as XLSX from "xlsx";
import { afterEach, beforeEach, describe, expect, it, vi } from "vite-plus/test";

import { AttachmentFilePreview } from "./AttachmentFilePreview";
import { WorkspaceOfficePreview } from "./WorkspaceOfficePreview";
import {
  parseOfficeDocument,
  type OfficeWorkerRequest,
  type OfficeWorkerResponse,
} from "./officeDocumentParser";

const mocks = vi.hoisted(() => ({
  url: "https://environment.test/api/assets/signed/book.xlsx",
  refresh: vi.fn<() => Promise<string | null>>(),
}));
vi.mock("~/assets/assetUrls", () => ({
  useAssetUrlRefresh: () => mocks.refresh,
  useAssetUrlState: () => ({ _tag: "Success", url: mocks.url }),
}));
vi.mock("~/hooks/useCopyToClipboard", () => ({
  useCopyToClipboard: () => ({ copyToClipboard: vi.fn(), isCopied: false }),
}));
vi.mock("~/hooks/useSettings", () => ({
  useClientSettings: () => true,
  useUpdateClientSettings: () => vi.fn(),
}));
vi.mock("~/components/ui/toast", () => ({ toastManager: { add: vi.fn() } }));
vi.mock("~/components/ChatMarkdown", () => ({ default: () => null }));
vi.mock("~/components/ui/scroll-area", () => ({
  ScrollArea: ({ children }: { children: ReactNode }) => children,
}));
vi.mock("~/components/ui/button", () => ({
  Button: ({
    variant: _variant,
    size: _size,
    ...props
  }: React.ComponentProps<"button"> & { variant?: string; size?: string }) => <button {...props} />,
}));
vi.mock("./fileSurfaceChrome", () => ({
  FILE_SURFACE_SUBHEADER_CLASS: "",
  FileSurfaceAction: ({ label, onPress }: { label: string; onPress: () => void }) => (
    <button aria-label={label} onClick={onPress} />
  ),
  FileSurfaceFailure: ({ message, onRetry }: { message: string; onRetry?: () => void }) => (
    <div role="alert">
      {message}
      {onRetry ? <button onClick={onRetry}>Try again</button> : null}
    </div>
  ),
  FileSurfaceLoading: () => <div role="status">Loading</div>,
  FileSurfaceNotice: ({ children }: { children: ReactNode }) => <div>{children}</div>,
}));

/** Only the browser worker transport is replaced: its request is parsed by the real decoder. */
class ParserWorker extends EventTarget {
  static instances: ParserWorker[] = [];
  static posted: ((worker: ParserWorker) => void) | undefined;
  request: OfficeWorkerRequest | undefined;
  terminate = vi.fn();
  constructor() {
    super();
    ParserWorker.instances.push(this);
    queueMicrotask(() =>
      this.dispatchEvent(
        new MessageEvent<OfficeWorkerResponse>("message", { data: { type: "ready" } }),
      ),
    );
  }
  postMessage(request: OfficeWorkerRequest) {
    this.request = request;
    ParserWorker.posted?.(this);
  }
  finish() {
    const request = this.request;
    if (!request) throw new Error("Worker has no document request");
    const document = parseOfficeDocument(new Uint8Array(request.buffer), request.format);
    this.dispatchEvent(
      new MessageEvent<OfficeWorkerResponse>("message", {
        data: { id: request.id, ok: true, document },
      }),
    );
  }
  fail() {
    this.dispatchEvent(new Event("error"));
  }
}

function workbookBytes(bookType: "xlsx" | "biff8" = "xlsx", value = "ねこ / real workbook") {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet([
      ["Product", "Total"],
      [value, 1800],
    ]),
    "売上",
  );
  return new Uint8Array(XLSX.write(workbook, { type: "array", bookType, compression: true }));
}

const environmentId = EnvironmentId.make("test-environment");
const threadRef = { environmentId, threadId: ThreadId.make("test-thread") };
const nextWorker = () =>
  new Promise<ParserWorker>((resolve) => {
    ParserWorker.posted = resolve;
  });
let root: Root;
let container: HTMLDivElement;

beforeEach(() => {
  vi.stubGlobal("IS_REACT_ACT_ENVIRONMENT", true);
  vi.stubGlobal("Worker", ParserWorker);
  // JSDOM does not implement Blob streaming or object URLs; retain Node's native byte APIs.
  vi.stubGlobal("Blob", NodeBuffer.Blob);
  vi.stubGlobal("URL", NodeURL.URL);
  container = document.createElement("div");
  document.body.append(container);
  root = createRoot(container);
  mocks.url = "https://environment.test/api/assets/signed/book.xlsx";
  mocks.refresh.mockReset().mockImplementation(async () => mocks.url);
  ParserWorker.posted = undefined;
  ParserWorker.instances = [];
  vi.spyOn(URL, "createObjectURL").mockReturnValue("blob:office-local");
  vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => undefined);
  vi.stubGlobal(
    "fetch",
    vi.fn(async () => new Response(workbookBytes())),
  );
});

afterEach(async () => {
  await act(() => root.unmount());
  container.remove();
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

const workspace = (
  name = "/workspace/book.xlsx",
  reloadRevision = 0,
  workspaceMutationId: string | null = null,
) => (
  <WorkspaceOfficePreview
    environmentId={environmentId}
    threadRef={threadRef}
    draft={false}
    absolutePath={name}
    workspaceRoot="/workspace"
    format="xlsx"
    workspaceMutationId={workspaceMutationId}
    reloadRevision={reloadRevision}
  />
);

function pendingAuthorization() {
  let complete: ((url: string | null) => void) | undefined;
  const promise = new Promise<string | null>((resolve) => {
    complete = resolve;
  });
  mocks.refresh.mockReturnValueOnce(promise);
  return { promise, complete: (url: string | null) => complete?.(url) };
}

const remoteAttachment = (attachmentId = "captured-xlsx") => (
  <AttachmentFilePreview
    name={`${attachmentId}.xlsx`}
    mimeType="text/plain"
    sizeBytes={100}
    asset={{ environmentId, attachmentId }}
  />
);

describe("Office entry points with the real document client and parser", () => {
  it("waits for reauthorization before starting one download per open for a cached completed mutation", async () => {
    const firstAuthorization = pendingAuthorization();
    const firstPosted = nextWorker();
    await act(() => root.render(workspace("/workspace/book.xlsx", 0, "mutation-a")));
    expect(mocks.refresh).toHaveBeenCalledOnce();
    expect(fetch).not.toHaveBeenCalled();
    expect(ParserWorker.instances).toHaveLength(0);
    await act(async () => {
      firstAuthorization.complete("https://environment.test/authorized-first.xlsx");
      await firstAuthorization.promise;
    });
    const first = await firstPosted;
    await act(() => first.finish());
    await act(() => root.render(null));
    const secondAuthorization = pendingAuthorization();
    const secondPosted = nextWorker();
    await act(() => root.render(workspace("/workspace/book.xlsx", 0, "mutation-a")));
    expect(mocks.refresh).toHaveBeenCalledTimes(2);
    expect(fetch).toHaveBeenCalledOnce();
    expect(ParserWorker.instances).toHaveLength(1);
    await act(async () => {
      secondAuthorization.complete("https://environment.test/authorized-second.xlsx");
      await secondAuthorization.promise;
    });
    const second = await secondPosted;
    await act(() => second.finish());
    expect(fetch).toHaveBeenCalledTimes(2);
    expect(vi.mocked(fetch).mock.calls.map(([url]) => String(url))).toEqual([
      "https://environment.test/authorized-first.xlsx?workspace-revision=mutation-a%3A0",
      "https://environment.test/authorized-second.xlsx?workspace-revision=mutation-a%3A0",
    ]);
    expect(first.terminate).toHaveBeenCalledOnce();
    expect(second.terminate).toHaveBeenCalledOnce();
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook");
  });

  it("gates new workspace mutations and manual reloads until each authorization completes", async () => {
    const initialPosted = nextWorker();
    await act(() => root.render(workspace()));
    await act(async () => {
      (await initialPosted).finish();
    });
    const mutationAuthorization = pendingAuthorization();
    const changedPosted = nextWorker();
    await act(() => root.render(workspace("/workspace/book.xlsx", 0, "mutation-new")));
    expect(fetch).toHaveBeenCalledOnce();
    expect(ParserWorker.instances).toHaveLength(1);
    expect(container.querySelector("table")).toBeNull();
    vi.mocked(fetch).mockImplementation(
      async () => new Response(workbookBytes("xlsx", "mutation value")),
    );
    await act(async () => {
      mutationAuthorization.complete(mocks.url);
      await mutationAuthorization.promise;
    });
    await act(async () => {
      (await changedPosted).finish();
    });
    expect(container.querySelector("tbody")?.textContent).toContain("mutation value");
    const manualAuthorization = pendingAuthorization();
    const manualPosted = nextWorker();
    await act(() => root.render(workspace("/workspace/book.xlsx", 1, "mutation-new")));
    expect(fetch).toHaveBeenCalledTimes(2);
    expect(ParserWorker.instances).toHaveLength(2);
    vi.mocked(fetch).mockImplementation(
      async () => new Response(workbookBytes("xlsx", "manual value")),
    );
    await act(async () => {
      manualAuthorization.complete(mocks.url);
      await manualAuthorization.promise;
    });
    await act(async () => {
      (await manualPosted).finish();
    });
    expect(container.querySelector("tbody")?.textContent).toContain("manual value");
    expect(mocks.refresh).toHaveBeenCalledTimes(2);
    expect(fetch).toHaveBeenCalledTimes(3);
    expect(ParserWorker.instances).toHaveLength(3);
  });

  it("reauthorizes a failed preview before retrying its reader", async () => {
    const initialPosted = nextWorker();
    await act(() => root.render(workspace()));
    const first = await initialPosted;
    await act(() => first.fail());
    expect(container.querySelector('[role="alert"]')).not.toBeNull();
    const retryAuthorization = pendingAuthorization();
    const retryPosted = nextWorker();
    await act(() => container.querySelector<HTMLButtonElement>('[role="alert"] button')!.click());
    expect(mocks.refresh).toHaveBeenCalledOnce();
    expect(fetch).toHaveBeenCalledOnce();
    expect(ParserWorker.instances).toHaveLength(1);
    await act(async () => {
      retryAuthorization.complete("https://environment.test/renewed.xlsx");
      await retryAuthorization.promise;
    });
    await act(async () => {
      (await retryPosted).finish();
    });
    expect(vi.mocked(fetch).mock.calls.map(([url]) => String(url))).toEqual([
      mocks.url,
      "https://environment.test/renewed.xlsx",
    ]);
    expect(ParserWorker.instances).toHaveLength(2);
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook");
  });

  it("ignores an older resource authorization that resolves after the current preview", async () => {
    const oldAuthorization = pendingAuthorization();
    await act(() => root.render(workspace("/workspace/old.xlsx", 0, "mutation-a")));
    const currentAuthorization = pendingAuthorization();
    const posted = nextWorker();
    await act(() => root.render(workspace("/workspace/current.xlsx", 0, "mutation-a")));
    await act(async () => {
      currentAuthorization.complete("https://environment.test/current.xlsx");
      await currentAuthorization.promise;
    });
    await act(async () => {
      (await posted).finish();
    });
    await act(async () => {
      oldAuthorization.complete("https://environment.test/old.xlsx");
      await oldAuthorization.promise;
    });
    expect(fetch).toHaveBeenCalledExactlyOnceWith(
      "https://environment.test/current.xlsx?workspace-revision=mutation-a%3A0",
      { signal: expect.any(AbortSignal) },
    );
    expect(ParserWorker.instances).toHaveLength(1);
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook");
  });

  it("waits for a new authorization when returning to a previously authorized resource", async () => {
    const firstPosted = nextWorker();
    await act(() => root.render(workspace("/workspace/a.xlsx", 0, "mutation-a")));
    const first = await firstPosted;
    await act(() => first.finish());
    const otherAuthorization = pendingAuthorization();
    await act(() => root.render(workspace("/workspace/b.xlsx", 0, "mutation-a")));
    const returnAuthorization = pendingAuthorization();
    const returnedPosted = nextWorker();
    await act(() => root.render(workspace("/workspace/a.xlsx", 0, "mutation-a")));
    expect(fetch).toHaveBeenCalledOnce();
    expect(ParserWorker.instances).toHaveLength(1);
    expect(container.querySelector("table")).toBeNull();
    await act(async () => {
      otherAuthorization.complete("https://environment.test/stale-b.xlsx");
      await otherAuthorization.promise;
    });
    expect(fetch).toHaveBeenCalledOnce();
    await act(async () => {
      returnAuthorization.complete("https://environment.test/returned-a.xlsx");
      await returnAuthorization.promise;
    });
    const returned = await returnedPosted;
    await act(() => returned.finish());
    expect(fetch).toHaveBeenCalledTimes(2);
    expect(ParserWorker.instances).toHaveLength(2);
    expect(vi.mocked(fetch).mock.calls[1]?.[0]).toBe(
      "https://environment.test/returned-a.xlsx?workspace-revision=mutation-a%3A0",
    );
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook");
  });

  it("recovers from authorization failure and ignores an authorization completed after closing", async () => {
    const failedAuthorization = pendingAuthorization();
    await act(() => root.render(workspace("/workspace/book.xlsx", 0, "mutation-a")));
    await act(async () => {
      failedAuthorization.complete(null);
      await failedAuthorization.promise;
    });
    expect(container.querySelector('[role="alert"]')).not.toBeNull();
    expect(fetch).not.toHaveBeenCalled();
    const retryAuthorization = pendingAuthorization();
    const retryPosted = nextWorker();
    await act(() => container.querySelector<HTMLButtonElement>('[role="alert"] button')!.click());
    expect(fetch).not.toHaveBeenCalled();
    await act(async () => {
      retryAuthorization.complete("https://environment.test/recovered.xlsx");
      await retryAuthorization.promise;
    });
    const retry = await retryPosted;
    await act(() => retry.finish());
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook");
    const closingAuthorization = pendingAuthorization();
    await act(() => root.render(workspace("/workspace/book.xlsx", 1, "mutation-a")));
    await act(() => root.render(null));
    await act(async () => {
      closingAuthorization.complete("https://environment.test/closed.xlsx");
      await closingAuthorization.promise;
    });
    expect(fetch).toHaveBeenCalledOnce();
    expect(ParserWorker.instances).toHaveLength(1);
    expect(container.textContent).toBe("");
  });

  it("reloads changed bytes on a manual revision even when the signed URL is unchanged", async () => {
    const firstPosted = nextWorker();
    await act(() => root.render(workspace()));
    const first = await firstPosted;
    await act(() => first.finish());
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook");
    vi.mocked(fetch).mockImplementation(
      async () => new Response(workbookBytes("xlsx", "updated value")),
    );
    const secondPosted = nextWorker();
    await act(() => root.render(workspace("/workspace/book.xlsx", 1)));
    const second = await secondPosted;
    await act(() => second.finish());
    expect(container.querySelector("tbody")?.textContent).toContain("updated value");
    expect(container.querySelector("tbody")?.textContent).not.toContain("ねこ / real workbook");
    expect(vi.mocked(fetch).mock.calls.map(([url]) => String(url))).toEqual([
      mocks.url,
      `${mocks.url}?workspace-revision=%3A1`,
    ]);
  });

  it("loads a signed workspace asset and displays decoded cells", async () => {
    const posted = nextWorker();
    await act(() => root.render(workspace()));
    const worker = await posted;
    await act(() => worker.finish());
    expect(fetch).toHaveBeenCalledExactlyOnceWith(mocks.url, { signal: expect.any(AbortSignal) });
    expect(container.querySelector("table")?.getAttribute("aria-label")).toBe("売上 cell values");
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook1800");
    expect(container.textContent).toContain(
      "Values only; styles, charts, and formula recalculation",
    );
    expect(worker.terminate).toHaveBeenCalledOnce();
  });

  it("routes a text/plain xlsx attachment through signed binary loading rather than UTF-8 source", async () => {
    const posted = nextWorker();
    await act(() =>
      root.render(
        <AttachmentFilePreview
          name="captured.xlsx"
          mimeType="text/plain"
          sizeBytes={100}
          asset={{ environmentId, attachmentId: "captured-xlsx" }}
        />,
      ),
    );
    const worker = await posted;
    await act(() => worker.finish());
    expect(mocks.refresh).toHaveBeenCalledOnce();
    expect(fetch).toHaveBeenCalledExactlyOnceWith(mocks.url, { signal: expect.any(AbortSignal) });
    expect(container.querySelector("td")?.textContent).toBe("Product");
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook1800");
    expect(container.querySelector("pre")).toBeNull();
    expect(container.querySelector('[role="alert"]')).toBeNull();
  });

  it("starts one attachment reader after retry authorization even when the signed URL is unchanged", async () => {
    const initialAuthorization = pendingAuthorization();
    const initialPosted = nextWorker();
    await act(() => root.render(remoteAttachment()));
    expect(fetch).not.toHaveBeenCalled();
    expect(ParserWorker.instances).toHaveLength(0);
    await act(async () => {
      initialAuthorization.complete(mocks.url);
      await initialAuthorization.promise;
    });
    const first = await initialPosted;
    await act(() => first.fail());
    const retryAuthorization = pendingAuthorization();
    const retryPosted = nextWorker();
    vi.mocked(fetch).mockImplementation(
      async () => new Response(workbookBytes("xlsx", "renewed attachment")),
    );
    await act(() => container.querySelector<HTMLButtonElement>('[role="alert"] button')!.click());
    expect(mocks.refresh).toHaveBeenCalledTimes(2);
    expect(fetch).toHaveBeenCalledOnce();
    expect(ParserWorker.instances).toHaveLength(1);
    expect(container.querySelector("table")).toBeNull();
    await act(async () => {
      retryAuthorization.complete(mocks.url);
      await retryAuthorization.promise;
    });
    const retry = await retryPosted;
    await act(() => retry.finish());
    expect(fetch).toHaveBeenCalledTimes(2);
    expect(vi.mocked(fetch).mock.calls.map(([url]) => String(url))).toEqual([mocks.url, mocks.url]);
    expect(ParserWorker.instances).toHaveLength(2);
    expect(first.terminate).toHaveBeenCalledOnce();
    expect(retry.terminate).toHaveBeenCalledOnce();
    expect(container.querySelector("tbody")?.textContent).toContain("renewed attachment");
  });

  it("does not reuse attachment authorization across A-to-B-to-A or accept results after closing", async () => {
    const firstPosted = nextWorker();
    await act(() => root.render(remoteAttachment("a")));
    const first = await firstPosted;
    await act(() => first.finish());
    const otherAuthorization = pendingAuthorization();
    await act(() => root.render(remoteAttachment("b")));
    const returnedAuthorization = pendingAuthorization();
    const returnedPosted = nextWorker();
    await act(() => root.render(remoteAttachment("a")));
    expect(fetch).toHaveBeenCalledOnce();
    expect(ParserWorker.instances).toHaveLength(1);
    expect(container.querySelector("table")).toBeNull();
    await act(async () => {
      otherAuthorization.complete("https://environment.test/obsolete-b.xlsx");
      await otherAuthorization.promise;
    });
    expect(fetch).toHaveBeenCalledOnce();
    await act(async () => {
      returnedAuthorization.complete("https://environment.test/returned-a.xlsx");
      await returnedAuthorization.promise;
    });
    const returned = await returnedPosted;
    await act(() => returned.finish());
    expect(fetch).toHaveBeenCalledTimes(2);
    expect(vi.mocked(fetch).mock.calls[1]?.[0]).toBe("https://environment.test/returned-a.xlsx");
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook");
    const closingAuthorization = pendingAuthorization();
    await act(() => root.render(remoteAttachment("closing")));
    await act(() => root.render(null));
    await act(async () => {
      closingAuthorization.complete("https://environment.test/closed.xlsx");
      await closingAuthorization.promise;
    });
    expect(fetch).toHaveBeenCalledTimes(2);
    expect(ParserWorker.instances).toHaveLength(2);
    expect(container.textContent).toBe("");
  });

  it("retries a local Office Blob without remote authorization or HTTP loading", async () => {
    const file = new Blob([workbookBytes()]);
    const firstPosted = nextWorker();
    await act(() =>
      root.render(
        <AttachmentFilePreview
          name="local.xlsx"
          mimeType="application/octet-stream"
          sizeBytes={file.size}
          file={file}
        />,
      ),
    );
    const first = await firstPosted;
    await act(() => first.fail());
    const retryPosted = nextWorker();
    await act(() => container.querySelector<HTMLButtonElement>('[role="alert"] button')!.click());
    const retry = await retryPosted;
    await act(() => retry.finish());
    expect(mocks.refresh).not.toHaveBeenCalled();
    expect(fetch).not.toHaveBeenCalled();
    expect(ParserWorker.instances).toHaveLength(2);
    expect(first.terminate).toHaveBeenCalledOnce();
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook");
  });

  it("reads real BIFF8 from a local Blob and releases its object URL when changing to legacy ppt fallback", async () => {
    const file = new Blob([workbookBytes("biff8")], { type: "application/vnd.ms-excel" });
    const posted = nextWorker();
    await act(() =>
      root.render(
        <AttachmentFilePreview
          name="local.xls"
          mimeType="application/vnd.ms-excel"
          sizeBytes={file.size}
          file={file}
        />,
      ),
    );
    const worker = await posted;
    await act(() => worker.finish());
    expect(container.querySelector("tbody")?.textContent).toContain("ねこ / real workbook1800");
    expect(fetch).not.toHaveBeenCalled();
    expect(mocks.refresh).not.toHaveBeenCalled();
    const legacy = new Blob([new Uint8Array([0xd0, 0xcf, 0x11, 0xe0])]);
    await act(() =>
      root.render(
        <AttachmentFilePreview
          name="legacy.ppt"
          mimeType="application/vnd.ms-powerpoint"
          sizeBytes={legacy.size}
          file={legacy}
        />,
      ),
    );
    expect(container.textContent).toContain("No preview for this file");
    expect(container.textContent).toContain("Save it to open in an app that supports ppt files.");
    expect(container.querySelector("table")).toBeNull();
    expect(URL.revokeObjectURL).toHaveBeenCalledWith("blob:office-local");
    expect(worker.terminate).toHaveBeenCalledOnce();
  });

  it("terminates pending work on route replacement and unmount, ignoring a stale result", async () => {
    const firstPosted = nextWorker();
    await act(() => root.render(workspace("/workspace/first.xlsx")));
    const first = await firstPosted;
    mocks.url = "https://environment.test/api/assets/second/book.xlsx";
    const secondPosted = nextWorker();
    await act(() => root.render(workspace("/workspace/second.xlsx")));
    const second = await secondPosted;
    expect(first.terminate).toHaveBeenCalledOnce();
    await act(() => first.finish());
    expect(container.querySelector("table")).toBeNull();
    expect(container.textContent).toContain("Loading");
    await act(() => root.render(null));
    expect(second.terminate).toHaveBeenCalledOnce();
    await act(() => second.finish());
    expect(container.textContent).toBe("");
  });
});
