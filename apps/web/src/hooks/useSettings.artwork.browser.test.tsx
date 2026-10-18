import { act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, beforeEach, expect, it, vi } from "vite-plus/test";

vi.mock("~/localApi", () => ({
  ensureLocalApi: () => ({
    persistence: { getClientSettings: async () => null, setClientSettings: async () => {} },
  }),
}));

import {
  ensureClientSettingsHydrated,
  persistClientSettingsPatch,
  useEnvironmentArtworkImageUrl,
} from "./useSettings";

let root: Root;
let container: HTMLDivElement;

function ArtworkReader() {
  const url = useEnvironmentArtworkImageUrl();
  return <output>{url}</output>;
}

beforeEach(async () => {
  vi.stubGlobal("IS_REACT_ACT_ENVIRONMENT", true);
  await ensureClientSettingsHydrated();
  await persistClientSettingsPatch({ environmentArtworkImageUrl: "" });
  container = document.createElement("div");
  document.body.append(container);
  root = createRoot(container);
});

afterEach(async () => {
  await act(() => root.unmount());
  container.remove();
  vi.unstubAllGlobals();
});

it("updates the rendered artwork URL when its setting changes and clears it on reset", async () => {
  await act(() => root.render(<ArtworkReader />));
  expect(container.querySelector("output")?.textContent).toBe("");
  await act(() =>
    persistClientSettingsPatch({ environmentArtworkImageUrl: "https://example.test/first.png" }),
  );
  expect(container.querySelector("output")?.textContent).toBe("https://example.test/first.png");
  await act(() =>
    persistClientSettingsPatch({ environmentArtworkImageUrl: "https://example.test/second.png" }),
  );
  expect(container.querySelector("output")?.textContent).toBe("https://example.test/second.png");
  await act(() => persistClientSettingsPatch({ environmentArtworkImageUrl: "" }));
  expect(container.querySelector("output")?.textContent).toBe("");
});
