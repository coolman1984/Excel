import { describe, expect, it } from "vitest";
import source from "../App.tsx?raw";

describe("editor mode retention structure", () => {
 it("keeps the editor subtree mounted and hides it in data mode", () => {

  const wrapper = source.indexOf('<div style={{ display: mode === "editor" ? "contents" : "none" }}>');
  const grid = source.indexOf('<Grid', wrapper); const data = source.indexOf('{mode === "data" && ', grid); const dataApp = source.indexOf('<DataApp />', data);
  expect(wrapper).toBeGreaterThan(0); expect(grid).toBeGreaterThan(wrapper); expect(data).toBeGreaterThan(grid); expect(dataApp).toBeGreaterThan(data);
  // the Data mode is loaded only when first opened
  expect(source).toContain('React.lazy(() => import("./data/DataApp")');
  expect(source).toContain('setMode("editor")'); expect(source).toContain('setMode("data")');
 });
});


