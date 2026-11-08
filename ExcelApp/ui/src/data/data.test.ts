import { describe, expect, it } from "vitest";
import { suggestPattern, validateStep } from "./IngestionPage";
import { predicateFromFields } from "./SpecBuilder";
import { toastReducer } from "./DataApp";
import { mockDataCall } from "./mockData";
import { isCurrentPage } from "./ResultGrid";
import type { IngestRecipe } from "./contracts";
describe("data interface helpers", () => {
 it("suggests a pattern", () => expect(suggestPattern("C:\\Reports\\sales-202610.xlsx")).toBe("sales-*.xlsx"));
 it("validates sample, date and keys", () => { const r = { dateColumn: "", keyColumns: [] } as unknown as IngestRecipe; expect(validateStep(0,r,"")).toBeTruthy(); expect(validateStep(2,r,"x")).toBeTruthy(); expect(validateStep(3,r,"x")).toBeTruthy(); });
 it("preserves text codes and parses numeric filters", () => { expect(predicateFromFields("Sales","gt","12","number")).toEqual({op:"gt",col:"Sales",value:12}); expect(predicateFromFields("Code","eq","001","text")).toEqual({op:"eq",col:"Code",value:"001"}); });
 it("limits notices", () => expect(toastReducer(["1","2","3","4","5"],"new")).toEqual(["new","1","2","3","4"]));
 it("rejects stale paging replies", () => { expect(isCurrentPage(1,2,"a","a")).toBe(false); expect(isCurrentPage(2,2,"a","b")).toBe(false); expect(isCurrentPage(2,2,"a","a")).toBe(true); });
 it("pages synthetic results", async () => { const r = await mockDataCall("data.run", { spec: { source: { kind: "report", id: "sales" } } }); const p = await mockDataCall("data.page", { resultId:r.resultId, from:200, count:200 }); expect(p.resultId).toBe(r.resultId); expect(p.from).toBe(200); expect(p.rows).toHaveLength(200); });
});
