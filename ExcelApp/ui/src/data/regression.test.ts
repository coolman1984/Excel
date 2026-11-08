import { describe, expect, it } from "vitest";
import * as React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { PredEditor, changePredOp } from "./SpecBuilder";
import { validateStep, suggestPattern } from "./IngestionPage";
import { requestCurrentPage } from "./ResultGrid";
import { toastReducer } from "./DataApp";
import { mockDataCall } from "./mockData";
import type { IngestRecipe, ResultPage } from "./contracts";
const recipe: IngestRecipe = { id:"x", name:"Test", folders:["C:\\Demo"], patterns:["*.xlsx"], headerSignature:["Date","Code"], headerRow:1, sheet:"Sheet1", dateColumn:"Date", granularity:"month", week:{firstDay:"monday",numbering:"iso"},keyColumns:["Code"],typeFixes:[],partialThreshold:.3,truth:"latest",lock:{workingDay:5,holidays:[]},keepVersions:0,dependentViews:[] };
describe("lane C regression", () => {
 it("round-trips and renders every predicate operator", () => { for (const op of ["and","or","not","eq","ne","gt","ge","lt","le","between","in","contains","begins","ends","blank","notBlank","duplicate","colCompare"]) { const pred=changePredOp(op); const round=JSON.parse(JSON.stringify(pred)); expect(round).toEqual(pred); const html=renderToStaticMarkup(React.createElement(PredEditor,{value:round,columns:[{name:"Code",type:"text"}],onChange:()=>{}})); expect(html).toContain(`<option selected="">${op}</option>`); } });
 it("validates all wizard steps",()=>{ for(let n=0;n<7;n++) expect(validateStep(n,recipe,"C:\\Demo\\sample.xlsx")).toBe(""); expect(validateStep(0,recipe,"")).toBeTruthy(); expect(validateStep(1,{...recipe,sheet:""},"x")).toBeTruthy(); expect(validateStep(2,{...recipe,dateColumn:""},"x")).toBeTruthy(); expect(validateStep(3,{...recipe,keyColumns:[]},"x")).toBeTruthy(); expect(validateStep(4,{...recipe,typeFixes:[{column:"Code",fix:"padZeros",width:0}]},"x")).toBeTruthy(); expect(validateStep(5,{...recipe,partialThreshold:1},"x")).toBeTruthy(); expect(validateStep(6,{...recipe,folders:[]},"x")).toBeTruthy(); });
 it("drops an old response after a newer result wins",async()=>{ let generation=1; let old!: (p:ResultPage)=>void; let fresh!: (p:ResultPage)=>void; const accepted:string[]=[]; const a=requestCurrentPage(new Promise(r=>old=r),1,()=>generation,"old",p=>accepted.push(p.resultId)); generation=2; const b=requestCurrentPage(new Promise(r=>fresh=r),2,()=>generation,"new",p=>accepted.push(p.resultId)); fresh({resultId:"new",from:0,rows:[],kinds:[]}); await b; old({resultId:"old",from:0,rows:[],kinds:[]}); await a; expect(accepted).toEqual(["new"]); });
 it("reduces toasts and suggests patterns",()=>{expect(toastReducer(["a","b","c","d","e"],"z")).toEqual(["z","a","b","c","d"]);expect(suggestPattern("C:\\x\\sales-202610.xlsx")).toBe("sales-*.xlsx");});
 it("recognizes all twenty sample questions",async()=>{const examples=["top 10 Product by Sales","bottom 5 Product by Sales","total Sales by Region","count by Region","average Sales by Region","max Sales","min Sales","rows where Sales > 100","Region is North","show Sales","duplicates in Product","blank Product","sort by Sales desc","Sales per Region in October","trend of Sales by month","compare Sales and Budget by Product","what changed in October","how many rows in Sales","last 7 days of Sales","Sales this month vs last month"]; for(const text of examples) expect((await mockDataCall("data.ask",{text})).ok,text).toBe(true);});
});

