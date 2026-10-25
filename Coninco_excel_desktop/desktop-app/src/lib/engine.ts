import { invoke } from "@tauri-apps/api/core";
import type { ArchiveEntry, AuditResult, ComparisonRow, FormulaAudit } from "../types";

const isDesktop = () => Boolean((window as Window & { __TAURI_INTERNALS__?: unknown }).__TAURI_INTERNALS__);

const demoRows: ComparisonRow[] = [
  {
    sourceRow: 21, itemNo: "2465", description: "Fresh air duct 800x400, G.I thk 0.75mm", unit: "m2",
    section: "standard", productType: "rectangular_duct", w: 800, h: 400, w1: 0, h1: 0, w2: 0, h2: 0,
    lm: 0, rd: 0, a: 0, t: 0, q: 9, plan: "0.120 + (2*0.400)", upDown: "", verticalAxis: "",
    formulaDetail: "S=((W+H)*2*K + joint allowance)", acousticFireproofDetail: "", note: "",
    contractorK: 5.325, calculatedK: 5.325, kMatch: true,
    contractorScn: 15.12, calculatedScn: 15.12, scnMatch: true,
    contractorScc: 0, calculatedScc: 0, sccMatch: true,
    contractorSta: 0, calculatedSta: 0, staMatch: true,
    formulaAudits: [],
  },
  {
    sourceRow: 24, itemNo: "2468", description: "Square elbow 360x360, G.I sheet thk 0.58mm", unit: "m2",
    section: "standard", productType: "square_elbow", w: 360, h: 360, w1: 350, h1: 350, w2: 0, h2: 0,
    lm: 175, rd: 0, a: 25, t: 0, q: 1, plan: "1.000", upDown: "", verticalAxis: "",
    formulaDetail: "S=(K*(2*((W+50)+(H+50))*L)+...)", acousticFireproofDetail: "", note: "",
    contractorK: 1, calculatedK: 1, kMatch: true,
    contractorScn: 1.15, calculatedScn: 1.15, scnMatch: true,
    contractorScc: 0, calculatedScc: 0, sccMatch: true,
    contractorSta: 0, calculatedSta: 0, staMatch: true,
    formulaAudits: [],
  },
];

const demoArchive: ArchiveEntry[] = [{
  id: "demo-1", submittedName: "L01GF0087.xls", processedName: "L01GF0087_calculated.xlsx",
  sheet: "02", status: "COMPLETED", rows: 15, mismatches: 1, createdAt: "Today, 10:24",
}];

export const chooseWorkbook = () =>
  isDesktop() ? invoke<string | null>("choose_workbook") : Promise.resolve("/demo/L01GF0087.xls");

export const listSheets = (path: string) =>
  isDesktop() ? invoke<string[]>("list_sheets", { path }) : Promise.resolve(["02", "COMBINE"]);

export const runAudit = (path: string, sheet: string) =>
  isDesktop()
    ? invoke<AuditResult>("run_audit", { path, sheet })
    : Promise.resolve<AuditResult>({
        id: "demo-1", submittedName: "L01GF0087.xls", processedName: "L01GF0087_calculated.xlsx",
        sheet, rows: 15, mismatches: 1,
      });

export const listArchive = () =>
  isDesktop() ? invoke<ArchiveEntry[]>("list_archive") : Promise.resolve(demoArchive);

const number = (value: string | undefined) => Number.parseFloat(value ?? "") || 0;
const bool = (value: string | undefined) => value === "True" || value === "true";
const formulaAudits = (value: string | undefined): FormulaAudit[] => {
  if (!value) return [];
  try {
    const parsed = JSON.parse(value) as Array<Record<string, unknown>>;
    return parsed.map((audit) => ({
      target: String(audit.target ?? ""),
      label: String(audit.label ?? ""),
      originalFormula: String(audit.original_formula ?? ""),
      detailFormula: String(audit.detail_formula ?? ""),
      detailBlock: String(audit.detail_block ?? audit.detail_formula ?? ""),
      originalValue: Number(audit.original_value ?? 0),
      calculatedValue: Number(audit.calculated_value ?? 0),
      difference: Number(audit.difference ?? 0),
      matches: Boolean(audit.matches),
      error: String(audit.error ?? ""),
      correction: String(audit.correction ?? ""),
    }));
  } catch {
    return [];
  }
};
const rowFromPreview = (row: Record<string, string>): ComparisonRow => ({
  sourceRow: Number.parseInt(row.source_row ?? "0", 10), itemNo: row.item_no ?? "", description: row.description ?? "",
  unit: row.unit ?? "", section: row.section ?? "", productType: row.product_type ?? "", w: number(row.w), h: number(row.h),
  w1: number(row.w1), h1: number(row.h1), w2: number(row.w2), h2: number(row.h2), lm: number(row.l_m), rd: number(row.r_d),
  a: number(row.a), t: number(row.t), q: number(row.q), plan: row.plan ?? "", upDown: row.up_down ?? "", verticalAxis: row.vertical_axis ?? "",
  formulaDetail: row.formula_detail ?? "", acousticFireproofDetail: row.acoustic_fireproof_detail ?? "", note: row.note ?? "",
  contractorK: number(row.contractor_k), calculatedK: number(row.calculated_k), kMatch: bool(row.k_match),
  contractorScn: number(row.contractor_scn ?? row.contractor_s), calculatedScn: number(row.calculated_scn ?? row.calculated_s), scnMatch: bool(row.scn_match ?? row.s_match),
  contractorScc: number(row.contractor_scc ?? row.contractor_fireproof_s), calculatedScc: number(row.calculated_scc ?? row.calculated_fireproof_s), sccMatch: bool(row.scc_match ?? row.fireproof_match),
  contractorSta: number(row.contractor_sta ?? row.contractor_inner_s), calculatedSta: number(row.calculated_sta ?? row.calculated_inner_s), staMatch: bool(row.sta_match ?? row.inner_match),
  formulaAudits: formulaAudits(row.formula_audits),
});

export const loadComparison = async (id: string) => {
  if (!isDesktop()) return demoRows;
  const raw = await invoke<Record<string, string>[]>("load_comparison", { id });
  return raw.map(rowFromPreview);
};

export const deleteArchive = (id: string) =>
  isDesktop() ? invoke<void>("delete_archive", { id }) : Promise.resolve();

export const downloadArchive = (id: string, kind: "submitted" | "processed") =>
  isDesktop() ? invoke<string>("download_archive", { id, kind }) : Promise.resolve("/demo/downloaded-file.xlsx");

export const openArchiveFile = (id: string, kind: "submitted" | "processed") =>
  isDesktop() ? invoke<void>("open_archive_file", { id, kind }) : Promise.resolve();
