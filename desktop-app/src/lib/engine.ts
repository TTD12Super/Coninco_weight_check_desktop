import { invoke } from "@tauri-apps/api/core";
import type { ArchiveEntry, AuditResult, ComparisonRow, FormulaAudit, FormulaCatalogEntry } from "../types";

const isDesktop = () => Boolean((window as Window & { __TAURI_INTERNALS__?: unknown }).__TAURI_INTERNALS__);

const demoRows: ComparisonRow[] = [
  {
    sourceRow: 21, itemNo: "2465", description: "Fresh air duct 800x400, G.I thk 0.75mm", unit: "m2",
    section: "standard", productType: "rectangular_duct", w: 800, h: 400, w1: 0, h1: 0, w2: 0, h2: 0,
    inputReferences: { W: "F21", H: "G21", W1: "H21", H1: "I21", W2: "J21", H2: "K21", L_M: "L21", R_D: "M21", a: "N21", t: "O21", Q: "P21", K: "Q21" },
    lm: 0, rd: 0, a: 0, t: 0, q: 9, plan: "0.120 + (2*0.400)", upDown: "", verticalAxis: "",
    formulaDetail: "S=((W+H)*2*K + joint allowance)", acousticFireproofDetail: "", note: "",
    contractorK: 5.325, calculatedK: 5.325, kMatch: true,
    contractorScn: 15.12, calculatedScn: 15.12, scnMatch: true,
    contractorScc: 0, calculatedScc: 0, sccMatch: true,
    contractorSta: 0, calculatedSta: 0, staMatch: true,
    quantityLayout: "two",
    sharedQuantityTarget: "",
    formulaAudits: [],
  },
  {
    sourceRow: 24, itemNo: "2468", description: "Square elbow 360x360, G.I sheet thk 0.58mm", unit: "m2",
    section: "standard", productType: "square_elbow", w: 360, h: 360, w1: 350, h1: 350, w2: 0, h2: 0,
    inputReferences: { W: "F24", H: "G24", W1: "H24", H1: "I24", W2: "J24", H2: "K24", L_M: "L24", R_D: "M24", a: "N24", t: "O24", Q: "P24", K: "Q24" },
    lm: 175, rd: 0, a: 25, t: 0, q: 1, plan: "1.000", upDown: "", verticalAxis: "",
    formulaDetail: "S=(K*(2*((W+50)+(H+50))*L)+...)", acousticFireproofDetail: "", note: "",
    contractorK: 1, calculatedK: 1, kMatch: true,
    contractorScn: 1.15, calculatedScn: 1.15, scnMatch: true,
    contractorScc: 0, calculatedScc: 0, sccMatch: true,
    contractorSta: 0, calculatedSta: 0, staMatch: true,
    quantityLayout: "two",
    sharedQuantityTarget: "",
    formulaAudits: [],
  },
];

const demoArchive: ArchiveEntry[] = [{
  id: "demo-1", submittedName: "L01GF0087.xls", processedName: "L01GF0087_calculated.xlsx",
  sheet: "02", status: "COMPLETED", rows: 15, mismatches: 1, createdAt: "Today, 10:24",
}];

export const chooseWorkbook = () =>
  isDesktop() ? invoke<string | null>("choose_workbook") : Promise.resolve("/demo/L01GF0087.xls");

export const chooseReferenceWorkbook = () =>
  isDesktop() ? invoke<string | null>("choose_reference_workbook") : Promise.resolve("/demo/formula-reference.xlsx");

export const loadReferenceWorkbook = () =>
  isDesktop() ? invoke<string | null>("get_reference_workbook") : Promise.resolve("/demo/formula-reference.xlsx");

export const saveReferenceWorkbook = (path: string) =>
  isDesktop() ? invoke<string>("save_reference_workbook", { path }) : Promise.resolve(path);

export const clearReferenceWorkbook = () =>
  isDesktop() ? invoke<void>("clear_reference_workbook") : Promise.resolve();

export const listSheets = (path: string) =>
  isDesktop() ? invoke<string[]>("list_sheets", { path }) : Promise.resolve(["02", "COMBINE"]);

export const listReferenceSheets = (path: string) =>
  isDesktop() ? invoke<string[]>("list_reference_sheets", { path }) : Promise.resolve(["COMBINE"]);

export const loadFormulaCatalog = (path: string, sheet: string) =>
  isDesktop()
    ? invoke<FormulaCatalogEntry[]>("load_reference_catalog", { path, sheet }).then((entries) => entries.map((entry) => ({
        ...entry,
        productType: String((entry as FormulaCatalogEntry & { product_type?: string }).product_type ?? entry.productType ?? ""),
        formulaBlock: String((entry as FormulaCatalogEntry & { formula_block?: string }).formula_block ?? entry.formulaBlock ?? ""),
        sourceSheet: String((entry as FormulaCatalogEntry & { source_sheet?: string }).source_sheet ?? entry.sourceSheet ?? ""),
        sourceRow: Number((entry as FormulaCatalogEntry & { source_row?: number }).source_row ?? entry.sourceRow ?? 0),
      })))
    : Promise.resolve<FormulaCatalogEntry[]>([]);

export const exportFormulaCatalog = (path: string, sheet: string) =>
  isDesktop()
    ? invoke<string | null>("export_formula_catalog", { path, sheet })
    : Promise.resolve<string | null>("/demo/formula-catalog.json");

export const runAudit = (path: string, sheet: string, referencePath: string, referenceSheet: string) =>
  isDesktop()
    ? invoke<AuditResult>("run_audit", { path, sheet, referencePath, referenceSheet })
    : Promise.resolve<AuditResult>({
        id: "demo-1", submittedName: "L01GF0087.xls", processedName: "L01GF0087_calculated.xlsx",
        sheet, rows: 15, mismatches: 1, referenceFormulas: 63,
      });

export const listArchive = () =>
  isDesktop() ? invoke<ArchiveEntry[]>("list_archive") : Promise.resolve(demoArchive);

const number = (value: string | undefined) => Number.parseFloat(value ?? "") || 0;
const bool = (value: string | undefined) => value === "True" || value === "true";
const formulaMatchBlue = "#1565C0";
const formulaMismatchRed = "#B91C1C";
const inputReferences = (value: string | undefined, sourceRow: number): Record<string, string> => {
  try {
    const parsed = JSON.parse(value || "{}");
    if (parsed && typeof parsed === "object") return parsed as Record<string, string>;
  } catch {
    // Fall back for previews created before input references were persisted.
  }
  return {
    W: `F${sourceRow}`, H: `G${sourceRow}`, W1: `H${sourceRow}`, H1: `I${sourceRow}`,
    W2: `J${sourceRow}`, H2: `K${sourceRow}`, L_M: `L${sourceRow}`, R_D: `M${sourceRow}`,
    a: `N${sourceRow}`, t: `O${sourceRow}`, Q: `P${sourceRow}`, K: `Q${sourceRow}`,
  };
};
const formulaAudits = (value: string | undefined): FormulaAudit[] => {
  if (!value) return [];
  try {
    // Older preview files were read with a manual TSV split. In that case the
    // JSON field can still contain CSV quotes ("" becomes "). Accept both
    // formats so archived audits remain readable after the TSV parser fix.
    const candidates = [value.trim()];
    const quoted = value.trim();
    if (quoted.startsWith('"') && quoted.endsWith('"')) {
      candidates.push(quoted.slice(1, -1).replace(/""/g, '"'));
    }
    let parsedValue: unknown;
    for (const candidate of candidates) {
      try {
        parsedValue = JSON.parse(candidate);
        if (typeof parsedValue === "string") parsedValue = JSON.parse(parsedValue);
        break;
      } catch {
        // Try the next representation.
      }
    }
    if (!Array.isArray(parsedValue)) return [];
    const parsed = parsedValue as Array<Record<string, unknown>>;
    const formulaComparison = (value: unknown): FormulaAudit["formulaComparison"] => {
      if (!value || typeof value !== "object") return undefined;
      const comparison = value as Record<string, unknown>;
      const blocks = Array.isArray(comparison.blocks) ? comparison.blocks : [];
      return {
        status: String(comparison.status ?? "unverifiable") as NonNullable<FormulaAudit["formulaComparison"]>["status"],
        originalFormula: String(comparison.original_formula ?? comparison.originalFormula ?? ""),
        originalFormulaValue: comparison.original_formula_value == null && comparison.originalFormulaValue == null ? null : Number(comparison.original_formula_value ?? comparison.originalFormulaValue),
        detailFormulaValue: comparison.detail_formula_value == null && comparison.detailFormulaValue == null ? null : Number(comparison.detail_formula_value ?? comparison.detailFormulaValue),
        difference: comparison.difference == null ? null : Number(comparison.difference),
        error: String(comparison.error ?? ""),
        blocks: blocks.map((entry) => {
          const block = entry as Record<string, unknown>;
          const spans = block.original_spans ?? block.originalSpans;
          const label = String(block.label ?? "");
          const status = String(block.status ?? "unmatched") as "matched" | "unmatched";
          const blockColor = status === "matched" ? formulaMatchBlue : formulaMismatchRed;
          const matchedComponents = Array.isArray(block.matched_components ?? block.matchedComponents)
            ? ((block.matched_components ?? block.matchedComponents) as Array<Record<string, unknown>>).map((component) => {
              const mapSpans = (raw: unknown) => Array.isArray(raw) ? raw.map((span) => {
                const item = span as Record<string, unknown>;
                return { start: Number(item.start ?? 0), end: Number(item.end ?? 0) };
              }) : [];
              return {
                color: blockColor,
                originalSpans: mapSpans(component.original_spans ?? component.originalSpans),
                detailSpans: mapSpans(component.detail_spans ?? component.detailSpans),
                originalExpression: String(component.original_expression ?? component.originalExpression ?? ""),
                detailExpression: String(component.detail_expression ?? component.detailExpression ?? ""),
              };
            }) : [];
          return {
            label: String(block.label ?? ""),
            originalLabel: String(block.original_label ?? block.originalLabel ?? block.label ?? ""),
            expression: String(block.expression ?? ""),
            value: block.value == null ? null : Number(block.value),
            status,
            originalExpression: String(block.original_expression ?? block.originalExpression ?? ""),
            astOriginalExpression: String(block.ast_original_expression ?? block.astOriginalExpression ?? ""),
            astStandardExpression: String(block.ast_standard_expression ?? block.astStandardExpression ?? ""),
            astBlockOriginalExpressions: Array.isArray(block.ast_block_original_expressions ?? block.astBlockOriginalExpressions)
              ? ((block.ast_block_original_expressions ?? block.astBlockOriginalExpressions) as unknown[]).map(String)
              : [],
            astBlockStandardExpressions: Array.isArray(block.ast_block_standard_expressions ?? block.astBlockStandardExpressions)
              ? ((block.ast_block_standard_expressions ?? block.astBlockStandardExpressions) as unknown[]).map(String)
              : [],
            astDifferences: Array.isArray(block.ast_differences ?? block.astDifferences)
              ? ((block.ast_differences ?? block.astDifferences) as Array<Record<string, unknown>>).map((difference) => ({
                path: String(difference.path ?? ""),
                original: String(difference.original ?? ""),
                standard: String(difference.standard ?? ""),
                originalKind: String(difference.original_kind ?? difference.originalKind ?? ""),
                standardKind: String(difference.standard_kind ?? difference.standardKind ?? ""),
              }))
              : [],
            color: blockColor,
            error: String(block.error ?? ""),
            originalSpans: Array.isArray(spans) ? spans.map((span) => {
              const item = span as Record<string, unknown>;
              return { start: Number(item.start ?? 0), end: Number(item.end ?? 0) };
            }) : [],
            matchedComponents,
          };
        }),
      };
    };
    return parsed.map((audit) => ({
      target: String(audit.target ?? ""),
      label: String(audit.label ?? ""),
      originalFormula: String(audit.original_formula ?? audit.originalFormula ?? ""),
      detailFormula: String(audit.detail_formula ?? audit.detailFormula ?? ""),
      detailBlock: String(audit.detail_block ?? audit.detailBlock ?? audit.detail_formula ?? audit.detailFormula ?? ""),
      originalValue: Number(audit.original_value ?? audit.originalValue ?? 0),
      originalFormulaValue: audit.original_formula_value == null ? null : Number(audit.original_formula_value),
      calculatedValue: Number(audit.calculated_value ?? audit.calculatedValue ?? 0),
      difference: Number(audit.difference ?? 0),
      matches: Boolean(audit.matches),
      error: String(audit.error ?? ""),
      correction: String(audit.correction ?? ""),
      referenceFormula: String(audit.reference_formula ?? audit.referenceFormula ?? ""),
      referenceFormulaBlock: String(audit.reference_formula_block ?? audit.referenceFormulaBlock ?? ""),
      referenceStatus: String(audit.reference_status ?? audit.referenceStatus ?? "") as FormulaAudit["referenceStatus"],
      formulaComparison: formulaComparison(audit.formula_comparison ?? audit.formulaComparison),
      sourceFormulaComparison: formulaComparison(audit.source_formula_comparison ?? audit.sourceFormulaComparison),
    }));
  } catch {
    return [];
  }
};
const rowFromPreview = (row: Record<string, string>): ComparisonRow => ({
  sourceRow: Number.parseInt(row.source_row ?? "0", 10), itemNo: row.item_no ?? "", description: row.description ?? "",
  unit: row.unit ?? "", section: row.section ?? "", productType: row.product_type ?? "", w: number(row.w), h: number(row.h),
  w1: number(row.w1), h1: number(row.h1), w2: number(row.w2), h2: number(row.h2), lm: number(row.l_m), rd: number(row.r_d),
  a: number(row.a), t: number(row.t), q: number(row.q), inputReferences: inputReferences(row.input_references, Number.parseInt(row.source_row ?? "0", 10)), plan: row.plan ?? "", upDown: row.up_down ?? "", verticalAxis: row.vertical_axis ?? "",
  formulaDetail: row.formula_detail ?? "", acousticFireproofDetail: row.acoustic_fireproof_detail ?? "", note: row.note ?? "",
  contractorK: number(row.contractor_k), calculatedK: number(row.calculated_k), kMatch: bool(row.k_match),
  contractorScn: number(row.contractor_scn ?? row.contractor_s), calculatedScn: number(row.calculated_scn ?? row.calculated_s), scnMatch: bool(row.scn_match ?? row.s_match),
  contractorScc: number(row.contractor_scc ?? row.contractor_fireproof_s), calculatedScc: number(row.calculated_scc ?? row.calculated_fireproof_s), sccMatch: bool(row.scc_match ?? row.fireproof_match),
  contractorSta: number(row.contractor_sta ?? row.contractor_inner_s), calculatedSta: number(row.calculated_sta ?? row.calculated_inner_s), staMatch: bool(row.sta_match ?? row.inner_match),
  quantityLayout: row.quantity_layout === "three" ? "three" : "two",
  sharedQuantityTarget: row.shared_target === "S_fireproof" || row.shared_target === "S_inner" ? row.shared_target : "",
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
