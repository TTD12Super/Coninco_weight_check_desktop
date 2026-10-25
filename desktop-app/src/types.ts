export type Page = "process" | "catalog" | "archive" | "comparison";

export type FormulaCatalogEntry = {
  name: string;
  category: string;
  target: string;
  productType: string;
  formula: string;
  formulaBlock: string;
  sourceSheet: string;
  sourceRow: number;
};

export type ArchiveEntry = {
  id: string;
  submittedName: string;
  processedName: string;
  sheet: string;
  status: "COMPLETED" | "PROCESSING" | "FAILED";
  rows: number;
  mismatches: number;
  createdAt: string;
};

export type ComparisonRow = {
  sourceRow: number;
  itemNo: string;
  description: string;
  unit: string;
  section: string;
  productType: string;
  w: number;
  h: number;
  w1: number;
  h1: number;
  w2: number;
  h2: number;
  lm: number;
  rd: number;
  a: number;
  t: number;
  q: number;
  inputReferences: Record<string, string>;
  plan: string;
  upDown: string;
  verticalAxis: string;
  formulaDetail: string;
  acousticFireproofDetail: string;
  note: string;
  contractorK: number;
  calculatedK: number;
  kMatch: boolean;
  contractorScn: number;
  calculatedScn: number;
  scnMatch: boolean;
  contractorScc: number;
  calculatedScc: number;
  sccMatch: boolean;
  contractorSta: number;
  calculatedSta: number;
  staMatch: boolean;
  formulaAudits: FormulaAudit[];
};

export type FormulaAudit = {
  target: string;
  label: string;
  originalFormula: string;
  detailFormula: string;
  detailBlock: string;
  originalValue: number;
  originalFormulaValue: number | null;
  calculatedValue: number;
  difference: number;
  matches: boolean;
  error: string;
  correction: string;
  referenceFormula?: string;
  referenceFormulaBlock?: string;
  referenceStatus?: "equivalent" | "different" | "unverifiable" | "";
  formulaComparison?: FormulaComparison;
};

export type FormulaComparison = {
  status: "equivalent" | "different" | "same_value_structure_diff" | "unverifiable";
  originalFormulaValue?: number | null;
  detailFormulaValue?: number | null;
  difference?: number | null;
  error?: string;
  blocks: Array<{
    label: string;
    expression: string;
    value?: number | null;
    status: "matched" | "unmatched";
    originalExpression: string;
    color?: string;
    originalSpans?: Array<{ start: number; end: number }>;
    error?: string;
    matchedComponents?: Array<{
      color: string;
      originalSpans: Array<{ start: number; end: number }>;
      detailSpans: Array<{ start: number; end: number }>;
      originalExpression: string;
      detailExpression: string;
    }>;
  }>;
};

export type AuditResult = {
  id: string;
  submittedName: string;
  processedName: string;
  sheet: string;
  rows: number;
  mismatches: number;
  referenceFormulas: number;
};
