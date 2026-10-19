export type Page = "process" | "archive" | "comparison";

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
  plan: string;
  upDown: string;
  verticalAxis: string;
  formulaDetail: string;
  acousticFireproofDetail: string;
  note: string;
  contractorK: number;
  calculatedK: number;
  kMatch: boolean;
  contractorS: number;
  calculatedS: number;
  sMatch: boolean;
  contractorInnerS: number;
  calculatedInnerS: number;
  innerMatch: boolean;
};

export type AuditResult = {
  id: string;
  submittedName: string;
  processedName: string;
  sheet: string;
  rows: number;
  mismatches: number;
};
