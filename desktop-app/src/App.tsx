/** Main React shell: process a workbook, browse formulas, and review archives. */

import { useEffect, useMemo, useState } from "react";
import {
  Archive, ArrowLeft, ArrowUpRight, BarChart3, Check, ChevronDown, CircleHelp, CloudUpload,
  BookOpen, Download, FileSpreadsheet, FolderOpen, History, LayoutDashboard, Loader2, MoreHorizontal,
  Play, RefreshCw, Search, Settings2, ShieldCheck, Trash2, X, Zap,
} from "lucide-react";
import {
  chooseReferenceWorkbook, chooseWorkbook, clearReferenceWorkbook, deleteArchive, downloadArchive, exportFormulaCatalog, listArchive, listReferenceSheets, listSheets, loadComparison, loadFormulaCatalog,
  loadReferenceWorkbook, saveReferenceWorkbook,
  openArchiveFile, runAudit,
} from "./lib/engine";
import type { ArchiveEntry, ComparisonRow, FormulaAudit, FormulaCatalogEntry, Page } from "./types";

/** Keep displayed quantities consistent with the workbook audit precision. */
const n = (value: number) => value.toFixed(3);

/** Own the page state and coordinate all user actions with the engine bridge. */
function App() {
  const [page, setPage] = useState<Page>("process");
  const [referencePath, setReferencePath] = useState("");
  const [referenceSheets, setReferenceSheets] = useState<string[]>([]);
  const [referenceSheet, setReferenceSheet] = useState("");
  const [path, setPath] = useState("");
  const [sheets, setSheets] = useState<string[]>([]);
  const [sheet, setSheet] = useState("");
  const [archive, setArchive] = useState<ArchiveEntry[]>([]);
  const [selected, setSelected] = useState<ArchiveEntry | null>(null);
  const [rows, setRows] = useState<ComparisonRow[]>([]);
  const [catalog, setCatalog] = useState<FormulaCatalogEntry[]>([]);
  const [catalogSearch, setCatalogSearch] = useState("");
  const [search, setSearch] = useState("");
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState(0);
  const [message, setMessage] = useState("Ready for a local audit");
  const [toast, setToast] = useState<string | null>(null);
  const [pendingDelete, setPendingDelete] = useState<ArchiveEntry | null>(null);

  /** Refresh the local archive list after startup or a completed audit. */
  const refreshArchive = async () => setArchive(await listArchive());
  /** Restore the saved standard workbook and its selectable sheet on startup. */
  const restoreReference = async () => {
    try {
      const saved = await loadReferenceWorkbook();
      if (saved) {
        setReferencePath(saved);
        const available = await listReferenceSheets(saved);
        setReferenceSheets(available);
        setReferenceSheet(available[0] ?? "");
        setMessage(`${available.length} reference worksheets loaded`);
      }
    } catch (error) {
      const detail = error instanceof Error ? error.message : "Could not load saved formula reference";
      setMessage(detail);
      setToast(detail);
    }
  };
  useEffect(() => { void refreshArchive(); void restoreReference(); }, []);

  const visibleArchive = useMemo(() => archive.filter((entry) =>
    `${entry.submittedName} ${entry.processedName} ${entry.sheet}`.toLowerCase().includes(search.toLowerCase())), [archive, search]);

  const visibleCatalog = useMemo(() => catalog.filter((entry) =>
    `${entry.name} ${entry.category} ${entry.productType} ${entry.target} ${entry.formula} ${entry.formulaBlock}`
      .toLowerCase().includes(catalogSearch.toLowerCase())), [catalog, catalogSearch]);

  /** Select the contractor workbook and populate its worksheet selector. */
  const selectWorkbook = async () => {
    const chosen = await chooseWorkbook();
    if (!chosen) return;
    setPath(chosen); setSheets([]); setSheet(""); setMessage("Reading workbook sheets…");
    try { const available = await listSheets(chosen); setSheets(available); setSheet(available[0] ?? ""); setMessage(`${available.length} worksheets ready`); }
    catch (error) { setMessage(error instanceof Error ? error.message : "Could not read workbook"); }
  };

  /** Save the standard workbook, then populate its worksheet selector. */
  const selectReference = async () => {
    const chosen = await chooseReferenceWorkbook();
    if (!chosen) return;
    try {
      const saved = await saveReferenceWorkbook(chosen);
      setReferencePath(saved);
      const available = await listReferenceSheets(saved);
      setReferenceSheets(available);
      setReferenceSheet(available[0] ?? "");
      setCatalog([]);
      setMessage(`${available.length} reference worksheets loaded`);
    } catch (error) {
      const detail = error instanceof Error ? error.message : "Could not save formula reference workbook";
      setMessage(detail);
      setToast(detail);
    }
  };

  /** Load the selected standard sheet into the Formula Catalog page. */
  const openCatalog = async () => {
    if (!referencePath || !referenceSheet) {
      setPage("catalog");
      setMessage("Select and load a formula reference worksheet first");
      return;
    }
    setPage("catalog");
    setBusy(true);
    setMessage("Loading standard product formulas…");
    try {
      setCatalog(await loadFormulaCatalog(referencePath, referenceSheet));
      setMessage("Formula catalog loaded");
    } catch (error) {
      const detail = error instanceof Error ? error.message : "Could not load formula catalog";
      setMessage(detail);
      setToast(detail);
    } finally {
      setBusy(false);
    }
  };

  /** Export the currently parsed catalog as a reusable JSON file. */
  const exportCatalog = async () => {
    if (!referencePath || !referenceSheet || !catalog.length || busy) return;
    setBusy(true);
    setMessage("Exporting formula catalog JSON…");
    try {
      const saved = await exportFormulaCatalog(referencePath, referenceSheet);
      if (saved) {
        setMessage("Formula catalog exported");
        setToast(`Formula catalog saved to ${saved}`);
      } else {
        setMessage("Formula catalog export cancelled");
      }
    } catch (error) {
      const detail = error instanceof Error ? error.message : "Could not export formula catalog";
      setMessage(detail);
      setToast(detail);
    } finally {
      setBusy(false);
    }
  };

  /** Remove the saved standard workbook from local app storage. */
  const removeReference = async () => {
    if (!referencePath || busy) return;
    try {
      await clearReferenceWorkbook();
      setReferencePath("");
      setReferenceSheets([]);
      setReferenceSheet("");
      setCatalog([]);
      setMessage("Saved formula reference removed");
    } catch (error) {
      const detail = error instanceof Error ? error.message : "Could not remove formula reference";
      setMessage(detail);
      setToast(detail);
    }
  };

  /** Start the end-to-end Run Audit workflow and refresh the archive. */
  const audit = async () => {
    if (!referencePath || !referenceSheet || !path || !sheet) return;
    setBusy(true); setProgress(12); setMessage("Copying submitted workbook into the local archive…");
    const timer = window.setInterval(() => setProgress((value) => Math.min(value + 11, 88)), 450);
    try {
      const result = await runAudit(path, sheet, referencePath, referenceSheet); setProgress(100); setMessage("Audit complete — processed workbook is ready");
      setToast(result.rows + " rows processed · " + result.referenceFormulas + " reference formulas loaded"); await refreshArchive();
    } catch (error) { setMessage(error instanceof Error ? error.message : "Audit failed"); }
    finally { window.clearInterval(timer); window.setTimeout(() => setProgress(0), 800); setBusy(false); }
  };

  /** Load one archived audit's preview rows for the comparison page. */
  const viewComparison = async (entry: ArchiveEntry) => {
    setSelected(entry); setBusy(true); setMessage("Loading comparison data…");
    try { setRows(await loadComparison(entry.id)); setPage("comparison"); }
    catch (error) { setMessage(error instanceof Error ? error.message : "Could not load comparison"); }
    finally { setBusy(false); }
  };

  /** Delete the selected archive entry and its stored files. */
  const removeEntry = async (entry: ArchiveEntry) => {
    try {
      await deleteArchive(entry.id);
      await refreshArchive();
      setToast(`Deleted ${entry.submittedName}`);
    } catch (error) {
      setToast(error instanceof Error ? error.message : "Could not delete archive entry");
    }
  };

  /** Confirm the pending destructive archive deletion. */
  const confirmDelete = async () => {
    if (!pendingDelete) return;
    const entry = pendingDelete;
    setPendingDelete(null);
    await removeEntry(entry);
  };

  return <div className="min-h-screen bg-[#F7F7F7] text-coninco-ink">
    <aside className="fixed inset-y-0 left-0 z-20 flex w-[248px] flex-col border-r border-coninco-line bg-white">
      <div className="flex items-center gap-3 border-b border-coninco-line px-7 py-6">
        <img src="/coninco-logo.png" alt="CONINCO" className="h-12 w-12 object-contain" />
        <div><div className="text-[15px] font-black tracking-[.18em] text-coninco-ink">CONINCO</div><div className="text-[10px] font-bold tracking-[.28em] text-coninco-red">EXCEL AUDITOR</div></div>
      </div>
      <div className="px-5 pt-8 text-[10px] font-black uppercase tracking-[.2em] text-coninco-muted">Workspace</div>
      <nav className="space-y-2 px-4 pt-3">
        <NavButton active={page === "process"} icon={<LayoutDashboard size={18}/>} label="Process" onClick={() => setPage("process")} />
        <NavButton active={page === "catalog"} icon={<BookOpen size={18}/>} label="Formula catalog" onClick={() => void openCatalog()} badge={catalog.length || undefined} />
        <NavButton active={page === "archive"} icon={<Archive size={18}/>} label="Archive" onClick={() => { void refreshArchive(); setPage("archive"); }} badge={archive.length} />
      </nav>
      <div className="mt-auto space-y-4 border-t border-coninco-line px-5 py-5">
        <div className="rounded-2xl bg-coninco-soft p-4"><div className="flex items-center gap-2 text-xs font-bold text-coninco-red"><ShieldCheck size={15}/> Offline secure mode</div><p className="mt-2 text-[11px] leading-5 text-coninco-muted">Files stay inside this computer. No cloud upload.</p></div>
        <div className="flex items-center justify-between text-[11px] text-coninco-muted"><span>v0.2.0</span><CircleHelp size={15}/></div>
      </div>
    </aside>
    <main className="ml-[248px] min-h-screen">
      <header className="flex h-[76px] items-center justify-between border-b border-coninco-line bg-white/90 px-10 backdrop-blur">
        <div><div className="text-[11px] font-black uppercase tracking-[.2em] text-coninco-red">Local quantity verification</div><div className="mt-1 text-sm text-coninco-muted">{page === "process" ? "Run a controlled workbook audit" : page === "catalog" ? "Browse standard products and formulas" : page === "archive" ? "Manage stored workbooks" : "Inspect source and engine values"}</div></div>
        <div className="flex items-center gap-3"><div className="flex items-center gap-2 rounded-full border border-emerald-200 bg-emerald-50 px-3 py-2 text-[11px] font-bold text-emerald-700"><span className="h-2 w-2 animate-pulse rounded-full bg-emerald-500"/> OFFLINE MODE</div><button className="rounded-xl border border-coninco-line p-2.5 text-coninco-muted hover:border-coninco-red hover:text-coninco-red"><Settings2 size={17}/></button></div>
      </header>
      <div className="p-10">{page === "process" && <ProcessPage referencePath={referencePath} referenceSheet={referenceSheet} referenceSheets={referenceSheets} path={path} sheet={sheet} sheets={sheets} busy={busy} progress={progress} message={message} selectReference={selectReference} removeReference={removeReference} selectWorkbook={selectWorkbook} setReferenceSheet={setReferenceSheet} setSheet={setSheet} audit={audit} setPage={setPage}/>} {page === "catalog" && <CatalogPage entries={visibleCatalog} search={catalogSearch} setSearch={setCatalogSearch} referencePath={referencePath} referenceSheet={referenceSheet} message={message} busy={busy} canExport={catalog.length > 0} exportCatalog={exportCatalog}/>} {page === "archive" && <ArchivePage entries={visibleArchive} search={search} setSearch={setSearch} viewComparison={viewComparison} removeEntry={setPendingDelete} setToast={setToast}/>} {page === "comparison" && selected && <ComparisonPage entry={selected} rows={rows} back={() => setPage("archive")} setToast={setToast}/>}</div>
    </main>
    {toast && <div className="fixed bottom-7 right-8 z-50 flex items-center gap-3 rounded-2xl bg-coninco-ink px-5 py-4 text-sm text-white shadow-2xl animate-enter"><Check size={17} className="text-emerald-300"/>{toast}<button onClick={() => setToast(null)}><X size={15}/></button></div>}
    {pendingDelete && <div className="fixed inset-0 z-50 flex items-center justify-center bg-coninco-ink/45 p-6 backdrop-blur-sm"><div role="dialog" aria-modal="true" className="w-full max-w-lg rounded-[24px] border border-coninco-line bg-white p-7 shadow-2xl animate-enter"><div className="flex min-w-0 items-start gap-4"><div className="shrink-0 rounded-2xl bg-red-50 p-3 text-coninco-red"><Trash2 size={22}/></div><div className="min-w-0 flex-1"><h2 className="text-lg font-black text-coninco-ink">Delete archived audit?</h2><p className="mt-2 text-sm leading-6 text-coninco-muted">This will permanently remove the submitted workbook, processed workbook, and comparison data from this application.</p><p className="mt-3 min-w-0 max-w-full truncate rounded-xl bg-[#f7f7f7] px-3 py-2 text-xs font-bold text-coninco-ink">{pendingDelete.submittedName}</p></div></div><div className="mt-7 flex justify-end gap-3"><button onClick={() => setPendingDelete(null)} className="rounded-xl border border-coninco-line px-4 py-3 text-xs font-black text-coninco-muted hover:border-coninco-red hover:text-coninco-red">CANCEL</button><button onClick={() => void confirmDelete()} className="flex items-center gap-2 rounded-xl bg-coninco-red px-4 py-3 text-xs font-black text-white shadow-red hover:bg-coninco-dark"><Trash2 size={15}/> DELETE</button></div></div></div>}
  </div>;
}

/** Render one sidebar navigation item with active state and optional count. */
function NavButton({ active, icon, label, badge, onClick }: { active: boolean; icon: React.ReactNode; label: string; badge?: number; onClick: () => void }) {
  return <button onClick={onClick} className={`group flex w-full items-center gap-3 rounded-2xl px-4 py-3 text-left text-sm font-bold transition ${active ? "bg-coninco-red text-white shadow-red" : "text-coninco-muted hover:bg-coninco-soft hover:text-coninco-red"}`}><span className={active ? "text-white" : "text-coninco-muted group-hover:text-coninco-red"}>{icon}</span>{label}{badge !== undefined && <span className={`ml-auto rounded-full px-2 py-0.5 text-[10px] ${active ? "bg-white/20" : "bg-[#f1f2f3]"}`}>{badge}</span>}</button>;
}

const catalogCategoryColors: Record<string, { text: string; background: string; border: string }> = {
  "Cách nhiệt": { text: "#2563EB", background: "#EFF6FF", border: "#BFDBFE" },
  "Chống cháy": { text: "#DC2626", background: "#FEF2F2", border: "#FECACA" },
  "Tiêu âm": { text: "#7C3AED", background: "#F5F3FF", border: "#DDD6FE" },
};
const catalogTargetColors: Record<string, string> = { S: "#0F766E", SCN: "#2563EB", SCC: "#DC2626", STA: "#7C3AED", "SCC/TA": "#C2410C" };
const catalogLineColors: Record<string, string> = { S: "#0F766E", SCN: "#2563EB", SCC: "#DC2626", STA: "#7C3AED", "SCC/TA": "#C2410C" };
/** Turn the internal product key into a readable catalog label. */
const prettyProductType = (value: string) => value ? value.replace(/_/g, " ").replace(/\b\w/g, (letter) => letter.toUpperCase()) : "Unknown product type";
/** Show both categories for a shared SCC/TA formula. */
const catalogCategoriesForEntry = (entry: FormulaCatalogEntry) => entry.target === "SCC/TA" ? ["Chống cháy", "Tiêu âm"] : [entry.category || "Uncategorized"];
/** Make extracted formula text readable without changing its meaning. */
const formatFormulaForDisplay = (value: string) => value
  .replace(/\s+/g, " ")
  .replace(/\(\s+/g, "(")
  .replace(/\s+\)/g, ")")
  .replace(/\s*([+*/^])\s*/g, " $1 ")
  .replace(/\s*-\s*/g, " - ")
  .replace(/\s*,\s*/g, ", ")
  .replace(/\s*(<=|>=|<>|=)\s*/g, " $1 ")
  .replace(/ {2,}/g, " ")
  .trim();

/** Display grouped standard products and their formula blocks. */
function CatalogPage({ entries, search, setSearch, referencePath, referenceSheet, message, busy, canExport, exportCatalog }: { entries: FormulaCatalogEntry[]; search: string; setSearch: (value: string) => void; referencePath: string; referenceSheet: string; message: string; busy: boolean; canExport: boolean; exportCatalog: () => void }) {
  const categories = new Set(entries.flatMap(catalogCategoriesForEntry));
  const groups = useMemo(() => {
    const grouped = new Map<string, { name: string; productType: string; sourceRow: number; formulas: FormulaCatalogEntry[] }>();
    entries.forEach((entry) => {
      const key = `${entry.name}\u0000${entry.productType}\u0000${entry.sourceRow}`;
      const current = grouped.get(key);
      if (current) current.formulas.push(entry);
      else grouped.set(key, { name: entry.name, productType: entry.productType, sourceRow: entry.sourceRow, formulas: [entry] });
    });
    return [...grouped.values()];
  }, [entries]);
  return <div className="animate-enter space-y-7">
    <div className="flex items-end justify-between gap-6"><div><div className="text-[11px] font-black uppercase tracking-[.2em] text-coninco-red">Standard formula library</div><h1 className="mt-2 text-3xl font-black tracking-tight">Product formula catalog</h1><p className="mt-2 max-w-2xl text-sm text-coninco-muted">Browse the products, categories and readable formula blocks extracted from the saved reference workbook.</p></div><div className="flex items-center gap-3"><button onClick={exportCatalog} disabled={!canExport || busy} className="flex items-center gap-2 rounded-2xl border border-coninco-red bg-white px-4 py-3 text-xs font-black text-coninco-red shadow-soft transition hover:bg-coninco-soft disabled:cursor-not-allowed disabled:opacity-50"><Download size={16}/>{busy ? "EXPORTING…" : "EXPORT JSON"}</button><div className="rounded-2xl bg-coninco-ink px-5 py-3 text-right text-white"><div className="text-2xl font-black">{entries.length}</div><div className="text-[10px] font-bold uppercase tracking-wider text-white/55">visible formulas</div></div></div></div>
    <div className="flex flex-wrap items-center gap-3 rounded-2xl border border-coninco-line bg-white px-5 py-4 shadow-soft"><BookOpen size={18} className="text-coninco-red"/><div className="min-w-0 flex-1"><div className="truncate text-xs font-black text-coninco-ink">{referencePath ? referencePath.split(/[\\/]/).pop() : "No formula reference workbook selected"}</div><div className="mt-1 text-[11px] text-coninco-muted">{referenceSheet ? `Worksheet: ${referenceSheet} · ${message}` : "Select a reference workbook and worksheet from Process first."}</div></div><div className="rounded-full bg-coninco-soft px-3 py-1.5 text-[10px] font-black uppercase tracking-wider text-coninco-red">{categories.size} categories</div></div>
    <div className="flex items-center gap-3 rounded-2xl border border-coninco-line bg-white p-3 shadow-soft"><Search size={18} className="ml-2 text-coninco-muted"/><input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Search product, category, type or formula…" className="flex-1 bg-transparent px-2 py-2 text-sm outline-none"/><span className="rounded-xl bg-[#f1f2f3] px-3 py-2 text-[10px] font-black text-coninco-muted">{entries.length} ITEMS</span></div>
    {!entries.length && <div className="rounded-[24px] border border-dashed border-coninco-red/30 bg-white px-6 py-16 text-center shadow-soft"><BookOpen className="mx-auto text-coninco-red" size={32}/><div className="mt-4 text-base font-black">No formula catalog entries</div><p className="mt-2 text-sm text-coninco-muted">Load the standard reference workbook and open this page again.</p></div>}
    <div className="grid gap-5">{groups.map((group) => {
      const groupCategories = [...new Set(group.formulas.flatMap(catalogCategoriesForEntry))];
      return <article key={`${group.sourceRow}-${group.name}-${group.productType}`} className="overflow-hidden rounded-[24px] border border-coninco-line bg-white shadow-soft transition hover:-translate-y-0.5 hover:shadow-lg">
        <div className="flex items-start justify-between gap-4 border-b border-coninco-line px-6 py-5"><div className="min-w-0"><div className="mb-2 flex flex-wrap items-center gap-2">{groupCategories.map((category) => { const color = catalogCategoryColors[category] ?? { text: "#475569", background: "#F8FAFC", border: "#CBD5E1" }; return <span key={category} className="rounded-full border px-3 py-1 text-[10px] font-black uppercase tracking-wider" style={{ color: color.text, backgroundColor: color.background, borderColor: color.border }}>{category || "Uncategorized"}</span>; })}</div><h2 className="truncate text-base font-black text-coninco-ink" title={group.name}>{group.name || "Unnamed product"}</h2><div className="mt-1 text-xs font-semibold text-coninco-muted">{prettyProductType(group.productType)} · {group.formulas.length} formula{group.formulas.length === 1 ? "" : "s"}</div></div><span className="shrink-0 rounded-xl bg-coninco-soft px-3 py-2 text-[10px] font-black text-coninco-red">ROW {group.sourceRow}</span></div>
        <div className="space-y-4 p-6">{[...group.formulas.reduce((blocks, entry) => { const key = entry.formulaBlock || `${entry.target} = ${entry.formula}`; const current = blocks.get(key); if (current) current.push(entry); else blocks.set(key, [entry]); return blocks; }, new Map<string, FormulaCatalogEntry[]>()).values()].map((blockEntries, blockIndex) => {
          const firstEntry = blockEntries[0];
          const blockCategories = [...new Set(blockEntries.flatMap(catalogCategoriesForEntry))];
          const blockTargets = [...new Set(blockEntries.map((entry) => entry.target).filter(Boolean))];
          const categoryColor = catalogCategoryColors[blockCategories[0]] ?? { text: "#475569", background: "#F8FAFC", border: "#CBD5E1" };
          const lines = (firstEntry.formulaBlock || `${firstEntry.target} = ${firstEntry.formula}`).split(/\r?\n/).filter(Boolean);
          return <div key={`${firstEntry.sourceRow}-${blockIndex}-${firstEntry.formulaBlock}`} className="overflow-hidden rounded-2xl border" style={{ borderColor: categoryColor.border }}>
            <div className="flex items-center justify-between gap-3 px-4 py-3" style={{ backgroundColor: categoryColor.background }}><div className="flex flex-wrap items-center gap-2">{blockCategories.map((category) => <span key={category} className="rounded-full border px-2.5 py-1 text-[10px] font-black uppercase tracking-wider" style={{ color: (catalogCategoryColors[category] ?? categoryColor).text, borderColor: (catalogCategoryColors[category] ?? categoryColor).border }}>{category || "Uncategorized"}</span>)}{blockTargets.map((target) => <span key={target} className="rounded-full px-2.5 py-1 text-[10px] font-black uppercase tracking-wider text-white" style={{ backgroundColor: catalogTargetColors[target] ?? "#475569" }}>{target}</span>)}</div><span className="text-[10px] font-bold uppercase tracking-wider" style={{ color: categoryColor.text }}>formula block</span></div>
            <div className="space-y-1 overflow-x-auto bg-[#fbfbfc] p-3">{lines.map((line, lineIndex) => { const equalsIndex = line.indexOf("="); const label = (equalsIndex >= 0 ? line.slice(0, equalsIndex) : firstEntry.target).trim().toUpperCase(); const rhs = equalsIndex >= 0 ? line.slice(equalsIndex + 1) : line; const color = catalogLineColors[label] ?? catalogTargetColors[firstEntry.target] ?? "#475569"; return <div key={`${lineIndex}-${line}`} className="min-w-max rounded-lg border-l-4 bg-white px-3 py-2 font-mono text-xs leading-5 text-coninco-ink shadow-sm" style={{ borderLeftColor: color }}><span className="mr-2 font-black" style={{ color }}>{label} =</span><span className="whitespace-nowrap">{formatFormulaForDisplay(rhs)}</span></div>; })}</div>
            <div className="space-y-2 overflow-x-auto border-t border-dashed px-4 py-3 text-[11px] text-coninco-muted" style={{ borderColor: categoryColor.border }}>{blockEntries.map((entry) => <div key={`${entry.target}-${entry.category}`} className="min-w-max whitespace-nowrap"><span className="font-black text-coninco-ink">{entry.target} output:</span> <span className="font-mono">{formatFormulaForDisplay(entry.formula)}</span></div>)}</div>
          </div>;
        })}</div>
      </article>;
    })}</div>
  </div>;
}

/** Render the four-step input screen: reference, contractor, sheet, audit. */
function ProcessPage({ referencePath, referenceSheet, referenceSheets, path, sheet, sheets, busy, progress, message, selectReference, removeReference, selectWorkbook, setReferenceSheet, setSheet, audit, setPage }: { referencePath: string; referenceSheet: string; referenceSheets: string[]; path: string; sheet: string; sheets: string[]; busy: boolean; progress: number; message: string; selectReference: () => void; removeReference: () => void; selectWorkbook: () => void; setReferenceSheet: (value: string) => void; setSheet: (value: string) => void; audit: () => void; setPage: (page: Page) => void }) {
  return <div className="animate-enter space-y-8">
    <div className="relative overflow-hidden rounded-[28px] bg-coninco-ink px-9 py-8 text-white shadow-soft"><div className="absolute -right-16 -top-24 h-72 w-72 rounded-full bg-coninco-red/25 blur-3xl"/><div className="absolute bottom-[-100px] right-[30%] h-52 w-52 rounded-full bg-white/10 blur-3xl"/><div className="relative flex items-end justify-between"><div><div className="mb-4 inline-flex items-center gap-2 rounded-full border border-white/15 bg-white/10 px-3 py-1.5 text-[10px] font-bold uppercase tracking-[.18em] text-white/75"><Zap size={13} className="text-coninco-red"/> Precision audit workspace</div><h1 className="max-w-xl text-4xl font-black leading-tight tracking-tight">Turn contractor workbooks into <span className="text-coninco-red">clear evidence.</span></h1><p className="mt-3 max-w-lg text-sm leading-6 text-white/60">Read, calculate and compare locally with the original workbook always preserved.</p></div><div className="hidden w-48 animate-float rounded-3xl border border-white/10 bg-white/5 p-5 md:block"><BarChart3 className="mb-4 text-coninco-red" size={30}/><div className="text-2xl font-black">100%</div><div className="mt-1 text-xs text-white/50">local processing</div></div></div></div>
    <div className="flex items-center gap-3 rounded-2xl border border-coninco-line bg-white px-5 py-4 text-xs text-coninco-muted shadow-soft"><CircleHelp size={16} className="shrink-0 text-coninco-red"/><span className="font-semibold">{message}</span></div>
    <section className="rounded-[24px] border border-coninco-red/20 bg-white p-7 shadow-soft"><SectionHeading number="01" title="Select formula reference workbook" subtitle="Read standard product names, categories and formulas only. This file is never calculated."/><button onClick={selectReference} disabled={busy} className="group mt-6 flex w-full items-center gap-5 rounded-2xl border-2 border-dashed border-coninco-red/30 bg-coninco-soft/35 p-6 text-left transition hover:border-coninco-red hover:bg-coninco-soft disabled:cursor-not-allowed disabled:opacity-60"><span className="flex h-14 w-14 items-center justify-center rounded-2xl bg-white text-coninco-red transition group-hover:scale-105"><FileSpreadsheet size={26}/></span><span className="min-w-0 flex-1"><span className="block truncate text-sm font-black text-coninco-ink">{referencePath ? referencePath.split(/[\\/]/).pop() : "Choose the standard formula file"}</span><span className="mt-1 block truncate text-xs text-coninco-muted">{referencePath || "Supports .xlsx, .xls and Apple .numbers files"}</span></span><span className="rounded-xl bg-coninco-red px-4 py-2.5 text-xs font-black text-white shadow-red">BROWSE</span></button>{referencePath && <div className="relative mt-5"><div className="mb-2 text-[10px] font-black uppercase tracking-wider text-coninco-muted">Reference worksheet</div><select value={referenceSheet} onChange={(event) => setReferenceSheet(event.target.value)} disabled={!referenceSheets.length || busy} className="w-full appearance-none rounded-2xl border border-coninco-line bg-[#fbfbfc] px-4 py-3.5 text-sm font-bold text-coninco-ink outline-none transition focus:border-coninco-red focus:ring-4 focus:ring-coninco-red/10"><option value="">Select a reference worksheet</option>{referenceSheets.map((item) => <option key={item}>{item}</option>)}</select><ChevronDown className="pointer-events-none absolute right-4 top-9 text-coninco-muted" size={18}/><div className="mt-2 text-[11px] text-coninco-muted">{referenceSheets.length ? `${referenceSheets.length} reference worksheets detected` : "Reference worksheets will appear here"}</div></div>}{referencePath && <button onClick={removeReference} disabled={busy} className="mt-4 inline-flex items-center gap-2 rounded-xl border border-red-200 bg-white px-4 py-2.5 text-[10px] font-black uppercase tracking-wider text-coninco-red transition hover:bg-red-50 disabled:cursor-not-allowed disabled:opacity-50"><Trash2 size={13}/> Remove saved reference file</button>}</section>
    <div className="grid grid-cols-[1.3fr_.7fr] gap-6">
      <section className="rounded-[24px] border border-coninco-line bg-white p-7 shadow-soft"><SectionHeading number="02" title="Select contractor workbook" subtitle="The submitted copy is stored inside the application archive."/><button onClick={selectWorkbook} className="group mt-6 flex w-full items-center gap-5 rounded-2xl border-2 border-dashed border-[#dfe1e3] bg-[#fbfbfc] p-6 text-left transition hover:border-coninco-red hover:bg-coninco-soft"><span className="flex h-14 w-14 items-center justify-center rounded-2xl bg-coninco-soft text-coninco-red transition group-hover:scale-105"><CloudUpload size={26}/></span><span className="min-w-0 flex-1"><span className="block text-sm font-black text-coninco-ink">{path ? path.split(/[\\/]/).pop() : "Drop an Excel workbook here"}</span><span className="mt-1 block truncate text-xs text-coninco-muted">{path || "Supports .xlsx and legacy .xls files"}</span></span><span className="rounded-xl bg-coninco-red px-4 py-2.5 text-xs font-black text-white shadow-red">BROWSE</span></button></section>
      <section className="rounded-[24px] border border-coninco-line bg-white p-7 shadow-soft"><SectionHeading number="03" title="Choose worksheet" subtitle="Select the source sheet to audit."/><div className="relative mt-6"><select value={sheet} onChange={(event) => setSheet(event.target.value)} disabled={!sheets.length || busy} className="w-full appearance-none rounded-2xl border border-coninco-line bg-[#fbfbfc] px-4 py-3.5 text-sm font-bold text-coninco-ink outline-none transition focus:border-coninco-red focus:ring-4 focus:ring-coninco-red/10"><option value="">Select a worksheet</option>{sheets.map((item) => <option key={item}>{item}</option>)}</select><ChevronDown className="pointer-events-none absolute right-4 top-4 text-coninco-muted" size={18}/></div><div className="mt-4 flex items-center gap-2 text-[11px] text-coninco-muted"><FileSpreadsheet size={14} className="text-coninco-red"/>{sheets.length ? `${sheets.length} worksheets detected` : "Workbook sheets will appear here"}</div></section>
    </div>
    <section className="relative overflow-hidden rounded-[24px] border border-coninco-red/20 bg-coninco-soft p-7 shadow-soft"><div className="absolute -right-10 -top-16 h-48 w-48 rounded-full bg-coninco-red/10 blur-3xl"/><div className="relative flex items-center justify-between"><SectionHeading number="04" title="Run audit" subtitle="Calculate from standard formulas; inspect formula text only when the saved result differs." accent/><button onClick={audit} disabled={!referencePath || !referenceSheet || !path || !sheet || busy} className="flex items-center gap-3 rounded-2xl bg-coninco-red px-7 py-4 text-sm font-black text-white shadow-red transition hover:bg-coninco-dark disabled:cursor-not-allowed disabled:opacity-40">{busy ? <Loader2 className="animate-spin" size={18}/> : <Play size={18} fill="currentColor"/>}{busy ? "PROCESSING…" : "RUN AUDIT"}</button></div>{busy && <div className="relative mt-6"><div className="mb-2 flex justify-between text-[11px] font-bold text-coninco-muted"><span>{message}</span><span>{progress}%</span></div><div className="h-2 overflow-hidden rounded-full bg-white"><div className="h-full rounded-full bg-coninco-red transition-all duration-500" style={{width: `${progress}%`}}/></div></div>}</section>
    <div className="flex items-center justify-between rounded-2xl border border-coninco-line bg-white px-6 py-5"><div className="flex items-center gap-3"><History size={18} className="text-coninco-red"/><div><div className="text-sm font-black">Need to review a previous audit?</div><div className="mt-1 text-xs text-coninco-muted">Open the local archive to compare submitted and processed files.</div></div></div><button onClick={() => setPage("archive")} className="flex items-center gap-2 text-xs font-black text-coninco-red hover:text-coninco-dark">OPEN ARCHIVE <ArrowUpRight size={15}/></button></div>
  </div>;
}

/** Render the numbered heading shared by each process section. */
function SectionHeading({ number, title, subtitle, accent = false }: { number: string; title: string; subtitle: string; accent?: boolean }) { return <div className="flex items-start gap-3"><span className={`flex h-9 w-9 shrink-0 items-center justify-center rounded-xl text-xs font-black ${accent ? "bg-white text-coninco-red" : "bg-coninco-soft text-coninco-red"}`}>{number}</span><div><h2 className={`text-base font-black ${accent ? "text-coninco-ink" : "text-coninco-ink"}`}>{title}</h2><p className="mt-1 text-xs leading-5 text-coninco-muted">{subtitle}</p></div></div>; }

/** Render archived audit entries and actions. */
function ArchivePage({ entries, search, setSearch, viewComparison, removeEntry, setToast }: { entries: ArchiveEntry[]; search: string; setSearch: (value: string) => void; viewComparison: (entry: ArchiveEntry) => void; removeEntry: (entry: ArchiveEntry) => void; setToast: (value: string) => void }) {
  return <div className="animate-enter space-y-7"><div className="flex items-end justify-between"><div><div className="text-[11px] font-black uppercase tracking-[.2em] text-coninco-red">Workspace archive</div><h1 className="mt-2 text-3xl font-black tracking-tight">Stored workbooks</h1><p className="mt-2 text-sm text-coninco-muted">Submitted originals and processed results stay paired on this computer.</p></div><div className="rounded-2xl bg-coninco-ink px-5 py-3 text-right text-white"><div className="text-2xl font-black">{entries.length}</div><div className="text-[10px] font-bold uppercase tracking-wider text-white/55">saved audits</div></div></div><div className="flex items-center gap-3 rounded-2xl border border-coninco-line bg-white p-3 shadow-soft"><Search size={18} className="ml-2 text-coninco-muted"/><input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Search by workbook or sheet…" className="flex-1 bg-transparent px-2 py-2 text-sm outline-none"/><button onClick={() => setToast("Archive is stored locally") } className="rounded-xl border border-coninco-line px-4 py-2 text-xs font-black text-coninco-muted hover:border-coninco-red hover:text-coninco-red"><RefreshCw size={14}/></button></div><div className="overflow-hidden rounded-[24px] border border-coninco-line bg-white shadow-soft"><div className="border-b border-coninco-line px-6 py-5"><div className="flex items-center gap-2 text-sm font-black"><Archive size={17} className="text-coninco-red"/> Audit history</div></div><div className="overflow-x-auto"><table className="table-grid w-full min-w-[1320px] border-collapse text-left"><thead className="bg-[#fbfbfc] text-[10px] font-black uppercase tracking-wider text-coninco-muted"><tr><th className="px-5 py-4">Submitted workbook</th><th className="px-5 py-4">Processed result</th><th className="px-5 py-4">Sheet</th><th className="px-5 py-4">Rows</th><th className="px-5 py-4">Status</th><th className="px-5 py-4 text-right">Actions</th></tr></thead><tbody>{entries.map((entry) => <tr key={entry.id} className="group hover:bg-[#fffafa]"><td className="px-5 py-5"><div className="flex items-center gap-3"><span className="rounded-xl bg-coninco-soft p-2.5 text-coninco-red"><FileSpreadsheet size={18}/></span><div><div className="text-sm font-bold">{entry.submittedName}</div><div className="mt-1 text-[11px] text-coninco-muted">{entry.createdAt}</div></div></div></td><td className="px-5 py-5"><div className="text-sm font-bold text-coninco-ink">{entry.processedName}</div><div className="mt-1 text-[11px] text-coninco-muted">{entry.mismatches} mismatched cells</div></td><td className="px-5 py-5"><span className="rounded-lg bg-[#f1f2f3] px-2.5 py-1 text-xs font-black">{entry.sheet}</span></td><td className="px-5 py-5 text-sm font-black">{entry.rows}</td><td className="px-5 py-5"><span className={`rounded-full px-3 py-1.5 text-[10px] font-black ${entry.status === "COMPLETED" ? "bg-emerald-50 text-emerald-700" : "bg-red-50 text-red-700"}`}>{entry.status}</span></td><td className="whitespace-nowrap px-5 py-5"><div className="flex justify-end gap-2"><button onClick={() => viewComparison(entry)} className="rounded-xl bg-coninco-red px-3 py-2 text-[10px] font-black text-white hover:bg-coninco-dark">VIEW COMPARISON</button><DownloadButton entry={entry} setToast={setToast}/><button onClick={() => void removeEntry(entry)} title="Delete" className="rounded-xl border border-red-100 p-2 text-red-400 hover:bg-red-50"><Trash2 size={15}/></button></div></td></tr>)}</tbody></table>{!entries.length && <div className="px-6 py-14 text-center text-sm text-coninco-muted">No archived workbooks found.</div>}</div></div></div>;
}

/** Render row-level values, match flags, and formula-error details. */
function ComparisonPage({ entry, rows, back, setToast }: { entry: ArchiveEntry; rows: ComparisonRow[]; back: () => void; setToast: (value: string) => void }) {
  const inputHeaders = ["ROW", "ITEM NO", "PRODUCT NAME", "UNIT", "SECTION", "PRODUCT TYPE", "W", "H", "W'", "H'", "W''", "H''", "L/M", "R/D", "a", "t", "Q", "PLAN", "UP-DOWN", "VERTICAL AXIS", "FORMULA DETAIL", "ACOUSTIC / FIREPROOF", "NOTE"];
  const hasSeparateFireproofAndAcoustic = rows.some((row) => row.quantityLayout === "three");
  const outputHeaders = hasSeparateFireproofAndAcoustic
    ? ["K SOURCE", "K ENGINE", "K CHECK", "SCN SOURCE", "SCN ENGINE", "SCN CHECK", "SCC SOURCE", "SCC ENGINE", "SCC CHECK", "STA SOURCE", "STA ENGINE", "STA CHECK"]
    : ["K SOURCE", "K ENGINE", "K CHECK", "S SOURCE", "S ENGINE", "S CHECK", "SCC/TA SOURCE", "SCC/TA ENGINE", "SCC/TA CHECK"];
  const formulaErrors = rows
    .flatMap((row) => mismatchedFormulaAudits(row).map((audit) => ({ row, audit })))
    .sort((left, right) => left.row.sourceRow - right.row.sourceRow || formulaTargetOrder(left.audit.target) - formulaTargetOrder(right.audit.target));

  return <div className="animate-enter space-y-7">
    <div className="flex items-center gap-4"><button onClick={back} className="rounded-xl border border-coninco-line bg-white p-3 text-coninco-muted hover:border-coninco-red hover:text-coninco-red"><ArrowLeft size={18}/></button><div className="flex-1"><div className="text-[11px] font-black uppercase tracking-[.2em] text-coninco-red">Comparison detail · Sheet {entry.sheet}</div><h1 className="mt-2 text-3xl font-black tracking-tight">{entry.processedName}</h1><p className="mt-2 text-sm text-coninco-muted">One row contains the source inputs and all calculated audit outputs.</p></div><button onClick={() => void openArchiveFile(entry.id, "processed")} className="flex items-center gap-2 rounded-xl border border-coninco-line bg-white px-4 py-3 text-xs font-black text-coninco-red hover:bg-coninco-soft"><FolderOpen size={16}/> OPEN WORKBOOK</button></div>
    <div className="grid grid-cols-4 gap-4"><Metric label="Rows compared" value={String(entry.rows)} icon={<BarChart3 size={17}/>} /><Metric label={hasSeparateFireproofAndAcoustic ? "K + SCN + SCC + STA" : "K + S + SCC/TA"} value={String(entry.rows * (hasSeparateFireproofAndAcoustic ? 4 : 3))} icon={<ShieldCheck size={17}/>} /><Metric label="Mismatches" value={String(entry.mismatches)} icon={<Zap size={17}/>} danger={entry.mismatches > 0}/><Metric label="Engine" value="Python" icon={<Settings2 size={17}/>} /></div>
    <FormulaErrorTable items={formulaErrors}/>
    <div className="overflow-hidden rounded-[24px] border border-coninco-line bg-white shadow-soft">
      <div className="border-b border-coninco-line px-6 py-5"><div className="flex items-center gap-2 text-sm font-black"><FileSpreadsheet size={17} className="text-coninco-red"/> Complete comparison table</div><p className="mt-1 text-xs text-coninco-muted">Each row combines contractor inputs, source quantities, engine quantities, and match status. Scroll horizontally to inspect all fields.</p></div>
      <div className="overflow-auto">
        <table className="table-grid min-w-[3600px] border-collapse text-left text-[11px]">
          <thead className="bg-[#fbfbfc] text-[10px] font-black uppercase tracking-wider text-coninco-muted">
            <tr><th colSpan={inputHeaders.length} className="border-r-2 border-coninco-line bg-coninco-soft px-4 py-3 text-coninco-red">INPUT DATA · SOURCE ROW VALUES</th><th colSpan={outputHeaders.length} className="bg-[#fff4f4] px-4 py-3 text-coninco-red">AUDIT OUTPUT · K / S / SCC-TA</th></tr>
            <tr>{inputHeaders.map((label) => <th key={label} className="whitespace-nowrap border-b border-coninco-line px-4 py-4">{label}</th>)}{outputHeaders.map((label) => <th key={label} className="whitespace-nowrap border-b border-l border-coninco-line px-4 py-4">{label}</th>)}</tr>
          </thead>
          <tbody>{rows.map((row) => <tr key={row.sourceRow} className="hover:bg-[#fffafa]">
            <td className="px-4 py-4 font-black text-coninco-red">{row.sourceRow}</td><td className="px-4 py-4">{row.itemNo}</td><td className="max-w-[260px] px-4 py-4 font-semibold">{row.description}</td><td className="px-4 py-4">{row.unit || "—"}</td><td className="px-4 py-4"><Tag value={row.section}/></td><td className="px-4 py-4">{row.productType}</td>
            {[row.w, row.h, row.w1, row.h1, row.w2, row.h2, row.lm, row.rd, row.a, row.t, row.q].map((value, index) => <td key={`input-number-${index}`} className="px-4 py-4 font-mono">{n(value)}</td>)}
            {[row.plan, row.upDown, row.verticalAxis, row.formulaDetail, row.acousticFireproofDetail, row.note].map((value, index) => <td key={`input-text-${index}`} className="max-w-[240px] px-4 py-4 text-coninco-muted">{value || "—"}</td>)}
            <td className="border-l-2 border-coninco-line px-4 py-4 font-mono">{n(row.contractorK)}</td><td className="px-4 py-4 font-mono">{n(row.calculatedK)}</td><td className="px-4 py-4"><CheckCell match={row.kMatch}/></td>
            <td className="border-l border-coninco-line px-4 py-4 font-mono">{n(row.contractorScn)}</td><td className="px-4 py-4 font-mono">{n(row.calculatedScn)}</td><td className="px-4 py-4"><CheckCell match={row.scnMatch}/></td>
            <td className="border-l border-coninco-line px-4 py-4 font-mono">{n(row.contractorScc)}</td><td className="px-4 py-4 font-mono">{n(row.calculatedScc)}</td><td className="px-4 py-4"><CheckCell match={row.sccMatch}/></td>
            {hasSeparateFireproofAndAcoustic && <><td className="border-l border-coninco-line px-4 py-4 font-mono">{n(row.contractorSta)}</td><td className="px-4 py-4 font-mono">{n(row.calculatedSta)}</td><td className="px-4 py-4"><CheckCell match={row.staMatch}/></td></>}
          </tr>)}</tbody>
        </table>
        {!rows.length && <div className="px-6 py-14 text-center text-sm text-coninco-muted">No comparison rows found.</div>}
      </div>
    </div>
    <div className="flex justify-end"><DownloadButton entry={entry} setToast={setToast}/></div>
  </div>;
}

/** Sort audit targets in the same order as the workbook outputs. */
const formulaTargetOrder = (target: string) => ({ K: 0, S: 1, S_fireproof: 2, S_inner: 2 }[target] ?? 9);
/** Convert internal audit target names into user-facing labels. */
const formulaTargetLabel = (target: string) => ({ K: "K", S: "S / SCN", S_fireproof: "SCC", S_inner: "STA / SCC-TA" }[target] ?? target.toUpperCase());

/** Add fallback audit rows when an older preview lacks formula metadata. */
function mismatchedFormulaAudits(row: ComparisonRow): FormulaAudit[] {
  const audits = new Map(row.formulaAudits.map((audit) => [audit.target, audit]));
  /** Preserve a visible mismatch even when only numeric legacy data exists. */
  const fallback = (target: string, label: string, originalValue: number, calculatedValue: number, detail: string): FormulaAudit => ({
    target,
    label,
    originalFormula: `Source value = ${n(originalValue)}`,
    detailFormula: detail,
    detailBlock: detail || "Formula audit metadata is unavailable. Run this audit again to load the original formula.",
    originalValue,
    originalFormulaValue: null,
    calculatedValue,
    difference: originalValue - calculatedValue,
    matches: false,
    error: "Legacy preview does not contain the original formula metadata",
    correction: "",
  });
  const specs: Array<[string, string, number, number, string]> = [
    ["K", "K", row.contractorK, row.calculatedK, `K = PLAN + UP-DOWN + VERTICAL AXIS = ${n(row.calculatedK)}`],
    ["S", "S / SCN", row.contractorScn, row.calculatedScn, row.formulaDetail],
    ["S_fireproof", "SCC", row.contractorScc, row.calculatedScc, row.acousticFireproofDetail || row.formulaDetail],
    ["S_inner", "STA / SCC-TA", row.contractorSta, row.calculatedSta, row.acousticFireproofDetail || row.formulaDetail],
  ];
  for (const [target, label, originalValue, calculatedValue, detail] of specs) {
    if (row.sharedQuantityTarget && target !== row.sharedQuantityTarget && target !== "K" && target !== "S") continue;
    const match = target === "K" ? row.kMatch : target === "S" ? row.scnMatch : target === "S_fireproof" ? row.sccMatch : row.staMatch;
    if (!match && !audits.has(target)) audits.set(target, fallback(target, label, originalValue, calculatedValue, detail));
  }
  return [...audits.values()]
    .filter((audit) => !audit.matches)
    .sort((left, right) => formulaTargetOrder(left.target) - formulaTargetOrder(right.target));
}

/** Render the formula-only review shown after a numeric mismatch. */
function FormulaErrorTable({ items }: { items: Array<{ row: ComparisonRow; audit: FormulaAudit }> }) {
  return <section className={`overflow-hidden rounded-[24px] border bg-white shadow-soft ${items.length ? "border-red-200" : "border-emerald-200"}`}>
    <div className={`flex items-center justify-between border-b px-6 py-5 ${items.length ? "border-red-100 bg-red-50/70" : "border-emerald-100 bg-emerald-50/70"}`}>
      <div className="flex items-center gap-3"><span className={`rounded-xl p-2.5 ${items.length ? "bg-red-100 text-coninco-red" : "bg-emerald-100 text-emerald-700"}`}><Zap size={17}/></span><div><div className="text-sm font-black">Formula error review</div><p className="mt-1 text-xs text-coninco-muted">Formula structure is checked only when the saved result differs.</p></div></div>
      <span className={`rounded-full px-3 py-1.5 text-[10px] font-black ${items.length ? "bg-red-100 text-red-700" : "bg-emerald-100 text-emerald-700"}`}>{items.length ? `${items.length} FORMULA CHECKS` : "NO DIFFERENCES"}</span>
    </div>
    {items.length ? <div className="overflow-auto"><table className="table-grid w-full min-w-[3100px] border-collapse text-left text-[11px]"><thead className="bg-[#fbfbfc] text-[10px] font-black uppercase tracking-wider text-coninco-muted"><tr><th className="px-4 py-3">ROW</th><th className="min-w-[320px] px-4 py-3">PRODUCT NAME / TYPE</th><th className="px-4 py-3">CHECK</th><th className="px-4 py-3">VERSION</th><th className="min-w-[680px] px-4 py-3">FORMULA</th><th className="min-w-[1600px] px-4 py-3">AST FORMULA DIFFERENCE</th><th className="px-4 py-3 text-right">VALUE</th><th className="px-4 py-3 text-right">DIFFERENCE</th></tr></thead><tbody>{items.map(({ row, audit }) => <FormulaErrorPair key={`${row.sourceRow}-${audit.target}`} row={row} audit={audit}/>)}</tbody></table></div> : <div className="flex items-center gap-3 px-6 py-6 text-sm font-semibold text-emerald-700"><Check size={17}/> All calculated quantities match the workbook values. Formula text was not checked for matching results.</div>}
  </section>;
}

/** Render original and standard versions of one failed target side by side. */
function FormulaErrorPair({ row, audit }: { row: ComparisonRow; audit: FormulaAudit }) {
  const label = audit.label || formulaTargetLabel(audit.target);
  const sourceComparison = audit.sourceFormulaComparison ?? audit.formulaComparison;
  const sourceStatus = new Map((sourceComparison?.blocks ?? []).map((block) => [block.label.toUpperCase(), block.status]));
  const blocks = [...(audit.formulaComparison?.blocks ?? sourceComparison?.blocks ?? [])]
    .map((block) => ({
      ...block,
      status: block.status === "unmatched"
        || (sourceStatus.get(block.label.toUpperCase()) === "unmatched" && (block.astDifferences ?? []).length > 0)
        ? "unmatched" as const
        : "matched" as const,
      color: block.status === "unmatched"
        || (sourceStatus.get(block.label.toUpperCase()) === "unmatched" && (block.astDifferences ?? []).length > 0)
        ? "#B91C1C"
        : "#1565C0",
    }))
    .sort((left, right) => left.label.localeCompare(right.label, undefined, { numeric: true }));
  const originalFormulaValue = sourceComparison?.originalFormulaValue ?? audit.originalFormulaValue;
  const calculatedValue = n(audit.calculatedValue);
  const formulaDelta = originalFormulaValue == null ? null : Math.abs(originalFormulaValue - audit.calculatedValue);
  const savedDelta = Math.abs(audit.difference);
  const valuesExplainCache = formulaDelta !== null && formulaDelta <= 0.0005 && savedDelta > 0.0005;
  const originalValue = <div className="space-y-1 text-right">
    <div>SAVED CELL = {n(audit.originalValue)}</div>
    <div>ORIGINAL FORMULA = {originalFormulaValue == null ? "—" : n(originalFormulaValue)}</div>
  </div>;
  const standardValue = <div className="space-y-1 text-right">STANDARD FORMULA = {calculatedValue}</div>;
  const difference = <div className="space-y-1 text-right">
    <div>SAVED vs STANDARD Δ={audit.error ? "ERROR" : n(savedDelta)}</div>
    <div>ORIGINAL FORMULA vs STANDARD Δ={formulaDelta == null ? "—" : n(formulaDelta)}</div>
    {valuesExplainCache && <div className="max-w-[220px] whitespace-normal text-[9px] font-bold text-coninco-muted">Formula results match; the saved Excel value is different.</div>}
  </div>;
  const inputReferenceEntries = Object.entries(row.inputReferences);
  return <>
    <tr className="bg-red-50/45">
      <td rowSpan={2} className="align-top px-4 py-4 font-black text-coninco-red">{row.sourceRow}</td>
      <td rowSpan={2} className="max-w-[320px] align-top break-words px-4 py-4">
        <div className="font-bold leading-5 text-coninco-ink">{row.description || "—"}</div>
        <div className="mt-2 text-[10px] font-black uppercase tracking-wider text-coninco-muted">{row.productType || "—"}</div>
      </td>
      <td rowSpan={2} className="align-top px-4 py-4"><span className="rounded-full bg-red-100 px-2.5 py-1 text-[10px] font-black text-red-700">{label}</span></td>
      <td className="px-4 py-3"><span className="rounded-full bg-[#eef4ff] px-2.5 py-1 text-[10px] font-black text-[#1565C0]">ORIGINAL</span></td>
      <td className="px-4 py-3"><OriginalFormulaText value={sourceComparison?.originalFormula || audit.originalFormula || "Original formula not available"} comparison={sourceComparison} /></td>
      <td rowSpan={2} className="align-top px-4 py-4"><AstDifferenceList blocks={blocks} originalFormula={sourceComparison?.originalFormula || audit.originalFormula || ""} /></td>
      <td className="px-4 py-3 text-right font-mono text-[10px] font-bold text-coninco-ink">{originalValue}</td>
      <td rowSpan={2} className="align-top px-4 py-4 text-right font-mono text-[10px] font-black text-red-700">{difference}</td>
    </tr>
    <tr className="bg-white">
      <td className="px-4 py-3"><span className="rounded-full bg-[#f1f2f3] px-2.5 py-1 text-[10px] font-black text-coninco-muted">STANDARD</span></td>
      <td className="px-4 py-3"><div className="min-w-max space-y-3">
        <div className="flex flex-wrap items-center gap-1.5 text-[10px] font-black uppercase tracking-wide text-coninco-muted">
          <span className="mr-1 text-coninco-red">INPUT CELLS</span>
          {inputReferenceEntries.map(([name, address]) => <span key={`${name}-${address}`} className="rounded-md bg-[#f1f2f3] px-1.5 py-1 font-mono normal-case tracking-normal text-coninco-ink">{name}: {address}</span>)}
        </div>
        <FormulaBlockList
          blocks={blocks}
          fallback={audit.referenceFormulaBlock
            || (audit.referenceFormula ? `${audit.label || label}: ${audit.referenceFormula}` : "")
            || audit.detailBlock
            || audit.detailFormula
            || row.acousticFireproofDetail
            || row.formulaDetail
            || audit.error}
        />
      </div></td>
      <td className="px-4 py-3 text-right font-mono text-[10px] font-bold text-coninco-ink">{standardValue}</td>
    </tr>
  </>;
}

type FormulaComparisonBlock = NonNullable<FormulaAudit["formulaComparison"]>["blocks"][number];

/** Color original formula blocks according to their AST match status. */
function OriginalFormulaText({ value, comparison }: { value: string; comparison?: FormulaAudit["sourceFormulaComparison"] }) {
  const spans = (comparison?.blocks ?? [])
    .flatMap((block) => (block.originalSpans ?? []).map((span) => ({
      ...span,
      color: block.status === "matched" ? "#1565C0" : "#B91C1C",
    })))
    .sort((left, right) => left.start - right.start || right.end - left.end);
  const baseColor = spans.length || comparison?.status === "equivalent" ? "#1565C0" : "#B91C1C";
  const parts: React.ReactNode[] = [];
  let cursor = 0;
  spans.forEach((span, index) => {
    const start = Math.max(cursor, span.start);
    const end = Math.max(start, span.end);
    if (start > cursor) parts.push(<span key={`original-gap-${index}`} style={{ color: baseColor }}>{value.slice(cursor, start)}</span>);
    if (end > start) parts.push(<span key={`original-block-${index}`} style={{ color: span.color, fontWeight: 900 }}>{value.slice(start, end)}</span>);
    cursor = Math.max(cursor, end);
  });
  if (cursor < value.length) parts.push(<span key="original-tail" style={{ color: baseColor }}>{value.slice(cursor)}</span>);
  return <div className="min-w-max space-y-1.5 font-mono text-[11px] leading-5">
    <div className="whitespace-nowrap rounded-lg border-l-4 bg-white px-3 py-2" style={{ borderLeftColor: baseColor }}>
      {spans.length ? parts : <span style={{ color: baseColor, fontWeight: 900 }}>{value}</span>}
    </div>
  </div>;
}

/** Render standard formula blocks with blue/red block-level status. */
function FormulaBlockList({ blocks, fallback }: { blocks: FormulaComparisonBlock[]; fallback?: string }) {
  if (!blocks.length) {
    return <div className="min-w-max whitespace-nowrap font-mono text-[11px] font-black" style={{ color: "#B91C1C" }}>{formatFormulaForDisplay(fallback || "Formula not available")}</div>;
  }
  return <div className="min-w-max space-y-1.5">
    {blocks.map((block, index) => {
      const color = block.status === "matched" ? "#1565C0" : "#B91C1C";
      const label = block.originalLabel || block.label;
      const expression = block.originalExpression || block.expression;
      return <div key={`${label}-${index}`} className="whitespace-nowrap rounded-lg border-l-4 bg-white px-3 py-2 font-mono text-[11px] leading-5" style={{ borderLeftColor: color, color }}>
        <span className="font-black">{label} = </span>
        <span className="font-black">{formatFormulaForDisplay(expression || block.error || "Formula block not available")}</span>
        {block.error && <span className="font-black"> · {block.error}</span>}
      </div>;
    })}
  </div>;
}

/** Highlight AST mismatch fragments inside one displayed formula string. */
function FormulaWithAstHighlights({ blockLabel, value, highlights, blockHighlights = [] }: { blockLabel: string; value: string; highlights: string[]; blockHighlights?: string[] }) {
  const text = formatFormulaForDisplay(value || highlights.join(" · ") || "Formula text missing");
  const isIntermediateBlock = /^S_V\d+$/i.test(blockLabel);
  /** Locate normalized highlight text while tolerating display whitespace. */
  const findRanges = (values: string[], kind: "light" | "dark") => values.flatMap((highlight) => {
    const target = formatFormulaForDisplay(highlight);
    if (!target) return [];
    const pattern = target.split(/\s+/).filter(Boolean).map((part) => part.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("\\s*");
    const match = new RegExp(pattern, "i").exec(text);
    return match ? [{ start: match.index, end: match.index + match[0].length, kind }] : [];
  });
  const lightRanges = findRanges(blockHighlights, "light");
  const darkRanges = findRanges(highlights, "dark");
  const boundaries = Array.from(new Set([0, text.length, ...lightRanges.flatMap((range) => [range.start, range.end]), ...darkRanges.flatMap((range) => [range.start, range.end])])).sort((left, right) => left - right);
  const parts: React.ReactNode[] = [];
  const baseClass = isIntermediateBlock ? "text-red-500" : "text-coninco-ink";
  for (let index = 0; index < boundaries.length - 1; index += 1) {
    const start = boundaries[index];
    const end = boundaries[index + 1];
    if (end <= start) continue;
    const dark = darkRanges.some((range) => range.start <= start && range.end >= end);
    const light = !dark && lightRanges.some((range) => range.start <= start && range.end >= end);
    const className = dark
      ? "font-black text-red-900 underline decoration-2 underline-offset-2"
      : light ? "rounded bg-red-50/80 text-red-500" : baseClass;
    parts.push(<span key={`ast-part-${index}`} className={className}>{text.slice(start, end)}</span>);
  }
  return <span className={`whitespace-nowrap font-mono text-[10px] leading-5 ${isIntermediateBlock ? "rounded bg-red-50/80 px-1" : "text-coninco-ink"}`}>{parts.length ? parts : text}</span>;
}

/** Show the exact AST branches that differ between original and standard text. */
function AstDifferenceList({ blocks, originalFormula }: { blocks: FormulaComparisonBlock[]; originalFormula: string }) {
  const differences = blocks
    .filter((block) => (block.astDifferences ?? []).length)
    .map((block) => {
      const blockDifferences = block.astDifferences ?? [];
      const originalFallback = blockDifferences.map((difference) => difference.original).filter(Boolean).join(" · ");
      const standardFallback = blockDifferences.map((difference) => difference.standard).filter(Boolean).join(" · ");
      return {
        block,
        key: `${block.originalLabel || block.label}-ast`,
        label: block.originalLabel || block.label,
        originalFormula: block.astOriginalExpression || originalFormula || originalFallback,
        standardFormula: block.astStandardExpression || block.originalExpression || block.expression || standardFallback,
        differences: blockDifferences,
      };
    });
  if (!differences.length) return null;
  return <div className="min-w-[700px] space-y-2">
    {differences.map(({ block, key, label, originalFormula: blockOriginalFormula, standardFormula, differences: blockDifferences }, index) => <div key={key} className="rounded-lg border-l-4 border-red-700 bg-red-50/70 px-3 py-2 leading-5">
      <div className="mb-1 font-black uppercase tracking-wide text-red-900">{index + 1}. {label} · AST mismatch</div>
      <div className="flex min-w-max items-start gap-2"><span className={`w-32 shrink-0 text-right font-black ${/^S_V\d+$/i.test(label) ? "text-red-500" : "text-red-900"}`}>Original {label}:</span><FormulaWithAstHighlights blockLabel={label} blockHighlights={block.astBlockOriginalExpressions ?? []} value={blockOriginalFormula} highlights={blockDifferences.map((difference) => difference.original)} /></div>
      <div className="flex min-w-max items-start gap-2"><span className={`w-32 shrink-0 text-right font-black ${/^S_V\d+$/i.test(label) ? "text-red-500" : "text-red-900"}`}>{label}:</span><FormulaWithAstHighlights blockLabel={label} blockHighlights={block.astBlockStandardExpressions ?? []} value={standardFormula} highlights={blockDifferences.map((difference) => difference.standard)} /></div>
    </div>)}
  </div>;
}

/** Render one summary metric card. */
function Metric({ label, value, icon, danger = false }: { label: string; value: string; icon: React.ReactNode; danger?: boolean }) { return <div className="rounded-2xl border border-coninco-line bg-white p-5 shadow-soft"><div className="flex items-center justify-between text-coninco-muted"><span className="text-[10px] font-black uppercase tracking-wider">{label}</span><span className={danger ? "text-coninco-red" : "text-coninco-red"}>{icon}</span></div><div className={`mt-3 text-2xl font-black ${danger ? "text-coninco-red" : "text-coninco-ink"}`}>{value}</div></div>; }
/** Render a compact section/category badge. */
function Tag({ value }: { value: string }) { return <span className="rounded-lg bg-coninco-soft px-2.5 py-1 text-[10px] font-black capitalize text-coninco-red">{value}</span>; }
/** Render a compact MATCH or DIFF status badge. */
function CheckCell({ match }: { match: boolean }) { return <span className={`inline-flex items-center gap-1 rounded-full px-2.5 py-1 text-[10px] font-black ${match ? "bg-emerald-50 text-emerald-700" : "bg-red-50 text-red-700"}`}>{match ? <Check size={12}/> : <X size={12}/>} {match ? "MATCH" : "DIFF"}</span>; }

/** Offer submitted and processed workbook download actions. */
function DownloadButton({ entry, setToast }: { entry: ArchiveEntry; setToast: (value: string) => void }) {
  const [open, setOpen] = useState(false);
  /** Download the requested archive version and report the resulting path. */
  const download = async (kind: "submitted" | "processed") => {
    setOpen(false);
    try {
      const savedPath = await downloadArchive(entry.id, kind);
      setToast(savedPath ? `File saved to ${savedPath}` : "Download cancelled");
    } catch (error) {
      setToast(error instanceof Error ? error.message : "Could not download file");
    }
  };
  return <><button onClick={() => setOpen(true)} title="Download workbook" className="flex items-center gap-2 rounded-xl border border-coninco-line px-3 py-2 text-[10px] font-black text-coninco-muted hover:border-coninco-red hover:text-coninco-red"><Download size={15}/> DOWNLOAD</button>{open && <div className="fixed inset-0 z-50 flex items-center justify-center bg-coninco-ink/45 p-6 backdrop-blur-sm" onClick={() => setOpen(false)}><div role="dialog" aria-modal="true" className="w-full max-w-lg rounded-[24px] border border-coninco-line bg-white p-7 text-left shadow-2xl animate-enter" onClick={(event) => event.stopPropagation()}><div className="flex items-start justify-between gap-4"><div><div className="text-[11px] font-black uppercase tracking-[.18em] text-coninco-red">Download workbook</div><h2 className="mt-2 text-xl font-black text-coninco-ink">Choose version</h2><p className="mt-2 text-sm text-coninco-muted">Select the file version you want to save.</p></div><button onClick={() => setOpen(false)} className="rounded-xl border border-coninco-line p-2 text-coninco-muted hover:border-coninco-red hover:text-coninco-red"><X size={17}/></button></div><div className="mt-6 grid gap-3"><button onClick={() => void download("submitted")} className="rounded-2xl border border-coninco-line px-4 py-4 text-left text-sm font-bold hover:border-coninco-red hover:bg-coninco-soft">Original submitted workbook<span className="mt-1 block text-xs font-normal text-coninco-muted">Download the file exactly as submitted.</span></button><button onClick={() => void download("processed")} className="rounded-2xl border border-coninco-line px-4 py-4 text-left text-sm font-bold hover:border-coninco-red hover:bg-coninco-soft">Processed calculated workbook<span className="mt-1 block text-xs font-normal text-coninco-muted">Download the file with audit results.</span></button></div><div className="mt-6 flex justify-end"><button onClick={() => setOpen(false)} className="rounded-xl border border-coninco-line px-4 py-3 text-xs font-black text-coninco-muted hover:border-coninco-red">CANCEL</button></div></div></div>}</>;
}

export default App;
