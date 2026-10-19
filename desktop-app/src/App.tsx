import { useEffect, useMemo, useState } from "react";
import {
  Archive, ArrowLeft, ArrowUpRight, BarChart3, Check, ChevronDown, CircleHelp, CloudUpload,
  Download, FileSpreadsheet, FolderOpen, History, LayoutDashboard, Loader2, MoreHorizontal,
  Play, RefreshCw, Search, Settings2, ShieldCheck, Trash2, X, Zap,
} from "lucide-react";
import {
  chooseWorkbook, deleteArchive, downloadArchive, listArchive, listSheets, loadComparison,
  openArchiveFile, runAudit,
} from "./lib/engine";
import type { ArchiveEntry, ComparisonRow, Page } from "./types";

const n = (value: number) => value.toFixed(3);

function App() {
  const [page, setPage] = useState<Page>("process");
  const [path, setPath] = useState("");
  const [sheets, setSheets] = useState<string[]>([]);
  const [sheet, setSheet] = useState("");
  const [archive, setArchive] = useState<ArchiveEntry[]>([]);
  const [selected, setSelected] = useState<ArchiveEntry | null>(null);
  const [rows, setRows] = useState<ComparisonRow[]>([]);
  const [search, setSearch] = useState("");
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState(0);
  const [message, setMessage] = useState("Ready for a local audit");
  const [toast, setToast] = useState<string | null>(null);
  const [pendingDelete, setPendingDelete] = useState<ArchiveEntry | null>(null);

  const refreshArchive = async () => setArchive(await listArchive());
  useEffect(() => { void refreshArchive(); }, []);

  const visibleArchive = useMemo(() => archive.filter((entry) =>
    `${entry.submittedName} ${entry.processedName} ${entry.sheet}`.toLowerCase().includes(search.toLowerCase())), [archive, search]);

  const selectWorkbook = async () => {
    const chosen = await chooseWorkbook();
    if (!chosen) return;
    setPath(chosen); setSheets([]); setSheet(""); setMessage("Reading workbook sheets…");
    try { const available = await listSheets(chosen); setSheets(available); setSheet(available[0] ?? ""); setMessage(`${available.length} worksheets ready`); }
    catch (error) { setMessage(error instanceof Error ? error.message : "Could not read workbook"); }
  };

  const audit = async () => {
    if (!path || !sheet) return;
    setBusy(true); setProgress(12); setMessage("Copying submitted workbook into the local archive…");
    const timer = window.setInterval(() => setProgress((value) => Math.min(value + 11, 88)), 450);
    try {
      const result = await runAudit(path, sheet); setProgress(100); setMessage("Audit complete — processed workbook is ready");
      setToast(`${result.rows} rows processed with ${result.mismatches} mismatched cells`); await refreshArchive();
    } catch (error) { setMessage(error instanceof Error ? error.message : "Audit failed"); }
    finally { window.clearInterval(timer); window.setTimeout(() => setProgress(0), 800); setBusy(false); }
  };

  const viewComparison = async (entry: ArchiveEntry) => {
    setSelected(entry); setBusy(true); setMessage("Loading comparison data…");
    try { setRows(await loadComparison(entry.id)); setPage("comparison"); }
    catch (error) { setMessage(error instanceof Error ? error.message : "Could not load comparison"); }
    finally { setBusy(false); }
  };

  const removeEntry = async (entry: ArchiveEntry) => {
    try {
      await deleteArchive(entry.id);
      await refreshArchive();
      setToast(`Deleted ${entry.submittedName}`);
    } catch (error) {
      setToast(error instanceof Error ? error.message : "Could not delete archive entry");
    }
  };

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
        <NavButton active={page === "archive"} icon={<Archive size={18}/>} label="Archive" onClick={() => { void refreshArchive(); setPage("archive"); }} badge={archive.length} />
      </nav>
      <div className="mt-auto space-y-4 border-t border-coninco-line px-5 py-5">
        <div className="rounded-2xl bg-coninco-soft p-4"><div className="flex items-center gap-2 text-xs font-bold text-coninco-red"><ShieldCheck size={15}/> Offline secure mode</div><p className="mt-2 text-[11px] leading-5 text-coninco-muted">Files stay inside this computer. No cloud upload.</p></div>
        <div className="flex items-center justify-between text-[11px] text-coninco-muted"><span>v0.2.0</span><CircleHelp size={15}/></div>
      </div>
    </aside>
    <main className="ml-[248px] min-h-screen">
      <header className="flex h-[76px] items-center justify-between border-b border-coninco-line bg-white/90 px-10 backdrop-blur">
        <div><div className="text-[11px] font-black uppercase tracking-[.2em] text-coninco-red">Local quantity verification</div><div className="mt-1 text-sm text-coninco-muted">{page === "process" ? "Run a controlled workbook audit" : page === "archive" ? "Manage stored workbooks" : "Inspect source and engine values"}</div></div>
        <div className="flex items-center gap-3"><div className="flex items-center gap-2 rounded-full border border-emerald-200 bg-emerald-50 px-3 py-2 text-[11px] font-bold text-emerald-700"><span className="h-2 w-2 animate-pulse rounded-full bg-emerald-500"/> OFFLINE MODE</div><button className="rounded-xl border border-coninco-line p-2.5 text-coninco-muted hover:border-coninco-red hover:text-coninco-red"><Settings2 size={17}/></button></div>
      </header>
      <div className="p-10">{page === "process" && <ProcessPage path={path} sheet={sheet} sheets={sheets} busy={busy} progress={progress} message={message} selectWorkbook={selectWorkbook} setSheet={setSheet} audit={audit} setPage={setPage}/>} {page === "archive" && <ArchivePage entries={visibleArchive} search={search} setSearch={setSearch} viewComparison={viewComparison} removeEntry={setPendingDelete} setToast={setToast}/>} {page === "comparison" && selected && <ComparisonPage entry={selected} rows={rows} back={() => setPage("archive")} setToast={setToast}/>}</div>
    </main>
    {toast && <div className="fixed bottom-7 right-8 z-50 flex items-center gap-3 rounded-2xl bg-coninco-ink px-5 py-4 text-sm text-white shadow-2xl animate-enter"><Check size={17} className="text-emerald-300"/>{toast}<button onClick={() => setToast(null)}><X size={15}/></button></div>}
    {pendingDelete && <div className="fixed inset-0 z-50 flex items-center justify-center bg-coninco-ink/45 p-6 backdrop-blur-sm"><div role="dialog" aria-modal="true" className="w-full max-w-lg rounded-[24px] border border-coninco-line bg-white p-7 shadow-2xl animate-enter"><div className="flex min-w-0 items-start gap-4"><div className="shrink-0 rounded-2xl bg-red-50 p-3 text-coninco-red"><Trash2 size={22}/></div><div className="min-w-0 flex-1"><h2 className="text-lg font-black text-coninco-ink">Delete archived audit?</h2><p className="mt-2 text-sm leading-6 text-coninco-muted">This will permanently remove the submitted workbook, processed workbook, and comparison data from this application.</p><p className="mt-3 min-w-0 max-w-full truncate rounded-xl bg-[#f7f7f7] px-3 py-2 text-xs font-bold text-coninco-ink">{pendingDelete.submittedName}</p></div></div><div className="mt-7 flex justify-end gap-3"><button onClick={() => setPendingDelete(null)} className="rounded-xl border border-coninco-line px-4 py-3 text-xs font-black text-coninco-muted hover:border-coninco-red hover:text-coninco-red">CANCEL</button><button onClick={() => void confirmDelete()} className="flex items-center gap-2 rounded-xl bg-coninco-red px-4 py-3 text-xs font-black text-white shadow-red hover:bg-coninco-dark"><Trash2 size={15}/> DELETE</button></div></div></div>}
  </div>;
}

function NavButton({ active, icon, label, badge, onClick }: { active: boolean; icon: React.ReactNode; label: string; badge?: number; onClick: () => void }) {
  return <button onClick={onClick} className={`group flex w-full items-center gap-3 rounded-2xl px-4 py-3 text-left text-sm font-bold transition ${active ? "bg-coninco-red text-white shadow-red" : "text-coninco-muted hover:bg-coninco-soft hover:text-coninco-red"}`}><span className={active ? "text-white" : "text-coninco-muted group-hover:text-coninco-red"}>{icon}</span>{label}{badge !== undefined && <span className={`ml-auto rounded-full px-2 py-0.5 text-[10px] ${active ? "bg-white/20" : "bg-[#f1f2f3]"}`}>{badge}</span>}</button>;
}

function ProcessPage({ path, sheet, sheets, busy, progress, message, selectWorkbook, setSheet, audit, setPage }: { path: string; sheet: string; sheets: string[]; busy: boolean; progress: number; message: string; selectWorkbook: () => void; setSheet: (value: string) => void; audit: () => void; setPage: (page: Page) => void }) {
  return <div className="animate-enter space-y-8">
    <div className="relative overflow-hidden rounded-[28px] bg-coninco-ink px-9 py-8 text-white shadow-soft"><div className="absolute -right-16 -top-24 h-72 w-72 rounded-full bg-coninco-red/25 blur-3xl"/><div className="absolute bottom-[-100px] right-[30%] h-52 w-52 rounded-full bg-white/10 blur-3xl"/><div className="relative flex items-end justify-between"><div><div className="mb-4 inline-flex items-center gap-2 rounded-full border border-white/15 bg-white/10 px-3 py-1.5 text-[10px] font-bold uppercase tracking-[.18em] text-white/75"><Zap size={13} className="text-coninco-red"/> Precision audit workspace</div><h1 className="max-w-xl text-4xl font-black leading-tight tracking-tight">Turn contractor workbooks into <span className="text-coninco-red">clear evidence.</span></h1><p className="mt-3 max-w-lg text-sm leading-6 text-white/60">Read, calculate and compare locally with the original workbook always preserved.</p></div><div className="hidden w-48 animate-float rounded-3xl border border-white/10 bg-white/5 p-5 md:block"><BarChart3 className="mb-4 text-coninco-red" size={30}/><div className="text-2xl font-black">100%</div><div className="mt-1 text-xs text-white/50">local processing</div></div></div></div>
    <div className="grid grid-cols-[1.3fr_.7fr] gap-6">
      <section className="rounded-[24px] border border-coninco-line bg-white p-7 shadow-soft"><SectionHeading number="01" title="Select contractor workbook" subtitle="The submitted copy is stored inside the application archive."/><button onClick={selectWorkbook} className="group mt-6 flex w-full items-center gap-5 rounded-2xl border-2 border-dashed border-[#dfe1e3] bg-[#fbfbfc] p-6 text-left transition hover:border-coninco-red hover:bg-coninco-soft"><span className="flex h-14 w-14 items-center justify-center rounded-2xl bg-coninco-soft text-coninco-red transition group-hover:scale-105"><CloudUpload size={26}/></span><span className="min-w-0 flex-1"><span className="block text-sm font-black text-coninco-ink">{path ? path.split(/[\\/]/).pop() : "Drop an Excel workbook here"}</span><span className="mt-1 block truncate text-xs text-coninco-muted">{path || "Supports .xlsx and legacy .xls files"}</span></span><span className="rounded-xl bg-coninco-red px-4 py-2.5 text-xs font-black text-white shadow-red">BROWSE</span></button></section>
      <section className="rounded-[24px] border border-coninco-line bg-white p-7 shadow-soft"><SectionHeading number="02" title="Choose worksheet" subtitle="Select the source sheet to audit."/><div className="relative mt-6"><select value={sheet} onChange={(event) => setSheet(event.target.value)} disabled={!sheets.length || busy} className="w-full appearance-none rounded-2xl border border-coninco-line bg-[#fbfbfc] px-4 py-3.5 text-sm font-bold text-coninco-ink outline-none transition focus:border-coninco-red focus:ring-4 focus:ring-coninco-red/10"><option value="">Select a worksheet</option>{sheets.map((item) => <option key={item}>{item}</option>)}</select><ChevronDown className="pointer-events-none absolute right-4 top-4 text-coninco-muted" size={18}/></div><div className="mt-4 flex items-center gap-2 text-[11px] text-coninco-muted"><FileSpreadsheet size={14} className="text-coninco-red"/>{sheets.length ? `${sheets.length} worksheets detected` : "Workbook sheets will appear here"}</div></section>
    </div>
    <section className="relative overflow-hidden rounded-[24px] border border-coninco-red/20 bg-coninco-soft p-7 shadow-soft"><div className="absolute -right-10 -top-16 h-48 w-48 rounded-full bg-coninco-red/10 blur-3xl"/><div className="relative flex items-center justify-between"><SectionHeading number="03" title="Run audit" subtitle="Calculate K, S and S-inner, compare source values, then store the result." accent/><button onClick={audit} disabled={!path || !sheet || busy} className="flex items-center gap-3 rounded-2xl bg-coninco-red px-7 py-4 text-sm font-black text-white shadow-red transition hover:bg-coninco-dark disabled:cursor-not-allowed disabled:opacity-40">{busy ? <Loader2 className="animate-spin" size={18}/> : <Play size={18} fill="currentColor"/>}{busy ? "PROCESSING…" : "RUN AUDIT"}</button></div>{busy && <div className="relative mt-6"><div className="mb-2 flex justify-between text-[11px] font-bold text-coninco-muted"><span>{message}</span><span>{progress}%</span></div><div className="h-2 overflow-hidden rounded-full bg-white"><div className="h-full rounded-full bg-coninco-red transition-all duration-500" style={{width: `${progress}%`}}/></div></div>}</section>
    <div className="flex items-center justify-between rounded-2xl border border-coninco-line bg-white px-6 py-5"><div className="flex items-center gap-3"><History size={18} className="text-coninco-red"/><div><div className="text-sm font-black">Need to review a previous audit?</div><div className="mt-1 text-xs text-coninco-muted">Open the local archive to compare submitted and processed files.</div></div></div><button onClick={() => setPage("archive")} className="flex items-center gap-2 text-xs font-black text-coninco-red hover:text-coninco-dark">OPEN ARCHIVE <ArrowUpRight size={15}/></button></div>
  </div>;
}

function SectionHeading({ number, title, subtitle, accent = false }: { number: string; title: string; subtitle: string; accent?: boolean }) { return <div className="flex items-start gap-3"><span className={`flex h-9 w-9 shrink-0 items-center justify-center rounded-xl text-xs font-black ${accent ? "bg-white text-coninco-red" : "bg-coninco-soft text-coninco-red"}`}>{number}</span><div><h2 className={`text-base font-black ${accent ? "text-coninco-ink" : "text-coninco-ink"}`}>{title}</h2><p className="mt-1 text-xs leading-5 text-coninco-muted">{subtitle}</p></div></div>; }

function ArchivePage({ entries, search, setSearch, viewComparison, removeEntry, setToast }: { entries: ArchiveEntry[]; search: string; setSearch: (value: string) => void; viewComparison: (entry: ArchiveEntry) => void; removeEntry: (entry: ArchiveEntry) => void; setToast: (value: string) => void }) {
  return <div className="animate-enter space-y-7"><div className="flex items-end justify-between"><div><div className="text-[11px] font-black uppercase tracking-[.2em] text-coninco-red">Workspace archive</div><h1 className="mt-2 text-3xl font-black tracking-tight">Stored workbooks</h1><p className="mt-2 text-sm text-coninco-muted">Submitted originals and processed results stay paired on this computer.</p></div><div className="rounded-2xl bg-coninco-ink px-5 py-3 text-right text-white"><div className="text-2xl font-black">{entries.length}</div><div className="text-[10px] font-bold uppercase tracking-wider text-white/55">saved audits</div></div></div><div className="flex items-center gap-3 rounded-2xl border border-coninco-line bg-white p-3 shadow-soft"><Search size={18} className="ml-2 text-coninco-muted"/><input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Search by workbook or sheet…" className="flex-1 bg-transparent px-2 py-2 text-sm outline-none"/><button onClick={() => setToast("Archive is stored locally") } className="rounded-xl border border-coninco-line px-4 py-2 text-xs font-black text-coninco-muted hover:border-coninco-red hover:text-coninco-red"><RefreshCw size={14}/></button></div><div className="overflow-hidden rounded-[24px] border border-coninco-line bg-white shadow-soft"><div className="border-b border-coninco-line px-6 py-5"><div className="flex items-center gap-2 text-sm font-black"><Archive size={17} className="text-coninco-red"/> Audit history</div></div><div className="overflow-x-auto"><table className="table-grid w-full min-w-[1320px] border-collapse text-left"><thead className="bg-[#fbfbfc] text-[10px] font-black uppercase tracking-wider text-coninco-muted"><tr><th className="px-5 py-4">Submitted workbook</th><th className="px-5 py-4">Processed result</th><th className="px-5 py-4">Sheet</th><th className="px-5 py-4">Rows</th><th className="px-5 py-4">Status</th><th className="px-5 py-4 text-right">Actions</th></tr></thead><tbody>{entries.map((entry) => <tr key={entry.id} className="group hover:bg-[#fffafa]"><td className="px-5 py-5"><div className="flex items-center gap-3"><span className="rounded-xl bg-coninco-soft p-2.5 text-coninco-red"><FileSpreadsheet size={18}/></span><div><div className="text-sm font-bold">{entry.submittedName}</div><div className="mt-1 text-[11px] text-coninco-muted">{entry.createdAt}</div></div></div></td><td className="px-5 py-5"><div className="text-sm font-bold text-coninco-ink">{entry.processedName}</div><div className="mt-1 text-[11px] text-coninco-muted">{entry.mismatches} mismatched cells</div></td><td className="px-5 py-5"><span className="rounded-lg bg-[#f1f2f3] px-2.5 py-1 text-xs font-black">{entry.sheet}</span></td><td className="px-5 py-5 text-sm font-black">{entry.rows}</td><td className="px-5 py-5"><span className={`rounded-full px-3 py-1.5 text-[10px] font-black ${entry.status === "COMPLETED" ? "bg-emerald-50 text-emerald-700" : "bg-red-50 text-red-700"}`}>{entry.status}</span></td><td className="whitespace-nowrap px-5 py-5"><div className="flex justify-end gap-2"><button onClick={() => viewComparison(entry)} className="rounded-xl bg-coninco-red px-3 py-2 text-[10px] font-black text-white hover:bg-coninco-dark">VIEW COMPARISON</button><DownloadButton entry={entry} setToast={setToast}/><button onClick={() => void removeEntry(entry)} title="Delete" className="rounded-xl border border-red-100 p-2 text-red-400 hover:bg-red-50"><Trash2 size={15}/></button></div></td></tr>)}</tbody></table>{!entries.length && <div className="px-6 py-14 text-center text-sm text-coninco-muted">No archived workbooks found.</div>}</div></div></div>;
}

function ComparisonPage({ entry, rows, back, setToast }: { entry: ArchiveEntry; rows: ComparisonRow[]; back: () => void; setToast: (value: string) => void }) {
  const inputHeaders = ["ROW", "ITEM NO", "DESCRIPTION", "UNIT", "SECTION", "PRODUCT", "W", "H", "W'", "H'", "W''", "H''", "L/M", "R/D", "a", "t", "Q", "PLAN", "UP-DOWN", "VERTICAL AXIS", "FORMULA DETAIL", "ACOUSTIC / FIREPROOF", "NOTE"];
  const outputHeaders = ["K SOURCE", "K ENGINE", "K CHECK", "S SOURCE", "S ENGINE", "S CHECK", "S INNER SOURCE", "S INNER ENGINE", "S INNER CHECK"];

  return <div className="animate-enter space-y-7">
    <div className="flex items-center gap-4"><button onClick={back} className="rounded-xl border border-coninco-line bg-white p-3 text-coninco-muted hover:border-coninco-red hover:text-coninco-red"><ArrowLeft size={18}/></button><div className="flex-1"><div className="text-[11px] font-black uppercase tracking-[.2em] text-coninco-red">Comparison detail · Sheet {entry.sheet}</div><h1 className="mt-2 text-3xl font-black tracking-tight">{entry.processedName}</h1><p className="mt-2 text-sm text-coninco-muted">One row contains the source inputs and all calculated audit outputs.</p></div><button onClick={() => void openArchiveFile(entry.id, "processed")} className="flex items-center gap-2 rounded-xl border border-coninco-line bg-white px-4 py-3 text-xs font-black text-coninco-red hover:bg-coninco-soft"><FolderOpen size={16}/> OPEN WORKBOOK</button></div>
    <div className="grid grid-cols-4 gap-4"><Metric label="Rows compared" value={String(entry.rows)} icon={<BarChart3 size={17}/>} /><Metric label="K / S checks" value={String(entry.rows * 3)} icon={<ShieldCheck size={17}/>} /><Metric label="Mismatches" value={String(entry.mismatches)} icon={<Zap size={17}/>} danger={entry.mismatches > 0}/><Metric label="Engine" value="Python" icon={<Settings2 size={17}/>} /></div>
    <div className="overflow-hidden rounded-[24px] border border-coninco-line bg-white shadow-soft">
      <div className="border-b border-coninco-line px-6 py-5"><div className="flex items-center gap-2 text-sm font-black"><FileSpreadsheet size={17} className="text-coninco-red"/> Complete comparison table</div><p className="mt-1 text-xs text-coninco-muted">Each row combines contractor inputs, source quantities, engine quantities, and match status. Scroll horizontally to inspect all fields.</p></div>
      <div className="overflow-auto">
        <table className="table-grid min-w-[3600px] border-collapse text-left text-[11px]">
          <thead className="bg-[#fbfbfc] text-[10px] font-black uppercase tracking-wider text-coninco-muted">
            <tr><th colSpan={inputHeaders.length} className="border-r-2 border-coninco-line bg-coninco-soft px-4 py-3 text-coninco-red">INPUT DATA · SOURCE ROW VALUES</th><th colSpan={outputHeaders.length} className="bg-[#fff4f4] px-4 py-3 text-coninco-red">AUDIT OUTPUT · SOURCE VS ENGINE</th></tr>
            <tr>{inputHeaders.map((label) => <th key={label} className="whitespace-nowrap border-b border-coninco-line px-4 py-4">{label}</th>)}{outputHeaders.map((label) => <th key={label} className="whitespace-nowrap border-b border-l border-coninco-line px-4 py-4">{label}</th>)}</tr>
          </thead>
          <tbody>{rows.map((row) => <tr key={row.sourceRow} className="hover:bg-[#fffafa]">
            <td className="px-4 py-4 font-black text-coninco-red">{row.sourceRow}</td><td className="px-4 py-4">{row.itemNo}</td><td className="max-w-[260px] px-4 py-4 font-semibold">{row.description}</td><td className="px-4 py-4">{row.unit || "—"}</td><td className="px-4 py-4"><Tag value={row.section}/></td><td className="px-4 py-4">{row.productType}</td>
            {[row.w, row.h, row.w1, row.h1, row.w2, row.h2, row.lm, row.rd, row.a, row.t, row.q].map((value, index) => <td key={`input-number-${index}`} className="px-4 py-4 font-mono">{n(value)}</td>)}
            {[row.plan, row.upDown, row.verticalAxis, row.formulaDetail, row.acousticFireproofDetail, row.note].map((value, index) => <td key={`input-text-${index}`} className="max-w-[240px] px-4 py-4 text-coninco-muted">{value || "—"}</td>)}
            <td className="border-l-2 border-coninco-line px-4 py-4 font-mono">{n(row.contractorK)}</td><td className="px-4 py-4 font-mono">{n(row.calculatedK)}</td><td className="px-4 py-4"><CheckCell match={row.kMatch}/></td>
            <td className="border-l border-coninco-line px-4 py-4 font-mono">{n(row.contractorS)}</td><td className="px-4 py-4 font-mono">{n(row.calculatedS)}</td><td className="px-4 py-4"><CheckCell match={row.sMatch}/></td>
            <td className="border-l border-coninco-line px-4 py-4 font-mono">{n(row.contractorInnerS)}</td><td className="px-4 py-4 font-mono">{n(row.calculatedInnerS)}</td><td className="px-4 py-4"><CheckCell match={row.innerMatch}/></td>
          </tr>)}</tbody>
        </table>
        {!rows.length && <div className="px-6 py-14 text-center text-sm text-coninco-muted">No comparison rows found.</div>}
      </div>
    </div>
    <div className="flex justify-end"><DownloadButton entry={entry} setToast={setToast}/></div>
  </div>;
}

function Metric({ label, value, icon, danger = false }: { label: string; value: string; icon: React.ReactNode; danger?: boolean }) { return <div className="rounded-2xl border border-coninco-line bg-white p-5 shadow-soft"><div className="flex items-center justify-between text-coninco-muted"><span className="text-[10px] font-black uppercase tracking-wider">{label}</span><span className={danger ? "text-coninco-red" : "text-coninco-red"}>{icon}</span></div><div className={`mt-3 text-2xl font-black ${danger ? "text-coninco-red" : "text-coninco-ink"}`}>{value}</div></div>; }
function Tag({ value }: { value: string }) { return <span className="rounded-lg bg-coninco-soft px-2.5 py-1 text-[10px] font-black capitalize text-coninco-red">{value}</span>; }
function CheckCell({ match }: { match: boolean }) { return <span className={`inline-flex items-center gap-1 rounded-full px-2.5 py-1 text-[10px] font-black ${match ? "bg-emerald-50 text-emerald-700" : "bg-red-50 text-red-700"}`}>{match ? <Check size={12}/> : <X size={12}/>} {match ? "MATCH" : "DIFF"}</span>; }

function DownloadButton({ entry, setToast }: { entry: ArchiveEntry; setToast: (value: string) => void }) {
  const [open, setOpen] = useState(false);
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
