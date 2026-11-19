import { useEffect, useState } from "react";
import {
  Activity, ArrowUpRight, Bell, BookOpen, Bot, BrainCircuit, ChevronDown,
  CircleHelp, FileText, FolderOpen, Gauge, HardDrive, LayoutDashboard,
  Plus, Search, Settings, ShieldCheck, Sparkles, UploadCloud, Users
} from "lucide-react";

type Health = { status: string; phase: number };

const navItems = [
  { label: "Overview", icon: LayoutDashboard, active: true },
  { label: "Documents", icon: FolderOpen, tag: "Phase 4" },
  { label: "AI Workspace", icon: Bot, tag: "Phase 6" },
  { label: "Reports", icon: FileText, tag: "Phase 7" },
  { label: "Knowledge Graph", icon: BrainCircuit, tag: "Phase 5" }
];

function App() {
  const [health, setHealth] = useState<Health | null>(null);
  const [active, setActive] = useState("Overview");

  useEffect(() => {
    fetch("/api/health")
      .then((r) => r.json())
      .then(setHealth)
      .catch(() => setHealth(null));
  }, []);

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-mark"><BrainCircuit size={23} /></div>
          <div><strong>GraphPilot<span> AI</span></strong><small>DOCUMENT INTELLIGENCE</small></div>
        </div>
        <div className="workspace-switch">
          <div className="workspace-avatar">G</div>
          <div><b>GraphPilot Workspace</b><small>Personal workspace</small></div>
          <ChevronDown size={15} />
        </div>
        <div className="nav-label">WORKSPACE</div>
        <nav>
          {navItems.map(({ label, icon: Icon, tag, active: isDefault }) => (
            <button key={label} className={`nav-item ${active === label ? "selected" : ""}`}
              onClick={() => setActive(label)}>
              <Icon size={18} /><span>{label}</span>
              {tag && <em>{tag}</em>}
              {isDefault && <i />}
            </button>
          ))}
        </nav>
        <div className="nav-label tools-label">MANAGE</div>
        <button className="nav-item" onClick={() => setActive("Team & Access")}><Users size={18}/><span>Team & Access</span><em>Later</em></button>
        <button className="nav-item" onClick={() => setActive("Settings")}><Settings size={18}/><span>Settings</span></button>
        <div className="sidebar-bottom">
          <div className="secure-note"><ShieldCheck size={17}/><div><b>Privacy-first by design</b><small>Workspace isolation is planned</small></div></div>
          <div className="profile"><div className="profile-avatar">GP</div><div><b>Workspace Admin</b><small>Free portfolio edition</small></div><Settings size={16}/></div>
        </div>
      </aside>

      <main className="main">
        <header className="topbar">
          <div className="breadcrumbs">Workspace <span>/</span> <b>{active}</b></div>
          <div className="top-actions"><div className="status"><i/> {health?.status === "ok" ? "System operational" : "Checking system"}</div><button className="icon-btn"><Search size={18}/></button><button className="icon-btn"><Bell size={18}/></button><div className="top-avatar">GP</div></div>
        </header>

        <section className="content">
          <div className="welcome-row">
            <div><div className="eyebrow"><Sparkles size={14}/> YOUR AI KNOWLEDGE WORKSPACE</div><h1>Good to see you<span>.</span></h1><p>Turn your organization's documents into trusted, actionable knowledge.</p></div>
            <button className="primary-btn" onClick={() => setActive("Documents")}><Plus size={17}/> Add documents</button>
          </div>

          {active !== "Overview" && <div className="notice"><Sparkles size={17}/><div><b>{active} is part of the development roadmap.</b><p>This Phase 1 release establishes the deployable interface and backend foundation. Features will be activated phase by phase.</p></div></div>}

          <div className="stats-grid">
            <Stat icon={<FileText/>} label="Documents indexed" value="—" note="Available in Phase 4" />
            <Stat icon={<BrainCircuit/>} label="Knowledge chunks" value="—" note="Available in Phase 5" />
            <Stat icon={<Bot/>} label="AI conversations" value="—" note="Available in Phase 6" />
            <Stat icon={<Activity/>} label="Workspace activity" value="—" note="Analytics coming later" />
          </div>

          <div className="section-heading"><div><h2>Get started</h2><p>Your document intelligence workflow, all in one place.</p></div><button className="text-btn">View roadmap <ArrowUpRight size={15}/></button></div>
          <div className="action-grid">
            <ActionCard icon={<UploadCloud/>} number="01" title="Bring your documents" desc="Upload PDFs, DOCX and text files, or connect a cloud source." status="Planned · Phase 4" />
            <ActionCard icon={<BrainCircuit/>} number="02" title="Build your knowledge base" desc="Organize content into topic-aware chunks with source metadata." status="Planned · Phase 5" />
            <ActionCard icon={<Sparkles/>} number="03" title="Ask and discover" desc="Get grounded answers with document and page-level citations." status="Planned · Phase 6" />
          </div>

          <div className="bottom-grid">
            <div className="panel">
              <div className="panel-title"><div><h3>Recent documents</h3><p>Your latest knowledge sources will appear here.</p></div><BookOpen size={19}/></div>
              <div className="empty-state"><div className="empty-icon"><FolderOpen size={25}/></div><b>Your workspace is ready</b><p>Document upload will be enabled in Phase 4.</p><button className="outline-btn" onClick={() => setActive("Documents")}>Explore documents</button></div>
            </div>
            <div className="panel assistant-panel">
              <div className="panel-title"><div><h3>AI assistant</h3><p>Grounded answers from your knowledge base.</p></div><div className="ai-badge"><Sparkles size={13}/> AI</div></div>
              <div className="assistant-preview"><div className="assistant-orb"><BrainCircuit size={27}/></div><b>Your AI workspace is taking shape</b><p>Once documents and retrieval are connected, ask questions across your entire knowledge base.</p><div className="fake-input">Ask anything about your documents… <span><ArrowUpRight size={16}/></span></div></div>
            </div>
          </div>
          <footer><span>© 2026 GraphPilot AI</span><span><ShieldCheck size={13}/> Built with privacy in mind <span className="dot">·</span> Phase 1 foundation</span></footer>
        </section>
      </main>
    </div>
  );
}

function Stat({ icon, label, value, note }: { icon: React.ReactNode; label: string; value: string; note: string }) {
  return <div className="stat-card"><div className="stat-top"><span className="stat-icon">{icon}</span><span className="stat-dash"><Gauge size={15}/></span></div><div className="stat-value">{value}</div><b>{label}</b><small>{note}</small></div>;
}
function ActionCard({ icon, number, title, desc, status }: { icon: React.ReactNode; number: string; title: string; desc: string; status: string }) {
  return <div className="action-card"><div className="action-top"><span className="action-icon">{icon}</span><small>{number}</small></div><h3>{title}</h3><p>{desc}</p><div className="planned">{status}</div></div>;
}
export default App;
