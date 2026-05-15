import { useState, useEffect } from "react";
import { supabase } from "./supabaseClient";
import App from "./App";

const WB = {
  bg:"#0d1b2a", card:"#152535", border:"#1e3448", borderLight:"#2a4560",
  primary:"#4a7c8e", primaryLight:"#8aabb8",
  text:"#f8fafc", textMuted:"#8aabb8", textDim:"#475569",
  green:"#22c55e", greenBg:"#052e16", greenBorder:"#14532d",
  red:"#ef4444", redBg:"#1a0505",
  yellow:"#eab308",
};

const ALLOWED_DOMAIN = "wbnation.com";

function LoginScreen({ onLogin }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [isSignUp, setIsSignUp] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [message, setMessage] = useState("");

  async function handleSubmit() {
    setError(""); setMessage("");
    if (!email.endsWith(`@${ALLOWED_DOMAIN}`)) {
      setError(`Only @${ALLOWED_DOMAIN} email addresses are allowed.`);
      return;
    }
    setLoading(true);
    try {
      if (isSignUp) {
        const { error } = await supabase.auth.signUp({ email, password });
        if (error) throw error;
        setMessage("Check your email for a confirmation link!");
      } else {
        const { data, error } = await supabase.auth.signInWithPassword({ email, password });
        if (error) throw error;
        onLogin(data.user);
      }
    } catch (err) {
      setError(err.message);
    }
    setLoading(false);
  }

  return (
    <div style={{ minHeight: "100vh", background: WB.bg, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Inter',system-ui,sans-serif", padding: 24 }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Syne:wght@700;800&display=swap'); *{box-sizing:border-box;} input::placeholder{color:${WB.textDim}!important;}`}</style>
      <div style={{ width: "100%", maxWidth: 420 }}>
        {/* Logo + title */}
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <img src="/Wright Brothers Logo.png" alt="Wright Brothers" style={{ height: 72, objectFit: "contain", marginBottom: 16 }} onError={e => { e.target.style.display = "none"; }} />
          <h1 style={{ fontFamily: "'Syne',sans-serif", fontSize: 28, fontWeight: 800, margin: 0, color: WB.text }}>Project Spend</h1>
          <p style={{ fontFamily: "'Syne',sans-serif", fontSize: 20, fontWeight: 700, color: WB.primaryLight, margin: "4px 0 0" }}>Baseline Tracker</p>
          <p style={{ color: WB.textDim, fontSize: 13, marginTop: 8 }}>Wright Brothers · The Building Company</p>
        </div>

        {/* Form */}
        <div style={{ background: WB.card, borderRadius: 16, border: `1px solid ${WB.border}`, padding: 28 }}>
          <h2 style={{ fontFamily: "'Syne',sans-serif", fontSize: 18, fontWeight: 700, margin: "0 0 20px", color: WB.text }}>
            {isSignUp ? "Create your account" : "Sign in to your account"}
          </h2>

          <div style={{ marginBottom: 14 }}>
            <label style={{ fontSize: 12, color: WB.textDim, display: "block", marginBottom: 5 }}>Work Email</label>
            <input type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder={`yourname@${ALLOWED_DOMAIN}`}
              onKeyDown={e => e.key === "Enter" && handleSubmit()}
              style={{ width: "100%", padding: "10px 12px", background: WB.bg, border: `1px solid ${WB.border}`, borderRadius: 8, color: WB.text, fontSize: 14, outline: "none" }} />
          </div>

          <div style={{ marginBottom: 20 }}>
            <label style={{ fontSize: 12, color: WB.textDim, display: "block", marginBottom: 5 }}>Password</label>
            <input type="password" value={password} onChange={e => setPassword(e.target.value)} placeholder="••••••••"
              onKeyDown={e => e.key === "Enter" && handleSubmit()}
              style={{ width: "100%", padding: "10px 12px", background: WB.bg, border: `1px solid ${WB.border}`, borderRadius: 8, color: WB.text, fontSize: 14, outline: "none" }} />
          </div>

          {error && <div style={{ background: WB.redBg, border: `1px solid ${WB.red}40`, borderRadius: 8, padding: "10px 12px", color: WB.red, fontSize: 13, marginBottom: 14 }}>{error}</div>}
          {message && <div style={{ background: WB.greenBg, border: `1px solid ${WB.greenBorder}`, borderRadius: 8, padding: "10px 12px", color: WB.green, fontSize: 13, marginBottom: 14 }}>{message}</div>}

          <button onClick={handleSubmit} disabled={loading}
            style={{ width: "100%", padding: "11px", background: WB.primary, border: "none", borderRadius: 8, color: WB.text, fontSize: 14, fontWeight: 600, cursor: loading ? "not-allowed" : "pointer", opacity: loading ? 0.7 : 1 }}>
            {loading ? "Please wait…" : isSignUp ? "Create Account" : "Sign In"}
          </button>

          <p style={{ textAlign: "center", marginTop: 16, fontSize: 13, color: WB.textDim }}>
            {isSignUp ? "Already have an account? " : "Don't have an account? "}
            <button onClick={() => { setIsSignUp(!isSignUp); setError(""); setMessage(""); }}
              style={{ background: "none", border: "none", color: WB.primaryLight, cursor: "pointer", fontSize: 13, fontWeight: 600 }}>
              {isSignUp ? "Sign in" : "Sign up"}
            </button>
          </p>

          <p style={{ textAlign: "center", marginTop: 8, fontSize: 11, color: WB.textDim }}>
            Access restricted to @{ALLOWED_DOMAIN} email addresses
          </p>
        </div>
      </div>
    </div>
  );
}

function ProjectManager({ user, onOpenProject, onSignOut }) {
  const [projects, setProjects] = useState([]);
  const [loading, setLoading] = useState(true);
  const [newProjectName, setNewProjectName] = useState("");
  const [creating, setCreating] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => { loadProjects(); }, []);

  async function loadProjects() {
    setLoading(true);
    const { data, error } = await supabase.from("projects").select("id, name, created_at, updated_at").order("updated_at", { ascending: false });
    if (!error) setProjects(data || []);
    setLoading(false);
  }

  async function createProject() {
    if (!newProjectName.trim()) return;
    setCreating(true); setError("");
    const { data, error } = await supabase.from("projects").insert({ name: newProjectName.trim(), user_id: user.id, data: {} }).select().single();
    if (error) { setError(error.message); }
    else { setNewProjectName(""); setProjects(p => [data, ...p]); onOpenProject(data); }
    setCreating(false);
  }

  async function deleteProject(id) {
    if (!confirm("Delete this project? This cannot be undone.")) return;
    await supabase.from("projects").delete().eq("id", id);
    setProjects(p => p.filter(proj => proj.id !== id));
  }

  return (
    <div style={{ minHeight: "100vh", background: WB.bg, fontFamily: "'Inter',system-ui,sans-serif", padding: 24 }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Syne:wght@700;800&display=swap'); *{box-sizing:border-box;}`}</style>
      <div style={{ maxWidth: 800, margin: "0 auto" }}>

        {/* Header */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 28, flexWrap: "wrap", gap: 12 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
            <img src="/Wright Brothers Logo.png" alt="Wright Brothers" style={{ height: 52, objectFit: "contain" }} onError={e => { e.target.style.display = "none"; }} />
            <div>
              <h1 style={{ fontFamily: "'Syne',sans-serif", fontSize: 22, fontWeight: 800, margin: 0, color: WB.text }}>My Projects</h1>
              <p style={{ color: WB.textDim, fontSize: 12, margin: 0 }}>{user.email}</p>
            </div>
          </div>
          <button onClick={onSignOut} style={{ padding: "7px 14px", background: "transparent", border: `1px solid ${WB.border}`, borderRadius: 8, color: WB.textDim, cursor: "pointer", fontSize: 12 }}>Sign Out</button>
        </div>

        {/* New project */}
        <div style={{ background: WB.card, borderRadius: 14, border: `1px solid ${WB.border}`, padding: 20, marginBottom: 20 }}>
          <h2 style={{ fontFamily: "'Syne',sans-serif", fontSize: 16, fontWeight: 700, margin: "0 0 14px", color: WB.text }}>New Project</h2>
          <div style={{ display: "flex", gap: 10 }}>
            <input value={newProjectName} onChange={e => setNewProjectName(e.target.value)} placeholder="e.g. Floating Feather MHP Clubhouse"
              onKeyDown={e => e.key === "Enter" && createProject()}
              style={{ flex: 1, padding: "9px 12px", background: WB.bg, border: `1px solid ${WB.border}`, borderRadius: 8, color: WB.text, fontSize: 13, outline: "none" }} />
            <button onClick={createProject} disabled={creating || !newProjectName.trim()}
              style={{ padding: "9px 20px", background: WB.primary, border: "none", borderRadius: 8, color: WB.text, fontSize: 13, fontWeight: 600, cursor: "pointer", opacity: creating ? 0.7 : 1 }}>
              {creating ? "Creating…" : "Create Project"}
            </button>
          </div>
          {error && <div style={{ color: WB.red, fontSize: 12, marginTop: 8 }}>{error}</div>}
        </div>

        {/* Projects list */}
        <div style={{ background: WB.card, borderRadius: 14, border: `1px solid ${WB.border}`, padding: 20 }}>
          <h2 style={{ fontFamily: "'Syne',sans-serif", fontSize: 16, fontWeight: 700, margin: "0 0 14px", color: WB.text }}>Your Projects</h2>
          {loading ? (
            <div style={{ color: WB.textDim, fontSize: 13, padding: 20, textAlign: "center" }}>Loading projects…</div>
          ) : projects.length === 0 ? (
            <div style={{ color: WB.textDim, fontSize: 13, padding: 20, textAlign: "center" }}>No projects yet — create one above!</div>
          ) : (
            <div style={{ display: "grid", gap: 8 }}>
              {projects.map(proj => (
                <div key={proj.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "12px 14px", background: WB.bg, borderRadius: 10, border: `1px solid ${WB.border}` }}>
                  <div style={{ flex: 1, overflow: "hidden" }}>
                    <div style={{ fontWeight: 600, color: WB.text, fontSize: 14, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{proj.name}</div>
                    <div style={{ fontSize: 11, color: WB.textDim, fontFamily: "monospace", marginTop: 2 }}>
                      Updated {new Date(proj.updated_at).toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" })}
                    </div>
                  </div>
                  <button onClick={() => onOpenProject(proj)}
                    style={{ padding: "7px 16px", background: WB.primary, border: "none", borderRadius: 7, color: WB.text, fontSize: 12, fontWeight: 600, cursor: "pointer", flexShrink: 0 }}>
                    Open
                  </button>
                  <button onClick={() => deleteProject(proj.id)}
                    style={{ padding: "7px 12px", background: "transparent", border: `1px solid ${WB.border}`, borderRadius: 7, color: WB.red, fontSize: 12, cursor: "pointer", flexShrink: 0 }}>
                    Delete
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

export default function AppWrapper() {
  const [user, setUser] = useState(null);
  const [currentProject, setCurrentProject] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);

  useEffect(() => {
    // Check existing session
    supabase.auth.getSession().then(({ data: { session } }) => {
      setUser(session?.user ?? null);
      setAuthLoading(false);
    });
    // Listen for auth changes
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setUser(session?.user ?? null);
      if (!session) setCurrentProject(null);
    });
    return () => subscription.unsubscribe();
  }, []);

  async function handleSignOut() {
    await supabase.auth.signOut();
    setCurrentProject(null);
  }

  if (authLoading) {
    return (
      <div style={{ minHeight: "100vh", background: WB.bg, display: "flex", alignItems: "center", justifyContent: "center" }}>
        <div style={{ color: WB.textDim, fontSize: 14, fontFamily: "sans-serif" }}>Loading…</div>
      </div>
    );
  }

  if (!user) return <LoginScreen onLogin={setUser} />;
  if (!currentProject) return <ProjectManager user={user} onOpenProject={setCurrentProject} onSignOut={handleSignOut} />;

  return (
    <App
      user={user}
      project={currentProject}
      onBackToProjects={() => setCurrentProject(null)}
      onSignOut={handleSignOut}
      supabase={supabase}
    />
  );
}
