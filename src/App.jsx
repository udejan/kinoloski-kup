import { useState, useEffect } from "react";
import { supabase } from "./supabase";

const loadXLSX = () => new Promise((resolve) => {
  if (window.XLSX) return resolve(window.XLSX);
  const s = document.createElement("script");
  s.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
  s.onload = () => resolve(window.XLSX);
  document.head.appendChild(s);
});

const CATEGORIES = ["Juniori", "Otvoreni razred", "Šampioni", "Veterani"];
const STATUS_LABEL = { pending: "Na čekanju", approved: "Odobreno", rejected: "Odbijeno" };
const STATUS_COLOR = {
  pending:  { bg: "#FEF3C7", color: "#92400E", border: "#F59E0B" },
  approved: { bg: "#D1FAE5", color: "#065F46", border: "#10B981" },
  rejected: { bg: "#FEE2E2", color: "#991B1B", border: "#EF4444" },
};
const slugify = (str) =>
  str.replace(/\s+/g, "_").replace(/[^a-zA-Z0-9_-]/g, "").slice(0, 40) || "Takmicenje";
const fmtDate = (d) => {
  if (!d) return "";
  try { return new Date(d).toLocaleDateString("sr-RS", { day: "numeric", month: "long", year: "numeric" }); }
  catch { return d; }
};

export default function App() {
  // core data
  const [competition, setCompetition]   = useState(null);
  const [applications, setApplications] = useState([]);
  const [users, setUsers]               = useState([]);
  const [loading, setLoading]           = useState(true);

  // session
  const [currentUser, setCurrentUser] = useState(null);
  const [userProfile, setUserProfile] = useState(null);
  const [viewRole, setViewRole]       = useState("guest");
  const [page, setPage]               = useState("home");

  // auth form
  const [authMode, setAuthMode]   = useState("login");
  const [authForm, setAuthForm]   = useState({ name: "", email: "", password: "" });
  const [authError, setAuthError] = useState("");
  const [authLoading, setAuthLoading] = useState(false);

  // dog form
  const [dogForm, setDogForm]       = useState({ dogName: "", breed: "", age: "", ageUnit: "meseci", microchip: "", pedigree: "", category: "Juniori" });
  const [dogError, setDogError]     = useState("");
  const [dogSuccess, setDogSuccess] = useState("");
  const [dogLoading, setDogLoading] = useState(false);

  // modals
  const [archiveModal, setArchiveModal]   = useState(false);
  const [newCompModal, setNewCompModal]   = useState(false);
  const [newCompForm, setNewCompForm]     = useState({ name: "", date: "", city: "", deadline: "" });
  const [newCompError, setNewCompError]   = useState("");
  const [editCompModal, setEditCompModal] = useState(false);
  const [editCompForm, setEditCompForm]   = useState({ name: "", date: "", city: "", deadline: "" });

  // misc
  const [exporting, setExporting]       = useState(false);
  const [resetMode, setResetMode] = useState(false);
  const [newPassword, setNewPassword] = useState("");
  const [resetComplete, setResetComplete] = useState(false);
  const [showPassword, setShowPassword] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  const [resetEmail, setResetEmail] = useState("");
  const [resetMsg, setResetMsg] = useState("");
  const [notification, setNotification] = useState(null);

  // ── COMPUTED ───────────────────────────────────────────────────────────────
  const effectiveRole = userProfile
    ? (viewRole === "admin" && userProfile.role === "admin" ? "admin" : "user")
    : "guest";
  const hasComp = competition !== null;
  const regOpen = hasComp && competition.registration_open && (
    !competition.deadline || new Date(competition.deadline) >= new Date(new Date().toDateString())
  );
  const daysToDeadline = hasComp && competition.deadline ? (() => {
    const diff = new Date(competition.deadline) - new Date(new Date().toDateString());
    return Math.ceil(diff / 86400000);
  })() : null;

  const showNotif = (msg, type = "success") => {
    setNotification({ msg, type });
    setTimeout(() => setNotification(null), 3500);
  };

  // ── INIT: auth listener + load data ───────────────────────────────────────
  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setCurrentUser(session?.user ?? null);
      if (session?.user) loadProfile(session.user.id);
    });
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_e, session) => {
      setCurrentUser(session?.user ?? null);
      if (session?.user) loadProfile(session.user.id);
      else { setUserProfile(null); setViewRole("guest"); }
    });
    loadCompetition();
    return () => subscription.unsubscribe();
  }, []);
  
  useEffect(() => {
  const hash = window.location.hash;
  if (hash && hash.includes("access_token")) {
    setPage("reset");
  }
  if (hash && hash.includes("error=access_denied")) {
    showNotif("Link za reset je istekao. Pokušajte ponovo.", "error");
    window.history.replaceState(null, "", window.location.pathname);
  }
}, []);

  useEffect(() => {
    if (competition) loadApplications();
  }, [competition]);

  // ── DATA LOADERS ──────────────────────────────────────────────────────────
  const loadProfile = async (uid) => {
    const { data } = await supabase.from("profiles").select("*").eq("id", uid).single();
    if (data) {
      setUserProfile(data);
      setViewRole(data.role === "admin" ? "admin" : "user");
    }
  };

  const loadCompetition = async () => {
    setLoading(true);
    const { data } = await supabase
      .from("competitions")
      .select("*")
      .order("created_at", { ascending: false })
      .limit(1)
      .maybeSingle();
    setCompetition(data ?? null);
    setLoading(false);
  };

  const loadApplications = async () => {
    const { data } = await supabase
      .from("applications")
      .select("*")
      .order("created_at", { ascending: true });
    setApplications(data ?? []);
  };

  const loadUsers = async () => {
    const { data } = await supabase.from("profiles").select("*").order("created_at");
    setUsers(data ?? []);
  };

  // ── AUTH ──────────────────────────────────────────────────────────────────
  const handleAuth = async () => {
    setAuthError(""); setAuthLoading(true);
    if (authMode === "register") {
      if (!authForm.name || !authForm.email || !authForm.password) {
        setAuthError("Sva polja su obavezna."); setAuthLoading(false); return;
      }
const { data, error } = await supabase.auth.signUp({
  email: authForm.email,
  password: authForm.password,
  options: { data: { name: authForm.name } },
});
if (error) { setAuthError(error.message); setAuthLoading(false); return; }
if (data?.user) {
  await supabase.from("profiles").insert({
    id: data.user.id,
    name: authForm.name,
    role: "user",
  });
  await supabase.auth.signOut();
}
setAuthMode("login");
setAuthForm({ name: "", email: "", password: "" });
showNotif("Registracija uspešna! Prijavite se sa vašim podacima.");
    } else {
      const { error } = await supabase.auth.signInWithPassword({
        email: authForm.email,
        password: authForm.password,
      });
      if (error) { setAuthError("Pogrešan email ili lozinka."); setAuthLoading(false); return; }
      showNotif(`Dobrodošli!`);
      setDogSuccess("");
      setDogError("");
      setPage("dashboard");
      if (competition) loadApplications();
    }
    setAuthForm({ name: "", email: "", password: "" });
    setAuthLoading(false);
  };

  const handleLogout = async () => {
    await supabase.auth.signOut();
    setPage("home"); setViewRole("guest");
    showNotif("Odjavili ste se.", "info");
  };

  const handleResetPassword = async () => {
    setResetMsg("");
    if (!resetEmail) return setResetMsg("Unesite email adresu.");
    const { error } = await supabase.auth.resetPasswordForEmail(resetEmail, {
      redirectTo: window.location.origin + "?reset=true",
      });
    if (error) return setResetMsg("Greška. Proverite email adresu.");
    setResetMsg("Link za reset lozinke je poslat na vaš email!");
  };

  const handleSetNewPassword = async () => {
  if (!newPassword || newPassword.length < 6)
    return showNotif("Lozinka mora imati najmanje 6 karaktera.", "error");
  const { error } = await supabase.auth.updateUser({ password: newPassword });
  if (error) return showNotif("Greška pri promeni lozinke.", "error");
  setResetComplete(true);
  showNotif("Lozinka uspešno promenjena! Možete se prijaviti.");
  await supabase.auth.signOut();
  setTimeout(() => {
    setPage("auth");
    setResetComplete(false);
    setNewPassword("");
    window.history.replaceState(null, "", window.location.pathname);
  }, 2000);
};

  // ── DOG SUBMIT ────────────────────────────────────────────────────────────
  const handleDogSubmit = async () => {
    setDogError(""); setDogSuccess("");
    if (!dogForm.dogName || !dogForm.breed || !dogForm.age || !dogForm.microchip)
      return setDogError("Ime, rasa, starost i broj mikročipa su obavezni.");
    if (!/^\d{15}$/.test(dogForm.microchip))
      return setDogError("Broj mikročipa mora imati tačno 15 cifara.");
    setDogLoading(true);
    const { error } = await supabase.from("applications").insert({
      competition_id: competition.id,
      owner_id: currentUser.id,
      owner_name: userProfile.name,
      dog_name: dogForm.dogName,
      breed: dogForm.breed,
      age: `${dogForm.age} ${dogForm.ageUnit}`,
      microchip: dogForm.microchip,
      pedigree: dogForm.pedigree,
      category: dogForm.category,
      status: "pending",
    });
    if (error) { setDogError("Greška pri slanju. Pokušajte ponovo."); setDogLoading(false); return; }
    await loadApplications();
    setDogForm({ dogName: "", breed: "", age: "", ageUnit: "meseci", microchip: "", pedigree: "", category: "Juniori" });
    setDogSuccess("Prijava uspešno poslata! Čeka se odobrenje administratora.");
    showNotif("Pas uspešno prijavljen!");
    setDogLoading(false);
  };

  // ── STATUS CHANGE ─────────────────────────────────────────────────────────
  const handleStatusChange = async (id, status) => {
    const { error } = await supabase.from("applications").update({ status }).eq("id", id);
    if (error) { showNotif("Greška pri promeni statusa.", "error"); return; }
    setApplications(applications.map(a => a.id === id ? { ...a, status } : a));
    showNotif(status === "approved" ? "Prijava odobrena." : "Prijava odbijena.",
              status === "approved" ? "success" : "error");
  };

  // ── COMPETITION CRUD ──────────────────────────────────────────────────────
  const openEditComp = () => {
    setEditCompForm({ name: competition.name, date: competition.date, city: competition.city, deadline: competition.deadline || "" });
    setEditCompModal(true);
  };
  const saveEditComp = async () => {
    if (!editCompForm.name || !editCompForm.date || !editCompForm.city) return;
    const { data, error } = await supabase.from("competitions")
      .update({ name: editCompForm.name, date: editCompForm.date, city: editCompForm.city, deadline: editCompForm.deadline || null })
      .eq("id", competition.id).select().single();
    if (!error) { setCompetition(data); setEditCompModal(false); showNotif("Podaci o takmičenju su ažurirani."); }
  };

  const openNewComp = () => {
    setNewCompForm({ name: "", date: "", city: "", deadline: "" });
    setNewCompError(""); setNewCompModal(true);
  };
  const handleCreateNewComp = async () => {
    if (!newCompForm.name || !newCompForm.date || !newCompForm.city)
      return setNewCompError("Naziv, datum i mesto su obavezni.");
    const { data, error } = await supabase.from("competitions").insert({
      name: newCompForm.name, date: newCompForm.date, city: newCompForm.city,
      deadline: newCompForm.deadline || null, registration_open: true,
    }).select().single();
    if (error) { setNewCompError("Greška pri kreiranju. Pokušajte ponovo."); return; }
    setCompetition(data); setApplications([]);
    setNewCompModal(false);
    showNotif(`Takmičenje "${data.name}" je kreirano. Prijave su otvorene!`);
  };

  const toggleRegistration = async () => {
    const { data } = await supabase.from("competitions")
      .update({ registration_open: !competition.registration_open })
      .eq("id", competition.id).select().single();
    if (data) { setCompetition(data); showNotif(data.registration_open ? "Prijave su ponovo otvorene." : "Prijave su zaključane.", "info"); }
  };

  // ── PROMOTE/DEMOTE USER ───────────────────────────────────────────────────
  const handleRoleChange = async (uid, newRole, name) => {
    const { error } = await supabase.from("profiles").update({ role: newRole }).eq("id", uid);
    if (error) { showNotif("Greška pri promeni uloge.", "error"); return; }
    setUsers(users.map(u => u.id === uid ? { ...u, role: newRole } : u));
    showNotif(newRole === "admin" ? `${name} je promovisan u admina!` : `${name} je degradiran na korisnika.`,
              newRole === "admin" ? "success" : "info");
  };

  // ── EXPORT XLSX ───────────────────────────────────────────────────────────
 const handleExportXLSX = async (comp = competition, apps = applications) => {
  setExporting(true);
  try {
    const XLSX = await loadXLSX();
    const approved = apps.filter(a => a.status === "approved");
    const wb = XLSX.utils.book_new();

    const headerStyle = {
      font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11 },
      fill: { fgColor: { rgb: "0369A1" } },
      alignment: { horizontal: "center", vertical: "center" },
      border: {
        top:    { style: "thin", color: { rgb: "CCCCCC" } },
        bottom: { style: "thin", color: { rgb: "CCCCCC" } },
        left:   { style: "thin", color: { rgb: "CCCCCC" } },
        right:  { style: "thin", color: { rgb: "CCCCCC" } },
      }
    };
    const titleStyle = {
      font: { bold: true, color: { rgb: "0C4A6E" }, sz: 14 },
      alignment: { horizontal: "left" },
    };
    const subtitleStyle = {
      font: { italic: true, color: { rgb: "475569" }, sz: 10 },
    };
    const rowEvenStyle = {
      fill: { fgColor: { rgb: "F0F9FF" } },
      border: {
        top:    { style: "thin", color: { rgb: "E2E8F0" } },
        bottom: { style: "thin", color: { rgb: "E2E8F0" } },
        left:   { style: "thin", color: { rgb: "E2E8F0" } },
        right:  { style: "thin", color: { rgb: "E2E8F0" } },
      }
    };
    const rowOddStyle = {
      fill: { fgColor: { rgb: "FFFFFF" } },
      border: {
        top:    { style: "thin", color: { rgb: "E2E8F0" } },
        bottom: { style: "thin", color: { rgb: "E2E8F0" } },
        left:   { style: "thin", color: { rgb: "E2E8F0" } },
        right:  { style: "thin", color: { rgb: "E2E8F0" } },
      }
    };
    const emptyStyle = {
      font: { italic: true, color: { rgb: "94A3B8" } },
      alignment: { horizontal: "center" },
    };

    const applyStyle = (ws, cellRef, style) => {
      if (!ws[cellRef]) ws[cellRef] = { v: "", t: "s" };
      ws[cellRef].s = style;
    };

    CATEGORIES.forEach(cat => {
      const rows = approved.filter(a => a.category === cat);
      const headers = ["Rb.", "Ime psa", "Rasa", "Starost", "Mikročip", "Broj rodovnika", "Vlasnik", "Status"];

      const data = [
        [`${comp.name} — Kategorija: ${cat}`, "", "", "", "", "", "", ""],
        [`Mesto: ${comp.city}   |   Datum: ${fmtDate(comp.date)}   |   Datum izvoza: ${new Date().toLocaleDateString("sr-RS")}`, "", "", "", "", "", "", ""],
        [],
        headers,
        ...rows.map((a, i) => [i + 1, a.dog_name, a.breed, a.age, a.microchip || "—", a.pedigree || "—", a.owner_name, "Odobreno"]),
      ];

      if (!rows.length) data.push(["", "Nema odobrenih prijava u ovoj kategoriji", "", "", "", "", "", ""]);

      const ws = XLSX.utils.aoa_to_sheet(data);
      ws["!cols"] = [{ wch: 5 }, { wch: 20 }, { wch: 22 }, { wch: 12 }, { wch: 18 }, { wch: 18 }, { wch: 24 }, { wch: 12 }];
      ws["!rows"] = [{ hpt: 24 }, { hpt: 16 }, { hpt: 8 }, { hpt: 20 }];

      // Merge title row
      ws["!merges"] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 7 } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: 7 } },
      ];

      // Style title and subtitle
      applyStyle(ws, "A1", titleStyle);
      applyStyle(ws, "A2", subtitleStyle);

      // Style headers (row 4, index 3)
      headers.forEach((_, ci) => {
        const cellRef = XLSX.utils.encode_cell({ r: 3, c: ci });
        applyStyle(ws, cellRef, headerStyle);
      });

      // Style data rows
      rows.forEach((_, ri) => {
        const style = ri % 2 === 0 ? rowEvenStyle : rowOddStyle;
        headers.forEach((_, ci) => {
          const cellRef = XLSX.utils.encode_cell({ r: 4 + ri, c: ci });
          applyStyle(ws, cellRef, style);
        });
      });

      // Style empty message
      if (!rows.length) {
        applyStyle(ws, "B5", emptyStyle);
      }

      XLSX.utils.book_append_sheet(wb, ws, cat);
    });

    // Summary sheet
    const summaryData = [
      [`${comp.name}`, "", ""],
      [`${comp.city}   |   ${fmtDate(comp.date)}`, "", ""],
      [],
      ["Kategorija", "Odobrenih prijava", ""],
      ...CATEGORIES.map(cat => [cat, approved.filter(a => a.category === cat).length, ""]),
      [],
      ["UKUPNO", approved.length, ""],
      [],
      ["Datum arhiviranja", new Date().toLocaleString("sr-RS"), ""],
    ];
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
    wsSummary["!cols"] = [{ wch: 28 }, { wch: 20 }, { wch: 10 }];
    wsSummary["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 2 } },
    ];

    applyStyle(wsSummary, "A1", titleStyle);
    applyStyle(wsSummary, "A2", subtitleStyle);
    applyStyle(wsSummary, "A4", headerStyle);
    applyStyle(wsSummary, "B4", headerStyle);

    CATEGORIES.forEach((_, i) => {
      const style = i % 2 === 0 ? rowEvenStyle : rowOddStyle;
      applyStyle(wsSummary, `A${5 + i}`, style);
      applyStyle(wsSummary, `B${5 + i}`, style);
    });

    const totalStyle = {
      font: { bold: true, color: { rgb: "FFFFFF" }, sz: 11 },
      fill: { fgColor: { rgb: "059669" } },
      alignment: { horizontal: "center" },
    };
    applyStyle(wsSummary, `A${5 + CATEGORIES.length + 1}`, totalStyle);
    applyStyle(wsSummary, `B${5 + CATEGORIES.length + 1}`, totalStyle);

    XLSX.utils.book_append_sheet(wb, wsSummary, "Pregled");
    XLSX.writeFile(wb, `${slugify(comp.name)}_Arhiva.xlsx`);
    showNotif(`Arhiva "${comp.name}" uspešno izvezena!`);
  } catch (e) {
    console.error(e);
    showNotif("Greška pri izvozu.", "error");
  }
  setExporting(false);
};

  // ── ARCHIVE & RESET ───────────────────────────────────────────────────────
  const handleArchiveAndReset = async () => {
    setArchiveModal(false);
    await handleExportXLSX(competition, applications);
    await supabase.from("applications").delete().eq("competition_id", competition.id);
    await supabase.from("competitions").delete().eq("id", competition.id);
    setCompetition(null); setApplications([]);
    showNotif("Takmičenje arhivirano. Kreirajte novo kada budete spremni.", "info");
  };

  const myApps = currentUser && hasComp
    ? applications.filter(a => a.owner_id === currentUser.id)
    : [];

  // ── STYLES ────────────────────────────────────────────────────────────────
  const s = {
    wrap: { fontFamily: "'Inter', system-ui, sans-serif", minHeight: "100vh", background: "linear-gradient(135deg,#E0F2FE 0%,#E8F5E9 100%)", color: "#1a1a2e" },
    nav:  { background: "rgba(255,255,255,0.92)", backdropFilter: "blur(8px)", borderBottom: "1px solid #e2e8f0", padding: "0 1.5rem", display: "flex", alignItems: "center", justifyContent: "space-between", height: 60, position: "sticky", top: 0, zIndex: 100 },
    logo: { display: "flex", alignItems: "center", gap: 8, fontWeight: 700, fontSize: 18, color: "#0369A1", cursor: "pointer" },
    navLinks: { display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" },
    navBtn: (active) => ({ background: active ? "#0369A1" : "transparent", color: active ? "#fff" : "#4B5563", border: active ? "none" : "1px solid #d1d5db", borderRadius: 8, padding: "6px 14px", fontSize: 13, fontWeight: 500, cursor: "pointer" }),
    roleToggle: { background: "#F0FDF4", border: "1px solid #86EFAC", color: "#15803D", borderRadius: 8, padding: "5px 12px", fontSize: 12, fontWeight: 600, cursor: "pointer" },
    card: { background: "#fff", borderRadius: 16, padding: "1.5rem", border: "1px solid #e2e8f0", boxShadow: "0 2px 8px rgba(0,0,0,0.06)" },
    infoCard: { background: "#fff", borderRadius: 16, padding: "1.5rem 1.25rem", border: "1px solid #e2e8f0", textAlign: "center", boxShadow: "0 2px 8px rgba(0,0,0,0.06)" },
    iconBox: (color) => ({ width: 48, height: 48, borderRadius: 12, background: color, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 12px", fontSize: 22 }),
    btn: (v = "primary") => ({
      background: v === "primary" ? "#0369A1" : v === "success" ? "#059669" : v === "danger" ? "#DC2626" : v === "warning" ? "#D97706" : v === "outline" ? "transparent" : "#6B7280",
      color: v === "outline" ? "#374151" : "#fff",
      border: v === "outline" ? "1px solid #d1d5db" : "none",
      borderRadius: 10, padding: "10px 20px", fontSize: 14, fontWeight: 600, cursor: "pointer",
    }),
    input:  { width: "100%", padding: "10px 14px", borderRadius: 10, border: "1px solid #d1d5db", fontSize: 14, outline: "none", boxSizing: "border-box", background: "#FAFAFA" },
    label:  { fontSize: 13, fontWeight: 600, color: "#374151", marginBottom: 5, display: "block" },
    fg:     { marginBottom: 16 },
    select: { width: "100%", padding: "10px 14px", borderRadius: 10, border: "1px solid #d1d5db", fontSize: 14, background: "#FAFAFA" },
    sectionTitle: { fontSize: 22, fontWeight: 700, color: "#0C4A6E", marginBottom: 4 },
    table: { width: "100%", borderCollapse: "collapse", fontSize: 14 },
    th:    { background: "#F1F5F9", padding: "10px 14px", textAlign: "left", fontWeight: 600, color: "#475569", fontSize: 13 },
    td:    { padding: "12px 14px", borderBottom: "1px solid #f1f5f9", verticalAlign: "middle" },
    badge: (st) => ({ ...STATUS_COLOR[st], border: `1px solid ${STATUS_COLOR[st]?.border}`, borderRadius: 20, padding: "3px 12px", fontSize: 12, fontWeight: 600, display: "inline-block" }),
    page:  { maxWidth: 960, margin: "0 auto", padding: "2rem 1.5rem" },
    notif: (type) => ({
      position: "fixed", top: 70, right: 20, zIndex: 999, maxWidth: 340,
      background: type === "error" ? "#FEE2E2" : type === "info" ? "#E0F2FE" : "#D1FAE5",
      color:      type === "error" ? "#991B1B" : type === "info" ? "#0C4A6E" : "#065F46",
      border:    `1px solid ${type === "error" ? "#FCA5A5" : type === "info" ? "#7DD3FC" : "#6EE7B7"}`,
      borderRadius: 12, padding: "12px 20px", fontSize: 14, fontWeight: 500, boxShadow: "0 4px 16px rgba(0,0,0,0.1)",
    }),
    modal:    { position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center", padding: "1rem" },
    modalBox: { background: "#fff", borderRadius: 20, padding: "2rem", maxWidth: 480, width: "100%", boxShadow: "0 8px 40px rgba(0,0,0,0.18)", maxHeight: "90vh", overflowY: "auto" },
  };

  // ── DEADLINE CHIP ─────────────────────────────────────────────────────────
  const deadlineChip = hasComp && competition.deadline ? (() => {
    if (daysToDeadline < 0)   return { text: "Rok prijava je istekao",    color: "#FEE2E2", textColor: "#991B1B" };
    if (daysToDeadline === 0) return { text: "Poslednji dan prijava!",    color: "#FEF3C7", textColor: "#92400E" };
    if (daysToDeadline <= 7)  return { text: `Još ${daysToDeadline} dana za prijavu`, color: "#FEF3C7", textColor: "#92400E" };
    return { text: `Prijave do ${fmtDate(competition.deadline)}`, color: "rgba(255,255,255,0.15)", textColor: "#fff" };
  })() : null;

  // ── STATUS BANNER ─────────────────────────────────────────────────────────
  const StatusBanner = () => hasComp ? (
    <div style={{ background: regOpen ? "linear-gradient(90deg,#0369A1,#0891B2)" : "linear-gradient(90deg,#475569,#64748B)", color: "#fff", padding: "10px 1.5rem", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <span style={{ fontSize: 20 }}>🏆</span>
        <div>
          <span style={{ fontWeight: 700, fontSize: 15 }}>{competition.name}</span>
          <span style={{ fontSize: 13, opacity: 0.85, marginLeft: 14 }}>📍 {competition.city}</span>
          <span style={{ fontSize: 13, opacity: 0.85, marginLeft: 14 }}>📅 {fmtDate(competition.date)}</span>
          {!regOpen && <span style={{ fontSize: 12, fontWeight: 700, marginLeft: 14, background: "rgba(255,255,255,0.2)", borderRadius: 20, padding: "2px 10px" }}>🔒 Prijave zaključane</span>}
          {regOpen && deadlineChip && <span style={{ fontSize: 12, fontWeight: 700, marginLeft: 14, background: deadlineChip.color, color: deadlineChip.textColor, borderRadius: 20, padding: "2px 10px" }}>{deadlineChip.text}</span>}
        </div>
      </div>
      {effectiveRole === "admin" && (
        <button onClick={openEditComp} style={{ background: "rgba(255,255,255,0.2)", color: "#fff", border: "1px solid rgba(255,255,255,0.4)", borderRadius: 8, padding: "4px 12px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>✏️ Uredi</button>
      )}
    </div>
  ) : (
    <div style={{ background: "#FEF3C7", borderBottom: "2px solid #F59E0B", padding: "12px 1.5rem", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
        <span style={{ fontSize: 20 }}>⏸️</span>
        <div>
          <span style={{ fontWeight: 700, color: "#92400E", fontSize: 14 }}>Nema aktivnog takmičenja</span>
          <span style={{ color: "#B45309", fontSize: 13, marginLeft: 10 }}>Prijave su trenutno zatvorene.</span>
        </div>
      </div>
      {effectiveRole === "admin" && (
        <button onClick={openNewComp} style={{ background: "#D97706", color: "#fff", border: "none", borderRadius: 8, padding: "6px 16px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>➕ Kreiraj takmičenje</button>
      )}
    </div>
  );

  // ── LOCKED SCREEN ─────────────────────────────────────────────────────────
  const LockedDashboard = ({ reason }) => (
    <div style={{ maxWidth: 560, margin: "0 auto", padding: "2rem 1.5rem" }}>
      <div style={{ ...s.card, textAlign: "center", padding: "3rem 2rem" }}>
        <div style={{ fontSize: 56, marginBottom: 16 }}>{reason === "deadline" ? "⏰" : reason === "manual" ? "🔒" : "⏸️"}</div>
        <h2 style={{ fontSize: 20, fontWeight: 700, color: "#0C4A6E", marginBottom: 10 }}>
          {reason === "deadline" ? "Rok za prijave je istekao" : reason === "manual" ? "Prijave su privremeno zaključane" : "Prijave su zatvorene"}
        </h2>
        <p style={{ color: "#64748B", lineHeight: 1.7, marginBottom: 0 }}>
          {reason === "deadline" ? `Rok za prijavljivanje na takmičenje „${competition?.name}" je istekao ${fmtDate(competition?.deadline)}.`
           : reason === "manual" ? "Administrator je privremeno zaključao prijave. Pratite obaveštenja za više informacija."
           : "Trenutno nema aktivnog takmičenja. Prijavljivanje pasa će biti dostupno čim administrator otvori novo takmičenje."}
        </p>
      </div>
    </div>
  );

  // ── LOADING ───────────────────────────────────────────────────────────────
  if (loading) return (
    <div style={{ ...s.wrap, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, color: "#0369A1" }}>
      🐕 Učitavanje...
    </div>
  );

  // ─────────────────────────────────────────────────────────────────────────
  return (
    <div style={s.wrap}>
      {notification && <div style={s.notif(notification.type)}>{notification.msg}</div>}

      {/* NAV */}
<nav style={s.nav}>
  <div style={s.logo} onClick={() => { setPage("home"); setMenuOpen(false); }}>🐕 KinološkiKup</div>

  {/* Hamburger dugme — vidljivo samo na mobilnom */}
  <button
    onClick={() => setMenuOpen(v => !v)}
    style={{ display: "none", background: "none", border: "none", fontSize: 24, cursor: "pointer", color: "#0369A1", padding: "4px 8px" }}
    className="hamburger"
  >
    {menuOpen ? "✕" : "☰"}
  </button>

  {/* Desktop navigacija */}
  <div style={s.navLinks} className="nav-desktop">
    {userProfile?.role === "admin" && (
      <button style={s.roleToggle} onClick={() => {
        const next = viewRole === "admin" ? "user" : "admin";
        setViewRole(next);
        setPage(next === "admin" ? "admin" : "dashboard");
      }}>
        {viewRole === "admin" ? "👤 Korisnik" : "🔧 Admin"}
      </button>
    )}
    {!currentUser ? (
      <>
        <button style={s.navBtn(page === "home")} onClick={() => setPage("home")}>Početna</button>
        <button style={s.navBtn(page === "auth")} onClick={() => { setPage("auth"); setAuthMode("login"); }}>Prijava</button>
        <button style={{ ...s.btn(), padding: "6px 16px", fontSize: 13 }} onClick={() => { setPage("auth"); setAuthMode("register"); }}>Registracija</button>
      </>
    ) : (
      <>
        <button style={s.navBtn(page === "home")} onClick={() => setPage("home")}>Početna</button>
        {effectiveRole === "admin"
          ? <button style={s.navBtn(page === "admin")} onClick={() => { setPage("admin"); loadUsers(); }}>Admin panel</button>
          : <button style={s.navBtn(page === "dashboard")} onClick={() => setPage("dashboard")}>Moje prijave</button>}
        <button style={{ ...s.btn("outline"), padding: "6px 14px", fontSize: 13 }} onClick={handleLogout}>Odjava</button>
      </>
    )}
  </div>
</nav>

{/* Mobilni meni */}
{menuOpen && (
  <div style={{ background: "#fff", borderBottom: "1px solid #e2e8f0", padding: "1rem 1.5rem", display: "flex", flexDirection: "column", gap: 10, boxShadow: "0 4px 12px rgba(0,0,0,0.08)", zIndex: 99, position: "relative" }} className="nav-mobile">
    {userProfile?.role === "admin" && (
      <button style={{ ...s.roleToggle, textAlign: "center" }} onClick={() => {
        const next = viewRole === "admin" ? "user" : "admin";
        setViewRole(next);
        setPage(next === "admin" ? "admin" : "dashboard");
        setMenuOpen(false);
      }}>
        {viewRole === "admin" ? "👤 Korisnik pogled" : "🔧 Admin pogled"}
      </button>
    )}
    {!currentUser ? (
      <>
        <button style={{ ...s.navBtn(page === "home"), textAlign: "left", padding: "10px 14px" }} onClick={() => { setPage("home"); setMenuOpen(false); }}>Početna</button>
        <button style={{ ...s.navBtn(page === "auth"), textAlign: "left", padding: "10px 14px" }} onClick={() => { setPage("auth"); setAuthMode("login"); setMenuOpen(false); }}>Prijava</button>
        <button style={{ ...s.btn(), width: "100%", padding: "10px" }} onClick={() => { setPage("auth"); setAuthMode("register"); setMenuOpen(false); }}>Registracija</button>
      </>
    ) : (
      <>
        <button style={{ ...s.navBtn(page === "home"), textAlign: "left", padding: "10px 14px" }} onClick={() => { setPage("home"); setMenuOpen(false); }}>Početna</button>
        {effectiveRole === "admin"
          ? <button style={{ ...s.navBtn(page === "admin"), textAlign: "left", padding: "10px 14px" }} onClick={() => { setPage("admin"); loadUsers(); setMenuOpen(false); }}>Admin panel</button>
          : <button style={{ ...s.navBtn(page === "dashboard"), textAlign: "left", padding: "10px 14px" }} onClick={() => { setPage("dashboard"); setMenuOpen(false); }}>Moje prijave</button>}
        <button style={{ ...s.btn("outline"), width: "100%", padding: "10px" }} onClick={() => { handleLogout(); setMenuOpen(false); }}>Odjava</button>
      </>
    )}
  </div>
)}

      <StatusBanner />

      {/* ── HOME ───────────────────────────────────────────────────────── */}
{page === "home" && (
  <>
    {/* Hero sekcija */}
    <div style={{
      position: "relative", minHeight: 480, display: "flex", alignItems: "center", justifyContent: "center",
      background: "linear-gradient(135deg, #0C4A6E 0%, #0369A1 100%)",
      overflow: "hidden",
    }}>
      <img
        src="https://images.unsplash.com/photo-1587300003388-59208cc962cb?w=1600&q=80"
        alt="Kinološko takmičenje"
        style={{ position: "absolute", inset: 0, width: "100%", height: "100%", objectFit: "cover", opacity: 0.25 }}
      />
      <div style={{ position: "relative", zIndex: 1, textAlign: "center", padding: "3rem 1.5rem", maxWidth: 700, margin: "0 auto" }}>
        <div style={{ fontSize: 56, marginBottom: 12 }}>🏆</div>
        {hasComp ? (
          <>
            <h1 style={{ fontSize: "clamp(1.8rem,5vw,2.8rem)", fontWeight: 800, color: "#fff", marginBottom: 10, lineHeight: 1.2 }}>{competition.name}</h1>
            <p style={{ fontSize: 16, color: "rgba(255,255,255,0.85)", marginBottom: 8 }}>📍 {competition.city} &nbsp;·&nbsp; 📅 {fmtDate(competition.date)}</p>
            <p style={{ fontSize: 15, color: "rgba(255,255,255,0.75)", marginBottom: 32, lineHeight: 1.6 }}>Prijavite svog ljubimca i takmičite se za titulu šampiona!</p>
          </>
        ) : (
          <p style={{ fontSize: 16, color: "rgba(255,255,255,0.8)", marginBottom: 32, lineHeight: 1.6 }}>Sledeće takmičenje još nije zakazano. Pratite nas za obaveštenja!</p>
        )}
        <div style={{ display: "flex", gap: 12, justifyContent: "center", flexWrap: "wrap" }}>
          {!currentUser ? (
            <>
              <button style={{ ...s.btn(), opacity: hasComp && regOpen ? 1 : 0.6, background: "#fff", color: "#0369A1", padding: "12px 28px", fontSize: 15 }}
                onClick={() => { setPage("auth"); setAuthMode("register"); }}>
                Prijavite psa →
              </button>
              <button style={{ background: "rgba(255,255,255,0.15)", color: "#fff", border: "1px solid rgba(255,255,255,0.4)", borderRadius: 10, padding: "12px 28px", fontSize: 15, fontWeight: 600, cursor: "pointer" }}
                onClick={() => { setPage("auth"); setAuthMode("login"); }}>
                Već imate nalog
              </button>
            </>
          ) : (
            <button style={{ ...s.btn(), background: "#fff", color: "#0369A1", padding: "12px 28px", fontSize: 15 }}
              onClick={() => setPage(effectiveRole === "admin" ? "admin" : "dashboard")}>
              Idi na dashboard →
            </button>
          )}
        </div>
      </div>
    </div>

    {/* Info kartice */}
    {hasComp && (
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 16, maxWidth: 780, margin: "2rem auto", padding: "0 1.5rem" }}>
        {[
          { icon: "📅", color: "#DBEAFE", title: "Datum",      text: fmtDate(competition.date) },
          { icon: "📍", color: "#D1FAE5", title: "Mesto",      text: competition.city },
          { icon: "🐾", color: "#FEF3C7", title: "Kategorije", text: `${CATEGORIES.length} kategorije` },
          { icon: "🎖️", color: "#FCE7F3", title: "Nagrade",    text: "Zlatna, srebrna, bronzana" },
        ].map((c, i) => (
          <div key={i} style={s.infoCard}>
            <div style={s.iconBox(c.color)}>{c.icon}</div>
            <div style={{ fontWeight: 700, fontSize: 13, color: "#1e293b", marginBottom: 4 }}>{c.title}</div>
            <div style={{ fontSize: 13, color: "#0369A1", fontWeight: 600 }}>{c.text}</div>
          </div>
        ))}
      </div>
    )}

    {/* Obaveštenje o uplati */}
    {hasComp && (
      <div style={{ maxWidth: 780, margin: "0 auto 2rem", padding: "0 1.5rem" }}>
        <div style={{ background: "#FEF3C7", border: "1px solid #F59E0B", borderRadius: 14, padding: "1.25rem 1.5rem" }}>
          <div style={{ fontWeight: 700, color: "#92400E", fontSize: 15, marginBottom: 8 }}>💳 Važno — uslovi prijave</div>
          <p style={{ color: "#78350F", fontSize: 14, lineHeight: 1.7, margin: 0 }}>
            Biće prihvaćene samo one prijave za koje je <strong>izvršena uplata kotizacije</strong>.
            Pri uplati kao poziv na broj navedite <strong>broj mikročipa psa</strong>.
            Administrator će odobriti prijavu nakon provere uplate.
            {competition.deadline && <><br /><strong>Rok za prijave: {fmtDate(competition.deadline)}</strong>{daysToDeadline !== null && daysToDeadline >= 0 && ` — još ${daysToDeadline} dana.`}</>}
          </p>
        </div>
      </div>
    )}

    {/* Galerija */}
    <div style={{ maxWidth: 780, margin: "0 auto 2rem", padding: "0 1.5rem" }}>
      <h2 style={{ fontSize: 20, fontWeight: 700, color: "#0C4A6E", marginBottom: 16, textAlign: "center" }}>Galerija</h2>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 14 }}>
        {[
          { url: "https://images.unsplash.com/photo-1558788353-f76d92427f16?w=600&q=80", alt: "Pas na takmičenju" },
          { url: "https://images.unsplash.com/photo-1477884213360-7e9d7dcc1e48?w=600&q=80", alt: "Kinološka izložba" },
          { url: "https://images.unsplash.com/photo-1450778869180-41d0601e046e?w=600&q=80", alt: "Psi na izložbi" },
          { url: "https://images.unsplash.com/photo-1537151625747-768eb6cf92b2?w=600&q=80", alt: "Pobednički pas" },
          { url: "https://images.unsplash.com/photo-1548199973-03cce0bbc87b?w=600&q=80", alt: "Psi u trčanju" },
          { url: "https://images.unsplash.com/photo-1552053831-71594a27632d?w=600&q=80", alt: "Pas sa medaljom" },
          
          ].map((img, i) => (
          <div key={i} style={{ borderRadius: 14, overflow: "hidden", boxShadow: "0 4px 12px rgba(0,0,0,0.1)", aspectRatio: "4/3" }}>
            <img src={img.url} alt={img.alt} style={{ width: "100%", height: "100%", objectFit: "cover", transition: "transform 0.3s", display: "block" }}
              onMouseEnter={e => e.target.style.transform = "scale(1.05)"}
              onMouseLeave={e => e.target.style.transform = "scale(1)"}
            />
          </div>
        ))}
      </div>
    </div>
  </>
)}
      {/* ── AUTH ───────────────────────────────────────────────────────── */}
      {page === "reset" && (
  <div style={{ maxWidth: 440, margin: "0 auto", padding: "2rem 1.5rem" }}>
    <div style={s.card}>
      <div style={{ textAlign: "center", marginBottom: 22 }}>
        <div style={{ fontSize: 38, marginBottom: 8 }}>🔑</div>
        <h2 style={{ fontSize: 19, fontWeight: 700, color: "#0C4A6E" }}>Nova lozinka</h2>
        <p style={{ fontSize: 13, color: "#64748B", marginTop: 6 }}>Unesite vašu novu lozinku.</p>
      </div>
      {!resetComplete && (
        <>
          <div style={s.fg}>
            <label style={s.label}>Nova lozinka</label>
            <div style={{ position: "relative" }}>
  <input style={{ ...s.input, paddingRight: 44 }} type={showPassword ? "text" : "password"}
    placeholder="Najmanje 6 karaktera" value={newPassword}
    onChange={e => setNewPassword(e.target.value)} />
  <button onClick={() => setShowPassword(v => !v)}
    style={{ position: "absolute", right: 12, top: "50%", transform: "translateY(-50%)", background: "none", border: "none", cursor: "pointer", fontSize: 16, color: "#94A3B8" }}>
    {showPassword ? "🙈" : "👁️"}
  </button>
</div>
          </div>
          <button style={{ ...s.btn(), width: "100%" }} onClick={handleSetNewPassword}>
            Sačuvaj novu lozinku
          </button>
        </>
      )}
      {resetComplete && (
        <div style={{ textAlign: "center", color: "#065F46", fontWeight: 600 }}>
          ✓ Lozinka promenjena! Preusmeravanje...
        </div>
      )}
    </div>
  </div>
)}
      {page === "auth" && (
  <div style={{ maxWidth: 440, margin: "0 auto", padding: "2rem 1.5rem" }}>
    <div style={s.card}>
      <div style={{ textAlign: "center", marginBottom: 22 }}>
        <div style={{ fontSize: 38, marginBottom: 8 }}>🐕</div>
        <h2 style={{ fontSize: 19, fontWeight: 700, color: "#0C4A6E" }}>
          {resetMode ? "Reset lozinke" : authMode === "login" ? "Prijavite se" : "Registrujte se"}
        </h2>
      </div>

      {resetMode ? (
        <>
          <p style={{ fontSize: 13, color: "#64748B", marginBottom: 16 }}>
            Unesite email adresu i poslaćemo vam link za reset lozinke.
          </p>
          <div style={s.fg}>
            <label style={s.label}>Email adresa</label>
            <input style={s.input} type="email" placeholder="vas@email.com"
              value={resetEmail} onChange={e => setResetEmail(e.target.value)} />
          </div>
          {resetMsg && (
            <div style={{ background: resetMsg.includes("poslat") ? "#D1FAE5" : "#FEE2E2", color: resetMsg.includes("poslat") ? "#065F46" : "#991B1B", borderRadius: 8, padding: "8px 12px", fontSize: 13, marginBottom: 12 }}>
              {resetMsg}
            </div>
          )}
          <button style={{ ...s.btn(), width: "100%", marginBottom: 12 }} onClick={handleResetPassword}>
            Pošalji link za reset
          </button>
          <div style={{ textAlign: "center", fontSize: 13, color: "#6B7280" }}>
            <span style={{ color: "#0369A1", cursor: "pointer", fontWeight: 600 }}
              onClick={() => { setResetMode(false); setResetMsg(""); setResetEmail(""); }}>
              ← Nazad na prijavu
            </span>
          </div>
        </>
      ) : (
        <>
          {authMode === "register" && (
            <div style={s.fg}>
              <label style={s.label}>Ime i prezime</label>
              <input style={s.input} placeholder="Vaše ime" value={authForm.name}
                onChange={e => setAuthForm({ ...authForm, name: e.target.value })} />
            </div>
          )}
          <div style={s.fg}>
            <label style={s.label}>Email adresa</label>
            <input style={s.input} type="email" placeholder="vas@email.com" value={authForm.email}
              onChange={e => setAuthForm({ ...authForm, email: e.target.value })} />
          </div>
          <div style={s.fg}>
            <label style={s.label}>Lozinka</label>
<div style={{ position: "relative" }}>
  <input style={{ ...s.input, paddingRight: 44 }} type={showPassword ? "text" : "password"}
    placeholder="••••••••" value={authForm.password}
    onChange={e => setAuthForm({ ...authForm, password: e.target.value })} />
  <button onClick={() => setShowPassword(v => !v)}
    style={{ position: "absolute", right: 12, top: "50%", transform: "translateY(-50%)", background: "none", border: "none", cursor: "pointer", fontSize: 16, color: "#94A3B8" }}>
    {showPassword ? "🙈" : "👁️"}
  </button>
</div>
          </div>
          {authMode === "login" && (
            <div style={{ textAlign: "right", marginTop: -10, marginBottom: 12 }}>
              <span style={{ fontSize: 12, color: "#0369A1", cursor: "pointer", fontWeight: 600 }}
                onClick={() => { setResetMode(true); setResetMsg(""); setResetEmail(authForm.email); }}>
                Zaboravili ste lozinku?
              </span>
            </div>
          )}
          {authError && (
            <div style={{ background: "#FEE2E2", color: "#991B1B", borderRadius: 8, padding: "8px 12px", fontSize: 13, marginBottom: 12 }}>
              {authError}
            </div>
          )}
          <button style={{ ...s.btn(), width: "100%", marginBottom: 12, opacity: authLoading ? 0.7 : 1 }}
            onClick={handleAuth} disabled={authLoading}>
            {authLoading ? "..." : authMode === "login" ? "Prijavite se" : "Registrujte se"}
          </button>
          <div style={{ textAlign: "center", fontSize: 13, color: "#6B7280" }}>
            {authMode === "login" ? "Nemate nalog? " : "Već imate nalog? "}
            <span style={{ color: "#0369A1", cursor: "pointer", fontWeight: 600 }}
              onClick={() => { setAuthMode(m => m === "login" ? "register" : "login"); setAuthError(""); }}>
              {authMode === "login" ? "Registrujte se" : "Prijavite se"}
            </span>
          </div>
        </>
      )}
    </div>
  </div>
)}
      {/* ── USER DASHBOARD ─────────────────────────────────────────────── */}
      {page === "dashboard" && currentUser && (
        !hasComp ? <LockedDashboard reason="none" /> :
        !regOpen ? <LockedDashboard reason={!competition.registration_open ? "manual" : "deadline"} /> : (
          <div style={s.page}>
            <div style={{ marginBottom: 22 }}>
              <h1 style={s.sectionTitle}>👋 Zdravo, {userProfile?.name}!</h1>
              <p style={{ color: "#64748B", fontSize: 14 }}>Upravljajte prijavama za <strong>{competition.name}</strong>.</p>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(120px,1fr))", gap: 12, marginBottom: 24 }}>
              {[
                { label: "Ukupno",     val: myApps.length,                                      color: "#DBEAFE" },
                { label: "Odobreno",   val: myApps.filter(a => a.status === "approved").length, color: "#D1FAE5" },
                { label: "Na čekanju", val: myApps.filter(a => a.status === "pending").length,  color: "#FEF3C7" },
                { label: "Odbijeno",   val: myApps.filter(a => a.status === "rejected").length, color: "#FEE2E2" },
              ].map((m, i) => (
                <div key={i} style={{ background: m.color, borderRadius: 14, padding: "1rem", textAlign: "center" }}>
                  <div style={{ fontSize: 26, fontWeight: 800, color: "#1e293b" }}>{m.val}</div>
                  <div style={{ fontSize: 12, color: "#475569", fontWeight: 500 }}>{m.label}</div>
                </div>
              ))}
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "minmax(0,1fr) minmax(0,330px)", gap: 22, alignItems: "start" }}>
              <div>
                <h2 style={{ fontSize: 16, fontWeight: 700, color: "#0C4A6E", marginBottom: 12 }}>Moje prijave</h2>
                {myApps.length === 0 ? (
                  <div style={{ ...s.card, textAlign: "center", padding: "2rem", color: "#94A3B8" }}>
                    <div style={{ fontSize: 34 }}>🐾</div>
                    <p style={{ marginTop: 10 }}>Još uvek niste prijavili nijednog psa.</p>
                  </div>
                ) : myApps.map(app => (
                  <div key={app.id} style={{ ...s.card, display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12, flexWrap: "wrap", marginBottom: 10 }}>
                    <div>
                      <div style={{ fontWeight: 700, fontSize: 15 }}>🐕 {app.dog_name}</div>
                      <div style={{ fontSize: 13, color: "#64748B", marginTop: 3 }}>{app.breed} · {app.age} · {app.category}</div>
                      <div style={{ fontSize: 12, color: "#94A3B8", marginTop: 2 }}>Mikročip: {app.microchip}</div>
                      {app.pedigree && <div style={{ fontSize: 12, color: "#94A3B8" }}>Rodovnik: {app.pedigree}</div>}
                    </div>
                    <span style={s.badge(app.status)}>{STATUS_LABEL[app.status]}</span>
                  </div>
                ))}
              </div>
              <div style={s.card}>
                <h2 style={{ fontSize: 16, fontWeight: 700, color: "#0C4A6E", marginBottom: 14 }}>➕ Prijavi psa</h2>
                <div style={{ background: "#FEF3C7", border: "1px solid #F59E0B", borderRadius: 10, padding: "10px 14px", marginBottom: 16 }}>
                  <div style={{ fontWeight: 700, color: "#92400E", fontSize: 13, marginBottom: 4 }}>💳 Pre slanja prijave</div>
                  <p style={{ color: "#78350F", fontSize: 12, lineHeight: 1.6, margin: 0 }}>
                    Uplatite kotizaciju i kao <strong>poziv na broj</strong> navedite broj mikročipa psa.
                    Prijava će biti odobrena tek nakon potvrde uplate.
                  </p>
                </div>
                {[
                  { label: "Ime psa *", key: "dogName", placeholder: "Npr. Max" },
                  { label: "Rasa *", key: "breed", placeholder: "Npr. Labrador" },
                ].map(f => (
                  <div key={f.key} style={s.fg}>
                    <label style={s.label}>{f.label}</label>
                    <input style={s.input} placeholder={f.placeholder} value={dogForm[f.key]} onChange={e => setDogForm({ ...dogForm, [f.key]: e.target.value })} />
                  </div>
                ))}
                <div style={s.fg}>
                  <label style={s.label}>Starost *</label>
                  <div style={{ display: "flex", gap: 8 }}>
                    <input style={{ ...s.input, flex: 1 }} type="number" min="1" placeholder="Npr. 18" value={dogForm.age} onChange={e => setDogForm({ ...dogForm, age: e.target.value })} />
                    <select style={{ ...s.select, width: "auto" }} value={dogForm.ageUnit} onChange={e => setDogForm({ ...dogForm, ageUnit: e.target.value })}>
                      <option value="meseci">meseci</option>
                      <option value="godina">godina</option>
                    </select>
                  </div>
                </div>
                <div style={s.fg}>
                  <label style={s.label}>Broj mikročipa * <span style={{ fontWeight: 400, color: "#94A3B8", fontSize: 12 }}>(15 cifara)</span></label>
                  <input style={s.input} placeholder="Npr. 688038000123456" maxLength={15} value={dogForm.microchip} onChange={e => setDogForm({ ...dogForm, microchip: e.target.value.replace(/\D/g, "") })} />
                </div>
                <div style={s.fg}>
                  <label style={s.label}>Broj rodovnika (opciono)</label>
                  <input style={s.input} placeholder="Npr. SRB-12345" value={dogForm.pedigree} onChange={e => setDogForm({ ...dogForm, pedigree: e.target.value })} />
                </div>
                <div style={s.fg}>
                  <label style={s.label}>Kategorija</label>
                  <select style={s.select} value={dogForm.category} onChange={e => setDogForm({ ...dogForm, category: e.target.value })}>
                    {CATEGORIES.map(c => <option key={c}>{c}</option>)}
                  </select>
                </div>
                {dogError   && <div style={{ background: "#FEE2E2", color: "#991B1B", borderRadius: 8, padding: "8px 12px", fontSize: 13, marginBottom: 12 }}>{dogError}</div>}
                {dogSuccess && <div style={{ background: "#D1FAE5", color: "#065F46", borderRadius: 8, padding: "8px 12px", fontSize: 13, marginBottom: 12 }}>{dogSuccess}</div>}
                <button style={{ ...s.btn(), width: "100%", opacity: dogLoading ? 0.7 : 1 }} onClick={handleDogSubmit} disabled={dogLoading}>
                  {dogLoading ? "Slanje..." : "Pošalji prijavu"}
                </button>
              </div>
            </div>
          </div>
        )
      )}

      {/* ── ADMIN PANEL ────────────────────────────────────────────────── */}
      {page === "admin" && currentUser && userProfile?.role === "admin" && (
        <div style={s.page}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12, marginBottom: 24 }}>
            <div>
              <h1 style={s.sectionTitle}>🔧 Admin panel</h1>
              {hasComp ? (
                <p style={{ color: "#64748B", fontSize: 14 }}>
                  {competition.name} · {competition.city} · {fmtDate(competition.date)}
                  {competition.deadline && <span style={{ marginLeft: 10, color: daysToDeadline < 0 ? "#DC2626" : daysToDeadline <= 7 ? "#D97706" : "#059669", fontWeight: 600 }}>
                    · Rok: {fmtDate(competition.deadline)} {daysToDeadline >= 0 ? `(još ${daysToDeadline} dana)` : "(istekao)"}
                  </span>}
                </p>
              ) : (
                <p style={{ color: "#B45309", fontSize: 14, fontWeight: 600 }}>⏸️ Nema aktivnog takmičenja.</p>
              )}
            </div>
            <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
              {!hasComp ? (
                <button style={{ ...s.btn("warning"), display: "flex", alignItems: "center", gap: 6 }} onClick={openNewComp}>➕ Kreiraj takmičenje</button>
              ) : (
                <>
                  <button style={{ ...s.btn(regOpen ? "warning" : "success"), display: "flex", alignItems: "center", gap: 6 }} onClick={toggleRegistration}>
                    {regOpen ? "🔒 Zaključaj prijave" : "🔓 Otvori prijave"}
                  </button>
                  <button style={{ ...s.btn("success"), display: "flex", alignItems: "center", gap: 6, opacity: exporting ? 0.7 : 1 }} onClick={() => handleExportXLSX()} disabled={exporting}>
                    {exporting ? "⏳ Izvoz..." : "📥 Izvezi odobrene"}
                  </button>
                  <button style={{ ...s.btn("danger"), display: "flex", alignItems: "center", gap: 6 }} onClick={() => setArchiveModal(true)}>
                    🗄️ Arhiviraj i resetuj
                  </button>
                </>
              )}
            </div>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(120px,1fr))", gap: 12, marginBottom: 22 }}>
            {[
              { label: "Ukupno",     val: applications.length,                                        color: "#DBEAFE" },
              { label: "Odobreno",   val: applications.filter(a => a.status === "approved").length,   color: "#D1FAE5" },
              { label: "Na čekanju", val: applications.filter(a => a.status === "pending").length,    color: "#FEF3C7" },
              { label: "Odbijeno",   val: applications.filter(a => a.status === "rejected").length,   color: "#FEE2E2" },
            ].map((m, i) => (
              <div key={i} style={{ background: m.color, borderRadius: 14, padding: "1rem", textAlign: "center" }}>
                <div style={{ fontSize: 26, fontWeight: 800, color: "#1e293b" }}>{m.val}</div>
                <div style={{ fontSize: 12, color: "#475569", fontWeight: 500 }}>{m.label}</div>
              </div>
            ))}
          </div>

          <div style={{ ...s.card, overflowX: "auto", marginBottom: 28 }}>
            <table style={s.table}>
              <thead>
                <tr>{["Vlasnik","Ime psa","Rasa","Starost","Mikročip","Kategorija","Status","Akcije"].map(h => (
                  <th key={h} style={s.th}>{h}</th>
                ))}</tr>
              </thead>
              <tbody>
                {applications.map(app => (
                  <tr key={app.id} style={{ background: app.status === "approved" ? "#F0FDF4" : app.status === "rejected" ? "#FFF5F5" : "white" }}>
                    <td style={s.td}><span style={{ fontWeight: 600 }}>{app.owner_name}</span></td>
                    <td style={s.td}>
                      <span style={{ fontWeight: 600 }}>🐕 {app.dog_name}</span>
                      {app.pedigree && <div style={{ fontSize: 11, color: "#94A3B8", marginTop: 2 }}>Rodovnik: {app.pedigree}</div>}
                    </td>
                    <td style={s.td}>{app.breed}</td>
                    <td style={s.td}>{app.age}</td>
                    <td style={s.td}><span style={{ fontFamily: "monospace", fontSize: 12, background: "#F1F5F9", borderRadius: 6, padding: "2px 8px" }}>{app.microchip}</span></td>
                    <td style={s.td}><span style={{ background: "#EFF6FF", color: "#1D4ED8", borderRadius: 20, padding: "2px 10px", fontSize: 12, fontWeight: 600 }}>{app.category}</span></td>
                    <td style={s.td}><span style={s.badge(app.status)}>{STATUS_LABEL[app.status]}</span></td>
                    <td style={s.td}>
                      <div style={{ display: "flex", gap: 6 }}>
                        <button style={{ ...s.btn("success"), padding: "5px 12px", fontSize: 12, opacity: app.status === "approved" ? 0.4 : 1 }}
                          onClick={() => handleStatusChange(app.id, "approved")} disabled={app.status === "approved"}>✓ Odobri</button>
                        <button style={{ ...s.btn("danger"), padding: "5px 12px", fontSize: 12, opacity: app.status === "rejected" ? 0.4 : 1 }}
                          onClick={() => handleStatusChange(app.id, "rejected")} disabled={app.status === "rejected"}>✕ Odbij</button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            {applications.length === 0 && (
              <div style={{ textAlign: "center", padding: "2.5rem", color: "#94A3B8" }}>
                <div style={{ fontSize: 34, marginBottom: 8 }}>{hasComp ? "🐾" : "⏸️"}</div>
                {hasComp ? "Nema prijava za ovo takmičenje." : "Kreirajte takmičenje da biste otvorili prijave."}
              </div>
            )}
          </div>

          <div style={{ marginBottom: 12 }}>
            <h2 style={{ fontSize: 17, fontWeight: 700, color: "#0C4A6E", marginBottom: 4 }}>👥 Upravljanje korisnicima</h2>
            <p style={{ fontSize: 13, color: "#64748B", marginBottom: 16 }}>Dodelite ili oduzimajte admin privilegije registrovanim korisnicima.</p>
          </div>
          <div style={{ ...s.card, overflowX: "auto" }}>
            <table style={s.table}>
              <thead>
                <tr>{["Ime i prezime","Email","Uloga","Akcija"].map(h => <th key={h} style={s.th}>{h}</th>)}</tr>
              </thead>
              <tbody>
                {users.map(u => {
                  const isSelf  = u.id === currentUser.id;
                  const isAdmin = u.role === "admin";
                  return (
                    <tr key={u.id} style={{ background: isAdmin ? "#F0FDF4" : "white" }}>
                      <td style={s.td}>
                        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                          <div style={{ width: 32, height: 32, borderRadius: "50%", background: isAdmin ? "#D1FAE5" : "#EFF6FF", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, fontWeight: 700, color: isAdmin ? "#065F46" : "#1D4ED8", flexShrink: 0 }}>
                            {u.name?.charAt(0)}
                          </div>
                          <span style={{ fontWeight: 600 }}>{u.name}{isSelf && <span style={{ fontSize: 11, color: "#94A3B8", fontWeight: 400, marginLeft: 6 }}>(vi)</span>}</span>
                        </div>
                      </td>
                      <td style={s.td}>{u.email || "—"}</td>
                      <td style={s.td}>
                        <span style={{ background: isAdmin ? "#D1FAE5" : "#F1F5F9", color: isAdmin ? "#065F46" : "#475569", border: `1px solid ${isAdmin ? "#10B981" : "#CBD5E1"}`, borderRadius: 20, padding: "3px 12px", fontSize: 12, fontWeight: 600 }}>
                          {isAdmin ? "🔧 Admin" : "👤 Korisnik"}
                        </span>
                      </td>
                      <td style={s.td}>
                        {isSelf ? <span style={{ fontSize: 12, color: "#CBD5E1" }}>— nije moguće</span> : (
                          <button style={{ ...s.btn(isAdmin ? "danger" : "primary"), padding: "5px 14px", fontSize: 12 }}
                            onClick={() => handleRoleChange(u.id, isAdmin ? "user" : "admin", u.name)}>
                            {isAdmin ? "↓ Ukloni admina" : "↑ Promoviši u admina"}
                          </button>
                        )}
                      </td>
                    </tr>
                  );
                })}
                {users.length === 0 && (
                  <tr><td colSpan={4} style={{ ...s.td, textAlign: "center", color: "#94A3B8", padding: "2rem" }}>Nema registrovanih korisnika.</td></tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* ── MODAL: Arhiviranje ──────────────────────────────────────────── */}
      {archiveModal && (
        <div style={s.modal}>
          <div style={s.modalBox}>
            <div style={{ textAlign: "center", marginBottom: 18 }}>
              <div style={{ fontSize: 46, marginBottom: 8 }}>🗄️</div>
              <h2 style={{ fontSize: 18, fontWeight: 700, color: "#0C4A6E", marginBottom: 8 }}>Arhiviranje takmičenja</h2>
              <p style={{ color: "#64748B", fontSize: 13, lineHeight: 1.6 }}>
                Odobrene prijave za <strong>„{competition?.name}"</strong> biće izvezene u Excel,
                a zatim <strong>sve prijave i takmičenje trajno obrisani</strong> iz baze.
              </p>
            </div>
            <div style={{ background: "#F8FAFC", borderRadius: 12, padding: "0.9rem", marginBottom: 12, border: "1px solid #e2e8f0" }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: "#374151", marginBottom: 10 }}>Pregled podataka:</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                {[
                  { label: "Ukupno",         val: applications.length,                                        color: "#DBEAFE" },
                  { label: "Odobrenih",       val: applications.filter(a => a.status === "approved").length,   color: "#D1FAE5" },
                  { label: "Na čekanju",      val: applications.filter(a => a.status === "pending").length,    color: "#FEF3C7" },
                  { label: "Odbijenih",       val: applications.filter(a => a.status === "rejected").length,   color: "#FEE2E2" },
                ].map((m, i) => (
                  <div key={i} style={{ background: m.color, borderRadius: 8, padding: "7px 12px", display: "flex", justifyContent: "space-between" }}>
                    <span style={{ fontSize: 12, color: "#475569" }}>{m.label}</span>
                    <span style={{ fontWeight: 700 }}>{m.val}</span>
                  </div>
                ))}
              </div>
            </div>
            <div style={{ background: "#FEF3C7", border: "1px solid #F59E0B", borderRadius: 10, padding: "9px 14px", marginBottom: 18, fontSize: 13, color: "#92400E" }}>
              ⚠️ Samo <strong>odobrene</strong> prijave ulaze u Excel arhivu. Nakon arhiviranja sistem prelazi u pauzirano stanje.
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button style={{ ...s.btn("outline"), flex: 1, padding: "11px" }} onClick={() => setArchiveModal(false)}>← Odustani</button>
              <button style={{ ...s.btn("danger"), flex: 1, padding: "11px", fontWeight: 700 }} onClick={handleArchiveAndReset}>✓ Da, arhiviraj i resetuj</button>
            </div>
          </div>
        </div>
      )}

      {/* ── MODAL: Novo takmičenje ──────────────────────────────────────── */}
      {newCompModal && (
        <div style={s.modal}>
          <div style={s.modalBox}>
            <div style={{ textAlign: "center", marginBottom: 20 }}>
              <div style={{ fontSize: 46, marginBottom: 8 }}>🏆</div>
              <h2 style={{ fontSize: 18, fontWeight: 700, color: "#0C4A6E", marginBottom: 6 }}>Kreirajte novo takmičenje</h2>
              <p style={{ color: "#64748B", fontSize: 13 }}>Nakon kreiranja, prijave će biti automatski otvorene.</p>
            </div>
            {[
              { label: "Naziv takmičenja *", key: "name", placeholder: "Npr. Jesenji kup Novog Sada", type: "text" },
              { label: "Datum održavanja *", key: "date", placeholder: "", type: "date" },
              { label: "Mesto/grad *",       key: "city", placeholder: "Npr. Novi Sad", type: "text" },
            ].map(f => (
              <div key={f.key} style={s.fg}>
                <label style={s.label}>{f.label}</label>
                <input style={s.input} type={f.type} placeholder={f.placeholder} value={newCompForm[f.key]}
                  onChange={e => setNewCompForm({ ...newCompForm, [f.key]: e.target.value })} />
              </div>
            ))}
            <div style={s.fg}>
              <label style={s.label}>Rok za prijave <span style={{ fontWeight: 400, color: "#94A3B8", fontSize: 12 }}>(opciono — automatsko zaključavanje)</span></label>
              <input style={s.input} type="date" value={newCompForm.deadline}
                onChange={e => setNewCompForm({ ...newCompForm, deadline: e.target.value })} />
            </div>
            {newCompError && <div style={{ background: "#FEE2E2", color: "#991B1B", borderRadius: 8, padding: "8px 12px", fontSize: 13, marginBottom: 12 }}>{newCompError}</div>}
            <div style={{ display: "flex", gap: 10 }}>
              <button style={{ ...s.btn("outline"), flex: 1 }} onClick={() => setNewCompModal(false)}>Odustani</button>
              <button style={{ ...s.btn(), flex: 1, padding: "12px" }} onClick={handleCreateNewComp}>✓ Kreiraj i otvori prijave</button>
            </div>
          </div>
        </div>
      )}

      {/* ── MODAL: Uredi takmičenje ─────────────────────────────────────── */}
      {editCompModal && (
        <div style={s.modal}>
          <div style={s.modalBox}>
            <div style={{ marginBottom: 18 }}>
              <h2 style={{ fontSize: 18, fontWeight: 700, color: "#0C4A6E", marginBottom: 4 }}>✏️ Uredi takmičenje</h2>
              <p style={{ color: "#64748B", fontSize: 13 }}>Ispravite naziv, datum, mesto ili rok prijava.</p>
            </div>
            {[
              { label: "Naziv takmičenja *", key: "name", type: "text",  placeholder: "" },
              { label: "Datum održavanja *", key: "date", type: "date",  placeholder: "" },
              { label: "Mesto/grad *",       key: "city", type: "text",  placeholder: "" },
            ].map(f => (
              <div key={f.key} style={s.fg}>
                <label style={s.label}>{f.label}</label>
                <input style={s.input} type={f.type} value={editCompForm[f.key]} onChange={e => setEditCompForm({ ...editCompForm, [f.key]: e.target.value })} />
              </div>
            ))}
            <div style={s.fg}>
              <label style={s.label}>Rok za prijave <span style={{ fontWeight: 400, color: "#94A3B8", fontSize: 12 }}>(opciono)</span></label>
              <input style={s.input} type="date" value={editCompForm.deadline} onChange={e => setEditCompForm({ ...editCompForm, deadline: e.target.value })} />
            </div>
            <div style={{ display: "flex", gap: 10 }}>
              <button style={{ ...s.btn("outline"), flex: 1 }} onClick={() => setEditCompModal(false)}>Odustani</button>
              <button style={{ ...s.btn(), flex: 1 }} onClick={saveEditComp}>Sačuvaj izmene</button>
            </div>
          </div>
        </div>
  )}
      {/* ── FOOTER ─────────────────────────────────────────────────────── */}
      <footer style={{ background: "#0C4A6E", color: "rgba(255,255,255,0.7)", textAlign: "center", padding: "1.5rem", fontSize: 13, marginTop: "auto" }}>
        <p style={{ margin: 0 }}>
          © {new Date().getFullYear()} KinološkiKup — Sva prava zadržana
        </p>
        <p style={{ margin: "6px 0 0", fontSize: 12 }}>
          Fotografije: <a href="https://unsplash.com" target="_blank" rel="noopener noreferrer" style={{ color: "rgba(255,255,255,0.5)", textDecoration: "none" }}>Unsplash</a>
        </p>
      </footer>
    </div>
  );
}