import { useState, useEffect, useCallback } from "react";

// ─────────────────────────────────────────────────────────────────
// CONSTANTS & HELPERS
// ─────────────────────────────────────────────────────────────────
const MANAGER_CODE = "LG-MGR-2026"; // Single shared code for all managers
const COMPANY_NAME = "Landis+Gyr";
const COMPANY_TAGLINE = "manage energy better";

// ─────────────────────────────────────────────────────────────────
// CONFIGURATION — Fill in after Azure App setup
// See GRAPH-SETUP-GUIDE.html and SHAREPOINT-SETUP-GUIDE.html
// ─────────────────────────────────────────────────────────────────
const GRAPH_CONFIG = {
  clientId:     "",   // Azure App (client) ID
  tenantId:     "",   // Azure Directory (tenant) ID
  clientSecret: "",   // Azure client secret
  senderEmail:  "edp-noreply@landis-gyr.com",  // Shared mailbox
  // SharePoint config
  siteUrl:      "",   // e.g. https://landis-gyr.sharepoint.com/sites/HR
  sitePath:     "",   // e.g. /sites/HR
};

const isGraphConfigured   = () => !!(GRAPH_CONFIG.clientId && GRAPH_CONFIG.tenantId && GRAPH_CONFIG.clientSecret);
const isSharePointEnabled = () => isGraphConfigured() && !!(GRAPH_CONFIG.siteUrl && GRAPH_CONFIG.sitePath);

// ─────────────────────────────────────────────────────────────────
// OTP STORE — in-memory, expires after 10 minutes
// ─────────────────────────────────────────────────────────────────
const OTP_STORE = {}; // { email: { code, expiry, userData } }
const OTP_EXPIRY_MS = 10 * 60 * 1000; // 10 minutes

function generateOTP() {
  return String(Math.floor(100000 + Math.random() * 900000));
}

function storeOTP(email, code, userData) {
  OTP_STORE[email.toLowerCase()] = { code, expiry: Date.now() + OTP_EXPIRY_MS, userData };
}

function verifyOTP(email, code) {
  const entry = OTP_STORE[email.toLowerCase()];
  if (!entry) return { valid: false, reason: "No code found. Please request a new one." };
  if (Date.now() > entry.expiry) {
    delete OTP_STORE[email.toLowerCase()];
    return { valid: false, reason: "Code has expired. Please request a new one." };
  }
  if (entry.code !== code.trim()) return { valid: false, reason: "Incorrect code. Please try again." };
  const userData = entry.userData;
  delete OTP_STORE[email.toLowerCase()];
  return { valid: true, userData };
}

async function sendOTPEmail(email, name, code, isNewAccount) {
  const subject = `[EDP 2026] Your verification code: ${code}`;
  const bodyHtml = `
    <p style="color:#374151">Dear <strong>${name || "User"}</strong>,</p>
    <p style="color:#374151">${isNewAccount ? "Your EDP account has been created." : "Here is your sign-in verification code."} Please enter it in the app within 10 minutes.</p>
    <div style="background:#4a4a4a;border-radius:12px;padding:24px;text-align:center;margin:24px 0">
      <div style="font-size:11px;color:#78be20;letter-spacing:3px;text-transform:uppercase;margin-bottom:8px">Your verification code</div>
      <div style="font-size:42px;font-weight:900;color:#fff;letter-spacing:10px;font-family:monospace">${code}</div>
      <div style="font-size:12px;color:rgba(255,255,255,0.5);margin-top:8px">Expires in 10 minutes</div>
    </div>
    <p style="color:#6b7280;font-size:13px">If you did not request this code, please ignore this email.</p>
    <p style="color:#6b7280;font-size:13px">Best regards,<br><strong>EDP Portal — Landis+Gyr</strong></p>`;

  if (isGraphConfigured()) {
    const html = htmlEmail("Verification Code", name, bodyHtml);
    await sendGraphEmail({ to: email, subject, body: html });
  } else {
    // Fallback: show code in alert (dev/local mode)
    setTimeout(() => alert("[DEV MODE — Graph not configured]\n\nOTP code for " + email + ":\n\n" + code + "\n\nIn production this would be sent by email."), 100);
  }
}

const EMPTY_FORM = {
  salarieNom: "", salarieSociete: "", salariePoste: "",
  responsableNom: "", dateEntretien: "", lieuEntretien: "",
  raisonNonRealisation: "", service: "",
  niveauSatisfaction: "", commentairesSatisfaction: "",
  meilleureRealisation: "", momentsDifficiles: "",
  formations: [{ intitule: "", certifiante: "", domaine: "", utiliteCollab: "", commentairesCollab: "", commentairesManager: "" }],
  rappelMissions: "",
  savoirFaire: [{ competence: "", niveau: "" }],
  savoirEtre: [{ competence: "", niveau: "" }],
  appreciationGlobale: "", pointsProgres: "", competencesNonUtilisees: "",
  evolutionMissions: "", remarqueEvolution: "",
  avenirProfessionnel: "", remarqueAvenir: "",
  besoinsFormation: [{ besoin: "", objectif: "", avisManager: "" }],
  bilanCompetences: "", bilanDelai: "", vae: "", vaeDelai: "",
  commentairesSalarie: "", commentairesResponsable: "",
};

const REQUIRED = ["salarieNom", "salarieSociete", "salariePoste", "responsableNom", "dateEntretien", "service", "niveauSatisfaction", "rappelMissions"];

const REQUIRED_LABELS = {
  salarieNom: "Employee Full Name",
  salarieSociete: "Company",
  salariePoste: "Job Title",
  responsableNom: "Interviewer Name",
  dateEntretien: "Interview Date",
  service: "Department",
  niveauSatisfaction: "Satisfaction Level",
  rappelMissions: "Main Missions"
};

const STATUS_CFG = {
  draft:     { bg: "#f3f4f6", color: "#374151", dot: "#9ca3af",  label: "Draft" },
  submitted: { bg: "#dbeafe", color: "#1d4ed8", dot: "#3b82f6",  label: "Pending Review" },
  approved:  { bg: "#dcfce7", color: "#15803d", dot: "#22c55e",  label: "Approved" },
  rejected:  { bg: "#fee2e2", color: "#b91c1c", dot: "#ef4444",  label: "Needs Correction" },
};

function uid() { return Math.random().toString(36).slice(2, 10); }
function now() { return new Date().toISOString(); }
function fmt(iso) { return iso ? new Date(iso).toLocaleString("en-GB", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }) : "—"; }

// ─────────────────────────────────────────────────────────────────
// GRAPH API TOKEN CACHE
// ─────────────────────────────────────────────────────────────────
let _tokenCache = { token: null, expiry: 0 };
async function getGraphToken() {
  if (_tokenCache.token && Date.now() < _tokenCache.expiry) return _tokenCache.token;
  const res = await fetch(
    `https://login.microsoftonline.com/${GRAPH_CONFIG.tenantId}/oauth2/v2.0/token`,
    { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ grant_type: "client_credentials", client_id: GRAPH_CONFIG.clientId,
        client_secret: GRAPH_CONFIG.clientSecret, scope: "https://graph.microsoft.com/.default" }) }
  );
  const d = await res.json();
  _tokenCache = { token: d.access_token, expiry: Date.now() + (d.expires_in - 60) * 1000 };
  return d.access_token;
}

async function graphGet(url) {
  const token = await getGraphToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0${url}`,
    { headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" } });
  return res.json();
}
async function graphPost(url, body) {
  const token = await getGraphToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0${url}`,
    { method: "POST", headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" }, body: JSON.stringify(body) });
  return res.ok ? res.json() : null;
}
async function graphPatch(url, body) {
  const token = await getGraphToken();
  await fetch(`https://graph.microsoft.com/v1.0${url}`,
    { method: "PATCH", headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" }, body: JSON.stringify(body) });
}
async function graphDelete(url) {
  const token = await getGraphToken();
  await fetch(`https://graph.microsoft.com/v1.0${url}`,
    { method: "DELETE", headers: { Authorization: `Bearer ${token}` } });
}

// ─────────────────────────────────────────────────────────────────
// SHAREPOINT HELPERS
// ─────────────────────────────────────────────────────────────────
// Encode site path for Graph API
function spBase() { return `/sites/${GRAPH_CONFIG.siteUrl.replace(/^https?:\/\//, "").replace(GRAPH_CONFIG.sitePath, "")}:${GRAPH_CONFIG.sitePath}`; }

// Ensure a SharePoint list exists, create it if not
async function ensureSPList(listName, columns) {
  try {
    await graphGet(`${spBase()}/lists/${listName}`);
  } catch {
    // Create list
    await graphPost(`${spBase()}/lists`, {
      displayName: listName, list: { template: "genericList" }
    });
    // Add columns
    for (const col of columns) {
      await graphPost(`${spBase()}/lists/${listName}/columns`, col);
    }
  }
}

// ── Initialize all 3 SharePoint lists ────────────────────────────
// ─────────────────────────────────────────────────────────────────
// SHAREPOINT DOCUMENT UPLOAD
// Uploads approved EDP PDF to SharePoint Documents library
// ─────────────────────────────────────────────────────────────────

// Generate a simple but complete PDF using pure JS (no library needed)
function generateEDPPdf(emp) {
  const f = emp.form || {};
  const genDate = new Date().toLocaleDateString("en-GB", { day: "2-digit", month: "long", year: "numeric" });

  // PDF coordinate helpers
  const lines = [];
  let y = 0;

  // We'll build a minimal but valid PDF structure
  // Using basic PDF text operations: BT/ET blocks, font sizing, page breaks
  const pageH = 841.89; // A4 height in points
  const pageW = 595.28; // A4 width in points
  const marginL = 50;
  const marginR = pageW - 50;
  const contentW = marginR - marginL;
  const lineH = 14;
  const pages = [];
  let currentPageLines = [];
  let cy = pageH - 60; // start from top

  function newPage() {
    pages.push([...currentPageLines]);
    currentPageLines = [];
    cy = pageH - 60;
  }

  function checkPage(needed = 20) {
    if (cy < 60 + needed) newPage();
  }

  function addText(text, x, fontSize, bold, color) {
    const safeText = String(text || "").replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)");
    const fontName = bold ? "Helvetica-Bold" : "Helvetica";
    const r = color ? color[0] / 255 : 0;
    const g = color ? color[1] / 255 : 0;
    const b = color ? color[2] / 255 : 0;
    currentPageLines.push(`BT /${fontName} ${fontSize} Tf ${r.toFixed(3)} ${g.toFixed(3)} ${b.toFixed(3)} rg ${x} ${cy.toFixed(1)} Td (${safeText}) Tj ET`);
  }

  function addLine(x1, y1, x2, y2, r, g, b, width) {
    currentPageLines.push(`${(r/255).toFixed(3)} ${(g/255).toFixed(3)} ${(b/255).toFixed(3)} RG ${width} w ${x1} ${y1.toFixed(1)} m ${x2} ${y2.toFixed(1)} l S`);
  }

  function addRect(x, y, w, h, fr, fg, fb) {
    currentPageLines.push(`${(fr/255).toFixed(3)} ${(fg/255).toFixed(3)} ${(fb/255).toFixed(3)} rg ${x} ${y.toFixed(1)} ${w} ${(-h).toFixed(1)} re f`);
  }

  function heading(title) {
    checkPage(35);
    cy -= 10;
    addRect(marginL, cy + 3, contentW, 22, 74, 74, 74);
    addText(title, marginL + 8, 11, true, [255, 255, 255]);
    cy -= 22;
    addLine(marginL, cy, marginR, cy, 120, 190, 32, 1);
    cy -= 8;
  }

  function row(label, value) {
    const valStr = String(value || "—");
    // Wrap long values
    const maxChars = 70;
    const chunks = [];
    let remaining = valStr;
    while (remaining.length > maxChars) {
      let cut = remaining.lastIndexOf(" ", maxChars);
      if (cut < 20) cut = maxChars;
      chunks.push(remaining.slice(0, cut));
      remaining = remaining.slice(cut).trim();
    }
    chunks.push(remaining);

    const rowH = Math.max(18, chunks.length * lineH + 6);
    checkPage(rowH + 4);

    // Alternating row background
    addRect(marginL, cy + 3, 160, rowH, 243, 244, 246);
    addText(label, marginL + 5, 9, true, [74, 74, 74]);
    chunks.forEach((chunk, i) => {
      if (i > 0) { cy -= lineH; checkPage(lineH); }
      addText(chunk, marginL + 168, 9, false, [30, 30, 30]);
    });
    cy -= rowH;
    // Row border
    addLine(marginL, cy + rowH - rowH, marginR, cy + rowH - rowH, 229, 231, 235, 0.5);
  }

  function paraLabel(txt) {
    checkPage(20);
    cy -= 4;
    addText(txt, marginL, 10, true, [74, 74, 74]);
    cy -= 14;
  }

  // ── Header ──
  addRect(0, pageH, pageW, 50, 74, 74, 74);
  addText("Professional Development Interview — EDP 2026", marginL, 14, true, [255, 255, 255]);
  cy = pageH - 30;
  addText("Landis+Gyr  |  BN4097a  |  " + genDate, marginL, 9, false, [120, 190, 32]);
  cy = pageH - 44;
  addText("Status: " + (STATUS_CFG[emp.status]?.label || emp.status) + "  |  Approved: " + fmt(emp.approvedAt), marginL, 9, false, [200, 200, 200]);
  cy = pageH - 68;

  // Sections
  heading("1. Identification");
  row("Employee Name", f.salarieNom);
  row("Company", f.salarieSociete);
  row("Job Title", f.salariePoste);
  row("Interviewer", f.responsableNom);
  row("Interview Date", f.dateEntretien);
  row("Location", f.lieuEntretien);
  row("Department", f.service);
  if (f.raisonNonRealisation) row("Reason (if any)", f.raisonNonRealisation);

  heading("2. Retrospective");
  row("Satisfaction Level", f.niveauSatisfaction ? f.niveauSatisfaction + " / 4" : "");
  row("Comments on Satisfaction", f.commentairesSatisfaction);
  row("Best Achievement", f.meilleureRealisation);
  row("Most Difficult Moments", f.momentsDifficiles);

  heading("3. Training Review");
  if ((f.formations || []).length === 0) { paraLabel("No training recorded."); }
  (f.formations || []).forEach((tr, i) => {
    paraLabel("Training " + (i + 1));
    row("Title", tr.intitule);
    row("Domain", tr.domaine);
    row("Certified", tr.certifiante);
    row("Usefulness (Employee)", tr.utiliteCollab);
    row("Employee Comments", tr.commentairesCollab);
    row("Manager Comments", tr.commentairesManager);
  });

  heading("4. Missions & Skills");
  row("Main Missions", f.rappelMissions);
  row("Overall Assessment", f.appreciationGlobale);
  row("Key Priorities", f.pointsProgres);
  row("Unused Skills", f.competencesNonUtilisees);
  paraLabel("Technical Skills (Know-How):");
  (f.savoirFaire || []).forEach(s => row(s.competence || "—", s.niveau ? "Level " + s.niveau : ""));
  paraLabel("Soft Skills (Behavioral):");
  (f.savoirEtre || []).forEach(s => row(s.competence || "—", s.niveau ? "Level " + s.niveau : ""));

  heading("5. Perspectives");
  row("Mission Evolution (Employee)", f.evolutionMissions);
  row("Manager Remarks", f.remarqueEvolution);
  row("Career Vision (Employee)", f.avenirProfessionnel);
  row("Manager Remarks", f.remarqueAvenir);
  paraLabel("Training & Development Needs:");
  (f.besoinsFormation || []).forEach(b => row(b.besoin || "—", "Obj: " + (b.objectif || "—") + "  Manager: " + (b.avisManager || "—")));
  row("Skills Assessment (Bilan)", f.bilanCompetences ? f.bilanCompetences + (f.bilanDelai ? " — " + f.bilanDelai : "") : "");
  row("VAE Project", f.vae ? f.vae + (f.vaeDelai ? " — " + f.vaeDelai : "") : "");

  heading("6. Summary & Validation");
  row("Employee Comments", f.commentairesSalarie);
  row("Manager Comments", f.commentairesResponsable);

  // Signature block
  checkPage(80);
  cy -= 20;
  addLine(marginL, cy, marginR, cy, 120, 190, 32, 1);
  cy -= 20;
  addText("Employee Signature: ________________________________", marginL, 10, false, [74, 74, 74]);
  addText("Date: _______________", marginL + 340, 10, false, [74, 74, 74]);
  cy -= 25;
  addText("Manager Signature:  ________________________________", marginL, 10, false, [74, 74, 74]);
  addText("Date: _______________", marginL + 340, 10, false, [74, 74, 74]);
  cy -= 20;
  addText("Generated by Landis+Gyr EDP Portal  |  BN4097a  |  Confidential", marginL, 8, false, [180, 180, 180]);

  // Flush last page
  pages.push([...currentPageLines]);

  // ── Build PDF binary ──
  const enc = s => {
    const bytes = [];
    for (let i = 0; i < s.length; i++) bytes.push(s.charCodeAt(i) & 0xFF);
    return new Uint8Array(bytes);
  };

  const pageCount = pages.length;
  const objects = [];
  const offsets = [];

  function obj(id, content) {
    offsets[id] = objects.reduce((s, o) => s + o.length, 0);
    objects.push(enc(`${id} 0 obj
${content}
endobj
`));
  }

  // Object 1: Catalog
  obj(1, `<< /Type /Catalog /Pages 2 0 R >>`);
  // Object 2: Pages (placeholder, patched below)
  const kidsRef = pages.map((_, i) => `${4 + i} 0 R`).join(" ");
  obj(2, `<< /Type /Pages /Kids [${kidsRef}] /Count ${pageCount} >>`);
  // Object 3: Font
  obj(3, `<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>`);
  // Bold font
  obj(4 + pageCount, `<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>`);

  // Page objects (starting at id 4)
  pages.forEach((pageLines, i) => {
    const streamContent = pageLines.join("\n");
    const streamBytes = enc(streamContent);
    const streamLen = streamBytes.length;
    const pageId = 4 + i;
    const contentId = 5 + pageCount + i;
    obj(pageId, `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pageW.toFixed(2)} ${pageH.toFixed(2)}] /Contents ${contentId} 0 R /Resources << /Font << /Helvetica 3 0 R /Helvetica-Bold ${4 + pageCount} 0 R >> >> >>`);
    obj(contentId, `<< /Length ${streamLen} >>\nstream\n${streamContent}\nendstream`);
  });

  // XRef table
  const xrefOffset = objects.reduce((s, o) => s + o.length, 0);
  const totalObjs = objects.length + 1;
  let xref = `xref
0 ${totalObjs}
0000000000 65535 f 
`;
  let runningOffset = 0;
  for (let i = 0; i < objects.length; i++) {
    xref += runningOffset.toString().padStart(10, "0") + " 00000 n \n";

    runningOffset += objects[i].length;
  }
  const trailer = `trailer
<< /Size ${totalObjs} /Root 1 0 R >>
startxref
${xrefOffset}
%%EOF`;

  const header = enc(`%PDF-1.4
%âãÏÓ
`);
  const allParts = [header, ...objects, enc(xref + trailer)];
  const totalLen = allParts.reduce((s, a) => s + a.length, 0);
  const out = new Uint8Array(totalLen);
  let off = 0;
  for (const p of allParts) { out.set(p, off); off += p.length; }
  return new Blob([out], { type: "application/pdf" });
}

// Upload PDF blob to SharePoint Documents library
async function uploadPDFToSharePoint(emp, pdfBlob) {
  if (!isSharePointEnabled()) return { ok: false, reason: "SharePoint not configured" };
  try {
    const token = await getGraphToken();
    const fileName = `EDP_${String(emp.name || "Employee").replace(/[^a-zA-Z0-9_-]/g, "_")}_${new Date().getFullYear()}_Approved.pdf`;
    const siteId = GRAPH_CONFIG.siteUrl.replace(/^https?:\/\//, "").replace(GRAPH_CONFIG.sitePath, "") + ":" + GRAPH_CONFIG.sitePath;

    // Upload to SharePoint Documents/EDP_Approved folder
    const uploadUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/EDP_Approved/${fileName}:/content`;
    const arrayBuffer = await pdfBlob.arrayBuffer();
    const res = await fetch(uploadUrl, {
      method: "PUT",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/pdf" },
      body: arrayBuffer,
    });
    if (res.ok) {
      const data = await res.json();
      return { ok: true, fileName, webUrl: data.webUrl };
    } else {
      const err = await res.text();
      console.error("[EDP] SP upload failed:", err);
      return { ok: false, reason: "Upload error: " + res.status };
    }
  } catch (e) {
    console.error("[EDP] SP upload error:", e);
    return { ok: false, reason: e.message };
  }
}

async function initSharePointLists() {
  if (!isSharePointEnabled()) return;
  try {
    await ensureSPList("EDP_Employees", [
      { name: "edpId",            text: {} },
      { name: "employeeName",     text: {} },
      { name: "employeeEmail",    text: {} },
      { name: "status",           text: {} },
      { name: "formData",         text: { allowMultipleLines: true } },
      { name: "historyData",      text: { allowMultipleLines: true } },
      { name: "submittedAt",      text: {} },
      { name: "approvedAt",       text: {} },
      { name: "rejectionReason",  text: { allowMultipleLines: true } },
    ]);
    await ensureSPList("EDP_Managers", [
      { name: "edpId",        text: {} },
      { name: "managerName",  text: {} },
      { name: "managerEmail", text: {} },
      { name: "registeredAt", text: {} },
    ]);
    await ensureSPList("EDP_Notifications", [
      { name: "edpId",     text: {} },
      { name: "userId",    text: {} },
      { name: "notifData", text: { allowMultipleLines: true } },
    ]);
    console.log("[EDP] SharePoint lists ready");
  } catch (e) { console.error("[EDP] SharePoint init failed:", e); }
}

// ── Get SharePoint item ID by edpId field ────────────────────────
async function getSPItemId(listName, edpId) {
  try {
    const res = await graphGet(`${spBase()}/lists/${listName}/items?$filter=fields/edpId eq '${edpId}'&$select=id`);
    return res?.value?.[0]?.id || null;
  } catch { return null; }
}

// ─────────────────────────────────────────────────────────────────
// STORAGE LAYER
// Tries SharePoint first, falls back to localStorage automatically
// ─────────────────────────────────────────────────────────────────

// ── localStorage fallback helpers ───────────────────────────────
const LS = {
  get: (key) => { try { return JSON.parse(localStorage.getItem("edp__" + key) || "null"); } catch { return null; } },
  set: (key, val) => { try { localStorage.setItem("edp__" + key, JSON.stringify(val)); } catch {} },
};

// ── EMPLOYEES ────────────────────────────────────────────────────
async function loadEmployees() {
  if (isSharePointEnabled()) {
    try {
      const res = await graphGet(`${spBase()}/lists/EDP_Employees/items?$expand=fields&$top=500`);
      return (res?.value || []).map(item => ({
        ...JSON.parse(item.fields.formData || "{}"),
        id: item.fields.edpId,
        name: item.fields.employeeName,
        email: item.fields.employeeEmail,
        status: item.fields.status,
        form: JSON.parse(item.fields.formData || "{}").form || {},
        history: JSON.parse(item.fields.historyData || "[]"),
        submittedAt: item.fields.submittedAt || null,
        approvedAt: item.fields.approvedAt || null,
        rejectionReason: item.fields.rejectionReason || "",
        _spId: item.id,
      }));
    } catch (e) { console.warn("[EDP] SP load failed, using localStorage:", e); }
  }
  return LS.get("employees") || [];
}

async function saveEmployees(list) {
  LS.set("employees", list); // Always keep local backup
  if (!isSharePointEnabled()) return;
  try {
    for (const emp of list) {
      const spId = emp._spId || await getSPItemId("EDP_Employees", emp.id);
      const fields = {
        edpId: emp.id, employeeName: emp.name, employeeEmail: emp.email,
        status: emp.status,
        formData: JSON.stringify({ form: emp.form }),
        historyData: JSON.stringify(emp.history || []),
        submittedAt: emp.submittedAt || "",
        approvedAt: emp.approvedAt || "",
        rejectionReason: emp.rejectionReason || "",
      };
      if (spId) {
        await graphPatch(`${spBase()}/lists/EDP_Employees/items/${spId}/fields`, fields);
      } else {
        await graphPost(`${spBase()}/lists/EDP_Employees/items`, { fields });
      }
    }
  } catch (e) { console.error("[EDP] SP saveEmployees failed:", e); }
}

async function deleteEmployeeSP(emp) {
  if (!isSharePointEnabled()) return;
  const spId = emp._spId || await getSPItemId("EDP_Employees", emp.id);
  if (spId) await graphDelete(`${spBase()}/lists/EDP_Employees/items/${spId}`);
}

// ── MANAGERS ────────────────────────────────────────────────────
async function loadManagers() {
  if (isSharePointEnabled()) {
    try {
      const res = await graphGet(`${spBase()}/lists/EDP_Managers/items?$expand=fields&$top=200`);
      return (res?.value || []).map(item => ({
        id: item.fields.edpId,
        name: item.fields.managerName,
        email: item.fields.managerEmail,
        registeredAt: item.fields.registeredAt,
        _spId: item.id,
      }));
    } catch (e) { console.warn("[EDP] SP loadManagers failed:", e); }
  }
  return LS.get("managers") || [];
}

async function saveManagers(list) {
  LS.set("managers", list);
  if (!isSharePointEnabled()) return;
  try {
    for (const mgr of list) {
      const spId = mgr._spId || await getSPItemId("EDP_Managers", mgr.id);
      const fields = { edpId: mgr.id, managerName: mgr.name, managerEmail: mgr.email, registeredAt: mgr.registeredAt || "" };
      if (spId) {
        await graphPatch(`${spBase()}/lists/EDP_Managers/items/${spId}/fields`, fields);
      } else {
        await graphPost(`${spBase()}/lists/EDP_Managers/items`, { fields });
      }
    }
  } catch (e) { console.error("[EDP] SP saveManagers failed:", e); }
}

async function deleteManagerSP(mgr) {
  if (!isSharePointEnabled()) return;
  const spId = mgr._spId || await getSPItemId("EDP_Managers", mgr.id);
  if (spId) await graphDelete(`${spBase()}/lists/EDP_Managers/items/${spId}`);
}

// ── NOTIFICATIONS ────────────────────────────────────────────────
async function loadNotifications(userId) {
  if (isSharePointEnabled()) {
    try {
      const res = await graphGet(`${spBase()}/lists/EDP_Notifications/items?$expand=fields&$filter=fields/userId eq '${userId}'&$top=50`);
      const item = res?.value?.[0];
      return item ? JSON.parse(item.fields.notifData || "[]") : [];
    } catch {}
  }
  return LS.get(`notifs_${userId}`) || [];
}

async function saveNotifications(userId, notifs) {
  LS.set(`notifs_${userId}`, notifs);
  if (!isSharePointEnabled()) return;
  try {
    const spId = await getSPItemId("EDP_Notifications", userId);
    const fields = { edpId: userId, userId, notifData: JSON.stringify(notifs) };
    if (spId) {
      await graphPatch(`${spBase()}/lists/EDP_Notifications/items/${spId}/fields`, fields);
    } else {
      await graphPost(`${spBase()}/lists/EDP_Notifications/items`, { fields });
    }
  } catch (e) { console.error("[EDP] SP saveNotifications failed:", e); }
}

async function addNotifForUser(userId, notif) {
  const existing = await loadNotifications(userId);
  const updated = [{ ...notif, id: uid(), ts: now(), read: false }, ...existing];
  await saveNotifications(userId, updated);
}

// ─────────────────────────────────────────────────────────────────
// EMAIL NOTIFICATIONS — Microsoft Graph API
// Sends real emails silently from edp-noreply@landis-gyr.com
// Falls back to Outlook mailto if Graph is not yet configured
// ─────────────────────────────────────────────────────────────────
async function sendGraphEmail({ to, subject, body }) {
  try {
    await graphPost(`/users/${GRAPH_CONFIG.senderEmail}/sendMail`, {
      message: {
        subject,
        body: { contentType: "HTML", content: body },
        toRecipients: [{ emailAddress: { address: to } }],
        from: { emailAddress: { address: GRAPH_CONFIG.senderEmail, name: "Landis+Gyr EDP Portal" } },
      },
      saveToSentItems: false,
    });
    return true;
  } catch (e) {
    console.error("Graph email failed:", e);
    return false;
  }
}

function mailtoFallback({ to, subject, body }) {
  const link = `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
  window.open(link, "_blank");
}

function htmlEmail(title, preheader, bodyHtml) {
  return `<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;background:#f5f5f5;margin:0;padding:20px">
<div style="max-width:560px;margin:0 auto;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.08)">
  <div style="background:#4a4a4a;padding:24px 32px;border-bottom:4px solid #78be20">
    <span style="font-size:22px;font-weight:900;color:#fff">Landis</span><span style="color:#78be20;font-size:22px;font-weight:900">+</span><span style="font-size:22px;font-weight:900;color:#fff">Gyr</span>
    <div style="color:#78be20;font-size:12px;margin-top:2px">manage energy better</div>
  </div>
  <div style="padding:28px 32px">
    <h2 style="color:#4a4a4a;margin:0 0 16px;font-size:20px">${title}</h2>
    ${bodyHtml}
  </div>
  <div style="background:#f5f5f5;padding:16px 32px;font-size:11px;color:#9ca3af;text-align:center">
    Landis+Gyr — EDP 2026 Portal &nbsp;·&nbsp; BN4097a &nbsp;·&nbsp; This is an automated message
  </div>
</div></body></html>`;
}

async function emailEmployeeSubmitted(empName, empEmail, managerEmails) {
  const subject = `[EDP 2026] New submission from ${empName}`;
  const bodyHtml = `<p style="color:#374151">Dear HR Team,</p>
    <p style="color:#374151"><strong>${empName}</strong> (${empEmail}) has submitted their Professional Development Interview form and it is awaiting your review.</p>
    <div style="background:#f0fdf4;border-left:4px solid #78be20;padding:12px 16px;margin:16px 0;border-radius:0 6px 6px 0">
      <strong style="color:#15803d">Action required:</strong> <span style="color:#374151">Please log in to the EDP portal to review and approve or return the form.</span>
    </div>
    <p style="color:#6b7280;font-size:13px">Best regards,<br><strong>EDP Portal — Landis+Gyr</strong></p>`;
  const html = htmlEmail("New EDP Submission", empName, bodyHtml);

  if (isGraphConfigured()) {
    for (const mgrEmail of (managerEmails || [])) {
      await sendGraphEmail({ to: mgrEmail, subject, body: html });
    }
  } else {
    mailtoFallback({ to: (managerEmails||[]).join(",") || "hr@landis-gyr.com", subject,
      body: `Dear HR Team,

${empName} (${empEmail}) has submitted their EDP form for review.

Please log in to the EDP portal.

Regards,
EDP Portal — Landis+Gyr` });
  }
}

async function emailEmployeeApproved(empName, empEmail) {
  const subject = `[EDP 2026] Your form has been approved ✓`;
  const bodyHtml = `<p style="color:#374151">Dear <strong>${empName}</strong>,</p>
    <p style="color:#374151">Great news! Your Professional Development Interview (EDP 2026) has been reviewed and <strong style="color:#15803d">approved</strong> by HR.</p>
    <div style="background:#f0fdf4;border-left:4px solid #78be20;padding:12px 16px;margin:16px 0;border-radius:0 6px 6px 0">
      <strong style="color:#15803d">Next step:</strong> <span style="color:#374151">Log back in to the EDP portal to download your completed document and upload it to SharePoint.</span>
    </div>
    <p style="color:#6b7280;font-size:13px">Best regards,<br><strong>HR Team — Landis+Gyr</strong></p>`;
  const html = htmlEmail("Your EDP Has Been Approved", empName, bodyHtml);

  if (isGraphConfigured()) {
    await sendGraphEmail({ to: empEmail, subject, body: html });
  } else {
    mailtoFallback({ to: empEmail, subject,
      body: `Dear ${empName},

Your EDP 2026 has been approved by HR.

Please log back in to download your document.

Best regards,
HR Team — Landis+Gyr` });
  }
}

async function emailEmployeeRejected(empName, empEmail, reason) {
  const subject = `[EDP 2026] Action required — Your form needs correction`;
  const bodyHtml = `<p style="color:#374151">Dear <strong>${empName}</strong>,</p>
    <p style="color:#374151">Your Professional Development Interview (EDP 2026) requires some corrections before it can be approved.</p>
    <div style="background:#fef2f2;border-left:4px solid #ef4444;padding:12px 16px;margin:16px 0;border-radius:0 6px 6px 0">
      <strong style="color:#b91c1c">Feedback from HR:</strong><br>
      <span style="color:#374151;font-style:italic">"${reason}"</span>
    </div>
    <p style="color:#374151">Please log back into the EDP portal, make the necessary corrections, and resubmit your form.</p>
    <p style="color:#6b7280;font-size:13px">Best regards,<br><strong>HR Team — Landis+Gyr</strong></p>`;
  const html = htmlEmail("EDP Correction Required", empName, bodyHtml);

  if (isGraphConfigured()) {
    await sendGraphEmail({ to: empEmail, subject, body: html });
  } else {
    mailtoFallback({ to: empEmail, subject,
      body: `Dear ${empName},

Your EDP needs corrections.

Feedback: "${reason}"

Please log back in and resubmit.

Best regards,
HR Team — Landis+Gyr` });
  }
}

// ─────────────────────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────────────
// DOCX GENERATION — Pure JS ZIP builder, no external libraries needed
// ─────────────────────────────────────────────────────────────────
function xmlEsc(s) {
  return String(s || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}
function docxPara(text, opts={}) {
  const {bold,size=20,color,center,spaceBefore,spaceAfter,borderBottom}=opts;
  const rpr=[bold?"<w:b/>":"", size!==20?`<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>`:"", color?`<w:color w:val="${color}"/>`:"", `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>`].join("");
  const ppr=[center?`<w:jc w:val="center"/>`:"", (spaceBefore||spaceAfter)?`<w:spacing${spaceBefore?` w:before="${spaceBefore}"`:""}${spaceAfter?` w:after="${spaceAfter}"`:""}/>`:"", borderBottom?`<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="${borderBottom}"/></w:pBdr>`:""].join("");
  const lines=String(text||"").split(/\n/);
  const runs=lines.map((line,i)=>`<w:r><w:rPr>${rpr}</w:rPr><w:t xml:space="preserve">${xmlEsc(line)}</w:t></w:r>${i<lines.length-1?"<w:br/>":""}`).join("");
  return `<w:p>${ppr?`<w:pPr>${ppr}</w:pPr>`:""}<w:r><w:rPr>${rpr}</w:rPr><w:t></w:t></w:r>${runs}</w:p>`;
}
function docxHeading(text) {
  return docxPara(text,{bold:true,size:24,color:"4a4a4a",spaceBefore:"240",spaceAfter:"60"})+docxPara("",{borderBottom:"78be20",spaceAfter:"120"});
}
function docxRow(label,value) {
  const tc=(w,fill,bold,val)=>{
    const lines=String(val||"—").split(/\n/);
    const runs=lines.map((l,i)=>`<w:r><w:rPr>${bold?"<w:b/>":""}<w:sz w:val="18"/><w:szCs w:val="18"/><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t xml:space="preserve">${xmlEsc(l)}</w:t></w:r>${i<lines.length-1?"<w:br/>":""}`).join("");
    return `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${fill?`<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`:""}<w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar></w:tcPr><w:p>${runs}</w:p></w:tc>`;
  };
  return `<w:tr>${tc(3120,"F3F4F6",true,label)}${tc(6240,"",false,value)}</w:tr>`;
}
function docxTable(rows) {
  if(!rows||!rows.length) return "";
  const b=`<w:top w:val="single" w:sz="4" w:space="0" w:color="E5E7EB"/><w:left w:val="single" w:sz="4" w:space="0" w:color="E5E7EB"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="E5E7EB"/><w:right w:val="single" w:sz="4" w:space="0" w:color="E5E7EB"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="E5E7EB"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="E5E7EB"/>`;
  return `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblBorders>${b}</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3120"/><w:gridCol w:w="6240"/></w:tblGrid>${rows.join("")}</w:tbl>`;
}
function sp(){return `<w:p><w:pPr><w:spacing w:after="80"/></w:pPr></w:p>`;}

// Pure JS CRC32
const crcTable=(()=>{const t=new Uint32Array(256);for(let i=0;i<256;i++){let c=i;for(let j=0;j<8;j++)c=c&1?0xEDB88320^(c>>>1):c>>>1;t[i]=c;}return t;})();
function crc32(buf){let c=0xFFFFFFFF;for(let i=0;i<buf.length;i++)c=crcTable[(c^buf[i])&0xFF]^(c>>>8);return(c^0xFFFFFFFF)>>>0;}
function strToBytes(s){return new TextEncoder().encode(s);}
function u16le(n){return[n&0xFF,(n>>8)&0xFF];}
function u32le(n){return[n&0xFF,(n>>8)&0xFF,(n>>16)&0xFF,(n>>24)&0xFF];}
function concatBytes(...arrs){const total=arrs.reduce((s,a)=>s+a.length,0),out=new Uint8Array(total);let off=0;for(const a of arrs){out.set(a,off);off+=a.length;}return out;}

function buildZip(files){
  // files: [{name, data: Uint8Array}]
  const localHeaders=[];
  const centralHeaders=[];
  let offset=0;
  for(const file of files){
    const nameBytes=strToBytes(file.name);
    const crc=crc32(file.data);
    const size=file.data.length;
    // Local file header
    const local=new Uint8Array([
      0x50,0x4B,0x03,0x04, // signature
      0x14,0x00,           // version needed
      0x00,0x00,           // flags
      0x00,0x00,           // compression (stored)
      0x00,0x00,           // mod time
      0x00,0x00,           // mod date
      ...u32le(crc),
      ...u32le(size),
      ...u32le(size),
      ...u16le(nameBytes.length),
      0x00,0x00,           // extra field length
    ]);
    const localEntry=concatBytes(local,nameBytes,file.data);
    localHeaders.push(localEntry);
    // Central directory entry
    const central=new Uint8Array([
      0x50,0x4B,0x01,0x02, // signature
      0x14,0x00,           // version made by
      0x14,0x00,           // version needed
      0x00,0x00,           // flags
      0x00,0x00,           // compression
      0x00,0x00,           // mod time
      0x00,0x00,           // mod date
      ...u32le(crc),
      ...u32le(size),
      ...u32le(size),
      ...u16le(nameBytes.length),
      0x00,0x00,           // extra field length
      0x00,0x00,           // comment length
      0x00,0x00,           // disk start
      0x00,0x00,           // internal attr
      0x00,0x00,0x00,0x00, // external attr
      ...u32le(offset),
    ]);
    centralHeaders.push(concatBytes(central,nameBytes));
    offset+=localEntry.length;
  }
  const centralDir=concatBytes(...centralHeaders);
  const eocd=new Uint8Array([
    0x50,0x4B,0x05,0x06, // signature
    0x00,0x00,           // disk number
    0x00,0x00,           // disk with central dir
    ...u16le(files.length),
    ...u16le(files.length),
    ...u32le(centralDir.length),
    ...u32le(offset),
    0x00,0x00,           // comment length
  ]);
  return concatBytes(...localHeaders,centralDir,eocd);
}

async function generateDocxBlob(emp) {
  const f=emp.form||{};
  const genDate=new Date().toLocaleDateString("en-GB",{day:"2-digit",month:"long",year:"numeric"});

  const formRows=(f.formations||[]).flatMap((tr,i)=>[
    docxRow("Training "+(i+1)+" — Title",tr.intitule),
    docxRow("Domain",tr.domaine),docxRow("Certified",tr.certifiante),
    docxRow("Usefulness (Employee)",tr.utiliteCollab),
    docxRow("Employee Comments",tr.commentairesCollab),
    docxRow("Manager Comments",tr.commentairesManager),
  ]);
  const skillRows=(f.savoirFaire||[]).map(s=>docxRow(s.competence||"—",s.niveau?"Level "+s.niveau:""));
  const softRows=(f.savoirEtre||[]).map(s=>docxRow(s.competence||"—",s.niveau?"Level "+s.niveau:""));
  const trainRows=(f.besoinsFormation||[]).map(b=>docxRow(b.besoin||"—","Obj: "+(b.objectif||"—")+"  Manager: "+(b.avisManager||"—")));

  const bodyXml=[
    docxPara("Professional Development Interview — EDP 2026",{bold:true,size:28,color:"4a4a4a",center:true,spaceAfter:"80"}),
    docxPara("Landis+Gyr  |  BN4097a  |  "+genDate,{size:18,color:"78be20",center:true,spaceAfter:"60"}),
    docxPara("Status: "+(STATUS_CFG[emp.status]?.label||emp.status)+"  |  Submitted: "+fmt(emp.submittedAt),{size:18,color:"9CA3AF",center:true,spaceAfter:"200"}),
    docxHeading("1. Identification"),
    docxTable([docxRow("Employee Name",f.salarieNom),docxRow("Company",f.salarieSociete),docxRow("Job Title",f.salariePoste),docxRow("Interviewer",f.responsableNom),docxRow("Interview Date",f.dateEntretien),docxRow("Location",f.lieuEntretien),docxRow("Department",f.service),docxRow("Reason (if any)",f.raisonNonRealisation)]),sp(),
    docxHeading("2. Retrospective"),
    docxTable([docxRow("Satisfaction Level",f.niveauSatisfaction?f.niveauSatisfaction+" / 4":""),docxRow("Comments",f.commentairesSatisfaction),docxRow("Best Achievement",f.meilleureRealisation),docxRow("Most Difficult Moments",f.momentsDifficiles)]),sp(),
    docxHeading("3. Training Review"),
    formRows.length?docxTable(formRows):docxPara("No training recorded.",{size:18,color:"9CA3AF"}),sp(),
    docxHeading("4. Missions & Skills"),
    docxTable([docxRow("Main Missions",f.rappelMissions),docxRow("Overall Assessment",f.appreciationGlobale),docxRow("Key Priorities",f.pointsProgres),docxRow("Unused Skills",f.competencesNonUtilisees)]),sp(),
    docxPara("Technical Skills (Know-How):",{bold:true,size:18}),
    skillRows.length?docxTable(skillRows):docxPara("None listed.",{size:18,color:"9CA3AF"}),sp(),
    docxPara("Soft Skills (Behavioral):",{bold:true,size:18}),
    softRows.length?docxTable(softRows):docxPara("None listed.",{size:18,color:"9CA3AF"}),sp(),
    docxHeading("5. Perspectives"),
    docxTable([docxRow("Mission Evolution (Employee)",f.evolutionMissions),docxRow("Manager Remarks",f.remarqueEvolution),docxRow("Career Vision (Employee)",f.avenirProfessionnel),docxRow("Manager Remarks",f.remarqueAvenir)]),sp(),
    docxPara("Training & Development Needs:",{bold:true,size:18}),
    trainRows.length?docxTable(trainRows):docxPara("None listed.",{size:18,color:"9CA3AF"}),sp(),
    docxTable([docxRow("Skills Assessment (Bilan)",f.bilanCompetences?f.bilanCompetences+(f.bilanDelai?" — "+f.bilanDelai:""):""),docxRow("VAE Project",f.vae?f.vae+(f.vaeDelai?" — "+f.vaeDelai:""):"")]),sp(),
    docxHeading("6. Summary & Validation"),
    docxTable([docxRow("Employee Comments",f.commentairesSalarie),docxRow("Manager Comments",f.commentairesResponsable)]),
    emp.rejectionReason?[sp(),docxTable([docxRow("Correction Note",emp.rejectionReason)])].join(""):"",sp(),
    docxPara("Signature — Employee: ________________________________     Date: _______________",{size:18,spaceBefore:"480"}),
    docxPara("Signature — Manager:  ________________________________     Date: _______________",{size:18,spaceBefore:"200"}),
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr>`,
  ].join("");

  const docXml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${bodyXml}</w:body></w:document>`;
  const stylesXml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;
  const docRels=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
  const rootRels=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;
  const contentTypes=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`;

  const enc=s=>strToBytes(s);
  const zipBytes=buildZip([
    {name:"[Content_Types].xml",    data:enc(contentTypes)},
    {name:"_rels/.rels",            data:enc(rootRels)},
    {name:"word/document.xml",      data:enc(docXml)},
    {name:"word/styles.xml",        data:enc(stylesXml)},
    {name:"word/_rels/document.xml.rels", data:enc(docRels)},
  ]);
  return new Blob([zipBytes],{type:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
}

// ─────────────────────────────────────────────────────────────────
// UI COMPONENTS
// ─────────────────────────────────────────────────────────────────
const C = { // Landis+Gyr brand colors
  navy: "#4a4a4a", gold: "#78be20", bg: "#f5f5f5", white: "#fff",
  gray50: "#f9fafb", gray100: "#f3f4f6", gray200: "#e5e7eb",
  gray400: "#9ca3af", gray600: "#4b5563", gray700: "#374151", gray900: "#111827",
  red: "#b91c1c", green: "#5a9e10", blue: "#1d4ed8", amber: "#b45309"
};

const inp = (disabled) => ({
  width: "100%", padding: "9px 12px", borderRadius: 6,
  border: `1.5px solid ${disabled ? C.gray200 : C.gray200}`,
  fontFamily: "'Lato', sans-serif", fontSize: 14, outline: "none",
  background: disabled ? C.gray50 : C.white, color: C.gray900,
  boxSizing: "border-box", transition: "border-color 0.15s",
});
const ta = (disabled) => ({ ...inp(disabled), minHeight: 82, resize: "vertical", lineHeight: 1.6 });

function StatusBadge({ status, size = "sm" }) {
  const c = STATUS_CFG[status] || STATUS_CFG.draft;
  return (
    <span style={{ display: "inline-flex", alignItems: "center", gap: 6,
      background: c.bg, color: c.color,
      padding: size === "lg" ? "6px 14px" : "3px 10px",
      borderRadius: 20, fontSize: size === "lg" ? 14 : 12, fontWeight: 600 }}>
      <span style={{ width: 7, height: 7, borderRadius: "50%", background: c.dot, display: "inline-block" }} />
      {c.label}
    </span>
  );
}

function Toast({ toasts, onDismiss }) {
  return (
    <div style={{ position: "fixed", top: 20, right: 20, zIndex: 9999, display: "flex", flexDirection: "column", gap: 10, maxWidth: 360 }}>
      {toasts.map(t => (
        <div key={t.id} style={{
          background: t.type === "success" ? "#064e3b" : t.type === "error" ? "#7f1d1d" : t.type === "warning" ? "#78350f" : "#1e3a5f",
          color: "#fff", padding: "13px 16px", borderRadius: 10, boxShadow: "0 8px 30px rgba(0,0,0,0.2)",
          display: "flex", gap: 10, fontSize: 13, lineHeight: 1.5,
          animation: "toastIn 0.25s ease"
        }}>
          <span style={{ fontSize: 17, marginTop: 1 }}>
            {t.type === "success" ? "✅" : t.type === "error" ? "❌" : t.type === "warning" ? "⚠️" : "🔔"}
          </span>
          <div style={{ flex: 1 }}>
            <div style={{ fontWeight: 700, marginBottom: 2 }}>{t.title}</div>
            <div style={{ opacity: 0.85 }}>{t.message}</div>
          </div>
          <button onClick={() => onDismiss(t.id)} style={{ background: "none", border: "none", color: "#fff", cursor: "pointer", fontSize: 16, opacity: 0.6, padding: 0, alignSelf: "flex-start" }}>×</button>
        </div>
      ))}
    </div>
  );
}

function Section({ title, children }) {
  return (
    <div style={{ marginBottom: 28 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16 }}>
        <h3 style={{ margin: 0, fontFamily: "'Playfair Display', serif", fontSize: 16, color: C.navy }}>{title}</h3>
        <div style={{ flex: 1, height: 1, background: `linear-gradient(to right, ${C.gold}, transparent)` }} />
      </div>
      {children}
    </div>
  );
}

function Field({ label, required, half, children }) {
  return (
    <div style={{ marginBottom: 14, width: half ? "calc(50% - 8px)" : "100%" }}>
      <label style={{ display: "block", fontSize: 12, color: C.gray600, marginBottom: 5, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5 }}>
        {label}{required && <span style={{ color: C.red, marginLeft: 3 }}>*</span>}
      </label>
      {children}
    </div>
  );
}

function Grid({ children, cols = 2 }) {
  return <div style={{ display: "flex", flexWrap: "wrap", gap: 16 }}>{children}</div>;
}

// ─────────────────────────────────────────────────────────────────
// EDP FORM
// ─────────────────────────────────────────────────────────────────
function EDPForm({ form, onChange, readOnly }) {
  const upd = (k, v) => onChange({ ...form, [k]: v });
  const R = readOnly;

  return (
    <div>
      <Section title="1. Identification">
        <Grid>
          <Field label="Employee Full Name" required half><input style={inp(R)} readOnly={R} value={form.salarieNom} onChange={e => upd("salarieNom", e.target.value)} placeholder="Last name, First name" /></Field>
          <Field label="Company" required half><input style={inp(R)} readOnly={R} value={form.salarieSociete} onChange={e => upd("salarieSociete", e.target.value)} /></Field>
          <Field label="Job Title" required half><input style={inp(R)} readOnly={R} value={form.salariePoste} onChange={e => upd("salariePoste", e.target.value)} /></Field>
          <Field label="Interviewer Name" required half><input style={inp(R)} readOnly={R} value={form.responsableNom} onChange={e => upd("responsableNom", e.target.value)} /></Field>
          <Field label="Interview Date" required half><input type="date" style={inp(R)} readOnly={R} value={form.dateEntretien} onChange={e => upd("dateEntretien", e.target.value)} /></Field>
          <Field label="Location" half><input style={inp(R)} readOnly={R} value={form.lieuEntretien} onChange={e => upd("lieuEntretien", e.target.value)} /></Field>
          <Field label="Department / Service" required half><input style={inp(R)} readOnly={R} value={form.service} onChange={e => upd("service", e.target.value)} /></Field>
          <Field label="Reason (if interview not held on schedule)" half><input style={inp(R)} readOnly={R} value={form.raisonNonRealisation} onChange={e => upd("raisonNonRealisation", e.target.value)} /></Field>
        </Grid>
      </Section>

      <Section title="2. Retrospective — Review of the Past Period">
        <Field label="Satisfaction Level" required>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
            {["1 — Not satisfying", "2 — Slightly satisfying", "3 — Satisfying", "4 — Very satisfying"].map((opt, i) => (
              <label key={i} style={{ display: "flex", alignItems: "center", gap: 7, cursor: R ? "default" : "pointer",
                padding: "7px 14px", border: `2px solid ${form.niveauSatisfaction === String(i+1) ? C.gold : C.gray200}`,
                borderRadius: 8, background: form.niveauSatisfaction === String(i+1) ? "#fef9ec" : C.white,
                fontSize: 13, transition: "all 0.15s" }}>
                <input type="radio" disabled={R} checked={form.niveauSatisfaction === String(i+1)} onChange={() => upd("niveauSatisfaction", String(i+1))} style={{ accentColor: C.gold }} />
                {opt}
              </label>
            ))}
          </div>
        </Field>
        <Grid>
          <Field label="Comments on Satisfaction" half><textarea style={ta(R)} readOnly={R} value={form.commentairesSatisfaction} onChange={e => upd("commentairesSatisfaction", e.target.value)} /></Field>
          <Field label="Best Achievement & Success Factors" half><textarea style={ta(R)} readOnly={R} value={form.meilleureRealisation} onChange={e => upd("meilleureRealisation", e.target.value)} /></Field>
          <Field label="Most Difficult Moments & Why" half><textarea style={ta(R)} readOnly={R} value={form.momentsDifficiles} onChange={e => upd("momentsDifficiles", e.target.value)} /></Field>
        </Grid>
      </Section>

      <Section title="3. Training Review (Past Year)">
        {(form.formations || []).map((f, idx) => (
          <div key={idx} style={{ background: C.gray50, border: `1px solid ${C.gray200}`, borderRadius: 8, padding: 16, marginBottom: 12 }}>
            <div style={{ fontSize: 12, color: C.gray400, fontWeight: 700, marginBottom: 10, textTransform: "uppercase" }}>Training #{idx + 1}</div>
            <Grid>
              <Field label="Training Name" half><input style={inp(R)} readOnly={R} value={f.intitule} onChange={e => { const a=[...form.formations]; a[idx].intitule=e.target.value; upd("formations",a); }} /></Field>
              <Field label="Domain" half><input style={inp(R)} readOnly={R} value={f.domaine} onChange={e => { const a=[...form.formations]; a[idx].domaine=e.target.value; upd("formations",a); }} /></Field>
              <Field label="Certified?" half>
                <select style={inp(R)} disabled={R} value={f.certifiante} onChange={e => { const a=[...form.formations]; a[idx].certifiante=e.target.value; upd("formations",a); }}>
                  <option value="">— Select —</option><option>Yes</option><option>No</option>
                </select>
              </Field>
              <Field label="Usefulness (Employee view)" half>
                <select style={inp(R)} disabled={R} value={f.utiliteCollab} onChange={e => { const a=[...form.formations]; a[idx].utiliteCollab=e.target.value; upd("formations",a); }}>
                  <option value="">— Select —</option><option>Not useful</option><option>Slightly useful</option><option>Partially useful</option><option>Fully useful</option>
                </select>
              </Field>
              <Field label="Employee Comments" half><textarea style={{...ta(R), minHeight:60}} readOnly={R} value={f.commentairesCollab} onChange={e => { const a=[...form.formations]; a[idx].commentairesCollab=e.target.value; upd("formations",a); }} /></Field>
              <Field label="Manager Comments" half><textarea style={{...ta(R), minHeight:60}} readOnly={R} value={f.commentairesManager} onChange={e => { const a=[...form.formations]; a[idx].commentairesManager=e.target.value; upd("formations",a); }} /></Field>
            </Grid>
            {!R && idx > 0 && <button onClick={() => { const a=form.formations.filter((_,i)=>i!==idx); upd("formations",a); }} style={{ fontSize: 12, color: C.red, background: "none", border: "none", cursor: "pointer", padding: 0, marginTop: 4 }}>Remove this training</button>}
          </div>
        ))}
        {!R && <button onClick={() => upd("formations", [...form.formations, { intitule:"", certifiante:"", domaine:"", utiliteCollab:"", commentairesCollab:"", commentairesManager:"" }])}
          style={{ fontSize: 13, color: C.navy, background: "none", border: `1.5px dashed ${C.navy}`, borderRadius: 6, padding: "6px 16px", cursor: "pointer" }}>
          + Add Training
        </button>}
      </Section>

      <Section title="4. Missions & Skills Profile">
        <Field label="Main Missions (filled by employee)" required><textarea style={ta(R)} readOnly={R} value={form.rappelMissions} onChange={e => upd("rappelMissions", e.target.value)} placeholder="Describe your main responsibilities…" /></Field>

        <div style={{ marginBottom: 14 }}>
          <div style={{ fontSize: 12, color: C.gray600, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 8 }}>Technical Skills (Know-How)</div>
          {(form.savoirFaire || []).map((s, idx) => (
            <div key={idx} style={{ display: "flex", gap: 10, marginBottom: 7, alignItems: "center" }}>
              <input style={{...inp(R), flex: 3}} readOnly={R} placeholder="Skill / Technique…" value={s.competence} onChange={e => { const a=[...form.savoirFaire]; a[idx].competence=e.target.value; upd("savoirFaire",a); }} />
              <select style={{...inp(R), flex: 1}} disabled={R} value={s.niveau} onChange={e => { const a=[...form.savoirFaire]; a[idx].niveau=e.target.value; upd("savoirFaire",a); }}>
                <option value="">Level</option><option value="1">1 — Beginner</option><option value="2">2 — Proficient</option><option value="3">3 — Expert</option>
              </select>
              {!R && idx > 0 && <button onClick={() => upd("savoirFaire", form.savoirFaire.filter((_,i)=>i!==idx))} style={{ color: C.red, background: "none", border: "none", cursor: "pointer", fontSize: 18, padding: 0 }}>×</button>}
            </div>
          ))}
          {!R && <button onClick={() => upd("savoirFaire", [...form.savoirFaire, {competence:"",niveau:""}])} style={{ fontSize: 12, color: C.navy, background: "none", border: `1px dashed ${C.navy}`, borderRadius: 5, padding: "4px 12px", cursor: "pointer", marginBottom: 14 }}>+ Row</button>}
        </div>

        <div style={{ marginBottom: 14 }}>
          <div style={{ fontSize: 12, color: C.gray600, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 8 }}>Soft Skills (Behavioral)</div>
          {(form.savoirEtre || []).map((s, idx) => (
            <div key={idx} style={{ display: "flex", gap: 10, marginBottom: 7, alignItems: "center" }}>
              <input style={{...inp(R), flex: 3}} readOnly={R} placeholder="Ability to…" value={s.competence} onChange={e => { const a=[...form.savoirEtre]; a[idx].competence=e.target.value; upd("savoirEtre",a); }} />
              <select style={{...inp(R), flex: 1}} disabled={R} value={s.niveau} onChange={e => { const a=[...form.savoirEtre]; a[idx].niveau=e.target.value; upd("savoirEtre",a); }}>
                <option value="">Level</option><option value="1">1 — Beginner</option><option value="2">2 — Proficient</option><option value="3">3 — Expert</option>
              </select>
              {!R && idx > 0 && <button onClick={() => upd("savoirEtre", form.savoirEtre.filter((_,i)=>i!==idx))} style={{ color: C.red, background: "none", border: "none", cursor: "pointer", fontSize: 18, padding: 0 }}>×</button>}
            </div>
          ))}
          {!R && <button onClick={() => upd("savoirEtre", [...form.savoirEtre, {competence:"",niveau:""}])} style={{ fontSize: 12, color: C.navy, background: "none", border: `1px dashed ${C.navy}`, borderRadius: 5, padding: "4px 12px", cursor: "pointer" }}>+ Row</button>}
        </div>

        <Grid>
          <Field label="Overall Job Mastery Assessment" half><textarea style={ta(R)} readOnly={R} value={form.appreciationGlobale} onChange={e => upd("appreciationGlobale", e.target.value)} /></Field>
          <Field label="Key Priorities / Areas for Progress" half><textarea style={ta(R)} readOnly={R} value={form.pointsProgres} onChange={e => upd("pointsProgres", e.target.value)} /></Field>
          <Field label="Skills not currently used in this role" half><textarea style={{...ta(R), minHeight:60}} readOnly={R} value={form.competencesNonUtilisees} onChange={e => upd("competencesNonUtilisees", e.target.value)} /></Field>
        </Grid>
      </Section>

      <Section title="5. Perspectives">
        <Grid>
          <Field label="How do you see your role evolving? (Employee)" half><textarea style={ta(R)} readOnly={R} value={form.evolutionMissions} onChange={e => upd("evolutionMissions", e.target.value)} /></Field>
          <Field label="Manager Remarks" half><textarea style={ta(R)} readOnly={R} value={form.remarqueEvolution} onChange={e => upd("remarqueEvolution", e.target.value)} /></Field>
          <Field label="Your long-term career vision? (Employee)" half><textarea style={ta(R)} readOnly={R} value={form.avenirProfessionnel} onChange={e => upd("avenirProfessionnel", e.target.value)} /></Field>
          <Field label="Manager Remarks" half><textarea style={ta(R)} readOnly={R} value={form.remarqueAvenir} onChange={e => upd("remarqueAvenir", e.target.value)} /></Field>
        </Grid>

        <div style={{ marginBottom: 14 }}>
          <div style={{ fontSize: 12, color: C.gray600, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 8 }}>Training & Development Needs</div>
          {(form.besoinsFormation || []).map((b, idx) => (
            <div key={idx} style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr auto", gap: 8, marginBottom: 8 }}>
              <input style={inp(R)} readOnly={R} placeholder="Training need…" value={b.besoin} onChange={e => { const a=[...form.besoinsFormation]; a[idx].besoin=e.target.value; upd("besoinsFormation",a); }} />
              <input style={inp(R)} readOnly={R} placeholder="Objective" value={b.objectif} onChange={e => { const a=[...form.besoinsFormation]; a[idx].objectif=e.target.value; upd("besoinsFormation",a); }} />
              <input style={inp(R)} readOnly={R} placeholder="Manager opinion" value={b.avisManager} onChange={e => { const a=[...form.besoinsFormation]; a[idx].avisManager=e.target.value; upd("besoinsFormation",a); }} />
              {!R && idx > 0 && <button onClick={() => upd("besoinsFormation", form.besoinsFormation.filter((_,i)=>i!==idx))} style={{ color: C.red, background: "none", border: "none", cursor: "pointer", fontSize: 18 }}>×</button>}
            </div>
          ))}
          {!R && <button onClick={() => upd("besoinsFormation", [...form.besoinsFormation, {besoin:"",objectif:"",avisManager:""}])} style={{ fontSize: 12, color: C.navy, background: "none", border: `1px dashed ${C.navy}`, borderRadius: 5, padding: "4px 12px", cursor: "pointer" }}>+ Row</button>}
        </div>

        <Grid>
          <div style={{ width: "calc(50% - 8px)" }}>
            <Field label="Skills Assessment Project (Bilan de compétences)?">
              <div style={{ display: "flex", gap: 14 }}>
                {["Yes","No"].map(o => <label key={o} style={{ display:"flex", alignItems:"center", gap:6, cursor: R?"default":"pointer", fontSize:14 }}><input type="radio" disabled={R} checked={form.bilanCompetences===o} onChange={() => upd("bilanCompetences",o)} style={{ accentColor: C.gold }} />{o}</label>)}
              </div>
            </Field>
            {form.bilanCompetences === "Yes" && <Field label="Timeline">
              <div style={{ display: "flex", gap: 14 }}>
                {["Within 1 year","Within 2 years"].map(o => <label key={o} style={{ display:"flex", alignItems:"center", gap:6, cursor: R?"default":"pointer", fontSize:13 }}><input type="radio" disabled={R} checked={form.bilanDelai===o} onChange={() => upd("bilanDelai",o)} style={{ accentColor: C.gold }} />{o}</label>)}
              </div>
            </Field>}
          </div>
          <div style={{ width: "calc(50% - 8px)" }}>
            <Field label="VAE Project (Prior Learning Assessment)?">
              <div style={{ display: "flex", gap: 14 }}>
                {["Yes","No"].map(o => <label key={o} style={{ display:"flex", alignItems:"center", gap:6, cursor: R?"default":"pointer", fontSize:14 }}><input type="radio" disabled={R} checked={form.vae===o} onChange={() => upd("vae",o)} style={{ accentColor: C.gold }} />{o}</label>)}
              </div>
            </Field>
            {form.vae === "Yes" && <Field label="Timeline">
              <div style={{ display: "flex", gap: 14 }}>
                {["Within 1 year","Within 2 years"].map(o => <label key={o} style={{ display:"flex", alignItems:"center", gap:6, cursor: R?"default":"pointer", fontSize:13 }}><input type="radio" disabled={R} checked={form.vaeDelai===o} onChange={() => upd("vaeDelai",o)} style={{ accentColor: C.gold }} />{o}</label>)}
              </div>
            </Field>}
          </div>
        </Grid>
      </Section>

      <Section title="6. Summary">
        <Grid>
          <Field label="Employee Comments" half><textarea style={ta(R)} readOnly={R} value={form.commentairesSalarie} onChange={e => upd("commentairesSalarie", e.target.value)} /></Field>
          <Field label="Manager Comments" half><textarea style={ta(R)} readOnly={R} value={form.commentairesResponsable} onChange={e => upd("commentairesResponsable", e.target.value)} /></Field>
        </Grid>
      </Section>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────
// MAIN APP
// ─────────────────────────────────────────────────────────────────
export default function App() {
  const [screen, setScreen] = useState("landing"); // landing | register | employee | manager | review
  const [currentUser, setCurrentUser] = useState(null); // { id, name, email, role }
  const [employees, setEmployees] = useState([]);
  const [managers, setManagers] = useState([]);
  const [selectedEmpId, setSelectedEmpId] = useState(null);
  const [localForm, setLocalForm] = useState(EMPTY_FORM);
  const [toasts, setToasts] = useState([]);
  const [toastId, setToastId] = useState(0);
  const [notifs, setNotifs] = useState([]);
  const [showNotifs, setShowNotifs] = useState(false);
  const [regName, setRegName] = useState("");
  const [regEmail, setRegEmail] = useState("");
  const [regInviteCode, setRegInviteCode] = useState("");
  const [regError, setRegError] = useState("");
  // OTP / MFA state
  const [otpScreen, setOtpScreen] = useState(false);      // show OTP entry screen
  const [otpEmail, setOtpEmail] = useState("");            // email awaiting verification
  const [otpInput, setOtpInput] = useState("");            // user-typed OTP
  const [otpSending, setOtpSending] = useState(false);    // loading state
  const [otpError, setOtpError] = useState("");            // error message
  const [otpResendTimer, setOtpResendTimer] = useState(0); // countdown seconds
  const [rejectReason, setRejectReason] = useState("");
  const [showRejectModal, setShowRejectModal] = useState(false);
  const [loading, setLoading] = useState(true);
  const [managerFilter, setManagerFilter] = useState("all");
  const [managerSearch, setManagerSearch] = useState("");
  const [managerTab, setManagerTab] = useState("employees"); // employees | admin
  const [confirmDeleteId, setConfirmDeleteId] = useState(null); // employee id to delete
  const [confirmDeleteMgrId, setConfirmDeleteMgrId] = useState(null);

  // Load data on mount — init SharePoint lists first if configured
  useEffect(() => {
    const init = async () => {
      await initSharePointLists();
      const [emps, mgrs] = await Promise.all([loadEmployees(), loadManagers()]);
      setEmployees(emps);
      setManagers(mgrs);
      setLoading(false);
    };
    init();
  }, []);

  // OTP resend countdown
  useEffect(() => {
    if (otpResendTimer <= 0) return;
    const t = setTimeout(() => setOtpResendTimer(s => s - 1), 1000);
    return () => clearTimeout(t);
  }, [otpResendTimer]);

  // Load user notifications when user changes
  useEffect(() => {
    if (currentUser) {
      loadNotifications(currentUser.id).then(setNotifs);
    }
  }, [currentUser]);

  const toast = useCallback((type, title, message) => {
    const id = toastId + 1;
    setToastId(id);
    setToasts(p => [...p, { id, type, title, message }]);
    setTimeout(() => setToasts(p => p.filter(t => t.id !== id)), 5500);
  }, [toastId]);

  const dismissToast = (id) => setToasts(p => p.filter(t => t.id !== id));

  const updateEmployees = async (updated) => {
    setEmployees(updated);
    await saveEmployees(updated);
  };

  const currentEmp = employees.find(e => e.id === currentUser?.id);
  const selectedEmp = employees.find(e => e.id === selectedEmpId);

  // ── REGISTER — Step 1: validate, then send OTP ──
  const handleRegister = async () => {
    setRegError("");
    if (!regName.trim() || !regEmail.trim()) { setRegError("Please fill in both fields."); return; }
    const emailLower = regEmail.trim().toLowerCase();
    const isManagerReg = regInviteCode.trim() === MANAGER_CODE;
    if (employees.find(e => e.email === emailLower)) { setRegError("This email is already registered as an employee."); return; }
    if (managers.find(m => m.email === emailLower)) { setRegError("This email is already registered as a manager."); return; }

    setOtpSending(true);
    const code = generateOTP();
    storeOTP(emailLower, code, {
      type: isManagerReg ? "manager" : "employee",
      name: regName.trim(), email: emailLower,
      inviteCode: regInviteCode.trim(), isNew: true
    });
    await sendOTPEmail(emailLower, regName.trim(), code, true);
    setOtpEmail(emailLower);
    setOtpInput("");
    setOtpError("");
    setOtpResendTimer(60);
    setOtpSending(false);
    setOtpScreen(true);
    toast("info", "Verification code sent", `A 6-digit code has been sent to ${emailLower}`);
  };

  // ── LOGIN — Step 1: find user, send OTP ──
  const handleLogin = async (email) => {
    const emailLower = email.trim().toLowerCase();
    const mgr = managers.find(m => m.email === emailLower);
    const emp = employees.find(e => e.email === emailLower);
    if (!mgr && !emp) { setRegError("Email not found. Please register first."); return; }

    const user = mgr || emp;
    setOtpSending(true);
    const code = generateOTP();
    storeOTP(emailLower, code, {
      type: mgr ? "manager" : "employee",
      name: user.name, email: emailLower, isNew: false
    });
    await sendOTPEmail(emailLower, user.name, code, false);
    setOtpEmail(emailLower);
    setOtpInput("");
    setOtpError("");
    setOtpResendTimer(60);
    setOtpSending(false);
    setOtpScreen(true);
    toast("info", "Verification code sent", `A 6-digit code has been sent to ${emailLower}`);
  };

  // ── OTP VERIFY — Step 2: check code and complete login/register ──
  const handleVerifyOTP = async () => {
    setOtpError("");
    const result = verifyOTP(otpEmail, otpInput);
    if (!result.valid) { setOtpError(result.reason); return; }

    const { type, name, email, isNew } = result.userData;

    if (type === "manager") {
      let mgr = managers.find(m => m.email === email);
      if (isNew) {
        mgr = { id: uid(), name, email, registeredAt: now() };
        const updatedMgrs = [...managers, mgr];
        setManagers(updatedMgrs);
        await saveManagers(updatedMgrs);
      }
      setCurrentUser({ id: mgr.id, name: mgr.name, email: mgr.email, role: "manager" });
      const n = await loadNotifications(mgr.id);
      setNotifs(n);
      setOtpScreen(false);
      setScreen("manager");
      toast("success", isNew ? "Manager account created!" : "Welcome back!", `Signed in as ${mgr.name}.`);
    } else {
      let emp = employees.find(e => e.email === email);
      if (isNew) {
        emp = {
          id: uid(), name, email,
          status: "draft", form: { ...EMPTY_FORM, salarieNom: name },
          history: [{ action: "Registered", date: now(), by: name }],
          submittedAt: null, approvedAt: null, rejectionReason: ""
        };
        await updateEmployees([...employees, emp]);
      }
      setCurrentUser({ id: emp.id, name: emp.name, email: emp.email, role: "employee" });
      setLocalForm({ ...emp.form });
      const n = await loadNotifications(emp.id);
      setNotifs(n);
      setOtpScreen(false);
      setScreen("employee");
      toast("success", isNew ? "Welcome!" : "Welcome back!", `Signed in as ${emp.name}.`);
    }
  };

  // ── RESEND OTP ──
  const handleResendOTP = async () => {
    if (otpResendTimer > 0) return;
    const entry = OTP_STORE[otpEmail];
    const name = entry?.userData?.name || "";
    setOtpSending(true);
    const code = generateOTP();
    if (entry) entry.code = code; // update existing entry
    else storeOTP(otpEmail, code, { type: "employee", name, email: otpEmail, isNew: false });
    await sendOTPEmail(otpEmail, name, code, false);
    setOtpResendTimer(60);
    setOtpSending(false);
    setOtpError("");
    toast("info", "New code sent", `A new verification code has been sent to ${otpEmail}`);
  };



  // ── SAVE DRAFT ──
  const handleSaveDraft = async () => {
    const updated = employees.map(e => e.id === currentUser.id ? { ...e, form: localForm } : e);
    await updateEmployees(updated);
    toast("info", "Draft saved", "Your form has been saved. You can continue later.");
  };

  // ── SUBMIT ──
  const handleSubmit = async () => {
    const missing = REQUIRED.filter(f => !localForm[f]);
    if (missing.length > 0) {
      toast("error", "Incomplete form", `Please fill in: ${missing.map(m => REQUIRED_LABELS[m]).join(", ")}.`);
      return;
    }
    const updated = employees.map(e => e.id === currentUser.id ? {
      ...e, form: localForm, status: "submitted", submittedAt: now(),
      history: [...(e.history||[]), { action: "Submitted", date: now(), by: e.name }]
    } : e);
    await updateEmployees(updated);
    // Notify all managers in-app
    const currentManagers = await loadManagers();
    for (const mgr of currentManagers) {
      await addNotifForUser(mgr.id, { type: "submitted", title: "New EDP Submitted", message: `${currentUser.name} has submitted their EDP form for review.`, empId: currentUser.id });
    }
    const mgrEmails = currentManagers.map(m => m.email);
    toast("success", "Form submitted!", isGraphConfigured() ? "Your EDP has been sent to HR. Email notifications sent automatically." : "Your EDP has been sent to HR. Outlook will open to notify managers.");
    setTimeout(() => emailEmployeeSubmitted(currentUser.name, currentUser.email, mgrEmails), 800);
  };

  // ── APPROVE ──
  const handleApprove = async () => {
    const approvedEmp = { ...selectedEmp, status: "approved", approvedAt: now(),
      history: [...(selectedEmp.history||[]), { action: "Approved", date: now(), by: currentUser?.name || "Manager" }] };
    const updated = employees.map(e => e.id === selectedEmpId ? approvedEmp : e);
    await updateEmployees(updated);
    await addNotifForUser(selectedEmpId, { type: "approved", title: "Your EDP has been approved ✓", message: "HR has reviewed and approved your Professional Development Interview form." });
    setTimeout(() => emailEmployeeApproved(selectedEmp?.name, selectedEmp?.email), 800);
    setScreen("manager");
    toast("success", "EDP Approved", `${selectedEmp?.name}'s form has been approved.`);

    // Upload PDF to SharePoint in background
    if (isSharePointEnabled()) {
      toast("info", "Uploading to SharePoint…", "Generating PDF and uploading to Documents/EDP_Approved…");
      setTimeout(async () => {
        try {
          const pdfBlob = generateEDPPdf(approvedEmp);
          const result = await uploadPDFToSharePoint(approvedEmp, pdfBlob);
          if (result.ok) {
            toast("success", "PDF uploaded to SharePoint ✓",
              `Saved as: ${result.fileName} in EDP_Approved folder.`);
          } else {
            toast("error", "SharePoint upload failed", result.reason + " — Use ⬇ Download to save locally.");
          }
        } catch(e) {
          console.error("[EDP] PDF/upload error:", e);
          toast("error", "PDF upload failed", "Check browser console. Use ⬇ Download as fallback.");
        }
      }, 1000);
    }
  };

  // ── REJECT ──
  const handleReject = async () => {
    if (!rejectReason.trim()) return;
    const updated = employees.map(e => e.id === selectedEmpId ? {
      ...e, status: "rejected", rejectionReason: rejectReason,
      history: [...(e.history||[]), { action: "Rejected", date: now(), by: "Manager", reason: rejectReason }]
    } : e);
    await updateEmployees(updated);
    await addNotifForUser(selectedEmpId, { type: "rejected", title: "Action Required — EDP Needs Correction", message: `Your EDP was returned for correction. Manager's note: "${rejectReason}"` });
    setShowRejectModal(false);
    setRejectReason("");
    toast("info", "EDP Returned", `${selectedEmp?.name}'s form sent back for correction. Opening Outlook to notify them…`);
    setTimeout(() => emailEmployeeRejected(selectedEmp?.name, selectedEmp?.email, rejectReason), 800);
    setScreen("manager");
  };

  // ── RE-EDIT ──
  const handleReEdit = async () => {
    const updated = employees.map(e => e.id === currentUser.id ? {
      ...e, status: "draft", rejectionReason: "",
      history: [...(e.history||[]), { action: "Re-opened for correction", date: now(), by: currentUser.name }]
    } : e);
    await updateEmployees(updated);
    toast("info", "Form re-opened", "You can now edit and resubmit your form.");
  };

  // ── DOWNLOAD DOCX ──
  const handleDownloadPDF = async (emp) => {
    toast("info", "Preparing PDF", "Generating PDF document…");
    try {
      const blob = generateEDPPdf(emp);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `EDP_${(emp.name || "form").replace(/\s+/g, "_")}_2026.pdf`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      toast("success", "PDF downloaded", "Your EDP PDF document is ready.");
    } catch(e) {
      console.error("PDF error:", e);
      toast("error", "PDF failed", "Could not generate PDF.");
    }
  };

  const handleDownload = async (emp) => {
    toast("info", "Preparing document", "Building your Word document…");
    try {
      const blob = await generateDocxBlob(emp);
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `EDP_${(emp.name || "form").replace(/\s+/g, "_")}_2026.docx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      toast("success", "Download ready", "Your EDP Word document has been downloaded.");
    } catch(e) {
      console.error("DOCX error:", e);
      toast("error", "Download failed", "Could not generate document. See browser console for details.");
    }
  };

  // ── MARK NOTIFS READ ──
  const markNotifsRead = async () => {
    if (!currentUser) return;
    const updated = notifs.map(n => ({ ...n, read: true }));
    setNotifs(updated);
    await saveNotifications(currentUser.id, updated);
  };

  const unreadCount = notifs.filter(n => !n.read).length;

  // ── FILTERED EMPLOYEES (manager dashboard) ──
  const filteredEmps = employees.filter(e => {
    const matchStatus = managerFilter === "all" || e.status === managerFilter;
    const matchSearch = !managerSearch || e.name.toLowerCase().includes(managerSearch.toLowerCase()) || e.email.toLowerCase().includes(managerSearch.toLowerCase());
    return matchStatus && matchSearch;
  });

  const stats = {
    total: employees.length,
    draft: employees.filter(e => e.status === "draft").length,
    submitted: employees.filter(e => e.status === "submitted").length,
    approved: employees.filter(e => e.status === "approved").length,
    rejected: employees.filter(e => e.status === "rejected").length,
  };

  // ── DELETE EMPLOYEE ──
  const handleDeleteEmployee = async (id) => {
    const emp = employees.find(e => e.id === id);
    if (emp) await deleteEmployeeSP(emp);
    const updated = employees.filter(e => e.id !== id);
    setEmployees(updated);
    LS.set("employees", updated);
    setConfirmDeleteId(null);
    toast("info", "Record deleted", "Employee record has been removed.");
  };

  // ── DELETE MANAGER ──
  const handleDeleteManager = async (id) => {
    const mgr = managers.find(m => m.id === id);
    if (mgr) await deleteManagerSP(mgr);
    const updated = managers.filter(m => m.id !== id);
    setManagers(updated);
    LS.set("managers", updated);
    setConfirmDeleteMgrId(null);
    toast("info", "Manager removed", "Manager account has been removed.");
  };

  // ── EXPORT CSV ──
  const handleExportCSV = () => {
    const headers = ["Name","Email","Status","Job Title","Department","Company","Interview Date","Submitted At","Approved At","Rejection Reason"];
    const rows = employees.map(e => [
      e.name, e.email, e.status,
      e.form?.salariePoste || "",
      e.form?.service || "",
      e.form?.salarieSociete || "",
      e.form?.dateEntretien || "",
      e.submittedAt ? fmt(e.submittedAt) : "",
      e.approvedAt ? fmt(e.approvedAt) : "",
      e.rejectionReason || ""
    ]);
    const csv = [headers, ...rows].map(r => r.map(v => `"${String(v).replace(/"/g,'""')}"`).join(",")).join("\n");
    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `EDP_2026_Export_${new Date().toISOString().slice(0,10)}.csv`; a.click();
    URL.revokeObjectURL(url);
    toast("success", "CSV exported", "All employee data downloaded as CSV.");
  };

  // ── EXPORT JSON ──
  const handleExportJSON = () => {
    const data = { exportedAt: now(), employees, managers };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `EDP_2026_Backup_${new Date().toISOString().slice(0,10)}.json`; a.click();
    URL.revokeObjectURL(url);
    toast("success", "JSON backup exported", "Full database backup downloaded.");
  };

  // ── RESET ALL DATA ──
  const handleResetAll = async () => {
    // Delete from SharePoint
    for (const emp of employees) await deleteEmployeeSP(emp);
    for (const mgr of managers) await deleteManagerSP(mgr);
    setEmployees([]);
    setManagers([]);
    LS.set("employees", []);
    LS.set("managers", []);
    toast("info", "All data cleared", "Database has been reset.");
    setConfirmDeleteId(null);
  };

  // ─── RENDER ───────────────────────────────────────────────────
  const Btn = ({ onClick, color, children, disabled, outline }) => (
    <button onClick={onClick} disabled={disabled} style={{
      padding: "9px 22px", borderRadius: 7, border: outline ? `2px solid ${color}` : "none",
      background: disabled ? C.gray200 : outline ? "transparent" : color,
      color: disabled ? C.gray400 : outline ? color : "#fff",
      cursor: disabled ? "not-allowed" : "pointer",
      fontFamily: "'Lato', sans-serif", fontSize: 14, fontWeight: 700,
      transition: "all 0.15s"
    }}>{children}</button>
  );

  // ── LANDIS+GYR LOGO SVG ──
  const LGLogo = ({ height = 36 }) => (
    <svg height={height} viewBox="0 0 220 80" xmlns="http://www.w3.org/2000/svg">
      {/* "Landis" text */}
      <text x="0" y="38" fontFamily="Arial, sans-serif" fontWeight="700" fontSize="34" fill="#ffffff">Landis</text>
      {/* Vertical bar */}
      <rect x="108" y="10" width="4" height="60" fill="#78be20"/>
      {/* "Gyr" text */}
      <text x="116" y="38" fontFamily="Arial, sans-serif" fontWeight="700" fontSize="34" fill="#ffffff">Gyr</text>
      {/* Green "+" plus sign */}
      <text x="190" y="28" fontFamily="Arial, sans-serif" fontWeight="700" fontSize="28" fill="#78be20">+</text>
      {/* Tagline */}
      <text x="2" y="68" fontFamily="Arial, sans-serif" fontSize="13" fill="#78be20">manage energy better</text>
    </svg>
  );

  // ── HEADER ──
  const Header = ({ title, subtitle }) => (
    <div style={{ background: C.navy, color: "#fff", padding: "0", borderBottom: `3px solid ${C.gold}` }}>
      <div style={{ maxWidth: 1100, margin: "0 auto", padding: "14px 24px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        {/* Left: Logo + page title */}
        <div style={{ display: "flex", alignItems: "center", gap: 24 }}>
          <LGLogo height={48} />
          <div style={{ borderLeft: "1px solid rgba(255,255,255,0.2)", paddingLeft: 24 }}>
            <div style={{ fontSize: 10, letterSpacing: 3, color: C.gold, textTransform: "uppercase", marginBottom: 3 }}>BN4097a — EDP 2026</div>
            <div style={{ fontSize: 18, fontFamily: "'Lato', sans-serif", fontWeight: 700 }}>{title}</div>
            {subtitle && <div style={{ fontSize: 12, color: "rgba(255,255,255,0.6)", marginTop: 2 }}>{subtitle}</div>}
          </div>
        </div>
        {/* Right: notifications + user */}
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          {currentUser && (
            <>
              <div style={{ position: "relative" }}>
                <button onClick={() => { setShowNotifs(!showNotifs); if (!showNotifs) markNotifsRead(); }}
                  style={{ background: "none", border: "none", color: "#fff", cursor: "pointer", fontSize: 20, position: "relative", padding: "4px 8px" }}>
                  🔔
                  {unreadCount > 0 && <span style={{ position: "absolute", top: 0, right: 0, background: "#ef4444", color: "#fff", borderRadius: "50%", width: 18, height: 18, fontSize: 11, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700 }}>{unreadCount}</span>}
                </button>
                {showNotifs && (
                  <div style={{ position: "absolute", right: 0, top: "100%", background: C.white, border: `1px solid ${C.gray200}`, borderRadius: 10, boxShadow: "0 8px 30px rgba(0,0,0,0.15)", width: 320, zIndex: 200, overflow: "hidden" }}>
                    <div style={{ padding: "12px 16px", borderBottom: `1px solid ${C.gray100}`, fontWeight: 700, fontSize: 13, color: C.navy, display: "flex", alignItems: "center", gap: 8 }}>
                      <span style={{ color: C.gold }}>🔔</span> Notifications
                    </div>
                    {notifs.length === 0 ? <div style={{ padding: "20px 16px", color: C.gray400, fontSize: 13, textAlign: "center" }}>No notifications yet</div> :
                      notifs.slice(0, 8).map(n => (
                        <div key={n.id} style={{ padding: "10px 16px", borderBottom: `1px solid ${C.gray100}`, background: n.read ? C.white : "#f0fdf4", borderLeft: n.read ? "none" : `3px solid ${C.gold}` }}>
                          <div style={{ fontWeight: 600, fontSize: 13, color: C.gray900, marginBottom: 2 }}>{n.title}</div>
                          <div style={{ fontSize: 12, color: C.gray600 }}>{n.message}</div>
                          <div style={{ fontSize: 11, color: C.gray400, marginTop: 4 }}>{fmt(n.ts)}</div>
                        </div>
                      ))
                    }
                  </div>
                )}
              </div>
              <div style={{ fontSize: 13, color: "rgba(255,255,255,0.7)", borderRight: "1px solid rgba(255,255,255,0.2)", paddingRight: 12 }}>{currentUser.name}</div>
              <button onClick={() => { setCurrentUser(null); setScreen("landing"); setShowNotifs(false); }}
                style={{ background: "rgba(255,255,255,0.08)", border: "1px solid rgba(255,255,255,0.2)", color: "#fff", borderRadius: 6, padding: "5px 12px", cursor: "pointer", fontSize: 12 }}>
                Sign out
              </button>
            </>
          )}
        </div>
      </div>
    </div>
  );

  // ── OTP VERIFICATION SCREEN ──
  if (otpScreen) return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "'Lato', sans-serif", display: "flex", alignItems: "center", justifyContent: "center" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Lato:wght@400;600;700;900&display=swap'); @keyframes toastIn{from{transform:translateX(30px);opacity:0}to{transform:translateX(0);opacity:1}} @keyframes pulse{0%,100%{opacity:1}50%{opacity:0.5}}`}</style>
      <Toast toasts={toasts} onDismiss={dismissToast} />

      <div style={{ width: "100%", maxWidth: 420, padding: "0 24px" }}>
        {/* Logo */}
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{ display: "inline-block", background: "#4a4a4a", borderRadius: 12, padding: "16px 28px", marginBottom: 12 }}>
            <svg height="40" viewBox="0 0 220 55" xmlns="http://www.w3.org/2000/svg">
              <text x="0" y="34" fontFamily="Arial" fontWeight="700" fontSize="30" fill="#ffffff">Landis</text>
              <rect x="108" y="6" width="3" height="44" fill="#78be20"/>
              <text x="115" y="34" fontFamily="Arial" fontWeight="700" fontSize="30" fill="#ffffff">Gyr</text>
              <text x="188" y="24" fontFamily="Arial" fontWeight="700" fontSize="22" fill="#78be20">+</text>
            </svg>
          </div>
        </div>

        {/* Card */}
        <div style={{ background: "#fff", borderRadius: 16, padding: 36, boxShadow: "0 8px 40px rgba(0,0,0,0.10)", border: `1px solid ${C.gray200}`, borderTop: `4px solid ${C.gold}` }}>

          {/* Shield icon */}
          <div style={{ textAlign: "center", marginBottom: 20 }}>
            <div style={{ width: 60, height: 60, borderRadius: "50%", background: "#f0fdf4", border: `2px solid #bbf7d0`, display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 28 }}>🔐</div>
          </div>

          <h2 style={{ textAlign: "center", fontWeight: 900, color: C.navy, fontSize: 20, margin: "0 0 8px" }}>Two-Factor Verification</h2>
          <p style={{ textAlign: "center", color: C.gray600, fontSize: 13, marginBottom: 24, lineHeight: 1.6 }}>
            A 6-digit code has been sent to<br/>
            <strong style={{ color: C.navy }}>{otpEmail}</strong>
          </p>

          {/* Code input */}
          <div style={{ marginBottom: 16 }}>
            <label style={{ display: "block", fontSize: 11, color: C.gray600, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 8 }}>Verification Code</label>
            <input
              value={otpInput}
              onChange={e => { setOtpInput(e.target.value.replace(/\D/g, "").slice(0, 6)); setOtpError(""); }}
              onKeyDown={e => e.key === "Enter" && otpInput.length === 6 && handleVerifyOTP()}
              placeholder="_ _ _ _ _ _"
              maxLength={6}
              autoFocus
              style={{
                width: "100%", padding: "14px 16px", borderRadius: 10,
                border: `2px solid ${otpError ? C.red : otpInput.length === 6 ? C.gold : C.gray200}`,
                fontFamily: "monospace", fontSize: 28, fontWeight: 700,
                textAlign: "center", letterSpacing: 10, outline: "none",
                background: "#fff", color: C.navy, transition: "border-color 0.2s",
                boxSizing: "border-box"
              }}
            />
            {otpError && (
              <div style={{ marginTop: 8, color: C.red, fontSize: 13, display: "flex", alignItems: "center", gap: 6 }}>
                <span>❌</span> {otpError}
              </div>
            )}
          </div>

          {/* Verify button */}
          <button
            onClick={handleVerifyOTP}
            disabled={otpInput.length !== 6 || otpSending}
            style={{
              width: "100%", padding: "13px", borderRadius: 10, border: "none",
              background: otpInput.length === 6 ? C.gold : C.gray200,
              color: otpInput.length === 6 ? "#fff" : C.gray400,
              fontFamily: "'Lato', sans-serif", fontSize: 15, fontWeight: 700,
              cursor: otpInput.length === 6 ? "pointer" : "not-allowed",
              transition: "all 0.2s", marginBottom: 16
            }}>
            {otpSending ? "Verifying…" : "✓ Verify & Sign In"}
          </button>

          {/* Resend / back */}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <button onClick={() => { setOtpScreen(false); setOtpInput(""); setOtpError(""); }}
              style={{ background: "none", border: "none", color: C.gray400, cursor: "pointer", fontSize: 13 }}>
              ← Back
            </button>
            <button
              onClick={handleResendOTP}
              disabled={otpResendTimer > 0 || otpSending}
              style={{ background: "none", border: "none", cursor: otpResendTimer > 0 ? "default" : "pointer",
                color: otpResendTimer > 0 ? C.gray400 : C.gold, fontSize: 13, fontWeight: 600 }}>
              {otpResendTimer > 0 ? `Resend in ${otpResendTimer}s` : "Resend code"}
            </button>
          </div>

          {/* Dev mode hint */}
          {!isGraphConfigured() && (
            <div style={{ marginTop: 16, background: "#fef9ec", border: `1px solid #fde68a`, borderRadius: 8, padding: "10px 14px", fontSize: 12, color: "#92400e" }}>
              ⚠️ <strong>Dev mode:</strong> Graph API not configured — a popup will show the OTP code instead of sending an email.
            </div>
          )}
        </div>

        <div style={{ textAlign: "center", marginTop: 20, color: C.gray400, fontSize: 12 }}>
          Landis+Gyr — EDP 2026 · Secure Portal
        </div>
      </div>
    </div>
  );

  // ── LOADING ──
  if (loading) return (
    <div style={{ minHeight: "100vh", background: C.bg, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ textAlign: "center", color: C.navy }}>
        <div style={{ fontSize: 40, marginBottom: 16 }}>⏳</div>
        <div style={{ fontFamily: "'Playfair Display', serif", fontSize: 18 }}>Loading…</div>
      </div>
    </div>
  );

  // ── LANDING ──
  if (screen === "landing") return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "'Lato', sans-serif" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Lato:wght@400;600;700;900&display=swap'); @keyframes toastIn{from{transform:translateX(30px);opacity:0}to{transform:translateX(0);opacity:1}} * { box-sizing: border-box; }`}</style>
      <Toast toasts={toasts} onDismiss={dismissToast} />
      <Header title="Professional Development Interview" subtitle="EDP 2026 — Employee Portal" />

      {/* Green accent bar */}
      <div style={{ height: 4, background: `linear-gradient(to right, ${C.gold}, #a8d86e)` }} />

      <div style={{ maxWidth: 920, margin: "50px auto", padding: "0 24px" }}>
        {/* Welcome */}
        <div style={{ textAlign: "center", marginBottom: 44 }}>
          <div style={{ display: "inline-block", background: C.gold, color: "#fff", borderRadius: 30, padding: "4px 18px", fontSize: 12, fontWeight: 700, letterSpacing: 1, textTransform: "uppercase", marginBottom: 16 }}>
            Landis+Gyr — EDP 2026
          </div>
          <h1 style={{ fontFamily: "'Lato', sans-serif", fontWeight: 900, color: C.navy, fontSize: 30, margin: "0 0 12px" }}>
            Professional Development Portal
          </h1>
          <p style={{ color: C.gray600, fontSize: 15, maxWidth: 520, margin: "0 auto", lineHeight: 1.7 }}>
            Complete your annual Professional Development Interview, track your approval status, and download your finalized document for SharePoint.
          </p>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 20 }}>
          {/* New employee */}
          <div style={{ background: C.white, borderRadius: 12, padding: 28, border: `1px solid ${C.gray200}`, boxShadow: "0 2px 12px rgba(0,0,0,0.05)", borderTop: `4px solid ${C.gold}` }}>
            <div style={{ fontSize: 30, marginBottom: 12 }}>🆕</div>
            <h3 style={{ fontWeight: 800, color: C.navy, margin: "0 0 8px", fontSize: 17 }}>First Time?</h3>
            <p style={{ color: C.gray600, fontSize: 13, marginBottom: 20, lineHeight: 1.6 }}>Register with your name and work email to create your EDP form.</p>
            <Btn onClick={() => { setScreen("register"); setRegError(""); setRegName(""); setRegEmail(""); }} color={C.gold}>Create Account →</Btn>
          </div>

          {/* Returning employee */}
          <div style={{ background: C.white, borderRadius: 12, padding: 28, border: `1px solid ${C.gray200}`, boxShadow: "0 2px 12px rgba(0,0,0,0.05)", borderTop: `4px solid ${C.navy}` }}>
            <div style={{ fontSize: 30, marginBottom: 12 }}>🔁</div>
            <h3 style={{ fontWeight: 800, color: C.navy, margin: "0 0 8px", fontSize: 17 }}>Returning Employee</h3>
            <p style={{ color: C.gray600, fontSize: 13, marginBottom: 12, lineHeight: 1.6 }}>Continue your saved form or check your approval status.</p>
            <input style={{ ...inp(false), marginBottom: 10 }} placeholder="Your work email address…" onKeyDown={e => e.key==="Enter" && handleLogin(e.target.value)} id="login-email" />
            <Btn onClick={() => { const el = document.getElementById("login-email"); if(el) handleLogin(el.value); }} color={C.navy}>Sign In →</Btn>
            {regError && <div style={{ color: C.red, fontSize: 12, marginTop: 8 }}>{regError}</div>}
          </div>

          {/* Manager */}
          <div style={{ background: C.navy, borderRadius: 12, padding: 28, boxShadow: "0 2px 12px rgba(0,0,0,0.12)", borderTop: `4px solid ${C.gold}` }}>
            <div style={{ fontSize: 30, marginBottom: 12 }}>🧑‍💼</div>
            <h3 style={{ fontWeight: 800, color: "#fff", margin: "0 0 8px", fontSize: 17 }}>HR / Manager</h3>
            <p style={{ color: "rgba(255,255,255,0.6)", fontSize: 13, marginBottom: 12, lineHeight: 1.6 }}>Already registered? Sign in with your email. First time? Click Register and enter the shared manager code.</p>
            <input style={{ ...inp(false), marginBottom: 10, background: "rgba(255,255,255,0.1)", border: "1px solid rgba(255,255,255,0.2)", color: "#fff" }}
              placeholder="Manager email address…" id="mgr-login-email" onKeyDown={e => e.key === "Enter" && handleLogin(e.target.value)} />
            <div style={{ display: "flex", gap: 8 }}>
              <Btn onClick={() => { const el = document.getElementById("mgr-login-email"); if(el) handleLogin(el.value); }} color={C.gold}>Sign In →</Btn>
              <Btn onClick={() => { setScreen("register"); setRegError(""); setRegName(""); setRegEmail(""); setRegInviteCode(""); }} color={C.white} outline>Register</Btn>
            </div>
          </div>
        </div>

        {/* Footer note */}
        <div style={{ textAlign: "center", marginTop: 40, color: C.gray400, fontSize: 12 }}>
          Landis+Gyr — Professional Development Interview System &nbsp;·&nbsp; BN4097a &nbsp;·&nbsp; 2026
        </div>
      </div>
    </div>
  );

  // ── REGISTER ──
  if (screen === "register") return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "'Lato', sans-serif" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Lato:wght@400;600;700;900&display=swap'); @keyframes toastIn{from{transform:translateX(30px);opacity:0}to{transform:translateX(0);opacity:1}}`}</style>
      <Toast toasts={toasts} onDismiss={dismissToast} />
      <Header title="Register — EDP 2026" />
      <div style={{ maxWidth: 480, margin: "60px auto", padding: "0 24px" }}>
        <div style={{ background: C.white, borderRadius: 12, padding: 36, boxShadow: "0 4px 20px rgba(0,0,0,0.08)", border: `1px solid ${C.gray200}`, borderTop: `4px solid ${C.gold}` }}>
          <h2 style={{ fontWeight: 900, color: C.navy, marginTop: 0, fontSize: 22 }}>Create Your Account</h2>
          <p style={{ color: C.gray600, fontSize: 13, marginBottom: 20, lineHeight: 1.6 }}>Employees: fill in your name and email only. <strong>HR Managers:</strong> also enter the shared manager code provided by your HR administrator.</p>
          <Field label="Full Name" required>
            <input style={inp(false)} value={regName} onChange={e => setRegName(e.target.value)} placeholder="Last name, First name" />
          </Field>
          <Field label="Work Email Address" required>
            <input style={inp(false)} type="email" value={regEmail} onChange={e => setRegEmail(e.target.value)} placeholder="your.name@landis-gyr.com" />
          </Field>
          <Field label="Manager Code (HR only — leave blank if employee)">
            <input style={{ ...inp(false),
                border: regInviteCode ? (regInviteCode === MANAGER_CODE ? `1.5px solid ${C.gold}` : `1.5px solid ${C.red}`) : inp(false).border }}
              type="password" value={regInviteCode} onChange={e => setRegInviteCode(e.target.value)}
              placeholder="Leave blank if you are an employee…" onKeyDown={e => e.key === "Enter" && handleRegister()} />
            {regInviteCode === MANAGER_CODE && (
              <div style={{ marginTop: 5, fontSize: 12, color: C.gold, fontWeight: 700 }}>✓ Valid — you will be registered as HR Manager</div>
            )}
            {regInviteCode && regInviteCode !== MANAGER_CODE && (
              <div style={{ marginTop: 5, fontSize: 12, color: C.red }}>✗ Incorrect manager code — registering as employee</div>
            )}
          </Field>
          {regError && <div style={{ color: C.red, fontSize: 13, marginBottom: 12, padding: "8px 12px", background: "#fef2f2", borderRadius: 6 }}>{regError}</div>}
          <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
            <Btn onClick={() => { setScreen("landing"); setRegError(""); setRegInviteCode(""); }} color={C.gray600} outline>← Back</Btn>
            <Btn onClick={handleRegister} color={regInviteCode === MANAGER_CODE ? C.gold : C.navy}>
              {regInviteCode === MANAGER_CODE ? "Create Manager Account →" : "Create Employee Account →"}
            </Btn>
          </div>
        </div>
      </div>
    </div>
  );

  // ── EMPLOYEE FORM ──
  if (screen === "employee" && currentUser?.role === "employee") {
    const emp = employees.find(e => e.id === currentUser.id);
    const status = emp?.status || "draft";
    const canEdit = status === "draft" || status === "rejected";
    const canSubmit = canEdit;

    return (
      <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "'Lato', sans-serif" }}>
        <style>{`@import url('https://fonts.googleapis.com/css2?family=Lato:wght@400;600;700;900&display=swap'); @keyframes toastIn{from{transform:translateX(30px);opacity:0}to{transform:translateX(0);opacity:1}}`}</style>
        <Toast toasts={toasts} onDismiss={dismissToast} />
        <Header title="My EDP Form" subtitle={`${currentUser.name} — ${currentUser.email}`} />

        {/* Status bar */}
        <div style={{ background: C.white, borderBottom: `1px solid ${C.gray200}`, padding: "10px 24px" }}>
          <div style={{ maxWidth: 900, margin: "0 auto", display: "flex", alignItems: "center", gap: 16, flexWrap: "wrap" }}>
            <StatusBadge status={status} size="lg" />
            {emp?.submittedAt && <span style={{ fontSize: 13, color: C.gray400 }}>Submitted: {fmt(emp.submittedAt)}</span>}
            {status === "approved" && emp?.approvedAt && <span style={{ fontSize: 13, color: C.green }}>✓ Approved: {fmt(emp.approvedAt)}</span>}
          </div>
        </div>

        {/* Rejection banner */}
        {status === "rejected" && emp?.rejectionReason && (
          <div style={{ background: "#fef2f2", borderBottom: `2px solid #fca5a5`, padding: "14px 24px" }}>
            <div style={{ maxWidth: 900, margin: "0 auto", display: "flex", gap: 12, alignItems: "flex-start" }}>
              <span style={{ fontSize: 22 }}>❌</span>
              <div>
                <div style={{ fontWeight: 700, color: C.red, fontSize: 15 }}>Your form needs correction</div>
                <div style={{ color: "#991b1b", fontSize: 14, marginTop: 3 }}>Manager's note: <em>"{emp.rejectionReason}"</em></div>
                <div style={{ marginTop: 10 }}>
                  <Btn onClick={handleReEdit} color={C.amber}>✏️ Re-open form for editing</Btn>
                </div>
              </div>
            </div>
          </div>
        )}

        <div style={{ maxWidth: 900, margin: "0 auto", padding: "32px 24px" }}>
          <div style={{ background: C.white, borderRadius: 12, padding: 32, boxShadow: "0 2px 12px rgba(0,0,0,0.06)", border: `1px solid ${C.gray200}` }}>
            <EDPForm form={localForm} onChange={canEdit ? setLocalForm : () => {}} readOnly={!canEdit} />
          </div>

          {/* Actions */}
          <div style={{ display: "flex", gap: 12, marginTop: 20, justifyContent: "flex-end", flexWrap: "wrap" }}>
            {(status === "approved" || status === "submitted") && (
              <Btn onClick={() => handleDownload(emp)} color={C.navy} outline>⬇ Download as Word (.docx)</Btn>
            )}
            {canEdit && (
              <>
                <Btn onClick={handleSaveDraft} color={C.gray600} outline>💾 Save Draft</Btn>
                <Btn onClick={handleSubmit} color={C.navy}>📤 Submit for Review</Btn>
              </>
            )}
          </div>
        </div>
      </div>
    );
  }

  // ── MANAGER DASHBOARD ──
  if (screen === "manager") return (
    <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "'Lato', sans-serif" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Lato:wght@400;600;700;900&display=swap'); @keyframes toastIn{from{transform:translateX(30px);opacity:0}to{transform:translateX(0);opacity:1}}`}</style>
      <Toast toasts={toasts} onDismiss={dismissToast} />
      <Header title="HR Manager Dashboard" subtitle={`EDP 2026 — Logged in as: ${currentUser?.name} · ${isSharePointEnabled() ? "🟢 SharePoint connected" : "🟡 Local storage"} · ${isGraphConfigured() ? "✅ Email active" : "⚠️ Email not configured"}`} />

      {/* Tab bar */}
      <div style={{ background: C.white, borderBottom: `1px solid ${C.gray200}` }}>
        <div style={{ maxWidth: 1100, margin: "0 auto", padding: "0 24px", display: "flex", gap: 0 }}>
          {[
            { id: "employees", label: "👥 Employees", count: employees.length },
            { id: "admin", label: "🗄️ Database Admin", count: null },
          ].map(tab => (
            <button key={tab.id} onClick={() => setManagerTab(tab.id)} style={{
              padding: "14px 20px", border: "none", background: "none", cursor: "pointer",
              fontFamily: "'Lato', sans-serif", fontSize: 14, fontWeight: managerTab === tab.id ? 700 : 400,
              color: managerTab === tab.id ? C.gold : C.gray600,
              borderBottom: managerTab === tab.id ? `3px solid ${C.gold}` : "3px solid transparent",
              transition: "all 0.15s", display: "flex", alignItems: "center", gap: 8
            }}>
              {tab.label}
              {tab.count !== null && <span style={{ background: managerTab === tab.id ? C.gold : C.gray200, color: managerTab === tab.id ? "#fff" : C.gray600, borderRadius: 10, padding: "1px 8px", fontSize: 12 }}>{tab.count}</span>}
            </button>
          ))}
        </div>
      </div>

      <div style={{ maxWidth: 1100, margin: "0 auto", padding: "28px 24px" }}>

      {/* ── ADMIN TAB ── */}
      {managerTab === "admin" && (
        <div>
          {/* SharePoint status banner */}
          <div style={{ background: isSharePointEnabled() ? "#f0fdf4" : "#fef9ec",
            border: `1px solid ${isSharePointEnabled() ? "#bbf7d0" : "#fde68a"}`,
            borderRadius: 10, padding: "14px 20px", marginBottom: 20,
            display: "flex", alignItems: "center", gap: 14 }}>
            <span style={{ fontSize: 24 }}>{isSharePointEnabled() ? "🟢" : "🟡"}</span>
            <div>
              <div style={{ fontWeight: 700, fontSize: 14, color: isSharePointEnabled() ? "#15803d" : "#92400e" }}>
                {isSharePointEnabled() ? "SharePoint Connected — data syncs automatically" : "SharePoint not configured — data stored in browser only"}
              </div>
              <div style={{ fontSize: 12, color: "#6b7280", marginTop: 2 }}>
                {isSharePointEnabled()
                  ? `Site: ${GRAPH_CONFIG.siteUrl} · Lists: EDP_Employees, EDP_Managers, EDP_Notifications`
                  : "See SHAREPOINT-SETUP-GUIDE.html to connect to SharePoint"}
              </div>
            </div>
          </div>

          {/* Summary cards */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14, marginBottom: 28 }}>
            {[
              { label: "Total Employees", value: employees.length, icon: "👥", color: C.navy },
              { label: "Registered Managers", value: managers.length, icon: "🧑‍💼", color: C.gold },
              { label: "Forms Submitted", value: employees.filter(e => e.status !== "draft").length, icon: "📤", color: C.blue },
              { label: "Forms Approved", value: employees.filter(e => e.status === "approved").length, icon: "✅", color: C.green },
            ].map(s => (
              <div key={s.label} style={{ background: C.white, borderRadius: 10, padding: "18px 20px", border: `1px solid ${C.gray200}`, borderLeft: `4px solid ${s.color}` }}>
                <div style={{ fontSize: 24, marginBottom: 6 }}>{s.icon}</div>
                <div style={{ fontSize: 28, fontWeight: 900, color: s.color }}>{s.value}</div>
                <div style={{ fontSize: 12, color: C.gray400, textTransform: "uppercase", letterSpacing: 0.5, marginTop: 2 }}>{s.label}</div>
              </div>
            ))}
          </div>

          {/* Export tools */}
          <div style={{ background: C.white, borderRadius: 12, padding: 24, border: `1px solid ${C.gray200}`, marginBottom: 20 }}>
            <div style={{ fontWeight: 800, fontSize: 16, color: C.navy, marginBottom: 6 }}>📤 Export Data</div>
            <div style={{ fontSize: 13, color: C.gray600, marginBottom: 16 }}>Download all employee data for reporting or backup.</div>
            <div style={{ display: "flex", gap: 10 }}>
              <Btn onClick={handleExportCSV} color={C.navy}>⬇ Export as CSV (Excel)</Btn>
              <Btn onClick={handleExportJSON} color={C.gray600} outline>⬇ Export as JSON (Backup)</Btn>
            </div>
          </div>

          {/* Managers table */}
          <div style={{ background: C.white, borderRadius: 12, padding: 24, border: `1px solid ${C.gray200}`, marginBottom: 20 }}>
            <div style={{ fontWeight: 800, fontSize: 16, color: C.navy, marginBottom: 16 }}>🧑‍💼 Registered Managers</div>
            {managers.length === 0 ? (
              <div style={{ color: C.gray400, fontSize: 13, textAlign: "center", padding: "20px 0" }}>No managers registered yet.</div>
            ) : (
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 14 }}>
                <thead>
                  <tr style={{ borderBottom: `2px solid ${C.gray200}` }}>
                    {["Name","Email","Registered","Actions"].map(h => (
                      <th key={h} style={{ padding: "8px 12px", textAlign: "left", fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5, color: C.gray400, fontWeight: 700 }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {managers.map(mgr => (
                    <tr key={mgr.id} style={{ borderBottom: `1px solid ${C.gray100}` }}>
                      <td style={{ padding: "10px 12px", fontWeight: 600, color: C.gray900 }}>
                        <span style={{ display: "inline-flex", alignItems: "center", gap: 8 }}>
                          <span style={{ width: 30, height: 30, borderRadius: "50%", background: C.gold, color: "#fff", display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 13, fontWeight: 700 }}>{mgr.name.charAt(0).toUpperCase()}</span>
                          {mgr.name}
                          {mgr.id === currentUser?.id && <span style={{ background: "#f0fdf4", color: C.green, fontSize: 11, padding: "1px 8px", borderRadius: 10, fontWeight: 600 }}>You</span>}
                        </span>
                      </td>
                      <td style={{ padding: "10px 12px", color: C.gray600 }}>{mgr.email}</td>
                      <td style={{ padding: "10px 12px", color: C.gray400, fontSize: 12 }}>{fmt(mgr.registeredAt)}</td>
                      <td style={{ padding: "10px 12px" }}>
                        {mgr.id !== currentUser?.id && (
                          confirmDeleteMgrId === mgr.id ? (
                            <span style={{ display: "flex", gap: 6, alignItems: "center" }}>
                              <span style={{ fontSize: 12, color: C.red }}>Confirm?</span>
                              <button onClick={() => handleDeleteManager(mgr.id)} style={{ padding: "3px 10px", background: C.red, color: "#fff", border: "none", borderRadius: 5, cursor: "pointer", fontSize: 12 }}>Yes</button>
                              <button onClick={() => setConfirmDeleteMgrId(null)} style={{ padding: "3px 10px", background: C.gray100, color: C.gray600, border: "none", borderRadius: 5, cursor: "pointer", fontSize: 12 }}>No</button>
                            </span>
                          ) : (
                            <button onClick={() => setConfirmDeleteMgrId(mgr.id)} style={{ color: C.red, background: "none", border: "none", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>🗑 Remove</button>
                          )
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>

          {/* Employees full table */}
          <div style={{ background: C.white, borderRadius: 12, padding: 24, border: `1px solid ${C.gray200}`, marginBottom: 20 }}>
            <div style={{ fontWeight: 800, fontSize: 16, color: C.navy, marginBottom: 16 }}>📋 All Employee Records</div>
            {employees.length === 0 ? (
              <div style={{ color: C.gray400, fontSize: 13, textAlign: "center", padding: "20px 0" }}>No employees registered yet.</div>
            ) : (
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead>
                    <tr style={{ borderBottom: `2px solid ${C.gray200}`, background: C.gray50 }}>
                      {["Name","Email","Job Title","Department","Status","Submitted","Approved","Actions"].map(h => (
                        <th key={h} style={{ padding: "10px 12px", textAlign: "left", fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5, color: C.gray400, fontWeight: 700, whiteSpace: "nowrap" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {employees.map(emp => (
                      <tr key={emp.id} style={{ borderBottom: `1px solid ${C.gray100}` }}>
                        <td style={{ padding: "10px 12px", fontWeight: 600, color: C.gray900, whiteSpace: "nowrap" }}>
                          <span style={{ display: "inline-flex", alignItems: "center", gap: 8 }}>
                            <span style={{ width: 28, height: 28, borderRadius: "50%", background: C.navy, color: "#fff", display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: 700, flexShrink: 0 }}>{emp.name.charAt(0).toUpperCase()}</span>
                            {emp.name}
                          </span>
                        </td>
                        <td style={{ padding: "10px 12px", color: C.gray600, fontSize: 12 }}>{emp.email}</td>
                        <td style={{ padding: "10px 12px", color: C.gray600 }}>{emp.form?.salariePoste || <span style={{ color: C.gray300 }}>—</span>}</td>
                        <td style={{ padding: "10px 12px", color: C.gray600 }}>{emp.form?.service || <span style={{ color: C.gray300 }}>—</span>}</td>
                        <td style={{ padding: "10px 12px" }}><StatusBadge status={emp.status} /></td>
                        <td style={{ padding: "10px 12px", color: C.gray400, fontSize: 12, whiteSpace: "nowrap" }}>{emp.submittedAt ? fmt(emp.submittedAt) : "—"}</td>
                        <td style={{ padding: "10px 12px", color: emp.approvedAt ? C.green : C.gray300, fontSize: 12, whiteSpace: "nowrap" }}>{emp.approvedAt ? fmt(emp.approvedAt) : "—"}</td>
                        <td style={{ padding: "10px 12px" }}>
                          <span style={{ display: "flex", gap: 6 }}>
                            <button onClick={() => { setSelectedEmpId(emp.id); setScreen("review"); }}
                              style={{ padding: "4px 10px", background: C.navy, color: "#fff", border: "none", borderRadius: 5, cursor: "pointer", fontSize: 12 }}>View</button>
                            {confirmDeleteId === emp.id ? (
                              <>
                                <button onClick={() => handleDeleteEmployee(emp.id)} style={{ padding: "4px 10px", background: C.red, color: "#fff", border: "none", borderRadius: 5, cursor: "pointer", fontSize: 12 }}>Confirm</button>
                                <button onClick={() => setConfirmDeleteId(null)} style={{ padding: "4px 10px", background: C.gray100, color: C.gray600, border: "none", borderRadius: 5, cursor: "pointer", fontSize: 12 }}>Cancel</button>
                              </>
                            ) : (
                              <button onClick={() => setConfirmDeleteId(emp.id)} style={{ padding: "4px 10px", background: "#fff", color: C.red, border: `1px solid ${C.red}`, borderRadius: 5, cursor: "pointer", fontSize: 12 }}>🗑 Delete</button>
                            )}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>

          {/* Danger zone */}
          <div style={{ background: "#fff5f5", borderRadius: 12, padding: 24, border: `1px solid #fca5a5` }}>
            <div style={{ fontWeight: 800, fontSize: 15, color: C.red, marginBottom: 6 }}>⚠️ Danger Zone</div>
            <div style={{ fontSize: 13, color: C.gray600, marginBottom: 16 }}>These actions are irreversible. All data will be permanently deleted.</div>
            {confirmDeleteId === "ALL" ? (
              <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                <span style={{ fontSize: 14, color: C.red, fontWeight: 600 }}>Are you sure? This will delete ALL employees and managers.</span>
                <Btn onClick={handleResetAll} color={C.red}>Yes, delete everything</Btn>
                <Btn onClick={() => setConfirmDeleteId(null)} color={C.gray600} outline>Cancel</Btn>
              </div>
            ) : (
              <Btn onClick={() => setConfirmDeleteId("ALL")} color={C.red} outline>🗑 Reset entire database</Btn>
            )}
          </div>
        </div>
      )}

      {/* ── EMPLOYEES TAB ── */}
      {managerTab === "employees" && <>
        {/* Stats */}}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 14, marginBottom: 28 }}>
          {[
            { label: "Total", value: stats.total, color: C.navy, icon: "👥" },
            { label: "Draft", value: stats.draft, color: C.gray600, icon: "📝" },
            { label: "Pending Review", value: stats.submitted, color: C.blue, icon: "⏳" },
            { label: "Approved", value: stats.approved, color: C.green, icon: "✅" },
            { label: "Needs Correction", value: stats.rejected, color: C.red, icon: "🔄" },
          ].map(s => (
            <div key={s.label} style={{ background: C.white, borderRadius: 10, padding: "18px 16px", border: `1px solid ${C.gray200}`, boxShadow: "0 1px 6px rgba(0,0,0,0.05)" }}>
              <div style={{ fontSize: 22, marginBottom: 6 }}>{s.icon}</div>
              <div style={{ fontSize: 26, fontWeight: 700, color: s.color, fontFamily: "'Playfair Display', serif" }}>{s.value}</div>
              <div style={{ fontSize: 12, color: C.gray400, textTransform: "uppercase", letterSpacing: 0.5 }}>{s.label}</div>
            </div>
          ))}
        </div>

        {/* Filters */}
        <div style={{ display: "flex", gap: 12, marginBottom: 20, flexWrap: "wrap", alignItems: "center" }}>
          <input style={{ ...inp(false), width: 240 }} placeholder="🔍 Search by name or email…" value={managerSearch} onChange={e => setManagerSearch(e.target.value)} />
          <div style={{ display: "flex", gap: 6 }}>
            {["all", "submitted", "approved", "rejected", "draft"].map(f => (
              <button key={f} onClick={() => setManagerFilter(f)} style={{
                padding: "7px 14px", borderRadius: 20, border: "none", cursor: "pointer",
                background: managerFilter === f ? C.navy : C.gray100,
                color: managerFilter === f ? "#fff" : C.gray700,
                fontFamily: "'Lato', sans-serif", fontSize: 13, fontWeight: managerFilter === f ? 700 : 400
              }}>
                {f === "all" ? "All" : STATUS_CFG[f]?.label}
                {f !== "all" && <span style={{ marginLeft: 6, background: "rgba(255,255,255,0.2)", borderRadius: 10, padding: "1px 6px", fontSize: 11 }}>{stats[f] || 0}</span>}
              </button>
            ))}
          </div>
        </div>

        {/* Employee list */}
        {filteredEmps.length === 0 ? (
          <div style={{ background: C.white, borderRadius: 12, padding: 48, textAlign: "center", color: C.gray400 }}>
            <div style={{ fontSize: 40, marginBottom: 12 }}>📭</div>
            <div style={{ fontSize: 16 }}>{employees.length === 0 ? "No employees registered yet." : "No results match your filters."}</div>
          </div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {filteredEmps.map(emp => (
              <div key={emp.id} style={{ background: C.white, borderRadius: 10, padding: "18px 22px", border: `1px solid ${C.gray200}`, boxShadow: "0 1px 6px rgba(0,0,0,0.04)", display: "flex", alignItems: "center", gap: 16, flexWrap: "wrap" }}>
                <div style={{ width: 42, height: 42, borderRadius: "50%", background: C.navy, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 700, fontSize: 16, fontFamily: "'Playfair Display', serif", flexShrink: 0 }}>
                  {emp.name.charAt(0).toUpperCase()}
                </div>
                <div style={{ flex: 1, minWidth: 180 }}>
                  <div style={{ fontWeight: 700, color: C.gray900, fontSize: 15 }}>{emp.name}</div>
                  <div style={{ fontSize: 12, color: C.gray400 }}>{emp.email}</div>
                  {emp.form?.salariePoste && <div style={{ fontSize: 12, color: C.gray600, marginTop: 2 }}>{emp.form.salariePoste} {emp.form.service && `— ${emp.form.service}`}</div>}
                </div>
                <div style={{ minWidth: 120 }}>
                  <StatusBadge status={emp.status} />
                  {emp.submittedAt && <div style={{ fontSize: 11, color: C.gray400, marginTop: 4 }}>Submitted {fmt(emp.submittedAt)}</div>}
                </div>
                <div style={{ display: "flex", gap: 8, flexShrink: 0 }}>
                  {emp.status === "submitted" && (
                    <button onClick={() => { setSelectedEmpId(emp.id); setScreen("review"); }}
                      style={{ padding: "7px 16px", borderRadius: 6, border: "none", background: C.navy, color: "#fff", cursor: "pointer", fontFamily: "'Lato', sans-serif", fontSize: 13, fontWeight: 700 }}>
                      Review →
                    </button>
                  )}
                  {(emp.status === "approved" || emp.status === "submitted") && (<>
                    <button onClick={() => handleDownload(emp)}
                      style={{ padding: "7px 14px", borderRadius: 6, border: `1px solid ${C.gray200}`, background: C.white, color: C.gray600, cursor: "pointer", fontFamily: "'Lato', sans-serif", fontSize: 13 }}>
                      ⬇ .docx
                    </button>
                    <button onClick={() => handleDownloadPDF(emp)}
                      style={{ padding: "7px 14px", borderRadius: 6, border: `1px solid #fecaca`, background: "#fff5f5", color: "#b91c1c", cursor: "pointer", fontFamily: "'Lato', sans-serif", fontSize: 13 }}>
                      ⬇ .pdf
                    </button>
                  </>)}
                  {(emp.status === "draft" || emp.status === "rejected" || emp.status === "approved") && (
                    <button onClick={() => { setSelectedEmpId(emp.id); setScreen("review"); }}
                      style={{ padding: "7px 14px", borderRadius: 6, border: `1px solid ${C.gray200}`, background: C.white, color: C.gray600, cursor: "pointer", fontFamily: "'Lato', sans-serif", fontSize: 13 }}>
                      View
                    </button>
                  )}
                </div>
              </div>
            ))}
          </div>
        )}
      </>}
      </div>
    </div>
  );

  // ── REVIEW SCREEN ──
  if (screen === "review" && selectedEmp) {
    const canAct = selectedEmp.status === "submitted";
    return (
      <div style={{ minHeight: "100vh", background: C.bg, fontFamily: "'Lato', sans-serif" }}>
        <style>{`@import url('https://fonts.googleapis.com/css2?family=Lato:wght@400;600;700;900&display=swap'); @keyframes toastIn{from{transform:translateX(30px);opacity:0}to{transform:translateX(0);opacity:1}}`}</style>
        <Toast toasts={toasts} onDismiss={dismissToast} />
        <Header title={`Reviewing: ${selectedEmp.name}`} subtitle={`${selectedEmp.form?.salariePoste || ""} ${selectedEmp.form?.service ? "— " + selectedEmp.form.service : ""}`} />

        {/* Sub-nav */}
        <div style={{ background: C.white, borderBottom: `1px solid ${C.gray200}`, padding: "10px 24px" }}>
          <div style={{ maxWidth: 900, margin: "0 auto", display: "flex", alignItems: "center", gap: 16 }}>
            <button onClick={() => setScreen("manager")} style={{ background: "none", border: "none", color: C.navy, cursor: "pointer", fontSize: 13, fontWeight: 600, padding: 0 }}>← Back to Dashboard</button>
            <StatusBadge status={selectedEmp.status} size="lg" />
            {canAct && <span style={{ fontSize: 13, color: C.amber, fontWeight: 600 }}>⚠️ Awaiting your review</span>}
          </div>
        </div>

        {/* History */}
        {selectedEmp.history?.length > 0 && (
          <div style={{ background: "#eff6ff", borderBottom: `1px solid #bfdbfe`, padding: "10px 24px" }}>
            <div style={{ maxWidth: 900, margin: "0 auto", display: "flex", gap: 16, flexWrap: "wrap", fontSize: 12, color: C.blue }}>
              <span style={{ fontWeight: 700 }}>History:</span>
              {selectedEmp.history.map((h, i) => (
                <span key={i}>{h.action} <span style={{ color: C.gray400 }}>({fmt(h.date)})</span>{h.reason ? ` — "${h.reason}"` : ""}</span>
              ))}
            </div>
          </div>
        )}

        <div style={{ maxWidth: 900, margin: "0 auto", padding: "32px 24px" }}>
          <div style={{ background: C.white, borderRadius: 12, padding: 32, boxShadow: "0 2px 12px rgba(0,0,0,0.06)", border: `1px solid ${C.gray200}` }}>
            <EDPForm form={selectedEmp.form || EMPTY_FORM} onChange={() => {}} readOnly={true} />
          </div>

          <div style={{ display: "flex", gap: 12, marginTop: 20, justifyContent: "flex-end", flexWrap: "wrap" }}>
            <Btn onClick={() => handleDownload(selectedEmp)} color={C.gray600} outline>⬇ Word (.docx)</Btn>
            <Btn onClick={() => handleDownloadPDF(selectedEmp)} color="#b91c1c" outline>⬇ PDF</Btn>
            {isSharePointEnabled() && selectedEmp?.status === "approved" && (
              <Btn onClick={async () => {
                toast("info", "Uploading…", "Sending PDF to SharePoint…");
                const blob = generateEDPPdf(selectedEmp);
                const r = await uploadPDFToSharePoint(selectedEmp, blob);
                r.ok ? toast("success", "Uploaded ✓", `Saved as ${r.fileName}`) : toast("error", "Upload failed", r.reason);
              }} color={C.gold}>↑ Upload PDF to SharePoint</Btn>
            )}
            {canAct && (
              <>
                <Btn onClick={() => setShowRejectModal(true)} color={C.red}>✕ Request Correction</Btn>
                <Btn onClick={handleApprove} color={C.green}>✓ Approve EDP</Btn>
              </>
            )}
          </div>
        </div>

        {/* Reject Modal */}
        {showRejectModal && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 500 }}>
            <div style={{ background: C.white, borderRadius: 14, padding: 36, maxWidth: 500, width: "90%", boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}>
              <h3 style={{ margin: "0 0 8px", fontFamily: "'Playfair Display', serif", color: C.red }}>Request Correction</h3>
              <p style={{ margin: "0 0 16px", fontSize: 14, color: C.gray600 }}>The employee will be notified with your feedback and must correct and resubmit the form.</p>
              <textarea style={{ ...ta(false), width: "100%" }} placeholder="Explain what needs to be corrected (missing fields, incorrect info…)" value={rejectReason} onChange={e => setRejectReason(e.target.value)} />
              <div style={{ display: "flex", gap: 10, marginTop: 16, justifyContent: "flex-end" }}>
                <Btn onClick={() => { setShowRejectModal(false); setRejectReason(""); }} color={C.gray600} outline>Cancel</Btn>
                <Btn onClick={handleReject} color={C.red} disabled={!rejectReason.trim()}>Send Feedback</Btn>
              </div>
            </div>
          </div>
        )}
      </div>
    );
  }

  return null;
}
