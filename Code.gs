// ============================================================
//  FOREVER MG — CRM Backend (Google Apps Script)
//  Modules : Dashboard · Devis · Purchase Orders · Gmail
//  Version : 1.0 — Juin 2026
// ============================================================

// ---------- CONFIG ----------
const CONFIG = {
  SHEET_DEVIS: "Devis",
  SHEET_PO:    "PurchaseOrders",
  SHEET_LOG:   "Log",
  GMAIL_LABEL: "CRM-FOREVER",
  VERSION:     "1.0"
};

// ============================================================
//  POINT D'ENTRÉE HTTP
// ============================================================
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("FOREVER MG — CRM")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action;
    let result;

    switch (action) {
      // --- DEVIS ---
      case "getDevis":       result = getDevis(payload);       break;
      case "createDevis":    result = createDevis(payload);    break;
      case "updateDevis":    result = updateDevis(payload);    break;
      case "deleteDevis":    result = deleteDevis(payload);    break;
      // --- PO ---
      case "getPO":          result = getPO(payload);          break;
      case "createPO":       result = createPO(payload);       break;
      case "updatePO":       result = updatePO(payload);       break;
      // --- DASHBOARD ---
      case "getDashboard":   result = getDashboard();          break;
      // --- GMAIL ---
      case "getGmailLeads":  result = getGmailLeads();         break;
      case "sendEmail":      result = sendEmail(payload);      break;
      case "createDraft":    result = createDraft(payload);    break;
      default:
        result = { ok: false, error: "Action inconnue : " + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
//  INITIALISATION DES FEUILLES
// ============================================================
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Feuille DEVIS ---
  let devis = ss.getSheetByName(CONFIG.SHEET_DEVIS);
  if (!devis) {
    devis = ss.insertSheet(CONFIG.SHEET_DEVIS);
    const headers = [
      "ID","Date","Client","Email Client","Téléphone",
      "Catégorie","Produit / Service","Montant (Ar)",
      "Statut","Priorité","Échéance","Responsable",
      "Problème identifié","Prochaine action","Commentaire","Créé le","Modifié le"
    ];
    devis.getRange(1, 1, 1, headers.length).setValues([headers]);
    devis.getRange(1, 1, 1, headers.length)
      .setBackground("#1a4a3a").setFontColor("#ffffff").setFontWeight("bold");
    devis.setFrozenRows(1);
  }

  // --- Feuille PURCHASE ORDERS ---
  let po = ss.getSheetByName(CONFIG.SHEET_PO);
  if (!po) {
    po = ss.insertSheet(CONFIG.SHEET_PO);
    const headers = [
      "ID PO","Numéro PO","Date réception","Client","Contact",
      "Description","Montant (Ar)","Conditions paiement",
      "Statut","Échéance acquittement","Lien Coupa","Facture créée","Commentaire","Créé le"
    ];
    po.getRange(1, 1, 1, headers.length).setValues([headers]);
    po.getRange(1, 1, 1, headers.length)
      .setBackground("#1a4a3a").setFontColor("#ffffff").setFontWeight("bold");
    po.setFrozenRows(1);
  }

  // --- Feuille LOG ---
  let log = ss.getSheetByName(CONFIG.SHEET_LOG);
  if (!log) {
    log = ss.insertSheet(CONFIG.SHEET_LOG);
    log.getRange(1,1,1,4).setValues([["Timestamp","Utilisateur","Action","Détail"]]);
    log.getRange(1,1,1,4)
      .setBackground("#1a4a3a").setFontColor("#ffffff").setFontWeight("bold");
  }

  return { ok: true, message: "Feuilles initialisées." };
}

// ============================================================
//  UTILITAIRES
// ============================================================
function genId(prefix) {
  return prefix + "-" + new Date().getTime().toString(36).toUpperCase();
}

function now() {
  return Utilities.formatDate(new Date(), "Indian/Antananarivo", "dd/MM/yyyy HH:mm");
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { initSheets(); sh = ss.getSheetByName(name); }
  return sh;
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map((row, i) => {
    const obj = { _row: i + 2 };
    headers.forEach((h, j) => { obj[h] = row[j]; });
    return obj;
  });
}

function logAction(action, detail) {
  try {
    const sh = getSheet(CONFIG.SHEET_LOG);
    sh.appendRow([now(), Session.getActiveUser().getEmail(), action, detail]);
  } catch(e) {}
}

// ============================================================
//  MODULE DEVIS
// ============================================================
function getDevis(payload) {
  const sh   = getSheet(CONFIG.SHEET_DEVIS);
  const rows = sheetToObjects(sh);
  const { statut, priorite, search } = payload || {};

  let result = rows.filter(r => r["ID"] !== "");

  if (statut && statut !== "tous")
    result = result.filter(r => r["Statut"] === statut);
  if (priorite && priorite !== "tous")
    result = result.filter(r => r["Priorité"] === priorite);
  if (search) {
    const s = search.toLowerCase();
    result = result.filter(r =>
      String(r["Client"]).toLowerCase().includes(s) ||
      String(r["Produit / Service"]).toLowerCase().includes(s)
    );
  }

  // Trier : urgents en premier, puis par date
  const prio = { "Urgent": 0, "Haute": 1, "Moyenne": 2, "Basse": 3 };
  result.sort((a, b) => {
    const pa = prio[a["Priorité"]] ?? 9;
    const pb = prio[b["Priorité"]] ?? 9;
    return pa !== pb ? pa - pb : 0;
  });

  return { ok: true, data: result };
}

function createDevis(payload) {
  const sh = getSheet(CONFIG.SHEET_DEVIS);
  const id = genId("DEV");
  const ts = now();
  const row = [
    id,
    payload.date        || ts.split(" ")[0],
    payload.client      || "",
    payload.emailClient || "",
    payload.telephone   || "",
    payload.categorie   || "Demande devis",
    payload.produit     || "",
    payload.montant     || "",
    payload.statut      || "Ouvert",
    payload.priorite    || "Moyenne",
    payload.echeance    || "",
    payload.responsable || "",
    payload.probleme    || "",
    payload.action      || "",
    payload.commentaire || "",
    ts, ts
  ];
  sh.appendRow(row);
  logAction("CREATE_DEVIS", id + " — " + payload.client);
  return { ok: true, id, message: "Devis créé." };
}

function updateDevis(payload) {
  const sh   = getSheet(CONFIG.SHEET_DEVIS);
  const rows = sheetToObjects(sh);
  const item = rows.find(r => r["ID"] === payload.id);
  if (!item) return { ok: false, error: "Devis introuvable : " + payload.id };

  const fields = {
    "Client":             payload.client,
    "Email Client":       payload.emailClient,
    "Téléphone":          payload.telephone,
    "Catégorie":          payload.categorie,
    "Produit / Service":  payload.produit,
    "Montant (Ar)":       payload.montant,
    "Statut":             payload.statut,
    "Priorité":           payload.priorite,
    "Échéance":           payload.echeance,
    "Responsable":        payload.responsable,
    "Problème identifié": payload.probleme,
    "Prochaine action":   payload.action,
    "Commentaire":        payload.commentaire,
    "Modifié le":         now()
  };

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  Object.entries(fields).forEach(([key, val]) => {
    if (val !== undefined) {
      const col = headers.indexOf(key) + 1;
      if (col > 0) sh.getRange(item._row, col).setValue(val);
    }
  });

  logAction("UPDATE_DEVIS", payload.id + " — statut: " + payload.statut);
  return { ok: true, message: "Devis mis à jour." };
}

function deleteDevis(payload) {
  const sh   = getSheet(CONFIG.SHEET_DEVIS);
  const rows = sheetToObjects(sh);
  const item = rows.find(r => r["ID"] === payload.id);
  if (!item) return { ok: false, error: "Devis introuvable." };
  sh.deleteRow(item._row);
  logAction("DELETE_DEVIS", payload.id);
  return { ok: true, message: "Devis supprimé." };
}

// ============================================================
//  MODULE PURCHASE ORDERS
// ============================================================
function getPO(payload) {
  const sh   = getSheet(CONFIG.SHEET_PO);
  const rows = sheetToObjects(sh);
  const { statut } = payload || {};

  let result = rows.filter(r => r["ID PO"] !== "");
  if (statut && statut !== "tous")
    result = result.filter(r => r["Statut"] === statut);

  result.sort((a, b) => {
    const order = { "En retard": 0, "À acquitter": 1, "À facturer": 2,
                    "Acquitté":  3, "Facturé":     4, "Clôturé":    5 };
    return (order[a["Statut"]] ?? 9) - (order[b["Statut"]] ?? 9);
  });

  return { ok: true, data: result };
}

function createPO(payload) {
  const sh = getSheet(CONFIG.SHEET_PO);
  const id = genId("PO");
  const ts = now();
  const row = [
    id,
    payload.numeroPO       || "",
    payload.dateReception  || ts.split(" ")[0],
    payload.client         || "",
    payload.contact        || "",
    payload.description    || "",
    payload.montant        || "",
    payload.conditions     || "NET 30",
    payload.statut         || "À acquitter",
    payload.echeance       || "",
    payload.lienCoupa      || "",
    payload.factureCreee   || "Non",
    payload.commentaire    || "",
    ts
  ];
  sh.appendRow(row);
  logAction("CREATE_PO", id + " — " + payload.numeroPO);
  return { ok: true, id, message: "PO créé." };
}

function updatePO(payload) {
  const sh   = getSheet(CONFIG.SHEET_PO);
  const rows = sheetToObjects(sh);
  const item = rows.find(r => r["ID PO"] === payload.id);
  if (!item) return { ok: false, error: "PO introuvable : " + payload.id };

  const fields = {
    "Statut":               payload.statut,
    "Facture créée":        payload.factureCreee,
    "Commentaire":          payload.commentaire,
    "Échéance acquittement":payload.echeance,
    "Lien Coupa":           payload.lienCoupa
  };

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  Object.entries(fields).forEach(([key, val]) => {
    if (val !== undefined) {
      const col = headers.indexOf(key) + 1;
      if (col > 0) sh.getRange(item._row, col).setValue(val);
    }
  });

  logAction("UPDATE_PO", payload.id + " → " + payload.statut);
  return { ok: true, message: "PO mis à jour." };
}

// ============================================================
//  MODULE DASHBOARD — KPIs agrégés
// ============================================================
function getDashboard() {
  const devisRows = sheetToObjects(getSheet(CONFIG.SHEET_DEVIS))
    .filter(r => r["ID"] !== "");
  const poRows    = sheetToObjects(getSheet(CONFIG.SHEET_PO))
    .filter(r => r["ID PO"] !== "");

  // -- KPIs Devis --
  const totalDevis       = devisRows.length;
  const devisOuverts     = devisRows.filter(r => r["Statut"] === "Ouvert").length;
  const devisUrgents     = devisRows.filter(r => r["Priorité"] === "Urgent").length;
  const devisEnCours     = devisRows.filter(r => r["Statut"] === "En cours").length;
  const devisValides     = devisRows.filter(r => r["Statut"] === "Validé").length;
  const devisNonRepondus = devisRows.filter(r => r["Statut"] === "Ouvert" && r["Priorité"] === "Urgent").length;

  const montantDevisOuverts = devisRows
    .filter(r => ["Ouvert","En cours"].includes(r["Statut"]))
    .reduce((s, r) => s + (parseFloat(String(r["Montant (Ar)"]).replace(/\s/g,"")) || 0), 0);

  // -- KPIs PO --
  const totalPO        = poRows.length;
  const poEnRetard     = poRows.filter(r => r["Statut"] === "En retard").length;
  const poAAcquitter   = poRows.filter(r => r["Statut"] === "À acquitter").length;
  const poAFacturer    = poRows.filter(r => r["Statut"] === "À facturer").length;
  const poActifs       = poRows.filter(r => !["Clôturé","Facturé"].includes(r["Statut"]));

  const montantPOActifs = poActifs
    .reduce((s, r) => s + (parseFloat(String(r["Montant (Ar)"]).replace(/\s/g,"")) || 0), 0);

  // -- Urgences (relances à faire aujourd'hui) --
  const today    = Utilities.formatDate(new Date(), "Indian/Antananarivo", "dd/MM/yyyy");
  const relances = devisRows.filter(r =>
    r["Échéance"] === today ||
    (r["Priorité"] === "Urgent" && ["Ouvert","En cours"].includes(r["Statut"]))
  );

  // -- 5 derniers devis --
  const recentDevis = devisRows
    .sort((a,b) => String(b["Créé le"]).localeCompare(String(a["Créé le"])))
    .slice(0, 5);

  // -- 5 derniers PO --
  const recentPO = poRows
    .sort((a,b) => String(b["Créé le"]).localeCompare(String(a["Créé le"])))
    .slice(0, 5);

  return {
    ok: true,
    kpi: {
      devis: { total: totalDevis, ouverts: devisOuverts, urgents: devisUrgents,
               enCours: devisEnCours, valides: devisValides,
               nonRepondus: devisNonRepondus, montantOuverts: montantDevisOuverts },
      po:    { total: totalPO, enRetard: poEnRetard, aAcquitter: poAAcquitter,
               aFacturer: poAFacturer, montantActifs: montantPOActifs }
    },
    relances,
    recentDevis,
    recentPO
  };
}

// ============================================================
//  MODULE GMAIL
// ============================================================
function getGmailLeads() {
  try {
    const queries = [
      "subject:devis is:unread",
      "subject:demande is:unread",
      "subject:commande is:unread",
      "subject:PO is:unread",
      "subject:urgent is:unread",
      "subject:prix is:unread",
      "subject:cotation is:unread"
    ];

    const seen    = {};
    const results = [];

    queries.forEach(q => {
      const threads = GmailApp.search(q, 0, 10);
      threads.forEach(thread => {
        if (seen[thread.getId()]) return;
        seen[thread.getId()] = true;
        const msg  = thread.getMessages()[thread.getMessageCount() - 1];
        const from = msg.getFrom();
        results.push({
          threadId:  thread.getId(),
          subject:   thread.getFirstMessageSubject(),
          from,
          date:      Utilities.formatDate(msg.getDate(), "Indian/Antananarivo", "dd/MM/yyyy HH:mm"),
          snippet:   thread.getMessages()[0].getPlainBody().substring(0, 200),
          unread:    thread.isUnread(),
          important: thread.isImportant()
        });
      });
    });

    // Trier : importants + non lus en tête
    results.sort((a,b) => {
      if (a.important && !b.important) return -1;
      if (!a.important && b.important) return  1;
      if (a.unread && !b.unread) return -1;
      if (!a.unread && b.unread) return  1;
      return 0;
    });

    return { ok: true, data: results.slice(0, 20) };
  } catch(err) {
    return { ok: false, error: "Gmail inaccessible : " + err.message };
  }
}

function sendEmail(payload) {
  try {
    GmailApp.sendEmail(
      payload.to,
      payload.subject,
      payload.body,
      {
        cc:       payload.cc       || "",
        replyTo:  payload.replyTo  || "",
        name:     "FOREVER MG"
      }
    );
    logAction("SEND_EMAIL", "À: " + payload.to + " | Sujet: " + payload.subject);
    return { ok: true, message: "E-mail envoyé." };
  } catch(err) {
    return { ok: false, error: err.message };
  }
}

function createDraft(payload) {
  try {
    GmailApp.createDraft(
      payload.to,
      payload.subject,
      payload.body,
      { cc: payload.cc || "", name: "FOREVER MG" }
    );
    logAction("CREATE_DRAFT", "À: " + payload.to + " | Sujet: " + payload.subject);
    return { ok: true, message: "Brouillon créé dans Gmail." };
  } catch(err) {
    return { ok: false, error: err.message };
  }
}

// ============================================================
//  TRIGGER : Vérification quotidienne des relances
// ============================================================
function checkRelancesQuotidiennes() {
  const devisRows = sheetToObjects(getSheet(CONFIG.SHEET_DEVIS))
    .filter(r => r["ID"] !== "");
  const today     = Utilities.formatDate(new Date(), "Indian/Antananarivo", "dd/MM/yyyy");

  const urgents = devisRows.filter(r =>
    r["Statut"] === "Ouvert" && r["Priorité"] === "Urgent"
  );

  if (urgents.length > 0) {
    const user  = Session.getActiveUser().getEmail();
    const lines = urgents.map(r =>
      `• ${r["Client"]} — ${r["Produit / Service"]} (${r["Échéance"] || "pas d'échéance"})`
    ).join("\n");

    GmailApp.sendEmail(user,
      `[FOREVER MG CRM] ${urgents.length} devis urgent(s) à traiter`,
      `Bonjour,\n\nVoici les devis urgents en attente :\n\n${lines}\n\n` +
      `Accédez au CRM pour traiter ces demandes.\n\n— FOREVER MG CRM`
    );
  }
}

// ============================================================
//  INSTALLATION DU TRIGGER QUOTIDIEN
// ============================================================
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("checkRelancesQuotidiennes")
    .timeBased().everyDays(1).atHour(8).create();
  return { ok: true, message: "Trigger quotidien installé (08h00)." };
}
