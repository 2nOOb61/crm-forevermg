// ============================================================
//  FOREVER MG — CRM + FACTURATION (Code.gs fusionné)
//  Modules : CRM (Dashboard · Devis · PO · Gmail)
//           + Facturation (Clients · Devis · Factures · BL · POS)
//  Version : 2.0 — Juin 2026
// ============================================================


var SPREADSHEET_ID = ""; // Laissez vide si le script est lié au Sheets

// ---------- CONFIG CRM ----------
const CONFIG = {
  SHEET_DEVIS: "Devis",
  SHEET_PO:    "PurchaseOrders",
  SHEET_LOG:   "Log",
  VERSION:     "2.0"
};

// ============================================================
//  POINT D'ENTRÉE HTTP — ROUTEUR UNIFIÉ
// ============================================================
function doGet(e) {
  // Mode facturation (FacturePro)
  if (e && e.parameter && e.parameter.interface === "1") {
    return HtmlService.createHtmlOutputFromFile('facturation')
      .setTitle('FOREVER MG — Facturation')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (e && e.parameter && e.parameter.action) {
    return handleSyncGet_(e);
  }
  if (e && e.parameter && e.parameter.data === "1") {
    var token = e.parameter.token || "";
    var user = requireAuth(token);
    var data = sanitizeDataForRole_(user.role, getAllDataInternal());
    return ContentService
      .createTextOutput(JSON.stringify(data, null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.ping === "1") {
    return ContentService
      .createTextOutput(JSON.stringify(healthCheck(), null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // Mode CRM (PWA principale)
  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("FOREVER MG — CRM")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onInstall(e) { onOpen(e); }

function doPost(e) {
  // Mode facturation (sync)
  if (e && e.parameter && e.parameter.action) {
    return handleSyncPost_(e);
  }
  // Mode CRM (JSON API)
  try {
    const payload = JSON.parse(e.postData.contents);
    const cmd     = payload.cmd;

    // --- Authentification : token requis sauf pour les commandes publiques ---
    const PUBLIC_CMDS = ["login", "hasUsers", "createFirstAdmin"];
    if (PUBLIC_CMDS.indexOf(cmd) === -1) {
      const authUser = validateToken(payload.token);
      if (!authUser) {
        return ContentService
          .createTextOutput(JSON.stringify({ ok: false, error: "AUTH_REQUIRED" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      payload.__user = authUser; // utilisateur authentifié disponible aux handlers
    }

    let result;
    switch (cmd) {
      case "login":                result = authenticate(payload.username, payload.password); break;
      case "hasUsers":             result = { ok: true, hasUsers: hasUsers() };               break;
      case "createFirstAdmin":     result = createFirstAdmin(payload.username, payload.password, payload.displayName); break;
      case "checkAuth":            result = { ok: true, user: payload.__user };               break;
      case "getUsers":             result = crmGetUsers(payload);          break;
      case "saveUser":             result = crmSaveUser(payload);          break;
      case "getDevis":             result = getDevis(payload);             break;
      case "createDevis":          result = createDevis(payload);          break;
      case "updateDevis":          result = updateDevis(payload);          break;
      case "deleteDevis":          result = deleteDevis(payload);          break;
      case "getPO":                result = getPO(payload);                break;
      case "createPO":             result = createPO(payload);             break;
      case "updatePO":             result = updatePO(payload);             break;
      case "getDashboard":         result = getDashboard();                break;
      case "getGmailLeads":        result = getGmailLeads();               break;
      case "createDevisFromEmail":   result = createDevisFromEmail(payload);         break;
      case "sendEmail":              result = sendEmail(payload);                    break;
      case "createDraft":            result = createDraft(payload);                  break;
      case "createFactDocFromCrm":   result = createFactDocFromCrm(payload);         break;
      default:
        result = { ok: false, error: "Commande inconnue : " + cmd };
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

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- Gestion utilisateurs côté CRM (réutilise listUsers/saveUser, admin only) ---
function crmGetUsers(payload) {
  try {
    return { ok: true, users: listUsers(payload.token) };
  } catch (e) {
    return { ok: false, error: e.message === "FORBIDDEN" ? "Réservé aux administrateurs." : e.message };
  }
}

function crmSaveUser(payload) {
  try {
    var res = saveUser(payload.token, payload.user || {});
    if (res && res.error) {
      var map = { EMAIL_REQUIRED: "Email/identifiant requis.", PASSWORD_REQUIRED: "Mot de passe requis pour un nouvel utilisateur." };
      res.error = map[res.error] || res.error;
    }
    return res;
  } catch (e) {
    return { ok: false, error: e.message === "FORBIDDEN" ? "Réservé aux administrateurs." : e.message };
  }
}

// ============================================================
//  INTÉGRATION CRM → FACTURATION
// ============================================================

function createFactDocFromCrm(payload) {
  var devisId = String(payload.devisId || "");
  var docType = String(payload.docType || "quote"); // "quote" | "invoice"
  if (!devisId) return { ok: false, error: "devisId requis" };

  var sh = getSheet(CONFIG.SHEET_DEVIS);
  var rows = sheetToObjects(sh);
  var devis = rows.find(function(r) { return r["ID"] === devisId; });
  if (!devis) return { ok: false, error: "Devis CRM introuvable : " + devisId };

  var clientId = findOrCreateFactClient_(devis);
  if (!clientId) return { ok: false, error: "Nom de client manquant" };

  var items = [];
  try {
    var raw = String(devis["Items"] || "").trim();
    if (raw.startsWith("[")) items = JSON.parse(raw);
  } catch(e) {}
  if (!items.length) {
    var produit = String(devis["Produit / Service"] || "").trim();
    if (produit) items.push({ description: produit, qty: 1, unitPrice: toNumber(devis["Montant (Ar)"] || 0) });
  }

  var noteParts = [];
  if (devis["Commentaire"])          noteParts.push(String(devis["Commentaire"]));
  if (devis["Problème identifié"])   noteParts.push("Problème : " + String(devis["Problème identifié"]));
  if (devis["Prochaine action"])     noteParts.push("Action : " + String(devis["Prochaine action"]));

  var doc = {
    clientId: clientId,
    date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"),
    status: docType === "invoice" ? "pending" : "draft",
    items: items,
    notes: noteParts.join(" | "),
    acompte: Math.max(toNumber(devis["Acompte"] || 0), 0),
    discountRate: Math.min(Math.max(toNumber(devis["Remise"] || 0), 0), 100)
  };

  var saved;
  try {
    saved = docType === "invoice" ? saveInvoiceInternal(doc) : saveQuoteInternal(doc);
  } catch(e) {
    return { ok: false, error: "Erreur création document : " + e.message };
  }

  updateDevisFactRef_(sh, devis._row, (docType === "invoice" ? "FA:" : "DV:") + saved.number, saved.id);
  return { ok: true, docId: saved.id, docNumber: saved.number, type: docType };
}

// Lie le devis à Facturation s'il ne l'est pas encore, sinon le synchronise.
function linkOrSyncDevisToFact_(devisId) {
  var sh = getSheet(CONFIG.SHEET_DEVIS);
  var devis = sheetToObjects(sh).find(function(r) { return r["ID"] === devisId; });
  if (!devis) return;
  var docId = String(devis["FacturationDocID"] || "").trim();
  if (docId) {
    syncDevisToFact_(devisId);            // déjà lié → synchroniser
    return;
  }
  // Pas encore lié → créer un Devis Facturation si le montant est significatif
  if (devisFactTotal_(devis) > 0) {
    createFactDocFromCrm({ devisId: devisId, docType: "quote" });
  }
}

// Total d'un devis CRM (depuis les articles JSON, sinon depuis Montant)
function devisFactTotal_(devis) {
  var total = 0;
  try {
    var raw = String(devis["Items"] || "").trim();
    if (raw.charAt(0) === "[") {
      JSON.parse(raw).forEach(function(it) {
        total += (Number(it.qty) || 0) * (Number(it.unitPrice != null ? it.unitPrice : it.pu) || 0);
      });
    }
  } catch (e) {}
  if (!total) total = toNumber(devis["Montant (Ar)"] || 0);
  return total;
}

function syncDevisToFact_(devisId) {
  // Lire le devis CRM
  var sh = getSheet(CONFIG.SHEET_DEVIS);
  var devis = sheetToObjects(sh).find(function(r) { return r["ID"] === devisId; });
  if (!devis) return;

  var factDocId = String(devis["FacturationDocID"] || "").trim();
  var factRef   = String(devis["FacturationRef"]   || "").trim();
  if (!factDocId || !factRef) return;

  // Accès direct à la feuille Facturation (Quotes ou Invoices)
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = factRef.substring(0, 3) === "FA:" ? "Invoices" : "Quotes";
  var docSheet  = ss.getSheetByName(sheetName);
  if (!docSheet) return;

  var data    = docSheet.getDataRange().getValues();
  if (data.length < 2) return;
  var headers = data[0];

  // Trouver la ligne par ID
  var idCol = headers.indexOf("ID");
  if (idCol < 0) return;
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol] || "").trim() === factDocId) { rowIndex = i + 1; break; }
  }
  if (rowIndex < 0) return;
  var rowData = data[rowIndex - 1]; // données actuelles (0-based)

  // Mettre à jour le client
  var clientId = findOrCreateFactClient_(devis);
  var clientCol = headers.indexOf("ClientID");
  if (clientId && clientCol >= 0) docSheet.getRange(rowIndex, clientCol + 1).setValue(clientId);

  // Mettre à jour les articles depuis le CRM (remplace tout)
  var syncItems = [];
  try {
    var rawItems = String(devis["Items"] || "").trim();
    if (rawItems.startsWith("[")) syncItems = JSON.parse(rawItems);
  } catch(e) {}
  if (!syncItems.length) {
    var sp = String(devis["Produit / Service"] || "").trim();
    if (sp) syncItems.push({ description: sp, qty: 1, unitPrice: toNumber(devis["Montant (Ar)"] || 0) });
  }
  // Remise / Acompte depuis le devis CRM
  var dr = Math.min(Math.max(toNumber(devis["Remise"] || 0), 0), 100);
  var acompte = Math.max(toNumber(devis["Acompte"] || 0), 0);
  var discountCol = headers.indexOf("DiscountRate");
  if (discountCol >= 0) docSheet.getRange(rowIndex, discountCol + 1).setValue(dr);
  var acompteCol = headers.indexOf("Acompte");
  if (acompteCol >= 0) docSheet.getRange(rowIndex, acompteCol + 1).setValue(acompte);

  if (syncItems.length) {
    var itemsCol = headers.indexOf("Items");
    if (itemsCol < 0) {
      itemsCol = docSheet.getLastColumn();
      docSheet.getRange(1, itemsCol + 1).setValue("Items");
    }
    docSheet.getRange(rowIndex, itemsCol + 1).setValue(JSON.stringify(syncItems));

    // Recalculer les totaux avec la remise du CRM
    var totalsCol = headers.indexOf("Totals");
    if (totalsCol >= 0) {
      docSheet.getRange(rowIndex, totalsCol + 1).setValue(JSON.stringify(computeTotalsServer(syncItems, dr)));
    }
  }

  // Mettre à jour les notes
  var notesCol = headers.indexOf("Notes");
  if (notesCol >= 0) {
    var parts = [];
    if (devis["Commentaire"])        parts.push(String(devis["Commentaire"]));
    if (devis["Problème identifié"]) parts.push("Problème : " + String(devis["Problème identifié"]));
    if (devis["Prochaine action"])   parts.push("Action : " + String(devis["Prochaine action"]));
    if (parts.length) docSheet.getRange(rowIndex, notesCol + 1).setValue(parts.join(" | "));
  }

  logAction("SYNC_DEVIS_FACT", devisId + " → " + factDocId);
}

function findOrCreateFactClient_(devis) {
  var name = String(devis["Client"] || "").trim();
  if (!name) return "";
  var sheet = getOrCreateSheet('Clients', ['ID','Name','Contact','Address','Phone','NIF','STAT']);
  var data = sheet.getDataRange().getValues();
  var nameIdx = data[0].indexOf('Name');
  if (nameIdx < 0) return "";
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][nameIdx] || "").trim().toLowerCase() === name.toLowerCase())
      return String(data[i][0]);
  }
  var id = 'CLI-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  sheet.appendRow([id, name, String(devis["Email Client"] || ""), "", String(devis["Téléphone"] || ""), "", ""]);
  return id;
}

function updateDevisFactRef_(sheet, rowIndex, factRef, factDocId) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var refCol = headers.indexOf("FacturationRef") + 1;
  var idCol  = headers.indexOf("FacturationDocID") + 1;
  if (refCol === 0) { refCol = ++lastCol; sheet.getRange(1, refCol).setValue("FacturationRef"); }
  if (idCol  === 0) { idCol  = ++lastCol; sheet.getRange(1, idCol).setValue("FacturationDocID"); }
  sheet.getRange(rowIndex, refCol).setValue(factRef);
  sheet.getRange(rowIndex, idCol).setValue(factDocId || "");
}

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
  var newRow = sh.getLastRow();
  if (payload.items) setDevisCol_(sh, newRow, "Items", payload.items);
  setDevisCol_(sh, newRow, "Remise", toNumber(payload.remise || 0));
  setDevisCol_(sh, newRow, "Acompte", toNumber(payload.acompte || 0));
  logAction("CREATE_DEVIS", id + " — " + payload.client);
  try { linkOrSyncDevisToFact_(id); } catch(e) { Logger.log("Link warn: " + e.message); }
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
  // Colonnes dynamiques (créées si absentes)
  if (payload.items !== undefined)   setDevisCol_(sh, item._row, "Items", payload.items);
  if (payload.remise !== undefined)  setDevisCol_(sh, item._row, "Remise", toNumber(payload.remise || 0));
  if (payload.acompte !== undefined) setDevisCol_(sh, item._row, "Acompte", toNumber(payload.acompte || 0));

  logAction("UPDATE_DEVIS", payload.id + " — statut: " + payload.statut);
  try { linkOrSyncDevisToFact_(payload.id); } catch(e) { Logger.log("Sync warn: " + e.message); }
  return { ok: true, message: "Devis mis à jour." };
}

// Écrit une valeur dans une colonne de la feuille Devis, en créant l'en-tête si absent
function setDevisCol_(sh, rowNum, header, value) {
  var hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var col = hdrs.indexOf(header) + 1;
  if (col === 0) { col = sh.getLastColumn() + 1; sh.getRange(1, col).setValue(header); }
  sh.getRange(rowNum, col).setValue(value);
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
//  MODULE GMAIL — v2 (mots-clés étendus + 2 boîtes + doublons)
// ============================================================

// Tous les mots-clés commerciaux détectés automatiquement
const GMAIL_KEYWORDS = [
  // Devis & prix
  "devis","cotation","proforma","pro-forma","offre de prix","tarif","prix",
  // Commandes
  "commande","purchase order","bon de commande","PO","order",
  // Demandes
  "demande","demande de","requête","besoin","souhait",
  // Produits FOREVER MG
  "impression","sticker","plaque","signalétique","signalisation",
  "lampion","claustra","étiquette","badge","kakémono","roll-up",
  "photo","vidéo","tournage","shooting","maquette","design",
  "personnalisé","logo","broderie","gravure","sérigraphie",
  // Facturation
  "facture","facturation","paiement","règlement","acompte","solde",
  // Urgences
  "urgent","urgence","asap","dès que possible","rapidement",
  // Livraison
  "livraison","délai","réception",
  // Relances
  "relance","suivi","rappel","sans réponse"
];

// Catégorisation automatique par mots-clés dans le sujet
function categorizeEmail(subject, snippet) {
  const txt = (subject + " " + snippet).toLowerCase();
  if (/devis|cotation|proforma|offre de prix|tarif/.test(txt))    return "Demande devis";
  if (/commande|purchase order|bon de commande|\bpo\b|order/.test(txt)) return "Commande";
  if (/facture|facturation|proforma/.test(txt))                   return "Facturation";
  if (/paiement|règlement|acompte|solde/.test(txt))               return "Paiement";
  if (/livraison|délai|réception/.test(txt))                      return "Livraison";
  if (/relance|rappel|sans réponse|suivi/.test(txt))              return "Relance";
  if (/urgent|urgence|asap/.test(txt))                            return "Urgent";
  if (/réclamation|problème|erreur|anomalie/.test(txt))           return "Réclamation";
  return "Autre";
}

// Priorité automatique
function autoPriority(subject, snippet, isImportant) {
  const txt = (subject + " " + snippet).toLowerCase();
  if (/urgent|urgence|asap|immédiat/.test(txt)) return "Urgent";
  if (isImportant)                               return "Haute";
  if (/devis|commande|proforma|facture/.test(txt)) return "Haute";
  return "Moyenne";
}

// Clé de déduplication : expéditeur normalisé + sujet normalisé
function dedupeKey(from, subject) {
  const emailMatch = from.match(/<(.+?)>/);
  const email = emailMatch ? emailMatch[1].toLowerCase() : from.toLowerCase();
  const subj  = subject.toLowerCase()
    .replace(/^(re|fwd|fw|tr|aw)[\s:]+/gi, "")
    .replace(/\s+/g, " ").trim()
    .substring(0, 60);
  return email + "|" + subj;
}

function getGmailLeads() {
  try {
    // Construire une requête globale OR avec tous les mots-clés
    const kwQuery = GMAIL_KEYWORDS
      .map(k => k.includes(" ") ? `"${k}"` : `subject:${k}`)
      .join(" OR ");

    // On cherche dans les 2 boîtes via l'alias (GAS opère sur le compte connecté)
    // + on cible aussi les e-mails envoyés À commercial ou contact
    const baseQueries = [
      `(${kwQuery}) newer_than:90d`,
      `(${kwQuery}) to:(commercial4evermg@gmail.com OR contact4evermg@gmail.com) newer_than:90d`,
      `is:important newer_than:30d`,
      `is:unread is:important newer_than:30d`
    ];

    const seen    = {}; // dédup par threadId
    const dedupeS = {}; // dédup par expéditeur+sujet (doublons cross-boîtes)
    const results = [];
    const doublons = [];

    baseQueries.forEach(q => {
      let threads = [];
      try { threads = GmailApp.search(q, 0, 25); } catch(e) {}

      threads.forEach(thread => {
        if (seen[thread.getId()]) return;
        seen[thread.getId()] = true;

        const msgs    = thread.getMessages();
        const lastMsg = msgs[msgs.length - 1];
        const firstMsg= msgs[0];
        const from    = firstMsg.getFrom();
        const subject = thread.getFirstMessageSubject();
        const snippet = firstMsg.getPlainBody().substring(0, 300);
        const dk      = dedupeKey(from, subject);
        const categorie = categorizeEmail(subject, snippet);
        const priorite  = autoPriority(subject, snippet, thread.isImportant());

        // Détection doublons (même expéditeur + sujet similaire)
        const isDuplicate = !!dedupeS[dk];
        if (isDuplicate) {
          doublons.push({ threadId: thread.getId(), subject, from, note: "Doublon détecté" });
          return; // on ne l'ajoute pas aux résultats principaux
        }
        dedupeS[dk] = true;

        // Détecter si l'e-mail cible contact ou commercial
        const toRecipients = lastMsg.getTo() + " " + lastMsg.getCc();
        const boite = toRecipients.includes("commercial4evermg") ? "commercial" :
                      toRecipients.includes("contact4evermg")    ? "contact"    : "principale";

        results.push({
          threadId:   thread.getId(),
          subject,
          from,
          boite,
          categorie,
          priorite,
          date:       Utilities.formatDate(lastMsg.getDate(), "Indian/Antananarivo", "dd/MM/yyyy HH:mm"),
          snippet:    snippet.substring(0, 250),
          unread:     thread.isUnread(),
          important:  thread.isImportant(),
          msgCount:   msgs.length,
          isDuplicate: false
        });
      });
    });

    // Trier : urgents > importants > non lus > date
    const prioOrder = { "Urgent": 0, "Haute": 1, "Moyenne": 2, "Basse": 3 };
    results.sort((a, b) => {
      const pa = prioOrder[a.priorite] ?? 9;
      const pb = prioOrder[b.priorite] ?? 9;
      if (pa !== pb) return pa - pb;
      if (a.important && !b.important) return -1;
      if (!a.important && b.important) return  1;
      if (a.unread && !b.unread) return -1;
      if (!a.unread && b.unread) return  1;
      return 0;
    });

    return {
      ok:       true,
      data:     results.slice(0, 40),
      doublons: doublons,
      stats: {
        total:    results.length,
        unread:   results.filter(r => r.unread).length,
        urgent:   results.filter(r => r.priorite === "Urgent").length,
        doublons: doublons.length,
        parBoite: {
          commercial: results.filter(r => r.boite === "commercial").length,
          contact:    results.filter(r => r.boite === "contact").length,
          principale: results.filter(r => r.boite === "principale").length
        },
        parCategorie: GMAIL_KEYWORDS.reduce((acc, _) => acc, {
          "Demande devis": results.filter(r => r.categorie === "Demande devis").length,
          "Commande":      results.filter(r => r.categorie === "Commande").length,
          "Facturation":   results.filter(r => r.categorie === "Facturation").length,
          "Urgent":        results.filter(r => r.categorie === "Urgent").length,
          "Relance":       results.filter(r => r.categorie === "Relance").length,
          "Autre":         results.filter(r => r.categorie === "Autre").length
        })
      }
    };
  } catch(err) {
    return { ok: false, error: "Gmail inaccessible : " + err.message };
  }
}

// Créer un devis directement depuis un e-mail Gmail
function createDevisFromEmail(payload) {
  try {
    const thread  = GmailApp.getThreadById(payload.threadId);
    if (!thread) return { ok: false, error: "Thread introuvable." };
    const msg     = thread.getMessages()[0];
    const from    = msg.getFrom();
    const emailMatch = from.match(/<(.+?)>/);
    const email   = emailMatch ? emailMatch[1] : from;
    const nameMatch  = from.match(/^([^<]+)</);
    const name    = nameMatch ? nameMatch[1].trim() : from;

    return createDevis({
      client:      payload.client      || name,
      emailClient: payload.emailClient || email,
      produit:     payload.produit     || thread.getFirstMessageSubject(),
      categorie:   payload.categorie   || "Demande devis",
      statut:      "Ouvert",
      priorite:    payload.priorite    || "Haute",
      commentaire: "Créé depuis Gmail — " + thread.getFirstMessageSubject(),
      action:      "Répondre et envoyer devis"
    });
  } catch(err) {
    return { ok: false, error: err.message };
  }
}

function sendEmail(payload) {
  try {
    if (!payload.to) return { ok: false, error: "Destinataire requis." };
    var atts = buildAttachments_(payload);
    var opts = { cc: payload.cc || "", replyTo: payload.replyTo || "", name: "FOREVER MG" };
    if (atts.length) opts.attachments = atts;
    GmailApp.sendEmail(payload.to, payload.subject, payload.body, opts);
    logAction("SEND_EMAIL", "À: " + payload.to + " | Sujet: " + payload.subject + (atts.length ? " | PJ: " + atts.length : ""));
    return { ok: true, message: "E-mail envoyé." + (atts.length ? " (" + atts.length + " pièce(s) jointe(s))" : "") };
  } catch(err) {
    return { ok: false, error: err.message };
  }
}

function createDraft(payload) {
  try {
    var atts = buildAttachments_(payload);
    var opts = { cc: payload.cc || "", name: "FOREVER MG" };
    if (atts.length) opts.attachments = atts;
    GmailApp.createDraft(payload.to, payload.subject, payload.body, opts);
    logAction("CREATE_DRAFT", "À: " + payload.to + " | Sujet: " + payload.subject);
    return { ok: true, message: "Brouillon créé dans Gmail." + (atts.length ? " (" + atts.length + " PJ)" : "") };
  } catch(err) {
    return { ok: false, error: err.message };
  }
}

// Construit la liste des pièces jointes : PDF du devis (si payload.devisId) + fichiers base64
function buildAttachments_(payload) {
  var atts = [];
  if (payload.devisId) {
    try {
      var pdf = buildDevisPdfBlob_(payload.devisId);
      if (pdf) atts.push(pdf);
    } catch (e) { Logger.log("PDF devis échec: " + e.message); }
  }
  (payload.attachments || []).forEach(function(a) {
    try {
      if (a && a.data) {
        var bytes = Utilities.base64Decode(a.data);
        atts.push(Utilities.newBlob(bytes, a.mimeType || "application/octet-stream", a.name || "piece-jointe"));
      }
    } catch (e) { Logger.log("PJ échec: " + e.message); }
  });
  return atts;
}

// Génère un PDF du devis à partir de la feuille Devis
function buildDevisPdfBlob_(devisId) {
  var devis = sheetToObjects(getSheet(CONFIG.SHEET_DEVIS)).find(function(r) { return r["ID"] === devisId; });
  if (!devis) return null;
  var html = buildDevisHtml_(devis);
  var name = "Devis-" + String(devis["ID"] || "").replace(/[^A-Za-z0-9_-]/g, "") + ".pdf";
  return Utilities.newBlob(html, "text/html", "devis.html").getAs("application/pdf").setName(name);
}

function escHtml_(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function fmtAr_(n) {
  var num = Math.round(Number(n) || 0);
  return String(num).replace(/\B(?=(\d{3})+(?!\d))/g, " ");
}
function fmtDate_(v) {
  if (!v && v !== 0) return "";
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return String(v);
}

function buildDevisHtml_(devis) {
  var items = [];
  try {
    var raw = String(devis["Items"] || "").trim();
    if (raw.charAt(0) === "[") items = JSON.parse(raw);
  } catch (e) {}
  if (!items.length) {
    var p = String(devis["Produit / Service"] || "");
    if (p) items = [{ description: p, qty: 1, unitPrice: toNumber(devis["Montant (Ar)"] || 0) }];
  }
  var normItems = items.map(function(it) {
    return { description: it.description || "", qty: Number(it.qty) || 0, unitPrice: Number(it.unitPrice != null ? it.unitPrice : it.pu) || 0 };
  });
  var remise  = Math.min(Math.max(toNumber(devis["Remise"] || 0), 0), 100);
  var acompte = Math.max(toNumber(devis["Acompte"] || 0), 0);
  var totals    = computeTotalsServer(normItems, remise);
  var subTotal  = totals.subTotal;
  var remiseAmt = Math.round(subTotal * remise / 100);
  var totalTTC  = totals.totalTTC;
  var reste     = Math.max(0, totalTTC - acompte);

  var company = getCompanyInfoInternal() || {};
  var cName  = company.name  || "FOREVER MG";
  var cAddr  = company.address || "Antananarivo, Madagascar";
  var cPhone = company.phone || "+261 34 03 767 69";
  var cEmail = company.email || "commercial4evermg@gmail.com";
  var ids = [];
  if (company.nif)  ids.push("NIF : " + company.nif);
  if (company.stat) ids.push("STAT : " + company.stat);
  var logoHtml = company.logoDataUrl
    ? '<img src="' + escHtml_(company.logoDataUrl) + '" style="max-height:64px;max-width:170px;object-fit:contain" alt="logo">'
    : '<div class="brand">4EVER<span>MG</span></div>';

  var rows = normItems.map(function(it) {
    var lt = it.qty * it.unitPrice;
    return '<tr><td>' + escHtml_(it.description) + '</td>'
      + '<td class="c">' + it.qty + '</td>'
      + '<td class="r">' + fmtAr_(it.unitPrice) + ' Ar</td>'
      + '<td class="r">' + fmtAr_(lt) + ' Ar</td></tr>';
  }).join("");
  if (!rows) rows = '<tr><td colspan="4" style="text-align:center;color:#9ca3af">Aucun article</td></tr>';

  var totalsRows = '<div class="trow"><span>Sous-total</span><strong>' + fmtAr_(subTotal) + ' Ar</strong></div>';
  if (remiseAmt > 0) totalsRows += '<div class="trow"><span>Remise (' + remise + ' %)</span><strong>&minus; ' + fmtAr_(remiseAmt) + ' Ar</strong></div>';
  totalsRows += '<div class="trow grand"><span>Total TTC</span><strong>' + fmtAr_(totalTTC) + ' Ar</strong></div>';
  if (acompte > 0) {
    totalsRows += '<div class="trow"><span>Acompte versé</span><strong>&minus; ' + fmtAr_(acompte) + ' Ar</strong></div>';
    totalsRows += '<div class="trow grand"><span>Reste à payer</span><strong>' + fmtAr_(reste) + ' Ar</strong></div>';
  }

  var bankLines = [];
  if (company.bankName)    bankLines.push(escHtml_(company.bankName));
  if (company.bankAccount) bankLines.push(escHtml_(company.bankAccount));
  if (company.bankIban)    bankLines.push("IBAN : " + escHtml_(company.bankIban));
  var footLeft = "";
  if (bankLines.length)       footLeft += '<div><strong>Banque</strong><br>' + bankLines.join("<br>") + '</div>';
  if (company.paymentTerms)   footLeft += '<div style="margin-top:6px"><strong>Conditions</strong><br>' + escHtml_(company.paymentTerms) + '</div>';

  return '<!DOCTYPE html><html lang="fr"><head><meta charset="utf-8"><style>'
    + 'body{font-family:Arial,Helvetica,sans-serif;color:#1e293b;background:#fff;margin:28px;font-size:13px}'
    + '.wrap{max-width:860px;margin:0 auto}'
    + '.grid{display:flex;justify-content:space-between;gap:20px;border-bottom:3px solid #1a4a3a;padding-bottom:16px;margin-bottom:18px}'
    + '.brand{font-size:26px;font-weight:800;color:#1a4a3a}.brand span{color:#e8834a}'
    + '.co{font-size:12px;color:#64748b;margin-top:8px;line-height:1.5}.co strong{color:#1e293b}'
    + '.right{text-align:right}'
    + '.title{font-size:24px;font-weight:800;color:#1a4a3a;letter-spacing:.04em}'
    + '.meta{font-size:12px;color:#64748b;margin-top:8px;line-height:1.6}.meta b{color:#1e293b}'
    + '.client{background:#f1f5f9;border-radius:10px;padding:12px 16px;margin-bottom:16px}'
    + '.client .lbl{font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#64748b;font-weight:700;margin-bottom:4px}'
    + 'table{width:100%;border-collapse:collapse}'
    + 'th{background:#1a4a3a;color:#fff;padding:9px 12px;text-align:left;font-size:12px}'
    + 'th.r,td.r{text-align:right}th.c,td.c{text-align:center}'
    + 'td{padding:9px 12px;border-bottom:1px solid #e5e7eb}'
    + 'tbody tr:nth-child(even){background:#fafbfc}'
    + '.totals{max-width:320px;margin-left:auto;margin-top:14px;border:1px solid #e5e7eb;border-radius:10px;padding:10px 14px;background:#f8fafc}'
    + '.trow{display:flex;justify-content:space-between;padding:5px 0;font-size:13px}'
    + '.trow.grand{border-top:1px dashed #cbd5e1;margin-top:4px;padding-top:9px;font-size:16px;color:#1a4a3a}.trow.grand strong{font-weight:800}'
    + '.foot{margin-top:24px;padding-top:14px;border-top:1px solid #e5e7eb;color:#64748b;font-size:11px;display:flex;justify-content:space-between;gap:16px;flex-wrap:wrap}'
    + '</style></head><body><div class="wrap">'
    + '<div class="grid"><div>' + logoHtml
    +   '<div class="co"><strong>' + escHtml_(cName) + '</strong><br>' + escHtml_(cAddr) + '<br>' + escHtml_(cPhone) + ' &middot; ' + escHtml_(cEmail) + (ids.length ? '<br>' + ids.join(" &middot; ") : "") + '</div>'
    + '</div><div class="right"><div class="title">DEVIS</div><div class="meta">N&deg; <b>' + escHtml_(devis["ID"] || "") + '</b><br>Date : <b>' + fmtDate_(devis["Date"]) + '</b>'
    +   (devis["Échéance"] ? '<br>Validité : ' + fmtDate_(devis["Échéance"]) : "") + '</div></div></div>'
    + '<div class="client"><div class="lbl">Client</div><strong>' + escHtml_(devis["Client"] || "") + '</strong>'
    +   (devis["Email Client"] ? '<br>' + escHtml_(devis["Email Client"]) : "")
    +   (devis["Téléphone"] ? '<br>' + escHtml_(devis["Téléphone"]) : "") + '</div>'
    + '<table><thead><tr><th>Désignation</th><th class="c">Qté</th><th class="r">P.U.</th><th class="r">Total</th></tr></thead><tbody>' + rows + '</tbody></table>'
    + '<div class="totals">' + totalsRows + '</div>'
    + '<div class="foot"><div>' + footLeft + '</div><div style="text-align:right">Devis valable 30 jours à compter de sa date d\'émission.<br>Merci de votre confiance.</div></div>'
    + '</div></body></html>';
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

// ============================================================
//  MODULE FACTURATION (FacturePro — intégré dans CRM)
// ============================================================

// ---------------------------
// Auth / Users
// ---------------------------
var USERS_HEADERS = ['Username', 'Email', 'Password', 'PasswordHash', 'Salt', 'Role', 'Active', 'DisplayName'];

function getAuthSecret() {
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty('AUTH_SECRET');
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('AUTH_SECRET', secret);
  }
  return secret;
}

function makeSalt() {
  var raw = Utilities.getUuid() + String(new Date().getTime());
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(digest).slice(0, 16);
}

function hashPassword(password, salt) {
  var raw = String(salt || '') + '|' + String(password || '');
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(digest);
}

function signAuthPayload(payload) {
  var raw = String(payload || '') + '|' + getAuthSecret();
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(digest);
}

function makeAuthToken(username, ttlMinutes) {
  var ttl = typeof ttlMinutes === 'number' ? ttlMinutes : 720;
  var expires = new Date().getTime() + ttl * 60 * 1000;
  var payload = String(username || '') + '|' + String(expires);
  var sig = signAuthPayload(payload);
  var tokenRaw = payload + '|' + sig;
  return Utilities.base64EncodeWebSafe(tokenRaw);
}

function verifyAuthToken(token) {
  if (!token) return null;
  try {
    var raw = Utilities.newBlob(Utilities.base64DecodeWebSafe(token)).getDataAsString();
    var parts = raw.split('|');
    if (parts.length < 3) return null;
    var username = parts[0];
    var expires = parseInt(parts[1], 10);
    var sig = parts.slice(2).join('|');
    if (!username || !expires) return null;
    if (new Date().getTime() > expires) return null;
    var payload = username + '|' + expires;
    var expected = signAuthPayload(payload);
    if (expected !== sig) return null;
    return username;
  } catch (e) {
    return null;
  }
}

function getUsersSheet() {
  return getOrCreateSheet('Users', USERS_HEADERS);
}

function ensureUsersHeaders(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    sheet.appendRow(USERS_HEADERS);
    return USERS_HEADERS.slice();
  }
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var changed = false;
  USERS_HEADERS.forEach(function(h){
    if (headers.indexOf(h) === -1) {
      headers.push(h);
      changed = true;
    }
  });
  if (changed) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return headers;
}

function getUsersMeta() {
  var sheet = getUsersSheet();
  var headers = ensureUsersHeaders(sheet);
  var map = headerMap(headers);
  return { sheet: sheet, headers: headers, map: map };
}

function hasUsers() {
  var meta = getUsersMeta();
  return meta.sheet.getLastRow() > 1;
}

function isActiveFlag(value) {
  var v = String(value || '').trim().toLowerCase();
  if (!v) return true;
  return !(v === 'false' || v === '0' || v === 'no' || v === 'inactive');
}

function findUserRow(username) {
  var meta = getUsersMeta();
  var data = meta.sheet.getDataRange().getValues();
  var target = String(username || '').trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var u = String(row[meta.map.Username] || '').trim().toLowerCase();
    var mail = "";
    if (typeof meta.map.Email !== "undefined") {
      mail = String(row[meta.map.Email] || '').trim().toLowerCase();
    }
    if ((u && u === target) || (mail && mail === target)) {
      return { rowIndex: i + 1, row: data[i], meta: meta };
    }
  }
  return null;
}

function userFromRow(row, map) {
  return {
    username: row[map.Username] || '',
    role: normalizeRole_(row[map.Role] || ''),
    active: isActiveFlag(row[map.Active]),
    displayName: row[map.DisplayName] || row[map.Username] || ''
  };
}

function validateToken(token) {
  var username = verifyAuthToken(token);
  if (!username) return null;
  var found = findUserRow(username);
  if (!found || !found.row) return null;
  var user = userFromRow(found.row, found.meta.map);
  if (!user.active) return null;
  return user;
}

function requireAuth(token) {
  var user = validateToken(token);
  if (!user) throw new Error('Accès refusé. Merci de vous reconnecter.');
  user.role = normalizeRole_(user.role);
  return user;
}

function createFirstAdmin(username, password, displayName) {
  if (hasUsers()) return { ok: false, error: 'Des utilisateurs existent déjà.' };
  var u = String(username || '').trim();
  var p = String(password || '');
  if (!u || !p) return { ok: false, error: "Nom d’utilisateur et mot de passe requis." };
  var meta = getUsersMeta();
  var salt = makeSalt();
  var hash = hashPassword(p, salt);
  var row = new Array(meta.headers.length);
  row[meta.map.Username] = u;
  if (typeof meta.map.Email !== "undefined") {
    row[meta.map.Email] = u;
  }
  row[meta.map.Password] = '';
  row[meta.map.PasswordHash] = hash;
  row[meta.map.Salt] = salt;
  row[meta.map.Role] = 'admin';
  row[meta.map.Active] = 'true';
  row[meta.map.DisplayName] = String(displayName || u);
  meta.sheet.appendRow(row);
  var token = makeAuthToken(u);
  return { ok: true, token: token, user: userFromRow(row, meta.map) };
}

function authenticate(username, password) {
  var found = findUserRow(username);
  if (!found) return { ok: false, error: 'Utilisateur introuvable.' };
  var row = found.row;
  var meta = found.meta;
  var map = meta.map;
  var active = isActiveFlag(row[map.Active]);
  if (!active) return { ok: false, error: 'Compte désactivé.' };
  var plain = row[map.Password] || '';
  var salt = row[map.Salt] || '';
  var expected = row[map.PasswordHash] || '';
  if (!expected && plain) {
    if (String(password || '') !== String(plain || '')) return { ok: false, error: 'Mot de passe incorrect.' };
    salt = makeSalt();
    expected = hashPassword(password, salt);
    row[map.Password] = '';
    row[map.PasswordHash] = expected;
    row[map.Salt] = salt;
    while (row.length < meta.headers.length) row.push('');
    meta.sheet.getRange(found.rowIndex, 1, 1, meta.headers.length).setValues([row]);
  } else {
    var actual = hashPassword(password, salt);
    if (expected !== actual) return { ok: false, error: 'Mot de passe incorrect.' };
  }
  var token = makeAuthToken(row[map.Username]);
  return { ok: true, token: token, user: userFromRow(row, map) };
}

// ---------------------------
// RBAC
// ---------------------------
var PERM_ROLES = {
  "sync.import": ["admin"],
  "sync.export": ["admin"],
  "sync.history": ["admin"],
  "settings.write": ["admin"],
  "users.manage": ["admin"],
  "finance.view_amounts": ["admin", "manager", "sales", "accounting"],
  "clients.read": ["admin", "manager", "sales", "accounting", "logistics", "viewer"],
  "clients.write": ["admin", "manager", "sales", "accounting"],
  "clients.delete": ["admin", "manager", "sales"],
  "quotes.read": ["admin", "manager", "sales", "accounting", "logistics", "viewer"],
  "quotes.write": ["admin", "manager", "sales"],
  "quotes.delete": ["admin", "manager", "sales"],
  "invoices.read": ["admin", "manager", "sales", "accounting", "logistics", "viewer"],
  "invoices.write": ["admin", "manager", "sales", "accounting"],
  "invoices.delete": ["admin", "manager"],
  "invoices.status.update": ["admin", "manager", "accounting"],
  "deliveries.read": ["admin", "manager", "sales", "accounting", "logistics", "viewer"],
  "deliveries.write": ["admin", "manager", "logistics"],
  "deliveries.delete": ["admin", "manager"]
};

function normalizeRole_(role) {
  var r = String(role || '').trim().toLowerCase();
  return r || "viewer";
}

function can_(role, perm) {
  var r = normalizeRole_(role);
  if (r === "admin") return true;
  var roles = PERM_ROLES[perm] || [];
  return roles.indexOf(r) !== -1;
}

function assertCan_(role, perm) {
  if (!can_(role, perm)) {
    throw new Error("FORBIDDEN");
  }
}

function getRole_(email) {
  var meta = getUsersMeta();
  var map = meta.map;
  var data = meta.sheet.getDataRange().getValues();
  var target = String(email || '').trim().toLowerCase();
  if (!target) return null;
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var mail = "";
    if (typeof map.Email !== "undefined") {
      mail = String(row[map.Email] || '').trim().toLowerCase();
    }
    if (!mail && typeof map.Username !== "undefined") {
      mail = String(row[map.Username] || '').trim().toLowerCase();
    }
    if (mail && mail === target) {
      if (!isActiveFlag(row[map.Active])) return null;
      return normalizeRole_(row[map.Role] || "");
    }
  }
  return null;
}

function sanitizeDataForRole_(role, data) {
  if (can_(role, "finance.view_amounts")) return data;
  var clone = JSON.parse(JSON.stringify(data || {}));
  function stripItem(it) {
    if (!it || typeof it !== "object") return it;
    delete it.unitPrice;
    delete it.price;
    delete it.amount;
    delete it.total;
    return it;
  }
  function stripDoc(doc) {
    if (!doc || typeof doc !== "object") return doc;
    if (Array.isArray(doc.items)) {
      doc.items = doc.items.map(stripItem);
    }
    delete doc.totals;
    delete doc.acompte;
    delete doc.discountRate;
    delete doc.discount;
    delete doc.remainingDue;
    delete doc.remaining_due;
    delete doc.totalTTC;
    delete doc.subTotal;
    return doc;
  }
  if (Array.isArray(clone.quotes)) clone.quotes = clone.quotes.map(stripDoc);
  if (Array.isArray(clone.invoices)) clone.invoices = clone.invoices.map(stripDoc);
  if (Array.isArray(clone.deliveries)) clone.deliveries = clone.deliveries.map(stripDoc);
  return clone;
}

function sanitizeHistoryForRole_(role, rows) {
  if (can_(role, "finance.view_amounts")) return rows;
  return (rows || []).map(function(row) {
    var out = {};
    for (var k in row) out[k] = row[k];
    if (out.payload && typeof out.payload === "object") {
      out.payload = sanitizeDataForRole_(role, out.payload);
    }
    return out;
  });
}

// ---------------------------
// Team Sync helpers
// ---------------------------
var HISTORY_HEADERS = ['Timestamp', 'UserEmail', 'Action', 'Rev', 'Payload'];

function getSyncToken_() {
  return PropertiesService.getScriptProperties().getProperty('SYNC_TOKEN') || '';
}

function validateSyncToken_(token) {
  var expected = getSyncToken_();
  if (!expected) return false;
  return String(token || '') === expected;
}

function getSyncRev_() {
  var raw = PropertiesService.getScriptProperties().getProperty('SYNC_REV');
  var n = parseInt(raw, 10);
  return isNaN(n) ? 0 : n;
}

function setSyncRev_(rev) {
  PropertiesService.getScriptProperties().setProperty('SYNC_REV', String(rev || 0));
}

function getHistorySheet_() {
  return getOrCreateSheet('History', HISTORY_HEADERS);
}

function appendHistory_(userEmail, action, rev, payload) {
  var sheet = getHistorySheet_();
  var row = [
    new Date().toISOString(),
    String(userEmail || ''),
    String(action || ''),
    rev || 0,
    JSON.stringify(payload || {})
  ];
  sheet.appendRow(row);
}

function getHistoryRows_() {
  var sheet = getHistorySheet_();
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var payloadRaw = data[i][4];
    var payload = payloadRaw;
    if (typeof payloadRaw === "string") {
      try { payload = JSON.parse(payloadRaw); } catch (e) {}
    }
    rows.push({
      timestamp: data[i][0],
      userEmail: data[i][1],
      action: data[i][2],
      rev: data[i][3],
      payload: payload
    });
  }
  return rows;
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleSyncGet_(e) {
  var action = e.parameter.action;
  var token = e.parameter.token || "";
  var userEmail = e.parameter.userEmail || e.parameter.email || "";
  if (!validateSyncToken_(token)) return jsonOutput_({ ok: false, error: "UNAUTHORIZED" });
  var role = getRole_(userEmail);
  if (!role) return jsonOutput_({ ok: false, error: "FORBIDDEN" });

  if (action === "import") {
    if (!can_(role, "sync.import")) return jsonOutput_({ ok: false, error: "FORBIDDEN" });
    var data = sanitizeDataForRole_(role, getAllDataInternal());
    return jsonOutput_({ ok: true, rev: getSyncRev_(), data: data });
  }

  if (action === "history") {
    if (!can_(role, "sync.history")) return jsonOutput_({ ok: false, error: "FORBIDDEN" });
    var rows = sanitizeHistoryForRole_(role, getHistoryRows_());
    return jsonOutput_({ ok: true, rev: getSyncRev_(), history: rows });
  }

  return jsonOutput_({ ok: false, error: "INVALID_ACTION" });
}

function handleSyncPost_(e) {
  var action = e.parameter.action;
  if (action !== "export") return jsonOutput_({ ok: false, error: "INVALID_ACTION" });
  var payload = {};
  if (e && e.postData && e.postData.contents) {
    try { payload = JSON.parse(e.postData.contents); } catch (err) {}
  }
  var token = payload.token || e.parameter.token || "";
  var userEmail = payload.userEmail || payload.email || e.parameter.userEmail || e.parameter.email || "";
  if (!validateSyncToken_(token)) return jsonOutput_({ ok: false, error: "UNAUTHORIZED" });
  var role = getRole_(userEmail);
  if (!role) return jsonOutput_({ ok: false, error: "FORBIDDEN" });
  if (!can_(role, "sync.export")) return jsonOutput_({ ok: false, error: "FORBIDDEN" });

  var baseRev = parseInt(payload.baseRev || 0, 10);
  var currentRev = getSyncRev_();
  if (baseRev < currentRev) {
    return jsonOutput_({ ok: false, error: "CONFLICT", currentRev: currentRev });
  }

  var data = payload.data || payload.state || payload.payload;
  if (!data || typeof data !== "object") {
    return jsonOutput_({ ok: false, error: "INVALID_DATA" });
  }

  importAllData_(data);
  var newRev = currentRev + 1;
  setSyncRev_(newRev);
  appendHistory_(userEmail, "export", newRev, data);
  return jsonOutput_({ ok: true, rev: newRev });
}

function listUsers(token) {
  var user = requireAuth(token);
  assertCan_(user.role, "users.manage");
  var meta = getUsersMeta();
  var map = meta.map;
  var data = meta.sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    out.push({
      username: row[map.Username] || "",
      email: typeof map.Email !== "undefined" ? (row[map.Email] || "") : "",
      role: normalizeRole_(row[map.Role] || ""),
      active: isActiveFlag(row[map.Active]),
      displayName: row[map.DisplayName] || ""
    });
  }
  return out;
}

function saveUser(token, user) {
  var current = requireAuth(token);
  assertCan_(current.role, "users.manage");
  var u = user && typeof user === "object" ? user : {};
  var lookup = String(u.lookup || "").trim().toLowerCase();
  var email = String(u.email || "").trim();
  var username = String(u.username || "").trim();
  var displayName = String(u.displayName || "").trim();
  var role = normalizeRole_(u.role || "");
  var active = (u.active === false || u.active === "0" || u.active === 0) ? "false" : "true";
  var password = String(u.password || "");

  if (!email && !username) return { ok: false, error: "EMAIL_REQUIRED" };
  if (!username) username = email.split("@")[0];
  if (!email) email = username;

  var found = findUserRow(lookup || email || username);
  if (!found && !password) return { ok: false, error: "PASSWORD_REQUIRED" };

  var meta = getUsersMeta();
  var map = meta.map;
  var row;
  var rowIndex;
  if (found && found.row) {
    row = found.row;
    rowIndex = found.rowIndex;
  } else {
    row = new Array(meta.headers.length);
    rowIndex = null;
  }

  row[map.Username] = username;
  if (typeof map.Email !== "undefined") row[map.Email] = email;
  row[map.DisplayName] = displayName || username;
  row[map.Role] = role;
  row[map.Active] = active;

  if (password) {
    var salt = makeSalt();
    var hash = hashPassword(password, salt);
    row[map.Password] = "";
    row[map.PasswordHash] = hash;
    row[map.Salt] = salt;
  }

  while (row.length < meta.headers.length) row.push("");

  if (rowIndex) {
    meta.sheet.getRange(rowIndex, 1, 1, meta.headers.length).setValues([row]);
  } else {
    meta.sheet.appendRow(row);
  }

  return { ok: true };
}

function importAllData_(data) {
  var d = data || {};
  writeClients_(d.clients || []);
  writeQuotes_(d.quotes || []);
  writeInvoices_(d.invoices || []);
  writeDeliveries_(d.deliveries || []);
}

function writeSheet_(sheetName, headers, rows) {
  var sheet = getOrCreateSheet(sheetName, headers);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows && rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

function writeClients_(clients) {
  var headers = ['ID', 'Name', 'Contact', 'Address', 'Phone', 'NIF', 'STAT'];
  var rows = (clients || []).map(function(c){
    return [
      c.id || '',
      c.name || '',
      c.contact || '',
      c.address || '',
      c.phone || '',
      c.nif || '',
      c.stat || ''
    ];
  });
  writeSheet_('Clients', headers, rows);
}

function writeQuotes_(quotes) {
  var headers = ['ID', 'Number', 'ClientID', 'Date', 'Status', 'Items', 'Notes', 'Totals', 'ConvertedTo', 'Acompte', 'DiscountRate'];
  var rows = (quotes || []).map(function(q){
    return [
      q.id || '',
      q.number || '',
      q.clientId || '',
      q.date || '',
      q.status || '',
      JSON.stringify(q.items || []),
      q.notes || '',
      JSON.stringify(q.totals || {}),
      JSON.stringify(q.convertedTo || {}),
      q.acompte || 0,
      q.discountRate || 0
    ];
  });
  writeSheet_('Quotes', headers, rows);
}

function writeInvoices_(invoices) {
  var headers = ['ID', 'Number', 'ClientID', 'Date', 'Status', 'SourceQuoteID', 'Items', 'Notes', 'Totals', 'Acompte', 'DiscountRate'];
  var rows = (invoices || []).map(function(inv){
    return [
      inv.id || '',
      inv.number || '',
      inv.clientId || '',
      inv.date || '',
      inv.status || '',
      inv.sourceQuoteId || '',
      JSON.stringify(inv.items || []),
      inv.notes || '',
      JSON.stringify(inv.totals || {}),
      inv.acompte || 0,
      inv.discountRate || 0
    ];
  });
  writeSheet_('Invoices', headers, rows);
}

function writeDeliveries_(deliveries) {
  var headers = ['ID', 'Number', 'ClientID', 'Date', 'Status', 'SourceQuoteID', 'Items', 'Notes', 'SourceInvoiceID'];
  var rows = (deliveries || []).map(function(d){
    return [
      d.id || '',
      d.number || '',
      d.clientId || '',
      d.date || '',
      d.status || '',
      d.sourceQuoteId || '',
      JSON.stringify(d.items || []),
      d.notes || '',
      d.sourceInvoiceId || ''
    ];
  });
  writeSheet_('Deliveries', headers, rows);
}

// ---------------------------
// Data API
// ---------------------------
function getAllDataInternal() {
  try {
    return {
      clients: getClientsInternal(),
      quotes: getQuotesInternal(),
      invoices: getInvoicesInternal(),
      deliveries: getDeliveriesInternal()
    };
  } catch (e) {
    return {
      clients: [],
      quotes: [],
      invoices: [],
      deliveries: [],
      error: String(e && e.message ? e.message : e)
    };
  }
}

function getAllDataText(token) {
  var user = requireAuth(token);
  var data = sanitizeDataForRole_(user.role, getAllDataInternal());
  return JSON.stringify(data);
}

// ---------------------------
// Company info (shared)
// ---------------------------
var COMPANY_HEADERS = ['Name', 'Address', 'Phone', 'Email', 'NIF', 'STAT', 'LogoDataUrl', 'BankName', 'BankAccount', 'BankIban', 'BankBic', 'PaymentTerms', 'ResponsibleName', 'ReferenceLabel', 'ReferenceValue', 'PurchaseOrderLabel', 'PurchaseOrderValue'];

function getCompanySheet() {
  return getOrCreateSheet('Company', COMPANY_HEADERS);
}

function normalizeCompanyInfo(info) {
  info = info && typeof info === 'object' ? info : {};
  return {
    name: String(info.name || ''),
    address: String(info.address || ''),
    phone: String(info.phone || ''),
    email: String(info.email || ''),
    nif: String(info.nif || ''),
    stat: String(info.stat || ''),
    logoDataUrl: String(info.logoDataUrl || ''),
    bankName: String(info.bankName || ''),
    bankAccount: String(info.bankAccount || ''),
    bankIban: String(info.bankIban || ''),
    bankBic: String(info.bankBic || ''),
    paymentTerms: String(info.paymentTerms || ''),
    responsibleName: String(info.responsibleName || ''),
    referenceLabel: String(info.referenceLabel || ''),
    referenceValue: String(info.referenceValue || ''),
    purchaseOrderLabel: String(info.purchaseOrderLabel || ''),
    purchaseOrderValue: String(info.purchaseOrderValue || '')
  };
}

function getCompanyInfoInternal() {
  var sheet = getCompanySheet();
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return null;
  }
  var row = data[1];
  return {
    name: row[0] || '',
    address: row[1] || '',
    phone: row[2] || '',
    email: row[3] || '',
    nif: row[4] || '',
    stat: row[5] || '',
    logoDataUrl: row[6] || '',
    bankName: row[7] || '',
    bankAccount: row[8] || '',
    bankIban: row[9] || '',
    bankBic: row[10] || '',
    paymentTerms: row[11] || '',
    responsibleName: row[12] || '',
    referenceLabel: row[13] || '',
    referenceValue: row[14] || '',
    purchaseOrderLabel: row[15] || '',
    purchaseOrderValue: row[16] || ''
  };
}

function saveCompanyInfoInternal(info) {
  var sheet = getCompanySheet();
  var clean = normalizeCompanyInfo(info);
  var row = [
    clean.name,
    clean.address,
    clean.phone,
    clean.email,
    clean.nif,
    clean.stat,
    clean.logoDataUrl,
    clean.bankName,
    clean.bankAccount,
    clean.bankIban,
    clean.bankBic,
    clean.paymentTerms,
    clean.responsibleName,
    clean.referenceLabel,
    clean.referenceValue,
    clean.purchaseOrderLabel,
    clean.purchaseOrderValue
  ];
  if (sheet.getLastRow() < 2) {
    sheet.appendRow(row);
  } else {
    sheet.getRange(2, 1, 1, row.length).setValues([row]);
  }
  return clean;
}

function getCompanyInfo(token) {
  requireAuth(token);
  return getCompanyInfoInternal();
}

function saveCompanyInfo(token, info) {
  requireAuth(token);
  return saveCompanyInfoInternal(info);
}


function getClientsInternal() {
  var sheet = getOrCreateSheet('Clients', ['ID', 'Name', 'Contact', 'Address', 'Phone', 'NIF', 'STAT']);
  var headers = ensureClientHeaders(sheet);
  var map = headerMap(headers);
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    list.push({
      id: data[i][map.ID] || "",
      name: data[i][map.Name] || "",
      contact: data[i][map.Contact] || "",
      address: data[i][map.Address] || "",
      phone: data[i][map.Phone] || "",
      nif: data[i][map.NIF] || "",
      stat: data[i][map.STAT] || ""
    });
  }
  return list;
}

function getQuotesInternal() {
  var sheet = getOrCreateSheet('Quotes', ['ID', 'Number', 'ClientID', 'Date', 'Status', 'Items', 'Notes', 'Totals', 'ConvertedTo', 'Acompte', 'DiscountRate']);
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    list.push({
      id: data[i][0],
      number: data[i][1],
      clientId: data[i][2],
      date: data[i][3],
      status: data[i][4],
      items: safeJsonParse(data[i][5], []),
      notes: data[i][6],
      totals: safeJsonParse(data[i][7], {}),
      convertedTo: safeJsonParse(data[i][8], {}),
      acompte: data[i][9],
      discountRate: data[i][10]
    });
  }
  return list;
}

function getInvoicesInternal() {
  var sheet = getOrCreateSheet('Invoices', ['ID', 'Number', 'ClientID', 'Date', 'Status', 'SourceQuoteID', 'Items', 'Notes', 'Totals', 'Acompte', 'DiscountRate']);
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    list.push({
      id: data[i][0],
      number: data[i][1],
      clientId: data[i][2],
      date: data[i][3],
      status: data[i][4],
      sourceQuoteId: data[i][5],
      items: safeJsonParse(data[i][6], []),
      notes: data[i][7],
      totals: safeJsonParse(data[i][8], {}),
      acompte: data[i][9],
      discountRate: data[i][10]
    });
  }
  return list;
}

function getDeliveriesInternal() {
  var sheet = getOrCreateSheet('Deliveries', ['ID', 'Number', 'ClientID', 'Date', 'Status', 'SourceQuoteID', 'Items', 'Notes']);
  var headers = ensureDeliveryHeaders(sheet);
  var map = headerMap(headers);
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    list.push({
      id: data[i][map.ID] || "",
      number: data[i][map.Number] || "",
      clientId: data[i][map.ClientID] || "",
      date: data[i][map.Date] || "",
      status: data[i][map.Status] || "",
      sourceQuoteId: data[i][map.SourceQuoteID] || "",
      sourceInvoiceId: typeof map.SourceInvoiceID !== "undefined" ? (data[i][map.SourceInvoiceID] || "") : "",
      items: safeJsonParse(data[i][map.Items], []),
      notes: data[i][map.Notes] || ""
    });
  }
  return list;
}

function getOrCreateSheet(name, headers) {
  var ss = getSpreadsheet();
  if (!ss) {
    throw new Error("Aucun classeur actif. Liez le script à un Google Sheets ou configurez un ID.");
  }
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  var propId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (propId) {
    return SpreadsheetApp.openById(propId);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function healthCheck() {
  try {
    var ss = getSpreadsheet();
    return {
      ok: true,
      id: ss ? ss.getId() : "",
      name: ss ? ss.getName() : "",
      sheets: ss ? ss.getSheets().length : 0
    };
  } catch (e) {
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}

function saveClientInternal(client) {
  var sheet = getOrCreateSheet('Clients', ['ID', 'Name', 'Contact', 'Address', 'Phone', 'NIF', 'STAT']);
  var headers = ensureClientHeaders(sheet);
  var map = headerMap(headers);
  if (!client.id) client.id = 'CL-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][map.ID] === client.id) {
      rowIndex = i + 1;
      break;
    }
  }
  var row = [];
  row[map.ID] = client.id;
  row[map.Name] = client.name;
  row[map.Contact] = client.contact;
  row[map.Address] = client.address;
  row[map.Phone] = client.phone;
  row[map.NIF] = client.nif;
  row[map.STAT] = client.stat;
  while (row.length < headers.length) row.push('');
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return client;
}

function saveQuoteInternal(quote) {
  var sheet = getOrCreateSheet('Quotes', ['ID', 'Number', 'ClientID', 'Date', 'Status', 'Items', 'Notes', 'Totals', 'ConvertedTo', 'Acompte', 'DiscountRate']);
  if (!quote.id) quote.id = 'Q-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  if (!quote.number) quote.number = generateNextNumber(sheet, 'DV');
  quote.discountRate = toNumber(quote.discountRate || 0);
  quote.totals = computeTotalsServer(quote.items || [], quote.discountRate);
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === quote.id) {
      rowIndex = i + 1;
      break;
    }
  }
  var row = [
    quote.id, quote.number, quote.clientId, quote.date, quote.status,
    JSON.stringify(quote.items || []), quote.notes, JSON.stringify(quote.totals || {}),
    JSON.stringify(quote.convertedTo || {}), quote.acompte || 0, quote.discountRate || 0
  ];
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return quote;
}

function saveInvoiceInternal(invoice) {
  var sheet = getOrCreateSheet('Invoices', ['ID', 'Number', 'ClientID', 'Date', 'Status', 'SourceQuoteID', 'Items', 'Notes', 'Totals', 'Acompte', 'DiscountRate']);
  if (!invoice.id) invoice.id = 'INV-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  if (!invoice.number) invoice.number = generateNextNumber(sheet, 'FA');
  invoice.discountRate = toNumber(invoice.discountRate || 0);
  invoice.totals = computeTotalsServer(invoice.items || [], invoice.discountRate);
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === invoice.id) {
      rowIndex = i + 1;
      break;
    }
  }
  var row = [
    invoice.id, invoice.number, invoice.clientId, invoice.date, invoice.status, invoice.sourceQuoteId || '',
    JSON.stringify(invoice.items || []), invoice.notes, JSON.stringify(invoice.totals || {}),
    invoice.acompte || 0, invoice.discountRate || 0
  ];
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return invoice;
}

function saveDeliveryInternal(delivery) {
  var sheet = getOrCreateSheet('Deliveries', ['ID', 'Number', 'ClientID', 'Date', 'Status', 'SourceQuoteID', 'Items', 'Notes']);
  var headers = ensureDeliveryHeaders(sheet);
  var map = headerMap(headers);
  if (!delivery.id) delivery.id = 'BL-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  if (!delivery.number) delivery.number = generateNextNumber(sheet, 'BL');
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][map.ID] === delivery.id) {
      rowIndex = i + 1;
      break;
    }
  }
  var row = new Array(headers.length);
  row[map.ID] = delivery.id;
  row[map.Number] = delivery.number;
  row[map.ClientID] = delivery.clientId;
  row[map.Date] = delivery.date;
  row[map.Status] = delivery.status;
  row[map.SourceQuoteID] = delivery.sourceQuoteId || '';
  if (typeof map.SourceInvoiceID !== "undefined") {
    row[map.SourceInvoiceID] = delivery.sourceInvoiceId || '';
  }
  row[map.Items] = JSON.stringify(delivery.items || []);
  row[map.Notes] = delivery.notes;

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return delivery;
}

function saveClient(token, client) {
  var user = requireAuth(token);
  assertCan_(user.role, "clients.write");
  return saveClientInternal(client);
}

function saveQuote(token, quote) {
  var user = requireAuth(token);
  assertCan_(user.role, "quotes.write");
  return saveQuoteInternal(quote);
}

function saveInvoice(token, invoice) {
  var user = requireAuth(token);
  assertCan_(user.role, "invoices.write");
  var status = String(invoice && invoice.status ? invoice.status : "").toLowerCase();
  if (status === "paid" && !can_(user.role, "invoices.status.update")) {
    throw new Error("FORBIDDEN");
  }
  return saveInvoiceInternal(invoice);
}

function saveDelivery(token, delivery) {
  var user = requireAuth(token);
  assertCan_(user.role, "deliveries.write");
  return saveDeliveryInternal(delivery);
}

function deleteClient(token, id) {
  var user = requireAuth(token);
  assertCan_(user.role, "clients.delete");
  deleteFromSheetInternal('Clients', id);
}

function deleteQuote(token, id) {
  var user = requireAuth(token);
  assertCan_(user.role, "quotes.delete");
  deleteFromSheetInternal('Quotes', id);
}

function deleteInvoice(token, id) {
  var user = requireAuth(token);
  assertCan_(user.role, "invoices.delete");
  deleteFromSheetInternal('Invoices', id);
}

function deleteDelivery(token, id) {
  var user = requireAuth(token);
  assertCan_(user.role, "deliveries.delete");
  deleteFromSheetInternal('Deliveries', id);
}

function deleteFromSheetInternal(sheetName, id) {
  var ss = getSpreadsheet();
  if (!ss) return;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex > 0) sheet.deleteRow(rowIndex);
}

function convertQuoteToInvoice(token, quoteId) {
  var user = requireAuth(token);
  assertCan_(user.role, "quotes.write");
  assertCan_(user.role, "invoices.write");
  var quotes = getQuotesInternal();
  var quote = quotes.find(function(q) { return q.id === quoteId; });
  if (!quote) throw 'Devis introuvable.';
  if (quote.status === 'converted') throw 'Ce devis est déjà converti en facture.';
  var invoice = {
    id: null,
    number: null,
    clientId: quote.clientId,
    date: new Date().toISOString().slice(0, 10),
    status: 'pending',
    sourceQuoteId: quoteId,
    items: quote.items,
    notes: quote.notes,
    totals: computeTotalsServer(quote.items || [], quote.discountRate || 0),
    acompte: quote.acompte,
    discountRate: quote.discountRate
  };
  invoice = saveInvoiceInternal(invoice);
  quote.status = 'converted';
  quote.convertedTo = quote.convertedTo || {};
  quote.convertedTo.invoiceId = invoice.id;
  saveQuoteInternal(quote);
  return invoice;
}

function convertQuoteToDelivery(token, quoteId) {
  var user = requireAuth(token);
  assertCan_(user.role, "deliveries.write");
  throw 'Veuillez déjà convertir le devis en facture.';
}

function convertInvoiceToDelivery(token, invoiceId) {
  var user = requireAuth(token);
  assertCan_(user.role, "deliveries.write");
  var invoices = getInvoicesInternal();
  var invoice = invoices.find(function(i) { return i.id === invoiceId; });
  if (!invoice) throw 'Facture introuvable.';
  var delivery = {
    id: null,
    number: null,
    clientId: invoice.clientId,
    date: new Date().toISOString().slice(0, 10),
    status: 'draft',
    sourceQuoteId: invoice.sourceQuoteId || '',
    sourceInvoiceId: invoiceId,
    items: invoice.items,
    notes: invoice.notes
  };
  delivery = saveDeliveryInternal(delivery);
  return delivery;
}

function generateNextNumber(sheet, prefix) {
  var year = new Date().getFullYear();
  var data = sheet.getDataRange().getValues();
  var max = 0;
  for (var i = 1; i < data.length; i++) {
    var num = data[i][1];
    if (typeof num === 'string' && num.startsWith(prefix + '-' + year + '-')) {
      var seq = parseInt(num.split('-')[2], 10);
      if (!isNaN(seq) && seq > max) max = seq;
    }
  }
  return prefix + '-' + year + '-' + (max + 1).toString().padStart(4, '0');
}

function ensureClientHeaders(sheet) {
  var required = ['ID', 'Name', 'Contact', 'Address', 'Phone', 'NIF', 'STAT'];
  var data = sheet.getDataRange().getValues();
  if (data.length === 0) {
    sheet.getRange(1, 1, 1, required.length).setValues([required]);
    return required;
  }
  var headers = data[0];
  var map = headerMap(headers);
  var hasAddress = typeof map.Address !== "undefined";
  if (!hasAddress) {
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      rows.push([
        row[map.ID] || row[0] || '',
        row[map.Name] || row[1] || '',
        row[map.Contact] || row[2] || '',
        '',
        row[map.Phone] || row[3] || '',
        row[map.NIF] || row[4] || '',
        row[map.STAT] || row[5] || ''
      ]);
    }
    writeSheet_('Clients', required, rows);
    return required;
  }
  if (headers.length < required.length) {
    while (headers.length < required.length) headers.push('');
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return headers;
}

function ensureDeliveryHeaders(sheet) {
  var required = ['ID', 'Number', 'ClientID', 'Date', 'Status', 'SourceQuoteID', 'Items', 'Notes', 'SourceInvoiceID'];
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    sheet.appendRow(required);
    return required;
  }
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var changed = false;
  for (var i = 0; i < required.length; i++) {
    var h = required[i];
    if (headers.indexOf(h) === -1) {
      headers.push(h);
      changed = true;
    }
  }
  if (changed) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return headers;
}

function headerMap(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    map[headers[i]] = i;
  }
  return map;
}

function toNumber(value) {
  if (value === null || typeof value === 'undefined') return 0;
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  var s = String(value).trim();
  if (!s) return 0;
  s = s.replace(/[\s\u00A0\u202F]/g, '');
  s = s.replace(/[^\d,.\-]/g, '');
  if (s.indexOf(',') !== -1 && s.indexOf('.') === -1) {
    s = s.replace(',', '.');
  }
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function computeTotalsServer(items, discountRate) {
  var subTotal = 0;
  for (var i = 0; i < (items || []).length; i++) {
    var it = items[i] || {};
    var qty = Math.max(0, toNumber(it.qty || 0));
    var unitPrice = Math.max(0, toNumber(it.unitPrice || 0));
    subTotal += qty * unitPrice;
  }
  var headerDiscount = Math.min(Math.max(toNumber(discountRate || 0), 0), 100);
  var afterDiscount = subTotal * (1 - headerDiscount / 100);
  return {
    subTotal: Math.round(subTotal),
    taxTotal: 0,
    totalTTC: Math.round(afterDiscount)
  };
}

function safeJsonParse(value, fallback) {
  if (value === null || typeof value === 'undefined' || value === '') return fallback;
  if (typeof value === 'object') return value;
  try {
    return JSON.parse(value);
  } catch (e) {
    return fallback;
  }
}

// ---------------------------
// UI: Menu + Sidebar (Mon Interface)
// ---------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Mon Interface")
    .addItem("Ouvrir l'interface", "showInterface")
    .addToUi();
}

function showInterface() {
  var html = HtmlService.createHtmlOutputFromFile("Interface")
    .setTitle("Mon Interface")
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Cette fonction permet d'enregistrer la signature dans une feuille "Settings"
 */
function saveCompanySettings(token, companyData) {
  const user = requireAuth(token);
  if (user.role !== 'admin') throw new Error("Accès refusé");

  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName("Settings");
  if (!sheet) {
    sheet = ss.insertSheet("Settings");
    sheet.appendRow(["Clé", "Valeur"]);
  }
  
  const configString = JSON.stringify(companyData);
  const data = sheet.getDataRange().getValues();
  let foundIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === "COMPANY_CONFIG") {
      foundIndex = i + 1;
      break;
    }
  }
  
  if (foundIndex > -1) {
    sheet.getRange(foundIndex, 2).setValue(configString);
  } else {
    sheet.appendRow(["COMPANY_CONFIG", configString]);
  }
  return { success: true };
}

// [ADD START: Company Signature]
function getCompanySignature(token) {
  requireAuth(token);
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Settings");
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === "COMPANY_SIGNATURE_PNG") {
      var val = String(data[i][1] || "").trim();
      if (!val) return null;
      return { signatureDataUrl: val };
    }
  }
  return null;
}

function setCompanySignature(token, signatureDataUrl) {
  var user = requireAuth(token);
  assertCan_(user.role, "settings.write");

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Settings");
  if (!sheet) {
    sheet = ss.insertSheet("Settings");
    sheet.appendRow(["Clé", "Valeur"]);
  }

  var data = sheet.getDataRange().getValues();
  var targetRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === "COMPANY_SIGNATURE_PNG") {
      targetRow = i + 1;
      break;
    }
  }

  var clean = signatureDataUrl && String(signatureDataUrl).trim() ? String(signatureDataUrl).trim() : "";
  if (targetRow === -1) {
    sheet.appendRow(["COMPANY_SIGNATURE_PNG", clean]);
  } else {
    sheet.getRange(targetRow, 2).setValue(clean);
  }
  return clean ? { signatureDataUrl: clean } : { signatureDataUrl: null };
}
// [ADD END: Company Signature]

// ==========================================
// EXTENSION : VENTE RAPIDE & POS
// ==========================================
function getSalesSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName("Ventes_Rapides");
  if (!sheet) {
    sheet = ss.insertSheet("Ventes_Rapides");
    sheet.appendRow(['Date', 'ID Client', 'Nom Client', 'Contact', 'Lieu', 'Détails', 'Total (MGA)', 'Avance', 'Reste']);
    sheet.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#f3f3f3");
  }
  return sheet;
}
function saveQuickSale(token, saleData) {
  requireAuth(token);
  var sheet = getSalesSheet();
  sheet.appendRow([
    new Date(),
    saleData.clientId,
    saleData.clientName,
    saleData.contact,
    saleData.lieu,
    saleData.items,
    saleData.total,
    saleData.avance,
    saleData.reste
  ]);
  return { success: true };
}
function getQuickSalesStats(token) {
  requireAuth(token);
  var sheet = getSalesSheet();
  var data = sheet.getDataRange().getValues();
  var stats = { totalVendu: 0, totalAvance: 0, totalReste: 0, count: 0, countRecouvrement: 0 };
  var now = new Date();
  var thisMonth = now.getMonth();
  var thisYear = now.getFullYear();

  for (var i = 1; i < data.length; i++) {
    var rowDate = new Date(data[i][0]);
    var isSameMonth = rowDate.getMonth() === thisMonth && rowDate.getFullYear() === thisYear;
    var reste = Number(data[i][8] || 0);
    if (isSameMonth) {
      stats.totalVendu += Number(data[i][6] || 0);
      stats.totalAvance += Number(data[i][7] || 0);
      stats.totalReste += reste;
      stats.count++;
    }
    if (reste > 0) stats.countRecouvrement++;
  }
  return stats;
}
function getQuickSalesList(token) {
  requireAuth(token);
  var sheet = getSalesSheet();
  var data = sheet.getDataRange().getValues();
  var list = [];
  var now = new Date();
  var thisMonth = now.getMonth();
  var thisYear = now.getFullYear();
  for (var i = 1; i < data.length; i++) {
    var rowDate = new Date(data[i][0]);
    var isSameMonth = rowDate.getMonth() === thisMonth && rowDate.getFullYear() === thisYear;
    if (isSameMonth) {
      list.push({
        rowIndex: i + 1,
        date: data[i][0],
        clientId: data[i][1],
        clientName: data[i][2],
        contact: data[i][3],
        lieu: data[i][4],
        items: data[i][5],
        total: Number(data[i][6] || 0),
        avance: Number(data[i][7] || 0),
        reste: Number(data[i][8] || 0)
      });
    }
  }
  return list;
}
function validatePayment(token, rowIndex, montant) {
  requireAuth(token);
  var sheet = getSalesSheet();
  var data = sheet.getDataRange().getValues();
  var i = rowIndex - 1; // 0-based
  if (i < 1 || i >= data.length) throw new Error("Ligne introuvable");
  var resteActuel = Number(data[i][8] || 0);
  var avanceActuel = Number(data[i][7] || 0);
  var paiement = Math.min(Number(montant || 0), resteActuel);
  var nouveauAvance = avanceActuel + paiement;
  var nouveauReste = Math.max(0, resteActuel - paiement);
  sheet.getRange(rowIndex, 8).setValue(nouveauAvance);
  sheet.getRange(rowIndex, 9).setValue(nouveauReste);
  return { success: true, nouveauReste: nouveauReste };
}
