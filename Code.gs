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
    const cmd     = payload.cmd; // "cmd" au lieu de "action" pour éviter le conflit avec le champ "Prochaine action"
    let result;

    switch (cmd) {
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
      case "getGmailLeads":        result = getGmailLeads();              break;
      case "createDevisFromEmail": result = createDevisFromEmail(payload);break;
      case "sendEmail":            result = sendEmail(payload);           break;
      case "createDraft":          result = createDraft(payload);         break;
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
