# FOREVER MG CRM — Guide de déploiement

## 📁 Structure des fichiers

```
forever-mg-crm/
├── index.html      → PWA frontend (toute l'app)
├── Code.gs         → Backend Google Apps Script
├── manifest.json   → Manifeste PWA (installation mobile)
├── sw.js           → Service Worker (offline)
└── README.md       → Ce guide
```

---

## 🚀 Déploiement en 5 étapes

### ÉTAPE 1 — Créer le Google Apps Script

1. Ouvrez un Google Sheets existant ou créez-en un nouveau
2. Menu : **Extensions → Apps Script**
3. Renommez le projet : `FOREVER MG CRM`
4. Copiez tout le contenu de `Code.gs` dans l'éditeur
5. Créez un nouveau fichier HTML : **+ → Fichier HTML** → nommez-le `index`
6. Copiez tout le contenu de `index.html` dans ce fichier HTML

### ÉTAPE 2 — Initialiser les feuilles Sheets

Dans l'éditeur Apps Script :
1. Sélectionnez la fonction `initSheets` dans le menu déroulant
2. Cliquez **Exécuter**
3. Autorisez les permissions demandées (Google Drive, Gmail, Sheets)
4. Vérifiez que les feuilles `Devis`, `PurchaseOrders` et `Log` ont été créées

### ÉTAPE 3 — Déployer comme application web

1. Cliquez **Déployer → Nouveau déploiement**
2. Type : **Application Web**
3. Description : `FOREVER MG CRM v1.0`
4. Exécuter en tant que : **Moi (votre compte Google)**
5. Accès autorisé à : **Moi uniquement** (ou Toute personne pour partager l'équipe)
6. Cliquez **Déployer**
7. **Copiez l'URL de déploiement** (format : `https://script.google.com/macros/s/XXXX/exec`)

### ÉTAPE 4 — Configurer la PWA

Dans le fichier `index.html` (ou dans le fichier HTML de l'Apps Script), remplacez :
```javascript
const GAS_URL = "VOTRE_URL_GAS_ICI";
```
Par votre URL copiée à l'étape 3 :
```javascript
const GAS_URL = "https://script.google.com/macros/s/XXXX/exec";
```

### ÉTAPE 5 — Installer le trigger quotidien

Dans l'éditeur Apps Script :
1. Sélectionnez la fonction `installTriggers`
2. Cliquez **Exécuter**
3. ✅ Vous recevrez un e-mail à 08h00 chaque matin si des devis urgents sont en attente

---

## 📱 Hébergement de la PWA standalone (optionnel)

Pour une PWA installable sur mobile indépendante du GAS :

### Via GitHub Pages (recommandé)
```bash
# Clonez votre repo existant
git clone https://github.com/2nOOb61/votre-repo.git
cd votre-repo

# Créez un dossier crm et copiez les fichiers
mkdir crm
cp index.html crm/
cp manifest.json crm/
cp sw.js crm/

git add .
git commit -m "feat: FOREVER MG CRM PWA"
git push

# Activez GitHub Pages sur le repo → Settings → Pages → main/root
# URL : https://2nOOb61.github.io/votre-repo/crm/
```

### Enregistrement du Service Worker
Ajoutez dans `index.html` avant `</body>` :
```html
<script>
if ("serviceWorker" in navigator) {
  navigator.serviceWorker.register("./sw.js").catch(console.error);
}
</script>
```

---

## 🔐 Permissions Google requises

Lors de la première exécution, autorisez :
- **Google Sheets** : lecture/écriture des données CRM
- **Gmail** : lecture des e-mails commerciaux, envoi, création de brouillons
- **Script** : exécution des triggers automatiques

---

## 📊 Structure des données (Google Sheets)

### Feuille `Devis`
| Colonne | Type | Description |
|---------|------|-------------|
| ID | Texte | DEV-XXXXX (auto-généré) |
| Date | Date | Date de la demande |
| Client | Texte | Nom du client |
| Email Client | Email | Email de contact |
| Téléphone | Texte | Numéro |
| Catégorie | Liste | Demande devis / Modification / etc. |
| Produit / Service | Texte | Description |
| Montant (Ar) | Nombre | Montant en Ariary |
| Statut | Liste | Ouvert / En cours / Validé / Perdu / Annulé |
| Priorité | Liste | Urgent / Haute / Moyenne / Basse |
| Échéance | Date | Date limite |
| Responsable | Texte | Nom |
| Problème identifié | Texte | Note interne |
| Prochaine action | Texte | Action à faire |
| Commentaire | Texte | Notes libres |
| Créé le | DateTime | Auto |
| Modifié le | DateTime | Auto |

### Feuille `PurchaseOrders`
| Colonne | Type | Description |
|---------|------|-------------|
| ID PO | Texte | PO-XXXXX (auto) |
| Numéro PO | Texte | N° Coupa (ex: 3480-140011) |
| Date réception | Date | |
| Client | Texte | |
| Contact | Texte | Nom du contact |
| Description | Texte | Détail commande |
| Montant (Ar) | Nombre | |
| Conditions paiement | Liste | NET 30 / NET 15 / etc. |
| Statut | Liste | En retard / À acquitter / À facturer / Acquitté / Facturé / Clôturé |
| Échéance acquittement | Date | |
| Lien Coupa | URL | Lien direct Coupa |
| Facture créée | Oui/Non | |
| Commentaire | Texte | |
| Créé le | DateTime | Auto |

---

## ⚡ Actions rapides disponibles

### Dans le Dashboard
- Voir les KPIs en temps réel
- Relances urgentes du jour avec bouton direct
- Aperçu des 5 derniers devis et PO

### Dans Devis
- Créer / modifier / supprimer un devis
- Filtrer par statut et priorité
- Recherche libre
- Bouton ✉ pour pré-composer une relance
- Bouton ✕ pour marquer comme Perdu

### Dans PO
- Créer / modifier un PO
- Acquitter en 1 clic (→ statut Acquitté)
- Lien direct vers Coupa

### Dans Gmail
- Lecture des e-mails commerciaux (devis, demande, urgent, PO…)
- Clic sur un e-mail → pré-compose la réponse

### Dans Envoyer
- Envoi direct via Gmail
- 4 templates prêts : accusé, devis, relance, confirmation
- Sauvegarde en brouillon

---

## 🔄 Mise à jour du déploiement

Après modification du code :
1. Apps Script → **Déployer → Gérer les déploiements**
2. Cliquez sur ✏️ (modifier) à côté de votre déploiement
3. Version : **Nouvelle version**
4. Cliquez **Déployer**
5. L'URL reste identique — pas besoin de la reconfigurer

---

## 🐛 Dépannage fréquent

| Problème | Solution |
|----------|----------|
| "Mode démo" affiché | Remplacez `GAS_URL` dans index.html |
| Erreur de permission Gmail | Ré-exécutez `initSheets` et ré-autorisez |
| Feuilles non créées | Exécutez `initSheets()` manuellement |
| Trigger absent | Exécutez `installTriggers()` |
| CORS error | Vérifiez que le déploiement est "Application Web" et non API |

---

## 📞 Support FOREVER MG

- Email : commercial4evermg@gmail.com
- Tél : +261 34 03 767 69
- Facebook : facebook.com/4evermg
