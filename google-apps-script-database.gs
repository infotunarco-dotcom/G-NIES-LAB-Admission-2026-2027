/*
  SCRIPT GOOGLE SHEETS — Base de données centralisée TUNARCO
  ============================================================
  Reçoit les inscriptions (formation) et les participations
  (compétitions) envoyées par le site, les ajoute comme lignes dans
  deux feuilles ("Inscriptions" et "Compétitions"), avec un rendu
  "tableau professionnel" :
    - une bannière-titre fusionnée en haut (ligne 1)
    - un en-tête de colonnes bleu marine, texte blanc, figé (ligne 2)
    - une colonne "N°" numérotée automatiquement
    - des bandes alternées bleu très clair / blanc sur les lignes de
      données, pour la lisibilité (comme un tableau Excel classique)
    - la cellule "Confirmé" elle-même colorée en rouge (Non) ou vert
      (Oui), pour repérer d'un coup d'œil qui n'a pas encore confirmé
    - bordures sur tout le tableau
    - la photo d'identité et l'extrait de naissance sont envoyés en
      base64 par le site, décodés ici, et enregistrés automatiquement
      dans un dossier Google Drive ("Genies Lab" > sous-dossier par
      enfant). Le lien Drive est inséré dans le Sheet sous forme de
      lien cliquable. Les fichiers restent PRIVÉS (pas de partage
      public — ce sont des documents d'enfants).
    - tu peux soit laisser le script créer ce dossier "Genies Lab"
      automatiquement à la racine de ton Drive, soit lui donner l'ID
      d'un dossier Drive que tu as déjà (voir DRIVE_ROOT_FOLDER_ID
      plus bas). Place aussi ton Google Sheet (le fichier "Excel")
      dans ce même dossier Drive pour tout regrouper au même endroit
      — glisser-déposer dans Drive suffit, aucun code à changer.

  MISE À JOUR (si vous aviez déjà installé une version précédente) :
  1. Ouvre ton Google Sheet → Extensions → Apps Script.
  2. Sélectionne tout le code existant et efface-le.
  3. Colle TOUT le contenu de ce fichier à la place.
  4. Enregistre (icône disquette).
  5. Déployer → Gérer les déploiements → icône crayon (Edit) sur le
     déploiement existant → Version : "Nouvelle version" → Déployer.
     (L'URL /exec reste identique, rien à changer côté site.)
  6. Ce script utilise Google Drive (DriveApp) : au redéploiement,
     Google va demander une NOUVELLE autorisation ("accéder à votre
     Google Drive") — acceptez avec le même compte que d'habitude.
  7. IMPORTANT : la mise en page a changé (bannière + en-tête sur 2
     lignes au lieu d'1). Supprimez les onglets "Inscriptions" et
     "Compétitions" existants avant de retester, pour qu'ils soient
     recréés proprement avec le nouveau format.

  INSTALLATION DE ZÉRO :
  Voir les instructions données précédemment (créer un Sheet, coller
  ce code, Déployer → Nouveau déploiement → Application Web →
  Exécuter en tant que Moi / Qui a accès : Tout le monde).
*/

const DRIVE_ROOT_FOLDER_NAME = "Genies Lab";
// Si tu as DÉJÀ un dossier Drive précis à utiliser (au lieu d'en créer un
// nouveau), colle son ID ici (visible dans l'URL du dossier Drive, après
// /folders/) — sinon laisse "" et le script crée/réutilise un dossier
// "Genies Lab" à la racine de ton Drive automatiquement.
const DRIVE_ROOT_FOLDER_ID = "11b2k2idS0maHBeC_XBHph-GgQ7IHHTLH";
const NAVY = "#1a1035";
const BANDING_BLUE = "#eaf1fb";
const RED = "#fde2e2";
const GREEN = "#dcfce7";
const BORDER_COLOR = "#c7ccd8";

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet;

  if (data.type === "competition") {
    sheet = getOrCreateSheet(
      ss, "Compétitions", "TUNARCO — Compétitions",
      ["N°", "Date", "Compétition", "Élève", "Email", "Téléphone", "Confirmé"]
    );
    sheet.appendRow(["=ROW()-2", data.date || "", data.competition || "", data.nomEleve || "", data.email || "", data.tel || "", "Non"]);
  } else {
    sheet = getOrCreateSheet(
      ss, "Inscriptions", "TUNARCO — Inscriptions 2026-2027",
      [
        "N°", "Date inscription", "Élève", "Naissance", "Parents", "Ancien adhérent",
        "Email", "Tél 1", "Tél 2", "Jour", "Horaire", "Compétition (accès)",
        "Opt. Compétiteur", "Kit Programmation", "Kit Électronique", "Pull officiel", "Taille pull",
        "Total (DT)", "Mode paiement", "Photo ID (Drive)", "Extrait naissance (Drive)",
        "Droit à l'image", "Confirmé"
      ]
    );

    const photoLink = saveFileToDrive(data.photoIdentiteData, data.photoIdentite, data.nomEleve);
    const extraitLink = saveFileToDrive(data.extraitNaissanceData, data.extraitNaissance, data.nomEleve);

    sheet.appendRow([
      "=ROW()-2",
      data.dateInscription || "", data.nomEleve || "", data.dateNaissance || "",
      data.nomParents || "", data.ancienAdherent || "", data.email || "",
      data.tel1 || "", data.tel2 || "", data.jour || "", data.horaire || "",
      data.competition || "", data.optCompetiteur || "", data.optKitProg || "",
      data.optKitElec || "", data.optPull || "", data.pullTaille || "",
      data.total || "", data.modePaiement || "",
      photoLink ? `=HYPERLINK("${photoLink}","📷 Photo")` : "",
      extraitLink ? `=HYPERLINK("${extraitLink}","📄 Extrait")` : "",
      data.droitImage || "", "Non"
    ]);
  }

  applyRowStyling(sheet);

  return ContentService
    .createTextOutput(JSON.stringify({ result: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Décode un fichier envoyé en base64 (data URL) et l'enregistre dans
// Drive > "TUNARCO - Documents Inscriptions" > <Nom de l'élève>.
// Renvoie l'URL du fichier Drive, ou "" si rien à enregistrer.
function saveFileToDrive(dataUrl, fileName, studentName) {
  if (!dataUrl || typeof dataUrl !== "string") return "";
  const match = dataUrl.match(/^data:(.*?);base64,(.*)$/);
  if (!match) return "";
  try {
    const mimeType = match[1] || "application/octet-stream";
    const bytes = Utilities.base64Decode(match[2]);
    const blob = Utilities.newBlob(bytes, mimeType, fileName || "document");
    const rootFolder = DRIVE_ROOT_FOLDER_ID
      ? DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID)
      : getOrCreateFolder(DriveApp.getRootFolder(), DRIVE_ROOT_FOLDER_NAME);
    const studentFolder = getOrCreateFolder(rootFolder, sanitizeFolderName(studentName));
    const file = studentFolder.createFile(blob);
    return file.getUrl();
  } catch (err) {
    return "";
  }
}

function sanitizeFolderName(name) {
  const cleaned = (name || "").replace(/[\\/:*?"<>|]/g, "").trim();
  return cleaned || "Sans nom";
}

function getOrCreateFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

// Crée la feuille avec : bannière-titre fusionnée (ligne 1) +
// en-tête de colonnes (ligne 2), figés, largeurs ajustées.
function getOrCreateSheet(ss, name, title, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    const lastCol = headers.length;

    // Ligne 1 — bannière-titre fusionnée
    sheet.getRange(1, 1, 1, lastCol).merge();
    const banner = sheet.getRange(1, 1);
    banner.setValue(title);
    banner.setBackground(NAVY);
    banner.setFontColor("#ffffff");
    banner.setFontWeight("bold");
    banner.setFontSize(13);
    banner.setHorizontalAlignment("center");
    banner.setVerticalAlignment("middle");
    sheet.setRowHeight(1, 34);

    // Ligne 2 — en-tête des colonnes
    sheet.getRange(2, 1, 1, lastCol).setValues([headers]);
    const headerRange = sheet.getRange(2, 1, 1, lastCol);
    headerRange.setBackground(NAVY);
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    headerRange.setFontSize(11);
    headerRange.setHorizontalAlignment("center");
    headerRange.setVerticalAlignment("middle");
    headerRange.setWrap(true);
    sheet.setRowHeight(2, 34);

    sheet.setFrozenRows(2);
    sheet.setFrozenColumns(2); // N° + Élève/Compétition toujours visibles au scroll
    sheet.autoResizeColumns(1, lastCol);
    sheet.setColumnWidth(1, 45); // colonne N° plus étroite
  }
  return sheet;
}

// Applique bandes alternées + couleur de la cellule "Confirmé" + bordures.
function applyRowStyling(sheet) {
  const lastRow = Math.max(sheet.getMaxRows(), 300);
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  const dataRange = sheet.getRange(3, 1, lastRow - 2, lastCol); // les données commencent ligne 3
  const confirmColLetter = columnToLetter(lastCol); // "Confirmé" est toujours la dernière colonne

  const bandingRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=ISEVEN(ROW())`)
    .setBackground(BANDING_BLUE)
    .setRanges([dataRange])
    .build();

  // Couleur uniquement la cellule "Confirmé" (pas toute la ligne), pour ne
  // pas masquer les bandes alternées.
  const confirmRange = sheet.getRange(3, lastCol, lastRow - 2, 1);
  const redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Non")
    .setBackground(RED)
    .setBold(true)
    .setRanges([confirmRange])
    .build();
  const greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Oui")
    .setBackground(GREEN)
    .setBold(true)
    .setRanges([confirmRange])
    .build();

  // Ordre = priorité : rouge/vert (cellule) d'abord, bandes ensuite.
  sheet.setConditionalFormatRules([redRule, greenRule, bandingRule]);

  const actualLastRow = sheet.getLastRow();
  if (actualLastRow >= 1) {
    sheet.getRange(1, 1, actualLastRow, lastCol)
      .setBorder(true, true, true, true, true, true, BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID);
  }
  if (actualLastRow > 2) {
    sheet.getRange(3, 1, actualLastRow - 2, 1).setHorizontalAlignment("center"); // centrer la colonne N°
  }
}

function columnToLetter(column) {
  let letter = "";
  while (column > 0) {
    const rem = (column - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    column = Math.floor((column - 1) / 26);
  }
  return letter;
}
