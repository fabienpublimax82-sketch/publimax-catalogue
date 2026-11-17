/**
 * Catalogue Carrefour Publimax : envoi des surveys vers Google Drive + Google Sheet de récap.
 * À déployer depuis le compte Google de Publimax (voir INSTALLATION.md).
 * Aucune clé à saisir : l'empreinte de la clé du catalogue est enregistrée au premier envoi (propriété H).
 * Tout le reste (dossier Drive racine, classeur de récap) est créé automatiquement au premier envoi.
 */
var P = PropertiesService.getScriptProperties();
var ROOT_NAME = 'Catalogue Carrefour - surveys', SS_NAME = 'Catalogue Carrefour - récap surveys';
var HL = ['Dossier', 'Magasin', 'N°', 'Date visite', 'Saisi par', 'Rubrique', 'Désignation', 'Qté', 'L (mm)', 'H (mm)', 'Matière', 'Catégorie', 'Options / composition', 'Remarque', 'Prix', 'Photos', 'Id ligne', 'Mis à jour'];
var HD = ['Dossier', 'Magasin', 'N°', 'Date visite', 'Saisi par', 'Lignes', 'Photos', 'Mis à jour', 'Dossier Drive'];

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function hex_(b) { return b.map(function (x) { return ('0' + ((x + 256) % 256).toString(16)).slice(-2); }).join(''); }
/** Le premier appareil qui synchronise enregistre son empreinte ; ensuite seuls les appareils qui ont la même clé de catalogue passent. Pour repartir de zéro : supprimer la propriété H. */
function auth_(h) { h = String(h || ''); if (!/^[0-9a-f]{64}$/.test(h)) return false; var s = P.getProperty('H'); if (!s) { P.setProperty('H', h); return true; } return s === h; }
function clean_(s) { return String(s || '').replace(/[\\\/:*?"<>|\[\]]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80); }

function root_() {
  var id = P.getProperty('ROOT'); if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder(ROOT_NAME); P.setProperty('ROOT', f.getId()); return f;
}
function folder_(d, create) {
  var id = P.getProperty('F_' + d.id); if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  if (!create) return null;
  var f = root_().createFolder(clean_(d.ville || 'Magasin') + ' ' + clean_(d.date || '') + ' [' + clean_(d.id) + ']'); P.setProperty('F_' + d.id, f.getId()); return f;
}
function ss_() {
  var id = P.getProperty('SS'), ss = null; if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) {} }
  if (!ss) { ss = SpreadsheetApp.create(SS_NAME); P.setProperty('SS', ss.getId()); try { DriveApp.getFileById(ss.getId()).moveTo(root_()); } catch (e) {} }
  return ss;
}
function sheet_(ss, name, head) {
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0 ? ss.getSheets()[0].setName(name) : ss.insertSheet(name); }
  if (sh.getLastRow() === 0) { sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
/** Remplace toutes les lignes d'un dossier (colonne A) par les nouvelles. */
function replace_(sh, did, rows, width) {
  var last = sh.getLastRow(), keep = [];
  if (last > 1) keep = sh.getRange(2, 1, last - 1, width).getValues().filter(function (r) { return String(r[0]) !== String(did); });
  var all = keep.concat(rows);
  if (last > 1) sh.getRange(2, 1, last - 1, width).clearContent();
  if (all.length) sh.getRange(2, 1, all.length, width).setValues(all);
}

function doPost(ev) {
  var lock = LockService.getScriptLock();
  try {
    var b = JSON.parse(ev.postData.contents);
    if (!auth_(b.h)) return json_({ ok: false, err: 'accès refusé (clé du catalogue différente de celle du script)' });
    lock.waitLock(25000);
    if (b.a === 'photo') {
      var m = /^data:(image\/(?:jpeg|png|webp));base64,(.+)$/.exec(String(b.data || '')); if (!m) return json_({ ok: false, err: 'photo illisible' });
      var f = folder_({ id: b.did, ville: b.ville, date: b.date }, true), name = clean_(b.k).replace(/ /g, '_') + '.jpg', it = f.getFilesByName(name);
      var file = it.hasNext() ? it.next() : f.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name));
      file.setDescription(String(b.lid) + ' | ' + String(b.des || ''));
      return json_({ ok: true, id: file.getId(), url: file.getUrl() });
    }
    if (b.a === 'photoDel') {
      var fd = folder_({ id: b.did }, false);
      if (fd) { var itd = fd.getFilesByName(clean_(b.k).replace(/ /g, '_') + '.jpg'); while (itd.hasNext()) itd.next().setTrashed(true); }
      return json_({ ok: true });
    }
    if (b.a === 'lignes') {
      var d = b.dossier || {}, L = b.lignes || [], fo = folder_(d, true), ph = {}, np = 0, fi = fo.getFiles();
      while (fi.hasNext()) { var x = fi.next(), lid = String(x.getDescription() || '').split(' | ')[0]; if (!lid) continue; (ph[lid] = ph[lid] || []).push(x.getUrl()); np++; }
      var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
      var rows = L.map(function (r) {
        var op = (r.options || []).map(function (o) { return o.l + (o.q > 1 ? ' × ' + o.q : ''); });
        if (r.composition) op.unshift(r.composition.contenu + ', ' + r.composition.eclairage + ', format ' + r.composition.format);
        if (r.visuel) op.push('Visuel ' + r.visuel);
        return [d.id, d.ville || '', d.num || '', d.date || '', d.aut || '', r.rubrique || '', r.des || '', r.q || '', r.L || '', r.H || '', r.mat || '', r.cat || '', op.join('\n'), r.note || '', r.prix || '', (ph[r.id] || []).join('\n'), r.id, now];
      });
      var ss = ss_();
      replace_(sheet_(ss, 'Lignes', HL), d.id, rows, HL.length);
      replace_(sheet_(ss, 'Dossiers', HD), d.id, [[d.id, d.ville || '', d.num || '', d.date || '', d.aut || '', L.length, np, now, fo.getUrl()]], HD.length);
      return json_({ ok: true, lignes: L.length, photos: np });
    }
    return json_({ ok: false, err: 'action inconnue' });
  } catch (x) { return json_({ ok: false, err: String(x && x.message || x) }); }
  finally { try { lock.releaseLock(); } catch (e) {} }
}

/** La demande d'accès par e-mail n'est pas en service dans cette version. */
function doGet() { return json_({ st: 'indispo' }); }

/** À lancer une fois à la main dans l'éditeur pour accorder les autorisations (Drive + Sheets). */
function autoriser() { var r = root_(), s = ss_(); sheet_(s, 'Lignes', HL); sheet_(s, 'Dossiers', HD); Logger.log('OK. Dossier : ' + r.getUrl() + ' — Classeur : ' + s.getUrl()); }
