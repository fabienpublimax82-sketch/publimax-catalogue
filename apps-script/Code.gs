/**
 * Catalogue Carrefour Publimax : envoi des surveys vers Google Drive + Google Sheet de récap.
 * À déployer depuis le compte Google de Publimax (voir INSTALLATION.md).
 * Aucune clé à saisir : l'empreinte (H) puis la clé du catalogue (K) sont déposées par la première tablette qui synchronise.
 * Demandes d'accès : voir plus bas.
 * Tout le reste (dossier Drive racine, classeur de récap) est créé automatiquement au premier envoi.
 */
var P = PropertiesService.getScriptProperties();
var ROOT_NAME = 'Catalogue Carrefour - surveys', SS_NAME = 'Catalogue Carrefour - récap surveys';
var HL = ['Dossier', 'Magasin', 'N°', 'Date visite', 'Saisi par', 'Rubrique', 'Désignation', 'Qté', 'L (mm)', 'H (mm)', 'Matière', 'Catégorie', 'Options / composition', 'Remarque', 'Prix', 'Photos', 'Id ligne', 'Mis à jour'];
var HD = ['Dossier', 'Magasin', 'N°', 'Date visite', 'Saisi par', 'Lignes', 'Photos', 'Mis à jour', 'Dossier Drive', 'Client', 'Interlocuteur', 'Téléphone', 'E-mail', 'Adresse', 'Code postal', 'Ville', 'Pose souhaitée', 'Accès / contraintes', 'Remarques'];

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
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold'); sh.setFrozenRows(1);
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
    if (b.a === 'req') return json_(req_(b));
    if (b.a === 'st') return json_(st_(b));
    if (!auth_(b.h)) return json_({ ok: false, err: 'accès refusé (clé du catalogue différente de celle du script)' });
    lock.waitLock(25000);
    if (b.a === 'photo') {
      var m = /^data:(image\/(?:jpeg|png|webp));base64,(.+)$/.exec(String(b.data || '')); if (!m) return json_({ ok: false, err: 'photo illisible' });
      var f = folder_({ id: b.did, ville: b.ville, date: b.date }, true), name = clean_(b.k).replace(/ /g, '_') + '.jpg', it = f.getFilesByName(name);
      var file = it.hasNext() ? it.next() : f.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name));
      file.setDescription(String(b.lid) + ' | ' + String(b.des || ''));
      return json_({ ok: true, id: file.getId(), url: file.getUrl() });
    }
    if (b.a === 'key') { if (/^[A-Za-z0-9_\-]{40,50}$/.test(String(b.k || '')) && sha_(b.k + '|sync') === b.h) { P.setProperty('K', b.k); return json_({ ok: true }); } return json_({ ok: false, err: 'clé refusée' }); }
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
        return [d.id, d.ville || '', d.num || '', d.date || '', d.aut || '', r.rubrique || '', r.des || '', r.q || '', r.L || '', r.H || '', r.mat || '', r.cat || '', op.join('\n'), r.note || '', r.prix || '', (ph[r.k] || ph[r.id] || []).join('\n'), r.id, now];
      });
      var ss = ss_();
      replace_(sheet_(ss, 'Lignes', HL), d.id, rows, HL.length);
      replace_(sheet_(ss, 'Dossiers', HD), d.id, [[d.id, d.ville || '', d.num || '', d.date || '', d.aut || '', L.length, np, now, fo.getUrl(), d.client || '', d.contact || '', d.tel || '', d.mail || '', d.adresse || '', d.cp || '', d.commune || '', d.pose || '', d.acces || '', d.note || '']], HD.length);
      return json_({ ok: true, lignes: L.length, photos: np });
    }
    return json_({ ok: false, err: 'action inconnue' });
  } catch (x) { return json_({ ok: false, err: String(x && x.message || x) }); }
  finally { try { lock.releaseLock(); } catch (e) {} }
}

/* ================= Demandes d'accès =================
 * L'appli envoie prénom + e-mail + jeton de l'appareil. Publimax reçoit un mail avec un lien vers une page de décision
 * (le lien seul ne décide rien : il faut appuyer sur Autoriser ou Refuser). Une fois autorisé, cet appareil reçoit la clé du catalogue.
 */
function sha_(s) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8).map(function (x) { return ('0' + ((x + 256) % 256).toString(16)).slice(-2); }).join(''); }
function rid_(e, t) { return 'R_' + sha_(e + '|' + t).slice(0, 24); }
function mailOk_(e) { return /^[^\s@<>"']+@[^\s@<>"']+\.[^\s@<>"']+$/.test(e) && e.length < 120; }
function html_(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
function req_(b) {
  var e = String(b.e || '').trim().toLowerCase(), t = String(b.t || ''), n = String(b.n || '').slice(0, 60);
  if (!mailOk_(e) || t.length < 16 || t.length > 64) return { ok: false, st: 'err', err: 'Demande invalide.' };
  var id = rid_(e, t), cur = P.getProperty(id), r = cur ? JSON.parse(cur) : null;
  if (r && r.st === 'ok') { var k = P.getProperty('K'); return k ? { ok: true, st: 'ok', k: k } : { ok: true, st: 'ok' }; }
  var cache = CacheService.getScriptCache();
  if (cache.get('rq_' + e)) return { ok: true, st: 'att' };
  var nb = P.getKeys().filter(function (x) { return x.indexOf('R_') === 0; }).length;
  if (!r && nb >= 60) return { ok: false, st: 'err', err: 'Trop de demandes en attente : contacte Publimax.' };
  r = { e: e, t: t, n: n, ua: String(b.ua || '').slice(0, 120), st: 'att', at: new Date().toISOString(), nonce: Utilities.getUuid() };
  P.setProperty(id, JSON.stringify(r)); cache.put('rq_' + e, '1', 90);
  var url = ScriptApp.getService().getUrl() + '?a=dec&id=' + id + '&n=' + r.nonce;
  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: 'Survey Publimax : demande d\'accès de ' + (n || e),
    htmlBody: '<div style="font:16px Arial,sans-serif;color:#10242B"><p><b>' + html_(n) + '</b> (' + html_(e) + ') demande l\'accès à l\'appli Survey.</p>' +
      '<p style="color:#5B6B72;font-size:13px">Appareil : ' + html_(r.ua) + '</p>' +
      '<p><a href="' + url + '" style="display:inline-block;padding:14px 22px;background:#00849B;color:#fff;border-radius:10px;text-decoration:none;font-weight:bold">Voir la demande et décider</a></p>' +
      '<p style="color:#5B6B72;font-size:13px">Ouvrir ce lien ne décide rien : il faut ensuite appuyer sur Autoriser ou Refuser.</p></div>' });
  return { ok: true, st: 'att' };
}
function st_(b) {
  var e = String(b.e || '').trim().toLowerCase(), t = String(b.t || ''), cur = P.getProperty(rid_(e, t));
  if (!cur) return { ok: true, st: 'inconnu' };
  var r = JSON.parse(cur);
  if (r.st === 'ok') { var k = P.getProperty('K'); return k ? { ok: true, st: 'ok', k: k } : { ok: true, st: 'ok' }; }
  return { ok: true, st: r.st };
}
/** Appelée par la page de décision (bouton Autoriser / Refuser). */
function decide(id, nonce, yes) {
  var cur = P.getProperty(String(id)); if (!cur || String(id).indexOf('R_') !== 0) throw new Error('Demande introuvable.');
  var r = JSON.parse(cur); if (r.nonce !== nonce) throw new Error('Lien invalide.');
  r.st = yes ? 'ok' : 'refus'; r.dec = new Date().toISOString(); P.setProperty(String(id), JSON.stringify(r));
  return r.st;
}
function doGet(ev) {
  var q = ev.parameter || {};
  if (q.a !== 'dec') return json_({ st: 'indispo' });
  var cur = P.getProperty(String(q.id || '')), r = cur ? JSON.parse(cur) : null;
  if (!r || r.nonce !== q.n) return HtmlService.createHtmlOutput('<p style="font:18px Arial;padding:30px">Lien invalide ou demande supprimée.</p>');
  var lab = { att: 'en attente', ok: 'AUTORISÉE', refus: 'REFUSÉE' }[r.st] || r.st;
  var h = '<div style="font:17px Arial,sans-serif;color:#10242B;max-width:520px;margin:40px auto;padding:0 20px">' +
    '<h2 style="color:#00849B">Survey Publimax : demande d\'accès</h2>' +
    '<p><b>' + html_(r.n) + '</b><br>' + html_(r.e) + '</p><p style="color:#5B6B72;font-size:14px">Demandée le ' + html_(r.at.slice(0, 16).replace('T', ' à ')) + ' (UTC)<br>Appareil : ' + html_(r.ua) + '</p>' +
    '<p>État actuel : <b id="st">' + lab + '</b></p>' +
    '<p><button onclick="go(true)" style="font:inherit;font-weight:bold;padding:14px 26px;border:0;border-radius:10px;background:#00849B;color:#fff;cursor:pointer">Autoriser</button> &nbsp; ' +
    '<button onclick="go(false)" style="font:inherit;font-weight:bold;padding:14px 26px;border:2px solid #C62828;border-radius:10px;background:#fff;color:#C62828;cursor:pointer">Refuser</button></p>' +
    '<p id="m" style="font-weight:bold"></p><p style="color:#5B6B72;font-size:13px">Tu peux changer d\'avis en revenant sur cette page. Attention : une fois la tablette ouverte, refuser ensuite ne retire pas le catalogue déjà reçu.</p></div>' +
    '<script>function go(y){document.getElementById("m").textContent="…";google.script.run.withSuccessHandler(function(s){document.getElementById("st").textContent=s==="ok"?"AUTORISÉE":"REFUSÉE";document.getElementById("m").textContent=s==="ok"?"Accès autorisé. La tablette va se débloquer toute seule.":"Accès refusé."}).withFailureHandler(function(e){document.getElementById("m").textContent="Erreur : "+e.message}).decide(' + JSON.stringify(String(q.id)) + ',' + JSON.stringify(String(q.n)) + ',y)}</' + 'script>';
  return HtmlService.createHtmlOutput(h).setTitle('Survey Publimax : demande d\'accès');
}

/** À lancer une fois à la main dans l'éditeur pour accorder les autorisations (Drive + Sheets). */
function autoriser() { MailApp.getRemainingDailyQuota(); CacheService.getScriptCache(); var r = root_(), s = ss_(); sheet_(s, 'Lignes', HL); sheet_(s, 'Dossiers', HD); Logger.log('OK. Dossier : ' + r.getUrl() + ' — Classeur : ' + s.getUrl()); }
