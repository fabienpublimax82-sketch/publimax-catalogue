/**
 * Catalogue Carrefour Publimax : autorisations d'accès + dépôt des dossiers.
 * À déployer depuis le compte Google de Publimax (voir INSTALLATION.md).
 * Propriété du script à créer : KEY = la clé du catalogue.
 */
var P = PropertiesService.getScriptProperties();
function secret_() { var s = P.getProperty('SECRET'); if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); P.setProperty('SECRET', s); } return s; }
function admin_() { return P.getProperty('ADMIN') || Session.getEffectiveUser().getEmail(); }
function get_(e) { var v = P.getProperty('U_' + e); return v ? JSON.parse(v) : null; }
function set_(e, o) { P.setProperty('U_' + e, JSON.stringify(o)); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function mailOk_(e) { return /^[^\s@<>"']+@[^\s@<>"']+\.[^\s@<>"']+$/.test(e) && e.length < 120; }

function doGet(ev) {
  var q = ev.parameter || {}, a = q.a, e = String(q.e || '').trim().toLowerCase(), t = String(q.t || '');
  if (a === 'req') {
    if (!mailOk_(e) || t.length < 16) return json_({ st: 'err' });
    var u = get_(e);
    if (u && u.st === 'ok' && u.t === t) return json_({ st: 'ok', k: P.getProperty('KEY') });
    var first = !u || u.t !== t || u.st === 'refus';
    set_(e, { st: 'att', t: t, n: String(q.n || '').slice(0, 60), at: new Date().toISOString() });
    if (first) {
      var base = ScriptApp.getService().getUrl() + '?e=' + encodeURIComponent(e) + '&s=' + secret_() + '&a=';
      MailApp.sendEmail({ to: admin_(), subject: 'Catalogue Carrefour : demande d\'accès de ' + e,
        htmlBody: '<p><b>' + e + '</b> demande l\'accès au catalogue Carrefour (tablette).</p>' +
          '<p><a href="' + base + 'ok" style="display:inline-block;padding:12px 20px;background:#008099;color:#fff;border-radius:8px;text-decoration:none;font-weight:bold">Autoriser</a> &nbsp; ' +
          '<a href="' + base + 'no" style="display:inline-block;padding:12px 20px;border:2px solid #C62828;color:#C62828;border-radius:8px;text-decoration:none;font-weight:bold">Refuser</a></p>' +
          '<p style="color:#666">Tu peux revenir sur ta décision plus tard en recliquant sur l\'un des deux liens.</p>' });
    }
    return json_({ st: 'att' });
  }
  if (a === 'st') {
    var u2 = get_(e);
    if (!u2 || u2.t !== t) return json_({ st: 'inconnu' });
    return json_(u2.st === 'ok' ? { st: 'ok', k: P.getProperty('KEY') } : { st: u2.st });
  }
  if (a === 'ok' || a === 'no') {
    if (q.s !== secret_()) return HtmlService.createHtmlOutput('Lien invalide.');
    var u3 = get_(e); if (!u3) return HtmlService.createHtmlOutput('Aucune demande pour ' + e + '.');
    u3.st = a === 'ok' ? 'ok' : 'refus'; u3.dec = new Date().toISOString(); set_(e, u3);
    return HtmlService.createHtmlOutput('<div style="font:18px Segoe UI,Arial;padding:30px"><b>' + e + '</b> : accès ' + (a === 'ok' ? 'AUTORISÉ' : 'REFUSÉ') + '.</div>');
  }
  return json_({ st: 'err' });
}

function doPost(ev) {
  try {
    var b = JSON.parse(ev.postData.contents), e = String(b.e || '').toLowerCase(), u = get_(e);
    if (b.a !== 'save' || !u || u.st !== 'ok' || u.t !== b.t) return json_({ ok: false, err: 'accès non autorisé' });
    var it = DriveApp.getFoldersByName('Catalogue Carrefour - dossiers'), f = it.hasNext() ? it.next() : DriveApp.createFolder('Catalogue Carrefour - dossiers');
    var d = b.d.dossier || {}, name = String(d.ville || 'dossier').replace(/[^A-Za-z0-9]+/g, '_') + '_' + (d.date || '') + '_' + (d.id || '') + '.json';
    var old = f.getFilesByName(name); while (old.hasNext()) old.next().setTrashed(true);
    f.createFile(name, JSON.stringify(b.d, null, 1), 'application/json');
    MailApp.sendEmail(admin_(), 'Catalogue Carrefour : dossier ' + (d.ville || '') + ' reçu', e + ' a envoyé le dossier « ' + (d.ville || '') + ' » (' + (b.d.lignes || []).length + ' lignes). Fichier : Drive > Catalogue Carrefour - dossiers > ' + name);
    return json_({ ok: true });
  } catch (x) { return json_({ ok: false, err: String(x) }); }
}

/** À lancer une fois à la main dans l'éditeur pour accorder les autorisations (mail + Drive). */
function autoriser() { MailApp.getRemainingDailyQuota(); DriveApp.getRootFolder(); secret_(); Logger.log('OK, compte : ' + admin_()); }
