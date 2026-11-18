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
var HD = ['Dossier', 'Magasin', 'N°', 'Date visite', 'Saisi par', 'Lignes', 'Photos', 'Mis à jour', 'Dossier Drive', 'Client', 'Interlocuteur', 'Téléphone', 'E-mail', 'Adresse', 'Code postal', 'Ville', 'Pose souhaitée', 'Accès / contraintes', 'Remarques', 'Compte'];

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
    if (b.a === 'code') return json_(code_(b));
    if (b.a === 'decinfo' || b.a === 'dec') return json_(dec_(b));
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
      if (b.etat && b.etat.id === d.id) putJson_(fo, 'survey.json', b.etat);
      var ss = ss_();
      replace_(sheet_(ss, 'Lignes', HL), d.id, rows, HL.length);
      replace_(sheet_(ss, 'Dossiers', HD), d.id, [[d.id, d.ville || '', d.num || '', d.date || '', d.aut || '', L.length, np, now, fo.getUrl(), d.client || '', d.contact || '', d.tel || '', d.mail || '', d.adresse || '', d.cp || '', d.commune || '', d.pose || '', d.acces || '', d.note || '', d.own || '']], HD.length);
      return json_({ ok: true, lignes: L.length, photos: np });
    }
    if (b.a === 'list') {
      var sh = ss_().getSheetByName('Dossiers'), out = [], last = sh ? sh.getLastRow() : 0;
      if (last > 1) sh.getRange(2, 1, last - 1, HD.length).getDisplayValues().forEach(function (r) { if (r[0]) out.push({ id: r[0], ville: r[1], num: r[2], date: r[3], aut: r[4], n: r[5], np: r[6], maj: r[7], own: r[HD.length - 1] || '' }); });
      return json_({ ok: true, list: out });
    }
    if (b.a === 'get') {
      var fg = folder_({ id: String(b.did || '') }, false), et = fg ? getJson_(fg, 'survey.json') : null;
      if (!et) return json_({ ok: false, err: 'Ce survey a été envoyé avec une ancienne version : il faut le rouvrir sur l\'appareil d\'origine pour qu\'il se renvoie.' });
      return json_({ ok: true, etat: et });
    }
    if (b.a === 'phget') {
      var fp = folder_({ id: String(b.did || '') }, false), ip = fp ? fp.getFilesByName(clean_(b.k).replace(/ /g, '_') + '.jpg') : null;
      if (!ip || !ip.hasNext()) return json_({ ok: false, err: 'photo absente' });
      var bl = ip.next().getBlob();
      return json_({ ok: true, data: 'data:' + bl.getContentType() + ';base64,' + Utilities.base64Encode(bl.getBytes()) });
    }
    if (b.a === 'cat') { var cm = cat_(); return json_(Number(b.v) === cm.v ? { ok: true, same: true, ed: isEd_(b) } : { ok: true, cat: cm, ed: isEd_(b) }); }
    if (b.a === 'catimg') {
      var all = getJson_(root_(), 'catalogue-images.json') || {}, im = {};
      (b.ids || []).slice(0, 8).forEach(function (k) { if (all[k]) im[k] = all[k]; });
      return json_({ ok: true, imgs: im });
    }
    if (b.a === 'catset') return json_(catset_(b));
    return json_({ ok: false, err: 'action inconnue' });
  } catch (x) { return json_({ ok: false, err: String(x && x.message || x) }); }
  finally { try { lock.releaseLock(); } catch (e) {} }
}

/* ================= Modifications du catalogue (partagées par toutes les tablettes) =================
 * catalogue-modifs.json : { v, maj, lock, lines: { idLigne: { c, n, p, img } } } ; catalogue-images.json : { idImage: dataURL }.
 * Seuls les appareils autorisés « avec modification du catalogue » peuvent tout changer ;
 * les autres ne peuvent que choisir une catégorie, tant que les catégories ne sont pas verrouillées.
 */
function putJson_(folder, name, obj) { var s = JSON.stringify(obj), it = folder.getFilesByName(name); if (it.hasNext()) it.next().setContent(s); else folder.createFile(name, s, 'application/json'); }
function getJson_(folder, name) { var it = folder.getFilesByName(name); if (!it.hasNext()) return null; try { return JSON.parse(it.next().getBlob().getDataAsString()); } catch (e) { return null; } }
function cat_() { var c = getJson_(root_(), 'catalogue-modifs.json'); if (!c || typeof c !== 'object') c = { v: 0, lock: false, lines: {} }; c.lines = c.lines || {}; c.add = c.add || {}; c.nodes = c.nodes || {}; c.v = Number(c.v) || 0; return c; }
function isEd_(b) { var e = String(b.e || '').trim().toLowerCase(), t = String(b.t || ''); if (!e || !t) return false; var cur = P.getProperty(rid_(e, t)); if (!cur) return false; var r = JSON.parse(cur); return r.st === 'ok' && r.ed === true; }
function line_(o, m) {
  var CATS = { '': 1, S: 1, M: 1, T: 1, L: 1 }, TXT = { des: 160, mat: 160, info: 600 };
  ['c', 'n', 'p', 'img', 'des', 'mat', 'info', 'L', 'H', 'rl', 'hid', 'comp'].forEach(function (k) { if (o[k] === null) delete m[k]; });
  if (o.del === 1) m.del = 1;
  if (o.comp && typeof o.comp === 'object') { var js = JSON.stringify(o.comp); if (js.length < 80000) m.comp = JSON.parse(js); }
  if (o.hid != null) m.hid = o.hid ? 1 : 0;
  if (o.c != null && CATS[o.c]) m.c = o.c;
  if (o.n != null && /^[0-9.]{1,12}$/.test(String(o.n))) m.n = String(o.n);
  if (o.p != null && isFinite(o.p)) m.p = Number(o.p);
  if (o.img != null && /^[a-z0-9_]{1,30}$/.test(String(o.img))) m.img = String(o.img);
  for (var k in TXT) if (typeof o[k] === 'string') m[k] = o[k].slice(0, TXT[k]);
  ['L', 'H'].forEach(function (k) { if (o[k] != null && isFinite(o[k])) m[k] = Math.max(0, Math.round(Number(o[k]))); });
  if (o.rl && typeof o.rl === 'object') { var q = {}; for (var r in o.rl) if (/^[a-z0-9_]{1,24}$/.test(r) && (o.rl[r] === 0 || o.rl[r] === 1)) q[r] = o.rl[r]; m.rl = q; }
  return m;
}
function catset_(b) {
  var ed = isEd_(b), cm = cat_(), ch = b.ch || {}, add = b.add || {}, nd = b.nodes || {}, n = 0, NO = { ok: false, err: 'réservé : la modification du catalogue demande une autorisation.' };
  if (!ed && (b.lock != null || cm.lock || Object.keys(add).length || Object.keys(nd).length)) return NO;
  for (var nc in nd) { if (/^[0-9.]{1,12}$/.test(nc) && nd[nc]) { cm.nodes[nc] = { hid: nd[nc].hid ? 1 : 0 }; if (nd[nc].del === 1) cm.nodes[nc].del = 1; n++; } }
  for (var id in ch) {
    if (!/^[A-Za-z0-9_\-]{1,80}$/.test(id)) continue;
    var o = ch[id] || {};
    if (!ed && Object.keys(o).some(function (k) { return k !== 'c'; })) return NO;
    cm.lines[id] = line_(o, cm.lines[id] || {}); n++;
  }
  for (var ia in add) {
    if (!/^x_[a-z0-9]{6,24}$/.test(ia)) continue;
    cm.add[ia] = add[ia] && add[ia].del ? { del: 1 } : line_(add[ia] || {}, {}); n++;
  }
  if (ed && b.lock != null) cm.lock = b.lock === true;
  var im = b.imgs || {}, keys = Object.keys(im);
  if (ed && keys.length) {
    var all = getJson_(root_(), 'catalogue-images.json') || {};
    keys.forEach(function (k) { var v = String(im[k] || ''); if (/^o_[a-z0-9]{6,24}$/.test(k) && /^data:image\/(jpeg|webp|png);base64,/.test(v) && v.length < 400000) all[k] = v; });
    putJson_(root_(), 'catalogue-images.json', all);
  }
  cm.v++; cm.maj = new Date().toISOString(); putJson_(root_(), 'catalogue-modifs.json', cm);
  return { ok: true, cat: cm, ed: ed, n: n };
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
  var cache = CacheService.getScriptCache();
  if (r && r.st === 'ok') {
    if (b.ed === true && !r.ed && !cache.get('rq_' + e)) { cache.put('rq_' + e, '1', 90); mail_(id, r, true); }
    var k = P.getProperty('K'); return k ? { ok: true, st: 'ok', k: k, ed: !!r.ed } : { ok: true, st: 'ok', ed: !!r.ed };
  }
  var known = findOk_(e);
  if (known) {
    if (cache.get('rq_' + e)) return { ok: true, st: 'code' };
    var code = ('00000' + Math.floor(Math.random() * 1000000)).slice(-6);
    r = { e: e, t: t, n: known.n || n, ua: String(b.ua || '').slice(0, 120), st: 'code', code: code, exp: Date.now() + 15 * 60000, tries: 0, at: new Date().toISOString(), nonce: Utilities.getUuid(), inh: !!known.ed };
    P.setProperty(id, JSON.stringify(r)); cache.put('rq_' + e, '1', 60);
    MailApp.sendEmail({ to: e, subject: 'Survey Publimax : ton code de connexion',
      htmlBody: '<div style="font:16px Arial,sans-serif;color:#10242B"><p>Bonjour ' + html_(r.n) + ',</p><p>Voici ton code pour connecter un nouvel appareil à l\'appli Survey :</p>' +
        '<p style="font-size:34px;font-weight:bold;letter-spacing:6px;color:#00849B;margin:18px 0">' + code + '</p>' +
        '<p style="color:#5B6B72;font-size:13px">Il est valable 15 minutes. Appareil : ' + html_(r.ua) + '<br>Si ce n\'est pas toi, ignore ce message et préviens Publimax.</p></div>' });
    return { ok: true, st: 'code' };
  }
  if (n.length < 2) return { ok: false, st: 'nom', err: 'Adresse inconnue : indique ton prénom et ton nom pour demander l\'accès.' };
  if (cache.get('rq_' + e)) return { ok: true, st: 'att' };
  var nb = P.getKeys().filter(function (x) { return x.indexOf('R_') === 0; }).length;
  if (!r && nb >= 60) return { ok: false, st: 'err', err: 'Trop de demandes en attente : contacte Publimax.' };
  r = { e: e, t: t, n: n, ua: String(b.ua || '').slice(0, 120), st: 'att', at: new Date().toISOString(), nonce: Utilities.getUuid() };
  P.setProperty(id, JSON.stringify(r)); cache.put('rq_' + e, '1', 90);
  mail_(id, r, b.ed === true);
  return { ok: true, st: 'att' };
}
/** Un appareil déjà autorisé pour cette adresse (de préférence avec le droit de modifier). */
function findOk_(e) {
  var all = P.getProperties(), best = null;
  for (var k in all) { if (k.indexOf('R_') !== 0) continue; try { var r = JSON.parse(all[k]); if (r.e === e && r.st === 'ok' && (!best || (r.ed && !best.ed))) best = r; } catch (x) {} }
  return best;
}
/** Nouvel appareil d'une personne déjà autorisée : le code reçu par e-mail prouve que c'est bien elle. */
function code_(b) {
  var e = String(b.e || '').trim().toLowerCase(), t = String(b.t || ''), id = rid_(e, t), cur = P.getProperty(id), r = cur ? JSON.parse(cur) : null;
  if (!r || r.st !== 'code') return { ok: false, err: 'Aucun code en attente : redemande-le.' };
  if (Date.now() > r.exp) return { ok: false, err: 'Code expiré : redemande-le.' };
  if (r.tries >= 5) return { ok: false, err: 'Trop d\'essais : redemande un code.' };
  if (String(b.code || '').replace(/\D/g, '') !== r.code) { r.tries++; P.setProperty(id, JSON.stringify(r)); return { ok: false, err: 'Code incorrect.' }; }
  r.st = 'ok'; r.ed = !!r.inh; delete r.code; delete r.inh; r.dec = new Date().toISOString(); P.setProperty(id, JSON.stringify(r));
  var k = P.getProperty('K');
  return k ? { ok: true, st: 'ok', k: k, ed: r.ed, n: r.n } : { ok: false, err: 'Clé pas encore disponible côté Publimax.' };
}
function mail_(id, r, edit) {
  var url = SITE + 'decision.html#id=' + id + '&n=' + r.nonce;
  MailApp.sendEmail({ to: Session.getEffectiveUser().getEmail(), subject: 'Survey Publimax : ' + (edit ? 'demande de modification du catalogue de ' : 'demande d\'accès de ') + (r.n || r.e),
    htmlBody: '<div style="font:16px Arial,sans-serif;color:#10242B"><p><b>' + html_(r.n) + '</b> (' + html_(r.e) + ') ' + (edit ? 'demande le droit de modifier le catalogue de l\'appli Survey.' : 'demande l\'accès à l\'appli Survey.') + '</p>' +
      '<p style="color:#5B6B72;font-size:13px">Appareil : ' + html_(r.ua) + '</p>' +
      '<p><a href="' + url + '" style="display:inline-block;padding:14px 22px;background:#00849B;color:#fff;border-radius:10px;text-decoration:none;font-weight:bold">Voir la demande et décider</a></p>' +
      '<p style="color:#5B6B72;font-size:13px">Ouvrir ce lien ne décide rien : il faut ensuite appuyer sur un des boutons.</p></div>' });
}
function st_(b) {
  var e = String(b.e || '').trim().toLowerCase(), t = String(b.t || ''), cur = P.getProperty(rid_(e, t));
  if (!cur) return { ok: true, st: 'inconnu' };
  var r = JSON.parse(cur);
  if (r.st === 'ok') { var k = P.getProperty('K'); return k ? { ok: true, st: 'ok', k: k, ed: !!r.ed } : { ok: true, st: 'ok', ed: !!r.ed }; }
  return { ok: true, st: r.st };
}
/** Appelée par la page de décision (bouton Autoriser / Refuser). */
function decide(id, nonce, yes) {
  var cur = P.getProperty(String(id)); if (!cur || String(id).indexOf('R_') !== 0) throw new Error('Demande introuvable.');
  var r = JSON.parse(cur); if (r.nonce !== nonce) throw new Error('Lien invalide.');
  r.st = yes ? 'ok' : 'refus'; r.dec = new Date().toISOString(); P.setProperty(String(id), JSON.stringify(r));
  return r.st;
}
/** Page de décision hébergée avec l'appli (évite la page Google qui casse quand plusieurs comptes sont connectés). */
var SITE = 'https://fabienpublimax82-sketch.github.io/publimax-catalogue/';
function dec_(b) {
  var id = String(b.id || ''), cur = id.indexOf('R_') === 0 ? P.getProperty(id) : null, r = cur ? JSON.parse(cur) : null;
  if (!r || !b.n || r.nonce !== String(b.n)) return { ok: false, err: 'Lien invalide ou demande supprimée.' };
  if (b.a === 'dec') { r.st = b.yes === true ? 'ok' : 'refus'; r.ed = b.yes === true && b.ed === true; r.dec = new Date().toISOString(); P.setProperty(id, JSON.stringify(r)); }
  return { ok: true, r: { n: r.n, e: r.e, ua: r.ua, at: r.at, st: r.st, ed: !!r.ed } };
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
