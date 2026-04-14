// ══════════════════════════════════════════════════════════════════════════════
// Gebouw 5B Digital Twin — Google Apps Script
// Versie: april 2026 — inclusief telefoonbeheer + mobiel medewerker
// ══════════════════════════════════════════════════════════════════════════════

const SHEET_REGS     = 'Registraties';
const SHEET_DAILY    = 'Dagoverzicht';
const SHEET_ANALYSIS = 'Periodeanalyse';
const SHEET_TELEFOON = 'Telefoon';

// ── E-mail coordinator (ontvangt CC bij elke ziekmelding) ─────────────────
const CC_COORDINATOR = 'esveldav@gmail.com';   // ← later bijwerken naar Rachel

// ── Mail test: voer deze functie 1x handmatig uit in de GAS editor ────────
function testMailToestemming() {
  try {
    MailApp.sendEmail({
      to:      CC_COORDINATOR,
      subject: 'GAS Mail test — Gebouw 5B',
      body:    'Als je dit ziet werkt MailApp correct.\n\n— Gebouw 5B Digital Twin'
    });
    Logger.log('Testmail verstuurd naar ' + CC_COORDINATOR);
  } catch(e) {
    Logger.log('Testmail mislukt: ' + e.toString());
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// DISPATCHER
// ══════════════════════════════════════════════════════════════════════════════

function doGet(e) {
  return handleRequest((e && e.parameter) || {}, null);
}

function doPost(e) {
  let body = {};
  try {
    const raw = (e && e.postData && e.postData.contents) || '{}';
    body = JSON.parse(raw);
  } catch(err) {}
  return handleRequest((e && e.parameter) || {}, body);
}

function handleRequest(params, body) {
  const action = params.action || (body && body.action) || '';
  try {
    let result;
    // ── Vocht & Voeding ──────────────────────────────────────────────────
    if      (action === 'ping')               result = handlePing(params);
    else if (action === 'push')               result = handlePush(body);
    else if (action === 'pull')               result = handlePull(params, body);
    else if (action === 'periode')            result = handlePeriode(params);
    // ── Dienst ──────────────────────────────────────────────────────────
    else if (action === 'pushDienst')         result = handlePushDienst(body);
    else if (action === 'pullDienst')         result = handlePullDienst();
    else if (action === 'uitDienst')          result = handleUitDienst(body);
    // ── Chantal ─────────────────────────────────────────────────────────
    else if (action === 'getChantal')         result = handleGetChantal();
    else if (action === 'setChantal')         result = handleSetChantal(body);
    // ── Chat ────────────────────────────────────────────────────────────
    else if (action === 'pushChat')           result = handlePushChat(body);
    else if (action === 'pullChat')           result = handlePullChat();
    else if (action === 'deleteChat')         result = handleDeleteChat(body);
    // ── Stagiairs ────────────────────────────────────────────────────────
    else if (action === 'pushStagiairs')      result = handlePushStagiairs(body);
    else if (action === 'pullStagiairs')      result = handlePullStagiairs();
    else if (action === 'deleteStagiair')     result = handleDeleteStagiair(body);
    // ── Ziekmeldingen ────────────────────────────────────────────────────
    else if (action === 'pushZiekmelding')    result = handlePushZiekmelding(body);
    else if (action === 'pullZiekmeldingen')  result = handlePullZiekmeldingen();
    else if (action === 'updateZiekmelding')  result = handleUpdateZiekmelding(body);
    // ── Roosterplanner ──────────────────────────────────────────────────────
    else if (action === 'pushRooster')       result = handlePushRooster(body);
    else if (action === 'pullRooster')        result = handlePullRooster(params);
    // ── Telefoon ─────────────────────────────────────────────────────────
    else if (action === 'pushTelefoon')       result = handlePushTelefoon(body);
    else if (action === 'pullTelefoon')       result = handlePullTelefoon();
    // ── Gebruikersbeheer (auth) ──────────────────────────────────────────
    else if (action === 'getUser')            result = handleGetUser(params);
    else if (action === 'createUser')         result = handleCreateUser(body);
    else if (action === 'updateUser')         result = handleUpdateUser(body);
    else if (action === 'listUsers')          result = handleListUsers();
    // ── Onbekend ─────────────────────────────────────────────────────────
    else result = { ok: false, error: 'Onbekende actie: ' + action };
    return jsonResponse(result);
  } catch(err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// GEBRUIKERSBEHEER (AUTH)
// ══════════════════════════════════════════════════════════════════════════════

function initGebruikersSheet() {
  const sheet = getOrCreateSheet('Gebruikers', ['naam','wachtwoord','weergavenaam','rol','actief']);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]) === 'beheerder') {
      Logger.log('Al een beheerder aanwezig: ' + data[i][0]);
      return;
    }
  }
  sheet.appendRow(['admin', 'Welkom01!', 'Beheerder', 'beheerder', true]);
  Logger.log('Beheerder "admin" aangemaakt met wachtwoord "Welkom01!" — verander dit direct!');
}

function handleGetUser(params) {
  const naam = (params.naam || '').trim().toLowerCase();
  const ww   = params.ww   || '';
  if (!naam || !ww) return { ok: false, error: 'Naam of wachtwoord ontbreekt' };

  const sheet = getOrCreateSheet('Gebruikers', ['naam','wachtwoord','weergavenaam','rol','actief']);
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rij = data[i];
    if (String(rij[0]).trim().toLowerCase() === naam) {
      if (!rij[4]) return { ok: false, error: 'Account is gedeactiveerd' };
      if (String(rij[1]) !== ww) return { ok: false, error: 'Wachtwoord onjuist' };
      return {
        ok:   true,
        user: {
          naam:         String(rij[0]),
          weergavenaam: String(rij[2] || rij[0]),
          rol:          String(rij[3] || 'begeleider'),
          actief:       true
        }
      };
    }
  }
  return { ok: false, error: 'Gebruiker niet gevonden' };
}

function handleCreateUser(body) {
  const naam     = (body.naam         || '').trim().toLowerCase();
  const weergave = (body.weergavenaam || '').trim();
  const ww       =  body.wachtwoord   || '';
  const rol      =  body.rol          || 'begeleider';

  if (!naam || !weergave || !ww) return { ok: false, error: 'Velden incompleet' };

  const geldigeRollen = ['beheerder','begeleider','stagebegeleider','stagiair','ciop'];
  if (!geldigeRollen.includes(rol)) return { ok: false, error: 'Ongeldig rol: ' + rol };

  const sheet = getOrCreateSheet('Gebruikers', ['naam','wachtwoord','weergavenaam','rol','actief']);
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === naam) {
      return { ok: false, error: 'Gebruikersnaam bestaat al' };
    }
  }

  sheet.appendRow([naam, ww, weergave, rol, true]);
  return { ok: true };
}

function handleUpdateUser(body) {
  const naam = (body.naam || '').trim().toLowerCase();
  if (!naam) return { ok: false, error: 'Geen naam opgegeven' };

  const sheet = getOrCreateSheet('Gebruikers', ['naam','wachtwoord','weergavenaam','rol','actief']);
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === naam) {
      if (body.wachtwoord !== undefined) sheet.getRange(i+1, 2).setValue(body.wachtwoord);
      if (body.actief     !== undefined) sheet.getRange(i+1, 5).setValue(body.actief);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Gebruiker niet gevonden' };
}

function handleListUsers() {
  const sheet = getOrCreateSheet('Gebruikers', ['naam','wachtwoord','weergavenaam','rol','actief']);
  const data  = sheet.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    users.push({
      naam:         String(data[i][0]),
      weergavenaam: String(data[i][2] || data[i][0]),
      rol:          String(data[i][3] || 'begeleider'),
      actief:       data[i][4] === true || String(data[i][4]).toUpperCase() === 'TRUE'
    });
  }
  return { ok: true, users };
}

// ══════════════════════════════════════════════════════════════════════════════
// ZIEKMELDINGEN
// ══════════════════════════════════════════════════════════════════════════════

function handlePushZiekmelding(body) {
  const z = body && body.ziekmelding;
  if (!z) return { ok: false, error: 'Geen ziekmelding' };

  const sheet = getOrCreateSheet('Ziekmeldingen', [
    'id','naam','email','etages','tijdIn','datum','status','aandachtOp','opgelostOp'
  ]);

  sheet.appendRow([
    z.id, z.naam, z.email || '',
    "'" + (z.etages || []).join(';'),
    z.tijdIn, z.datum, 'actief', '', ''
  ]);

  const etageLabels = (z.etages || []).map(function(e) {
    return ['BG','1e','2e','3e'][e] || e;
  }).join(', ');

  if (z.email) {
    try {
      MailApp.sendEmail({
        to:      z.email,
        subject: 'Ziekmelding ontvangen — Gebouw 5B',
        body:    'Hallo ' + z.naam + ',\n\n' +
                 'Je ziekmelding is ontvangen op ' + z.datum + ' om ' + z.tijdIn + '.\n' +
                 'Verdieping(en): ' + etageLabels + '.\n\n' +
                 'De coördinator is op de hoogte gesteld. Beterschap!\n\n' +
                 '— Gebouw 5B Digital Twin'
      });
    } catch(e) { Logger.log('Mail medewerker fout: ' + e.toString()); }
  }

  try {
    MailApp.sendEmail({
      to:      CC_COORDINATOR,
      subject: '🤒 Ziekmelding: ' + z.naam + ' (' + etageLabels + ')',
      body:    'Nieuwe ziekmelding binnengekomen.\n\n' +
               'Naam:           ' + z.naam + '\n' +
               'E-mail:         ' + (z.email || '—') + '\n' +
               'Datum:          ' + z.datum + '\n' +
               'Tijdstip:       ' + z.tijdIn + '\n' +
               'Verdieping(en): ' + etageLabels + '\n\n' +
               '— Gebouw 5B Digital Twin'
    });
  } catch(e) { Logger.log('Mail coördinator fout: ' + e.toString()); }

  return { ok: true };
}

function handlePullZiekmeldingen() {
  const sheet = getOrCreateSheet('Ziekmeldingen', [
    'id','naam','email','etages','tijdIn','datum','status','aandachtOp','opgelostOp'
  ]);
  const data  = sheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const lijst = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const datum = row[5] instanceof Date
      ? Utilities.formatDate(row[5], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(row[5]).substring(0, 10);
    if (datum !== today) continue;
    lijst.push({
      id:         String(row[0]),
      naam:       String(row[1]),
      email:      String(row[2] || ''),
      etages:     row[3] ? String(row[3]).replace(/^'/,'').split(';')
                    .map(function(x){ return parseInt(x.trim(), 10); })
                    .filter(function(n){ return !isNaN(n); }) : [],
      tijdIn:     String(row[4]),
      datum:      datum,
      status:     String(row[6] || 'actief'),
      aandachtOp: String(row[7] || ''),
      opgelostOp: String(row[8] || '')
    });
  }
  return { ok: true, ziekmeldingen: lijst };
}

function handleUpdateZiekmelding(body) {
  const id     = body && body.id;
  const status = body && body.status;
  if (!id || !status) return { ok: false, error: 'Geen id of status' };

  const sheet = getOrCreateSheet('Ziekmeldingen', [
    'id','naam','email','etages','tijdIn','datum','status','aandachtOp','opgelostOp'
  ]);
  const data = sheet.getDataRange().getValues();
  const now  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i+1, 7).setValue(status);
      if (status === 'aandacht') sheet.getRange(i+1, 8).setValue(now);
      if (status === 'opgelost') sheet.getRange(i+1, 9).setValue(now);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Niet gevonden' };
}

// ══════════════════════════════════════════════════════════════════════════════
// VOCHT & VOEDING
// ══════════════════════════════════════════════════════════════════════════════

function handlePing(params) {
  ensureAllSheets();
  return {
    ok:       true,
    sheet:    SpreadsheetApp.getActiveSpreadsheet().getName(),
    location: params.location || ''
  };
}

function handlePush(body) {
  const records  = body.records  || [];
  const location = body.location || 'Onbekend';
  if (!records.length && !(body.people||[]).length) return { ok: true, added: 0 };

  ensureAllSheets();
  const sheet    = getOrCreateSheet(SHEET_REGS);
  const existing = getExistingIds(sheet);
  let added = 0;

  records.forEach(r => {
    if (existing.has(String(r.id))) return;
    sheet.appendRow([
      r.id, r.person, r.iconId, r.cat, r.type,
      r.amount || 0, r.time || 0, r.note || '',
      r.date || '', r.dropGroup || '', r.dropIdx || 0,
      location, new Date().toISOString()
    ]);
    added++;
  });

  if (added > 0) updateDailyOverview();

  const people = body.people || [];
  if (people.length) {
    const pSheet = getOrCreateSheet('Personen', ['id','name','initials','color','goal','room','createdBy']);
    const pData  = pSheet.getDataRange().getValues();
    const existingPids = new Set();
    for (let i = 1; i < pData.length; i++) if (pData[i][0]) existingPids.add(String(pData[i][0]));
    people.forEach(p => {
      if (!existingPids.has(String(p.id))) {
        pSheet.appendRow([p.id, p.name||'', p.initials||'', p.color||'', p.goal||1500, p.room||'', p.createdBy||'']);
      } else {
        const fresh = pSheet.getDataRange().getValues();
        for (let i = 1; i < fresh.length; i++) {
          if (String(fresh[i][0]) === String(p.id)) {
            if (p.room      !== undefined) pSheet.getRange(i+1, 6).setValue(p.room);
            if (p.goal      !== undefined) pSheet.getRange(i+1, 5).setValue(p.goal||1500);
            if (p.initials  !== undefined) pSheet.getRange(i+1, 3).setValue(p.initials||'');
            if (p.createdBy !== undefined && !fresh[i][6]) pSheet.getRange(i+1, 7).setValue(p.createdBy);
            break;
          }
        }
      }
    });
  }

  const updates = body.updates || [];
  if (updates.length) {
    const data2 = sheet.getDataRange().getValues();
    updates.forEach(u => {
      for (let i = 1; i < data2.length; i++) {
        if (String(data2[i][0]) === String(u.id)) {
          sheet.getRange(i+1, 6).setValue(u.amount||0);
          sheet.getRange(i+1, 7).setValue(u.time||0);
          sheet.getRange(i+1, 8).setValue(u.note||'');
          sheet.getRange(i+1, 2).setValue(u.person||'');
          break;
        }
      }
    });
  }

  const deleteIds = (body.deletes || []).map(String);
  if (deleteIds.length) {
    const delData = sheet.getDataRange().getValues();
    for (let i = delData.length - 1; i >= 1; i--) {
      if (deleteIds.indexOf(String(delData[i][0])) !== -1) sheet.deleteRow(i + 1);
    }
  }

  return { ok: true, added, updated: updates.length, deleted: deleteIds.length };
}

function handlePull(params, body) {
  const date = params.date || (body && body.date) || '';
  ensureAllSheets();
  const sheet = getOrCreateSheet(SHEET_REGS);
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return { ok: true, records: [], people: [] };

  const records = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const rec = {
      id: row[0], person: row[1], iconId: row[2], cat: row[3], type: row[4],
      amount: Number(row[5]), time: Number(row[6]), note: row[7],
      date: row[8] instanceof Date
        ? Utilities.formatDate(row[8], Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(row[8]).substring(0, 10)
    };
    if (date && String(rec.date).substring(0, 10) !== date) continue;
    records.push(rec);
  }

  const pSheet = getOrCreateSheet('Personen', ['id','name','initials','color','goal','room','createdBy']);
  const pData  = pSheet.getDataRange().getValues();
  const people = [];
  for (let i = 1; i < pData.length; i++) {
    if (!pData[i][0]) continue;
    people.push({
      id:        String(pData[i][0]),
      name:      String(pData[i][1]),
      initials:  String(pData[i][2]),
      color:     String(pData[i][3]),
      goal:      Number(pData[i][4]) || 1500,
      room:      String(pData[i][5] || ''),
      createdBy: String(pData[i][6] || '')
    });
  }

  return { ok: true, records, people, count: records.length };
}

function handlePeriode(params) {
  const fromDate = params.from || '';
  const toDate   = params.to   || '';
  if (!fromDate || !toDate) return { ok: false, error: 'Geef from en to op' };
  ensureAllSheets();
  const sheet = getOrCreateSheet(SHEET_REGS);
  const data  = sheet.getDataRange().getValues();
  const byPersonDay = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const personId = String(row[1]), type = String(row[4]), amount = Number(row[5]), date = String(row[8]);
    if (type !== 'drink') continue;
    if (date < fromDate || date > toDate) continue;
    if (!byPersonDay[personId]) byPersonDay[personId] = {};
    byPersonDay[personId][date] = (byPersonDay[personId][date] || 0) + amount;
  }
  const result = Object.keys(byPersonDay).map(pid => {
    const days = byPersonDay[pid], values = Object.values(days).filter(v => v > 0);
    return {
      person:       pid,
      days,
      total:        values.reduce((s,v) => s+v, 0),
      avg:          values.length ? Math.round(values.reduce((s,v) => s+v, 0) / values.length) : 0,
      daysWithData: values.length,
      min:          values.length ? Math.min(...values) : 0,
      max:          values.length ? Math.max(...values) : 0
    };
  });
  return { ok: true, data: result, from: fromDate, to: toDate };
}

// ══════════════════════════════════════════════════════════════════════════════
// DIENST
// Uitgebreid met mobiel + mobielZichtbaar velden
// ══════════════════════════════════════════════════════════════════════════════

function handlePushDienst(body) {
  const r = body.record;
  if (!r) return { ok: false, error: 'Geen record' };
  const sheet = getOrCreateSheet('InDienst', [
    'id','naam','afkorting','functie','kleur','etage',
    'koppels','stgBegeleider','stgNrs','tijdIn','datum',
    'mobiel','mobielZichtbaar'          // ← nieuw
  ]);
  const data  = sheet.getDataRange().getValues();
  const today = r.datum || '';
  // Verwijder eventuele eerdere rij van dezelfde persoon vandaag
  for (let i = data.length-1; i >= 1; i--) {
    if (String(data[i][1]) === String(r.naam) && String(data[i][10]).substring(0,10) === today)
      sheet.deleteRow(i+1);
  }
  sheet.appendRow([
    r.id, r.naam, r.afkorting, r.functie, r.kleur, r.etage,
    (r.koppels||[]).join(','),
    r.stgBegeleider ? 'ja' : 'nee',
    (r.stgNrs||[]).join(','),
    r.tijdIn, r.datum,
    r.mobiel          || '',            // ← nieuw
    r.mobielZichtbaar ? 'ja' : 'nee'   // ← nieuw
  ]);
  return { ok: true };
}

function handlePullDienst() {
  const sheet = getOrCreateSheet('InDienst', [
    'id','naam','afkorting','functie','kleur','etage',
    'koppels','stgBegeleider','stgNrs','tijdIn','datum',
    'mobiel','mobielZichtbaar'
  ]);
  const data  = sheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const records = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const datum = row[10] instanceof Date
      ? Utilities.formatDate(row[10], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(row[10]).substring(0, 10);
    if (datum !== today) continue;
    records.push({
      id:              String(row[0]),
      naam:            String(row[1]),
      afkorting:       String(row[2]),
      functie:         String(row[3]),
      kleur:           String(row[4]),
      etage:           Number(row[5]),
      koppels:         row[6] ? String(row[6]).split(',') : [],
      stgBegeleider:   row[7] === 'ja',
      stgNrs:          row[8] ? String(row[8]).split(',') : [],
      tijdIn:          String(row[9]),
      mobiel:          String(row[11] || ''),          // ← nieuw
      mobielZichtbaar: row[12] === 'ja'                // ← nieuw
    });
  }
  return { ok: true, records };
}

function handleUitDienst(body) {
  const id = body && body.id;
  if (!id) return { ok: false, error: 'Geen id' };
  const sheet = getOrCreateSheet('InDienst', [
    'id','naam','afkorting','functie','kleur','etage',
    'koppels','stgBegeleider','stgNrs','tijdIn','datum',
    'mobiel','mobielZichtbaar'
  ]);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length-1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { ok: true }; }
  }
  return { ok: true };
}

// ══════════════════════════════════════════════════════════════════════════════
// CHANTAL
// ══════════════════════════════════════════════════════════════════════════════

function handleGetChantal() {
  const sheet = getOrCreateSheet('Chantal', ['status','timestamp']);
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2 || !data[1][0]) {
    if (data.length < 2) sheet.appendRow(['welkom', new Date().toISOString()]);
    else sheet.getRange(2, 1, 1, 2).setValues([['welkom', new Date().toISOString()]]);
    return { ok: true, status: 'welkom' };
  }
  return { ok: true, status: String(data[1][0]) };
}

function handleSetChantal(body) {
  const status = body && body.status;
  if (status !== 'welkom' && status !== 'bezet') return { ok: false, error: 'Ongeldige status' };
  const sheet = getOrCreateSheet('Chantal', ['status','timestamp']);
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) sheet.appendRow([status, new Date().toISOString()]);
  else sheet.getRange(2, 1, 1, 2).setValues([[status, new Date().toISOString()]]);
  return { ok: true, status };
}

// ══════════════════════════════════════════════════════════════════════════════
// CHAT
// ══════════════════════════════════════════════════════════════════════════════

function handlePushChat(body) {
  const msg = body && body.message;
  if (!msg || !msg.naam || !msg.tekst) return { ok: false, error: 'Onvolledig bericht' };
  const sheet = getOrCreateSheet('Chat', ['id','naam','afkorting','kleur','tekst','tijd','datum']);
  sheet.appendRow([msg.id, msg.naam, msg.afkorting||'', msg.kleur||'#888888', msg.tekst, msg.tijd, msg.datum]);
  return { ok: true };
}

function handlePullChat() {
  const sheet = getOrCreateSheet('Chat', ['id','naam','afkorting','kleur','tekst','tijd','datum']);
  const data  = sheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const messages = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const datum = row[6] instanceof Date
      ? Utilities.formatDate(row[6], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(row[6]).substring(0, 10);
    if (datum !== today) continue;
    messages.push({
      id:        String(row[0]),
      naam:      String(row[1]),
      afkorting: String(row[2]),
      kleur:     String(row[3]),
      tekst:     String(row[4]),
      tijd:      String(row[5]),
      datum
    });
  }
  return { ok: true, messages };
}

function handleDeleteChat(body) {
  const id = body && body.id;
  if (!id) return { ok: false, error: 'Geen id' };
  const sheet = getOrCreateSheet('Chat', ['id','naam','afkorting','kleur','tekst','tijd','datum']);
  const data  = sheet.getDataRange().getValues();
  for (let i = data.length-1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { ok: true }; }
  }
  return { ok: true };
}

// ══════════════════════════════════════════════════════════════════════════════
// STAGIAIRS
// ══════════════════════════════════════════════════════════════════════════════

function handlePushStagiairs(body) {
  const stagiairs = body.stagiairs;
  if (!Array.isArray(stagiairs)) return { ok: false, error: 'Geen stagiairs array' };
  const sheet = getOrCreateSheet('Stagiairs', ['id','naam','kamer','begeleider','school','start','eind','fases','kamerHistorie','aangemaaktOp','driveLink']);
  const data  = sheet.getDataRange().getValues();
  const bestaand = {};
  for (let i = 1; i < data.length; i++) { if (data[i][0]) bestaand[String(data[i][0])] = i + 1; }
  stagiairs.forEach(s => {
    const rij = [
      s.id, s.naam||'', s.kamer||'', s.begeleider||'', s.school||'',
      s.start||'', s.eind||'', JSON.stringify(s.fases||[]),
      JSON.stringify(s.kamerHistorie||[]), s.aangemaaktOp||new Date().toISOString(), s.driveLink||''
    ];
    if (bestaand[String(s.id)]) sheet.getRange(bestaand[String(s.id)], 1, 1, rij.length).setValues([rij]);
    else sheet.appendRow(rij);
  });
  return { ok: true, count: stagiairs.length };
}

function handlePullStagiairs() {
  const sheet = getOrCreateSheet('Stagiairs', ['id','naam','kamer','begeleider','school','start','eind','fases','kamerHistorie','aangemaaktOp','driveLink']);
  const data  = sheet.getDataRange().getValues();
  const stagiairs = [];
  const tz = Session.getScriptTimeZone();
  function fmtDatum(val) {
    if (!val) return '';
    if (val instanceof Date) return Utilities.formatDate(val, tz, 'dd-MM-yyyy');
    const s = String(val).trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(8,10)+'-'+s.substring(5,7)+'-'+s.substring(0,4);
    return s;
  }
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    let fases = [], kamerHistorie = [];
    try { fases         = JSON.parse(row[7]||'[]'); } catch(e) {}
    try { kamerHistorie = JSON.parse(row[8]||'[]'); } catch(e) {}
    stagiairs.push({
      id:           String(row[0]),
      naam:         String(row[1]),
      kamer:        String(row[2]),
      begeleider:   String(row[3]),
      school:       String(row[4]),
      start:        fmtDatum(row[5]),
      eind:         fmtDatum(row[6]),
      fases,
      kamerHistorie,
      aangemaaktOp: String(row[9]  || ''),
      driveLink:    String(row[10] || '')
    });
  }
  return { ok: true, stagiairs };
}

function handleDeleteStagiair(body) {
  const id = body && body.id;
  if (!id) return { ok: false, error: 'Geen id' };
  const sheet = getOrCreateSheet('Stagiairs', ['id','naam','kamer','begeleider','school','start','eind','fases','kamerHistorie','aangemaaktOp','driveLink']);
  const data  = sheet.getDataRange().getValues();
  for (let i = data.length-1; i >= 1; i--) {
    if (String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { ok: true }; }
  }
  return { ok: false, error: 'Stagiair niet gevonden' };
}

// ══════════════════════════════════════════════════════════════════════════════
// ROOSTERPLANNER
// Records met datum in de toekomst — pullDienst filtert op vandaag dus ze
// worden pas zichtbaar op de juiste dag.
// ══════════════════════════════════════════════════════════════════════════════

function handlePushRooster(body) {
  const records = body && body.records;
  if (!Array.isArray(records) || !records.length) return { ok: false, error: 'Geen records' };

  const sheet = getOrCreateSheet('InDienst', [
    'id','naam','afkorting','functie','kleur','etage',
    'koppels','stgBegeleider','stgNrs','tijdIn','datum',
    'mobiel','mobielZichtbaar'
  ]);
  const data = sheet.getDataRange().getValues();

  // Verwijder bestaande rooster-records voor dezelfde datum+naam+etage combinaties
  const teVerwijderen = new Set(records.map(r => r.datum + '|' + r.etage + '|' + r.naam));
  for (let i = data.length - 1; i >= 1; i--) {
    const rij = data[i];
    const key = String(rij[10]).substring(0,10) + '|' + String(rij[5]) + '|' + String(rij[1]);
    if (teVerwijderen.has(key)) sheet.deleteRow(i + 1);
  }

  // Schrijf nieuwe records
  records.forEach(r => {
    sheet.appendRow([
      r.id, r.naam, r.afkorting||'', r.functie||'BGLB', r.kleur||'#888888', r.etage,
      (r.koppels||[]).join(','),
      r.stgBegeleider ? 'ja' : 'nee',
      (r.stgNrs||[]).join(','),
      r.tijdIn||'07:00', r.datum,
      r.mobiel||'', r.mobielZichtbaar ? 'ja' : 'nee'
    ]);
  });

  return { ok: true, geschreven: records.length };
}

function handlePullRooster(params) {
  const van = params && params.van;
  const tot = params && params.tot;

  const sheet = getOrCreateSheet('InDienst', [
    'id','naam','afkorting','functie','kleur','etage',
    'koppels','stgBegeleider','stgNrs','tijdIn','datum',
    'mobiel','mobielZichtbaar'
  ]);
  const data    = sheet.getDataRange().getValues();
  const records = [];
  const tz      = Session.getScriptTimeZone();

  for (let i = 1; i < data.length; i++) {
    const row   = data[i];
    if (!row[0]) continue;
    const datum = row[10] instanceof Date
      ? Utilities.formatDate(row[10], tz, 'yyyy-MM-dd')
      : String(row[10]).substring(0, 10);
    if (van && datum < van) continue;
    if (tot && datum > tot) continue;
    records.push({
      id:              String(row[0]),
      naam:            String(row[1]),
      afkorting:       String(row[2]),
      functie:         String(row[3]),
      kleur:           String(row[4]),
      etage:           Number(row[5]),
      koppels:         row[6] ? String(row[6]).split(',') : [],
      stgBegeleider:   row[7] === 'ja',
      stgNrs:          row[8] ? String(row[8]).split(',') : [],
      tijdIn:          String(row[9]),
      datum,
      mobiel:          String(row[11] || ''),
      mobielZichtbaar: row[12] === 'ja'
    });
  }
  return { ok: true, records };
}


// ══════════════════════════════════════════════════════════════════════════════
// TELEFOON
// Nieuw tabblad 'Telefoon': per etage/kant (0-A t/m 3-B) intern + extern nummer
// ══════════════════════════════════════════════════════════════════════════════

function handlePushTelefoon(body) {
  const data = body && body.data;
  if (!data) return { ok: false, error: 'Geen data meegestuurd' };

  const sheet = ensureTelefoonSheet();
  const rows  = sheet.getDataRange().getValues();

  // Bouw lookup: sleutel → rijnummer (1-based)
  const rijVoor = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) rijVoor[String(rows[i][0])] = i + 1;
  }

  Object.keys(data).forEach(sleutel => {
    const val    = data[sleutel] || {};
    const intern = String(val.intern || '').trim();
    const extern = String(val.extern || '').trim();
    if (rijVoor[sleutel]) {
      sheet.getRange(rijVoor[sleutel], 1, 1, 3).setValues([[sleutel, intern, extern]]);
    } else {
      sheet.appendRow([sleutel, intern, extern]);
    }
  });

  return { ok: true, geschreven: Object.keys(data).length };
}

function handlePullTelefoon() {
  const sheet  = ensureTelefoonSheet();
  const rows   = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < rows.length; i++) {
    const sleutel = String(rows[i][0] || '').trim();
    const intern  = String(rows[i][1] || '').trim();
    const extern  = String(rows[i][2] || '').trim();
    if (sleutel) result[sleutel] = { intern, extern };
  }
  return { ok: true, telefoon: result };
}

function ensureTelefoonSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_TELEFOON);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TELEFOON);
    const hdrs = ['sleutel', 'intern', 'extern'];
    sheet.appendRow(hdrs);
    sheet.getRange(1, 1, 1, hdrs.length).setFontWeight('bold')
         .setBackground('#1a73e8').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 80);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 180);
    // Alle 8 sleutels alvast leeg aanmaken
    ['0-A','0-B','1-A','1-B','2-A','2-B','3-A','3-B'].forEach((s, i) => {
      // BG alvast vooringevuld met bekende nummers
      const intern = s === '0-A' ? '8606' : s === '0-B' ? '8604' : '';
      sheet.appendRow([s, intern, '']);
    });
  }
  return sheet;
}

// ══════════════════════════════════════════════════════════════════════════════
// HULPFUNCTIES
// ══════════════════════════════════════════════════════════════════════════════

function updateDailyOverview() {
  const regSheet   = getOrCreateSheet(SHEET_REGS);
  const dailySheet = getOrCreateSheet(SHEET_DAILY);
  const data    = regSheet.getDataRange().getValues();
  const summary = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i]; if (!row[0]) continue;
    const key    = row[1] + '|' + row[8];
    const type   = String(row[4]);
    const amount = Number(row[5]);
    const iconId = String(row[2]);
    if (!summary[key]) summary[key] = { person: row[1], date: row[8], drinkMl: 0, foodCount: 0, icons: [] };
    if (type === 'drink') summary[key].drinkMl += amount;
    else summary[key].foodCount++;
    if (!summary[key].icons.includes(iconId)) summary[key].icons.push(iconId);
  }
  dailySheet.clearContents();
  const headers = ['Persoon ID','Datum','Vocht (ml)','Maaltijden/snacks','Iconen'];
  dailySheet.appendRow(headers);
  const hr = dailySheet.getRange(1, 1, 1, headers.length);
  hr.setBackground('#0d904f'); hr.setFontColor('#ffffff'); hr.setFontWeight('bold');
  Object.values(summary)
    .sort((a,b) => {
      const da = String(a.date), db = String(b.date);
      return da !== db ? (da < db ? -1 : 1) : (String(a.person) < String(b.person) ? -1 : 1);
    })
    .forEach(r => dailySheet.appendRow([r.person, r.date, r.drinkMl, r.foodCount, r.icons.join(', ')]));
  dailySheet.autoResizeColumns(1, headers.length);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 Vocht App')
    .addItem('Periodeanalyse vernieuwen', 'buildPeriodeAnalyse')
    .addItem('Dagoverzicht vernieuwen',   'updateDailyOverview')
    .addToUi();
}

function ensureAllSheets() {
  getOrCreateSheet(SHEET_REGS,     ['id','person','iconId','cat','type','amount','time','note','date','dropGroup','dropIdx','location','opgeslagen']);
  getOrCreateSheet(SHEET_DAILY,    ['Persoon ID','Datum','Vocht (ml)','Maaltijden/snacks','Iconen']);
  getOrCreateSheet(SHEET_ANALYSIS, ['Overzicht — gebruik menu "📊 Vocht App" → Periodeanalyse vernieuwen']);
  getOrCreateSheet('Ziekmeldingen', ['id','naam','email','etages','tijdIn','datum','status','aandachtOp','opgelostOp']);
  getOrCreateSheet('Gebruikers',   ['naam','wachtwoord','weergavenaam','rol','actief']);
  ensureTelefoonSheet();  // ← nieuw
}

function getOrCreateSheet(name, headers) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers) {
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
      const hr = sheet.getRange(1, 1, 1, headers.length);
      hr.setBackground('#1a73e8'); hr.setFontColor('#ffffff'); hr.setFontWeight('bold');
    }
  }
  return sheet;
}

function getExistingIds(sheet) {
  const data = sheet.getDataRange().getValues();
  const ids  = new Set();
  for (let i = 1; i < data.length; i++) if (data[i][0]) ids.add(String(data[i][0]));
  return ids;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
