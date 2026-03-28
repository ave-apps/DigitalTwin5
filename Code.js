const SHEET_REGS     = 'Registraties';
const SHEET_DAILY    = 'Dagoverzicht';
const SHEET_ANALYSIS = 'Periodeanalyse';

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
    if      (action === 'ping')        result = handlePing(params);
    else if (action === 'push')        result = handlePush(body);
    else if (action === 'pull')        result = handlePull(params, body);
    else if (action === 'periode')     result = handlePeriode(params);
    else if (action === 'pushDienst')  result = handlePushDienst(body);
    else if (action === 'pullDienst')  result = handlePullDienst();
    else if (action === 'uitDienst')   result = handleUitDienst(body);
    else if (action === 'getChantal')  result = handleGetChantal();
    else if (action === 'setChantal')  result = handleSetChantal(body);
    else if (action === 'pushChat')    result = handlePushChat(body);
    else if (action === 'pullChat')    result = handlePullChat();
    else if (action === 'deleteChat')  result = handleDeleteChat(body);
    else if (action === 'pushStagiairs')      result = handlePushStagiairs(body);
    else if (action === 'pullStagiairs')      result = handlePullStagiairs();
    else if (action === 'deleteStagiair')     result = handleDeleteStagiair(body);
    else if (action === 'pushZiekmelding')    result = handlePushZiekmelding(body);
    else if (action === 'pullZiekmeldingen')  result = handlePullZiekmeldingen();
    else if (action === 'updateZiekmelding')  result = handleUpdateZiekmelding(body);
    else                               result = { ok: false, error: 'Onbekende actie: ' + action };
    return jsonResponse(result);
  } catch(err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

// ── Ziekmeldingen ──────────────────────────────────────────────────────────

function handlePushZiekmelding(body) {
  const z = body && body.ziekmelding;
  if(!z) return { ok: false, error: 'Geen ziekmelding' };

  const sheet = getOrCreateSheet('Ziekmeldingen', [
    'id','naam','email','etages','tijdIn','datum','status','aandachtOp','opgelostOp'
  ]);

  sheet.appendRow([
    z.id,
    z.naam,
    z.email || '',
    (z.etages || []).join(','),
    z.tijdIn,
    z.datum,
    'actief',
    '',
    ''
  ]);

  // Stuur bevestigingsmail naar medewerker
  if(z.email) {
    try {
      const subject = 'Ziekmelding ontvangen — Gebouw 5B';
      const body_   = 'Hallo ' + z.naam + ',\n\n' +
        'Je ziekmelding is ontvangen op ' + z.datum + ' om ' + z.tijdIn + '.\n' +
        'Verdiepingen: ' + (z.etages || []).map(function(e){ return ['BG','1e','2e','3e'][e]||e; }).join(', ') + '.\n\n' +
        'De coördinator is op de hoogte gesteld. Beterschap!\n\n' +
        '— Gebouw 5B Digital Twin';
      MailApp.sendEmail(z.email, subject, body_);
    } catch(e) {
      Logger.log('Mail fout: ' + e.toString());
    }
  }

  return { ok: true };
}

function handlePullZiekmeldingen() {
  const sheet = getOrCreateSheet('Ziekmeldingen', [
    'id','naam','email','etages','tijdIn','datum','status','aandachtOp','opgelostOp'
  ]);
  const data  = sheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const lijst = [];
  for(let i = 1; i < data.length; i++) {
    const row = data[i];
    if(!row[0]) continue;
    const datum = row[5] instanceof Date
      ? Utilities.formatDate(row[5], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(row[5]).substring(0, 10);
    if(datum !== today) continue;
    lijst.push({
      id:         String(row[0]),
      naam:       String(row[1]),
      email:      String(row[2] || ''),
      etages:     row[3] ? String(row[3]).split(',').map(Number) : [],
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
  const status = body && body.status; // 'aandacht' | 'opgelost'
  if(!id || !status) return { ok: false, error: 'Geen id of status' };

  const sheet = getOrCreateSheet('Ziekmeldingen', [
    'id','naam','email','etages','tijdIn','datum','status','aandachtOp','opgelostOp'
  ]);
  const data = sheet.getDataRange().getValues();
  const now  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');

  for(let i = 1; i < data.length; i++) {
    if(String(data[i][0]) === String(id)) {
      sheet.getRange(i+1, 7).setValue(status);
      if(status === 'aandacht')  sheet.getRange(i+1, 8).setValue(now);
      if(status === 'opgelost')  sheet.getRange(i+1, 9).setValue(now);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Niet gevonden' };
}

function handlePing(params) {
  ensureAllSheets();
  return {
    ok: true,
    sheet: SpreadsheetApp.getActiveSpreadsheet().getName(),
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
        : String(row[8]).substring(0, 10),
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
      id: String(pData[i][0]), name: String(pData[i][1]), initials: String(pData[i][2]),
      color: String(pData[i][3]), goal: Number(pData[i][4]) || 1500,
      room: String(pData[i][5] || ''), createdBy: String(pData[i][6] || '')
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
    return { person: pid, days, total: values.reduce((s,v)=>s+v, 0),
      avg: values.length ? Math.round(values.reduce((s,v)=>s+v,0)/values.length) : 0,
      daysWithData: values.length, min: values.length ? Math.min(...values) : 0, max: values.length ? Math.max(...values) : 0 };
  });
  return { ok: true, data: result, from: fromDate, to: toDate };
}

function handlePushDienst(body) {
  const r = body.record;
  if(!r) return { ok: false, error: 'Geen record' };
  const sheet = getOrCreateSheet('InDienst', ['id','naam','afkorting','functie','kleur','etage','koppels','stgBegeleider','stgNrs','tijdIn','datum']);
  const data = sheet.getDataRange().getValues();
  const today = r.datum || '';
  for(let i = data.length-1; i >= 1; i--) {
    if(String(data[i][1]) === String(r.naam) && String(data[i][10]).substring(0,10) === today) sheet.deleteRow(i+1);
  }
  sheet.appendRow([r.id, r.naam, r.afkorting, r.functie, r.kleur, r.etage,
    (r.koppels||[]).join(','), r.stgBegeleider ? 'ja' : 'nee',
    (r.stgNrs||[]).join(','), r.tijdIn, r.datum]);
  return { ok: true };
}

function handlePullDienst() {
  const sheet = getOrCreateSheet('InDienst', ['id','naam','afkorting','functie','kleur','etage','koppels','stgBegeleider','stgNrs','tijdIn','datum']);
  const data  = sheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const records = [];
  for(let i = 1; i < data.length; i++) {
    const row = data[i];
    if(!row[0]) continue;
    const datum = row[10] instanceof Date
      ? Utilities.formatDate(row[10], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(row[10]).substring(0,10);
    if(datum !== today) continue;
    records.push({ id: String(row[0]), naam: String(row[1]), afkorting: String(row[2]),
      functie: String(row[3]), kleur: String(row[4]), etage: Number(row[5]),
      koppels: row[6] ? String(row[6]).split(',') : [],
      stgBegeleider: row[7] === 'ja',
      stgNrs: row[8] ? String(row[8]).split(',') : [],
      tijdIn: String(row[9]) });
  }
  return { ok: true, records };
}

function handleUitDienst(body) {
  const id = body && body.id;
  if(!id) return { ok: false, error: 'Geen id' };
  const sheet = getOrCreateSheet('InDienst', ['id','naam','afkorting','functie','kleur','etage','koppels','stgBegeleider','stgNrs','tijdIn','datum']);
  const data  = sheet.getDataRange().getValues();
  for(let i = data.length-1; i >= 1; i--) {
    if(String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { ok: true }; }
  }
  return { ok: true };
}

function handleGetChantal() {
  const sheet = getOrCreateSheet('Chantal', ['status','timestamp']);
  const data  = sheet.getDataRange().getValues();
  if(data.length < 2 || !data[1][0]) {
    if(data.length < 2) sheet.appendRow(['welkom', new Date().toISOString()]);
    else sheet.getRange(2, 1, 1, 2).setValues([['welkom', new Date().toISOString()]]);
    return { ok: true, status: 'welkom' };
  }
  return { ok: true, status: String(data[1][0]) };
}

function handleSetChantal(body) {
  const status = body && body.status;
  if(status !== 'welkom' && status !== 'bezet') return { ok: false, error: 'Ongeldige status' };
  const sheet = getOrCreateSheet('Chantal', ['status','timestamp']);
  const data  = sheet.getDataRange().getValues();
  if(data.length < 2) sheet.appendRow([status, new Date().toISOString()]);
  else sheet.getRange(2, 1, 1, 2).setValues([[status, new Date().toISOString()]]);
  return { ok: true, status };
}

function handlePushChat(body) {
  const msg = body && body.message;
  if(!msg || !msg.naam || !msg.tekst) return { ok: false, error: 'Onvolledig bericht' };
  const sheet = getOrCreateSheet('Chat', ['id','naam','afkorting','kleur','tekst','tijd','datum']);
  sheet.appendRow([msg.id, msg.naam, msg.afkorting || '', msg.kleur || '#888888', msg.tekst, msg.tijd, msg.datum]);
  return { ok: true };
}

function handlePullChat() {
  const sheet = getOrCreateSheet('Chat', ['id','naam','afkorting','kleur','tekst','tijd','datum']);
  const data  = sheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const messages = [];
  for(let i = 1; i < data.length; i++) {
    const row = data[i];
    if(!row[0]) continue;
    const datum = row[6] instanceof Date
      ? Utilities.formatDate(row[6], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(row[6]).substring(0, 10);
    if(datum !== today) continue;
    messages.push({ id: String(row[0]), naam: String(row[1]), afkorting: String(row[2]),
      kleur: String(row[3]), tekst: String(row[4]), tijd: String(row[5]), datum });
  }
  return { ok: true, messages };
}

function handleDeleteChat(body) {
  const id = body && body.id;
  if(!id) return { ok: false, error: 'Geen id' };
  const sheet = getOrCreateSheet('Chat', ['id','naam','afkorting','kleur','tekst','tijd','datum']);
  const data  = sheet.getDataRange().getValues();
  for(let i = data.length-1; i >= 1; i--){
    if(String(data[i][0]) === String(id)){ sheet.deleteRow(i+1); return { ok: true }; }
  }
  return { ok: true };
}

function handlePushStagiairs(body) {
  const stagiairs = body.stagiairs;
  if(!Array.isArray(stagiairs)) return { ok: false, error: 'Geen stagiairs array' };
  const sheet = getOrCreateSheet('Stagiairs', ['id','naam','kamer','begeleider','school','start','eind','fases','kamerHistorie','aangemaaktOp','driveLink']);
  const data  = sheet.getDataRange().getValues();
  const bestaand = {};
  for(let i = 1; i < data.length; i++){ if(data[i][0]) bestaand[String(data[i][0])] = i + 1; }
  stagiairs.forEach(s => {
    const rij = [s.id, s.naam||'', s.kamer||'', s.begeleider||'', s.school||'',
      s.start||'', s.eind||'', JSON.stringify(s.fases||[]),
      JSON.stringify(s.kamerHistorie||[]), s.aangemaaktOp||new Date().toISOString(), s.driveLink||''];
    if(bestaand[String(s.id)]) sheet.getRange(bestaand[String(s.id)], 1, 1, rij.length).setValues([rij]);
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
    if(!val) return '';
    if(val instanceof Date) return Utilities.formatDate(val, tz, 'dd-MM-yyyy');
    const s = String(val).trim();
    if(/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(8,10)+'-'+s.substring(5,7)+'-'+s.substring(0,4);
    return s;
  }
  for(let i = 1; i < data.length; i++){
    const row = data[i];
    if(!row[0]) continue;
    let fases = [], kamerHistorie = [];
    try { fases = JSON.parse(row[7]||'[]'); } catch(e){}
    try { kamerHistorie = JSON.parse(row[8]||'[]'); } catch(e){}
    stagiairs.push({ id: String(row[0]), naam: String(row[1]), kamer: String(row[2]),
      begeleider: String(row[3]), school: String(row[4]),
      start: fmtDatum(row[5]), eind: fmtDatum(row[6]),
      fases, kamerHistorie, aangemaaktOp: String(row[9]||''), driveLink: String(row[10]||'') });
  }
  return { ok: true, stagiairs };
}

function handleDeleteStagiair(body) {
  const id = body && body.id;
  if(!id) return { ok: false, error: 'Geen id' };
  const sheet = getOrCreateSheet('Stagiairs', ['id','naam','kamer','begeleider','school','start','eind','fases','kamerHistorie','aangemaaktOp','driveLink']);
  const data  = sheet.getDataRange().getValues();
  for(let i = data.length-1; i >= 1; i--){
    if(String(data[i][0]) === String(id)){ sheet.deleteRow(i+1); return { ok: true }; }
  }
  return { ok: false, error: 'Stagiair niet gevonden' };
}

function updateDailyOverview() {
  const regSheet   = getOrCreateSheet(SHEET_REGS);
  const dailySheet = getOrCreateSheet(SHEET_DAILY);
  const data = regSheet.getDataRange().getValues();
  const summary = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i]; if (!row[0]) continue;
    const key = `${row[1]}|${row[8]}`, type = String(row[4]), amount = Number(row[5]), iconId = String(row[2]);
    if (!summary[key]) summary[key] = { person: row[1], date: row[8], drinkMl: 0, foodCount: 0, icons: [] };
    if (type === 'drink') summary[key].drinkMl += amount;
    else summary[key].foodCount++;
    if (!summary[key].icons.includes(iconId)) summary[key].icons.push(iconId);
  }
  dailySheet.clearContents();
  const headers = ['Persoon ID','Datum','Vocht (ml)','Maaltijden/snacks','Iconen'];
  dailySheet.appendRow(headers);
  const headerRange = dailySheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#0d904f'); headerRange.setFontColor('#ffffff'); headerRange.setFontWeight('bold');
  Object.values(summary).sort((a,b)=>{ const da=String(a.date),db=String(b.date); return da!==db?(da<db?-1:1):(String(a.person)<String(b.person)?-1:1); })
    .forEach(r => dailySheet.appendRow([r.person, r.date, r.drinkMl, r.foodCount, r.icons.join(', ')]));
  dailySheet.autoResizeColumns(1, headers.length);
}

function buildPeriodeAnalyse() {
  const regSheet = getOrCreateSheet(SHEET_REGS);
  const anaSheet = getOrCreateSheet(SHEET_ANALYSIS);
  const data = regSheet.getDataRange().getValues();
  const persons = new Set(), dates = new Set(), cells = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i]; if (!row[0] || row[4] !== 'drink') continue;
    const pid = String(row[1]), date = String(row[8]), ml = Number(row[5]);
    persons.add(pid); dates.add(date); cells[`${pid}|${date}`] = (cells[`${pid}|${date}`] || 0) + ml;
  }
  const sortedDates = [...dates].sort(), sortedPersons = [...persons].sort();
  anaSheet.clearContents();
  const headerRow = ['Persoon →  Datum ↓', ...sortedDates, 'Totaal', 'Gemiddelde', 'Dagen met data'];
  anaSheet.appendRow(headerRow);
  const hr = anaSheet.getRange(1, 1, 1, headerRow.length);
  hr.setBackground('#1a73e8'); hr.setFontColor('#fff'); hr.setFontWeight('bold');
  sortedPersons.forEach(pid => {
    const row = [pid], values = [];
    sortedDates.forEach(d => { const ml = cells[`${pid}|${d}`] || 0; row.push(ml || ''); if (ml > 0) values.push(ml); });
    row.push(values.reduce((s,v)=>s+v,0), values.length ? Math.round(values.reduce((s,v)=>s+v,0)/values.length) : 0, values.length);
    anaSheet.appendRow(row);
  });
  const GOAL = 1500;
  for (let r = 2; r <= sortedPersons.length + 1; r++) {
    for (let col = 2; col <= sortedDates.length + 1; col++) {
      const cell = anaSheet.getRange(r, col), val = cell.getValue();
      if (!val || val === '') continue;
      cell.setBackground(val >= GOAL ? '#c8e6c9' : val >= GOAL*.7 ? '#fff9c4' : '#ffcdd2');
    }
  }
  anaSheet.autoResizeColumns(1, headerRow.length);
  SpreadsheetApp.getUi().alert(`Periodeanalyse bijgewerkt!\n${sortedPersons.length} personen, ${sortedDates.length} dagen.`);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 Vocht App')
    .addItem('Periodeanalyse vernieuwen', 'buildPeriodeAnalyse')
    .addItem('Dagoverzicht vernieuwen',   'updateDailyOverview')
    .addToUi();
}

function ensureAllSheets() {
  getOrCreateSheet(SHEET_REGS, ['id','person','iconId','cat','type','amount','time','note','date','dropGroup','dropIdx','location','opgeslagen']);
  getOrCreateSheet(SHEET_DAILY,    ['Persoon ID','Datum','Vocht (ml)','Maaltijden/snacks','Iconen']);
  getOrCreateSheet(SHEET_ANALYSIS, ['Overzicht — gebruik menu "📊 Vocht App" → Periodeanalyse vernieuwen']);
  getOrCreateSheet('Ziekmeldingen', ['id','naam','email','etages','tijdIn','datum','status','aandachtOp','opgelostOp']);
}

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers) {
      sheet.appendRow(headers); sheet.setFrozenRows(1);
      const hr = sheet.getRange(1, 1, 1, headers.length);
      hr.setBackground('#1a73e8'); hr.setFontColor('#ffffff'); hr.setFontWeight('bold');
    }
  }
  return sheet;
}

function getExistingIds(sheet) {
  const data = sheet.getDataRange().getValues(), ids = new Set();
  for (let i = 1; i < data.length; i++) if (data[i][0]) ids.add(String(data[i][0]));
  return ids;
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
