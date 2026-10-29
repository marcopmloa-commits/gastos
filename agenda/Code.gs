/**
 * Agenda e Pendências — servidor (Google Apps Script)
 * Guarda os itens na planilha-espelho (a mesma lida pelo bot do Telegram às 6h),
 * cria/apaga eventos no Google Agenda e lê as provas do QTS.
 */
const SENHA = 'TROQUE-ESTA-SENHA';
const PLANILHA_ID = '108AjFzfaZAqGBLWBv3itJwLfC7MxkZTLqu5ETsGr_AM';
const ABA = 'Untitled';
const QTS_ID = '1ZzWxPZALV77MbtZnjtHHgSz0DSeXuJg7d3lam386LkI';
const QTS_ABA = '3º CFO_ 2025-2026 PM';
const QTS_LINHA_INI = 3, QTS_LINHA_FIM = 3561;
const CAB = ['id', 'texto', 'categoria', 'tipo', 'data', 'hora', 'feito', 'eventId', 'criado', 'horaFim', 'origem', 'feitoEm'];
const CATS = { geral: 'Geral', compras: 'Compras', cfo: 'CFO' };
const TZ = Session.getScriptTimeZone();

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var b = JSON.parse(e.postData.contents);
    if (SENHA === 'TROQUE-ESTA-SENHA' || b.token !== SENHA) return out({ error: 'Senha incorreta.', code: 'auth' });
    lock.waitLock(20000);
    var f = OPS[b.op];
    if (!f) return out({ error: 'Operação desconhecida.', code: 'tool_error' });
    return out(f(b));
  } catch (err) {
    return out({ error: String(err && err.message || err), code: 'tool_error' });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
function doGet() { return out({ ok: true, app: 'agenda' }); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- utilidades ---------- */
function pad(n) { return n < 10 ? '0' + n : '' + n; }
function sData(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v || '').trim().slice(0, 10); }
function sHora(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'HH:mm') : String(v || '').trim().slice(0, 5); }
function sTxt(v) { return v instanceof Date ? v.toISOString() : String(v == null ? '' : v); }
function primeiraLinha(t) { return String(t).split('\n')[0].trim(); }
function dt(d, h) { var p = d.split('-'), q = (h || '00:00').split(':'); return new Date(+p[0], +p[1] - 1, +p[2], +q[0], +q[1]); }
function novoId() { return 'i' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }
function slug(s) { return String(s).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, ''); }
function hoje() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

/* ---------- planilha ---------- */
function aba() {
  var sh = SpreadsheetApp.openById(PLANILHA_ID).getSheetByName(ABA);
  if (!sh) throw new Error('Aba "' + ABA + '" não encontrada na planilha.');
  if (String(sh.getRange(1, 12).getValue()) !== 'feitoEm') {
    sh.getRange(1, 1, 1, CAB.length).setValues([CAB]);
    sh.getRange(1, 1, sh.getMaxRows(), CAB.length).setNumberFormat('@');
  }
  return sh;
}
function lerItens(sh) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return [];
  var v = sh.getRange(2, 1, n, CAB.length).getValues(), r = [];
  v.forEach(function (l, i) {
    if (!l[0] || !l[1]) return;
    var it = {
      id: String(l[0]), titulo: String(l[1]), cat: String(l[2] || 'geral'), tipo: String(l[3] || 'semprazo'),
      data: sData(l[4]), hora: sHora(l[5]), feito: String(l[6]) === 'sim', eventId: String(l[7] || ''),
      criado: sTxt(l[8]), horaFim: sHora(l[9]), origem: String(l[10] || ''), feitoEm: sTxt(l[11]), _linha: i + 2
    };
    it.agenda = it.data ? (it.eventId ? 'enviado' : 'pendente') : '';
    r.push(it);
  });
  return r;
}
function linhaDe(it) {
  return [it.id, it.titulo, it.cat, it.tipo, it.data || '', it.hora || '', it.feito ? 'sim' : 'não', it.eventId || '', it.criado || '', it.horaFim || '', it.origem || '', it.feitoEm || ''];
}
function limpo(it) { var o = {}; Object.keys(it).forEach(function (k) { if (k !== '_linha') o[k] = it[k]; }); return o; }
function achar(sh, id) {
  var it = lerItens(sh).filter(function (x) { return x.id === id; })[0];
  if (!it) throw new Error('Item não encontrado.');
  return it;
}
function anexar(sh, it) {
  var l = sh.getLastRow() + 1;
  sh.getRange(l, 1, 1, CAB.length).setNumberFormat('@').setValues([linhaDe(it)]);
  it._linha = l;
}

/* ---------- Google Agenda ---------- */
function criarEvento(it) {
  var cal = CalendarApp.getDefaultCalendar();
  var titulo = (it.cat === 'geral' ? '' : '[' + (CATS[it.cat] || 'Geral') + '] ') + primeiraLinha(it.titulo);
  var resto = String(it.titulo).split('\n').slice(1).join('\n').trim();
  var opt = resto ? { description: resto } : {};
  var ev;
  if (!it.hora) {
    ev = cal.createAllDayEvent(titulo, dt(it.data), opt);
  } else {
    var ini = dt(it.data, it.hora);
    var fim = it.horaFim && it.horaFim > it.hora ? dt(it.data, it.horaFim) : new Date(ini.getTime() + 3600000);
    ev = cal.createEvent(titulo, ini, fim, opt);
  }
  return ev.getId();
}
function apagarEvento(id) {
  if (!id) return;
  var cal = CalendarApp.getDefaultCalendar();
  var ev = cal.getEventById(id) || cal.getEventById(id + '@google.com');
  if (ev) ev.deleteEvent();
}
function enviarEvento(sh, it) {
  try {
    it.eventId = criarEvento(it);
    sh.getRange(it._linha, 8).setValue(it.eventId);
    it.agenda = 'enviado';
    return null;
  } catch (e) {
    return String(e && e.message || e);
  }
}

/* ---------- operações ---------- */
const OPS = {
  list: function () {
    return { items: lerItens(aba()).map(limpo) };
  },
  add: function (b) {
    var sh = aba(), x = b.item || {};
    var it = {
      id: novoId(), titulo: String(x.titulo || '').trim(), cat: CATS[x.cat] ? x.cat : 'geral',
      tipo: x.tipo === 'data' ? 'data' : 'semprazo', data: '', hora: '', horaFim: '', feito: false, feitoEm: '',
      eventId: '', criado: new Date().toISOString(), origem: '', agenda: ''
    };
    if (!it.titulo) throw new Error('Texto vazio.');
    if (it.tipo === 'data') {
      it.data = sData(x.data); it.hora = sHora(x.hora);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(it.data)) throw new Error('Data inválida.');
      it.agenda = 'pendente';
    }
    anexar(sh, it);
    var erro = it.tipo === 'data' ? enviarEvento(sh, it) : null;
    return { item: limpo(it), agendaErro: erro };
  },
  toggle: function (b) {
    var sh = aba(), it = achar(sh, b.id);
    it.feito = !!b.feito; it.feitoEm = it.feito ? new Date().toISOString() : '';
    sh.getRange(it._linha, 7).setValue(it.feito ? 'sim' : 'não');
    sh.getRange(it._linha, 12).setValue(it.feitoEm);
    return { item: limpo(it) };
  },
  remove: function (b) {
    var sh = aba(), it = achar(sh, b.id), erro = null;
    try { apagarEvento(it.eventId); } catch (e) { erro = String(e && e.message || e); }
    sh.deleteRow(it._linha);
    return { ok: true, agendaErro: erro };
  },
  send: function (b) {
    var sh = aba(), it = achar(sh, b.id);
    if (!it.data) throw new Error('Item sem data.');
    var erro = it.eventId ? null : enviarEvento(sh, it);
    return { item: limpo(it), agendaErro: erro };
  },
  events: function (b) {
    var p = String(b.ym).split('-'), y = +p[0], m = +p[1];
    var evs = CalendarApp.getDefaultCalendar().getEvents(new Date(y, m - 1, 1), new Date(y, m, 1));
    return {
      events: evs.map(function (e) {
        var todo = e.isAllDayEvent();
        return {
          id: e.getId(), titulo: e.getTitle() || '(sem título)', todoDia: todo,
          ini: todo ? sData(e.getAllDayStartDate()) : sData(e.getStartTime()),
          fim: todo ? sData(e.getAllDayEndDate()) : sData(e.getEndTime()),
          hora: todo ? '' : sHora(e.getStartTime())
        };
      })
    };
  },
  qts: function () {
    var sh = aba(), itens = lerItens(sh), hj = hoje();
    var q = SpreadsheetApp.openById(QTS_ID).getSheetByName(QTS_ABA);
    if (!q) throw new Error('Aba do QTS não encontrada.');
    var fim = Math.min(QTS_LINHA_FIM, q.getLastRow());
    var rg = q.getRange(QTS_LINHA_INI, 3, fim - QTS_LINHA_INI + 1, 12);
    var V = rg.getDisplayValues(), W = rg.getFontWeights();
    var data = null, atual = null, provas = [];
    for (var i = 0; i < V.length; i++) {
      var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(V[i][0] || '');
      if (m) { data = m[3] + '-' + pad(+m[2]) + '-' + pad(+m[1]); atual = null; }
      var disc = String(V[i][11] || '').trim(), neg = W[i][11] === 'bold';
      if (data && disc && neg) {
        var ini = hh(V[i][9]), f = hh(V[i][10]);
        if (atual && atual.disc === disc && atual.linha === i - 1) { if (f) atual.fim = f; atual.linha = i; }
        else { atual = { data: data, disc: disc, ini: ini, fim: f, linha: i }; provas.push(atual); }
      }
    }
    provas = provas.filter(function (p) { return p.data >= hj; });
    provas.forEach(function (p) { p.id = 'qts-' + p.data + '-' + (p.ini ? p.ini.replace(':', '') : 'dia') + '-' + slug(p.disc); });
    var achados = {}, tenho = {};
    provas.forEach(function (p) { achados[p.id] = true; });
    itens.forEach(function (i) { tenho[i.id] = true; });
    var novas = provas.filter(function (p) { return !tenho[p.id]; });
    var fora = itens.filter(function (i) { return i.origem === 'qts' && !i.feito && i.data >= hj && !achados[i.id]; })
      .map(function (i) { return i.titulo + ' (' + i.data + ')'; });
    var res = { novas: novas.length, ok: 0, semAgenda: 0, erroAgenda: '', fora: fora };
    novas.forEach(function (p) {
      var it = {
        id: p.id, titulo: 'Prova ' + p.disc, cat: 'cfo', tipo: 'data', data: p.data, hora: p.ini, horaFim: p.fim,
        feito: false, feitoEm: '', eventId: '', criado: new Date().toISOString(), origem: 'qts'
      };
      anexar(sh, it); res.ok++;
      var erro = enviarEvento(sh, it);
      if (erro) { res.semAgenda++; res.erroAgenda = erro; }
    });
    return res;
  }
};
function hh(s) { var m = /^(\d{1,2}):(\d{2})/.exec(String(s || '')); return m ? pad(+m[1]) + ':' + m[2] : ''; }
