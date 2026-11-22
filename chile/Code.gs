/**
 * GASTOS CHILE — backend (Google Apps Script ligado a uma Planilha Google)
 *
 * Como usar: crie uma Planilha, abra Extensões > Apps Script, cole este código,
 * troque o CODIGO abaixo e implante como "App da Web"
 * (Executar como: Eu · Quem pode acessar: Qualquer pessoa).
 */
const CODIGO = 'chile2026'; // <- TROQUE por um código só de vocês (os 6 digitam uma vez)
const ABA = 'Lancamentos';
const ABA_CFG = 'Config';
const CAB = ['id', 'tipo', 'data', 'descricao', 'categoria', 'pagador', 'moeda', 'valor',
             'cotacao', 'valorBRL', 'dividir', 'por', 'atualizadoEm', 'deletado', 'pagos'];
const CFG_PADRAO = {
  casais: [
    { id: 'c1', nome: 'Marco e Nicole' },
    { id: 'c2', nome: 'Norton e Má' },
    { id: 'c3', nome: 'Ana e João' }
  ],
  cotacao: 5.8 // R$ por 1.000 pesos chilenos (ajustável no app)
};

function planilha_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function aba_() {
  const ss = planilha_();
  let sh = ss.getSheetByName(ABA);
  if (!sh) {
    sh = ss.insertSheet(ABA);
    sh.getRange(1, 1, 1, CAB.length).setValues([CAB]).setFontWeight('bold');
    sh.getRange('C:C').setNumberFormat('@'); // data como texto (aaaa-mm-dd)
    sh.setFrozenRows(0);
  }
  return sh;
}

function lerConfig_() {
  const ss = planilha_();
  let sh = ss.getSheetByName(ABA_CFG);
  if (!sh) {
    sh = ss.insertSheet(ABA_CFG);
    sh.getRange('A1').setValue(JSON.stringify(CFG_PADRAO));
  }
  try { return JSON.parse(sh.getRange('A1').getValue()); } catch (e) { return CFG_PADRAO; }
}

function gravarConfig_(cfg) {
  const ss = planilha_();
  let sh = ss.getSheetByName(ABA_CFG) || ss.insertSheet(ABA_CFG);
  sh.getRange('A1').setValue(JSON.stringify(cfg));
}

function lerLinhas_() {
  const sh = aba_();
  const n = sh.getLastRow();
  if (n < 2) return [];
  const tz = Session.getScriptTimeZone();
  return sh.getRange(2, 1, n - 1, CAB.length).getValues().map(function (r) {
    const o = {};
    CAB.forEach(function (c, i) {
      let v = r[i];
      if (v instanceof Date) v = Utilities.formatDate(v, tz, c === 'atualizadoEm' ? "yyyy-MM-dd'T'HH:mm:ss" : 'yyyy-MM-dd');
      o[c] = v;
    });
    o.deletado = o.deletado === true || o.deletado === 1 || o.deletado === '1' || o.deletado === 'TRUE';
    return o;
  }).filter(function (o) { return o.id; });
}

function salvarLinha_(l) {
  const sh = aba_();
  const agora = new Date().toISOString();
  const vals = CAB.map(function (c) {
    if (c === 'atualizadoEm') return agora;
    if (c === 'deletado') return l.deletado ? 1 : 0;
    const v = l[c];
    return v === undefined || v === null ? '' : v;
  });
  const n = sh.getLastRow();
  let linha = 0;
  if (n >= 2) {
    const ids = sh.getRange(2, 1, n - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(l.id)) { linha = i + 2; break; }
  }
  if (!linha) linha = n + 1;
  sh.getRange(linha, 3).setNumberFormat('@');
  sh.getRange(linha, 1, 1, CAB.length).setValues([vals]);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.c !== CODIGO) return json_({ ok: false, erro: 'codigo' });
  return json_({ ok: true, config: lerConfig_(), linhas: lerLinhas_(), agora: new Date().toISOString() });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, erro: 'json' }); }
  if (b.c !== CODIGO) return json_({ ok: false, erro: 'codigo' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (b.acao === 'salvar') salvarLinha_(b.linha);
    else if (b.acao === 'apagar') {
      const alvo = lerLinhas_().filter(function (o) { return o.id === b.id; })[0];
      if (alvo) { alvo.deletado = true; salvarLinha_(alvo); }
    } else if (b.acao === 'config') gravarConfig_(b.config);
    else return json_({ ok: false, erro: 'acao' });
    return json_({ ok: true });
  } finally { lock.releaseLock(); }
}
