/**
 * GASTOS CHILE — backend (Google Apps Script ligado a uma Planilha Google)
 *
 * Como usar:
 *  1. Abra a planilha "Gastos Chile" > Extensões > Apps Script e cole este código (substituindo tudo).
 *  2. Escolha a função "configurar" no topo e clique em Executar → autorize com sua conta.
 *  3. Implantar > Nova implantação > tipo "App da Web"
 *       Executar como: Eu  ·  Quem pode acessar: Qualquer pessoa  (NÃO "Qualquer pessoa com conta do Google")
 *  4. Copie o endereço que termina em /exec e cole no app.
 *  Sempre que mudar este código: Implantar > Gerenciar implantações > ✏️ > Versão: "Nova versão" > Implantar.
 */
const CODIGO = 'chile2026'; // <- TROQUE por um código só de vocês (os 6 digitam uma vez)
const PLANILHA_ID = '1TUVDtKdWA6EVSaovWMPIsOlT_gLmOphLMfWMg-BJC6A'; // planilha "Gastos Chile"
const VERSAO = 5;
const ABA = 'Lancamentos';
const ABA_CFG = 'Config';
const CAB = ['id', 'tipo', 'data', 'descricao', 'categoria', 'pagador', 'moeda', 'valor',
             'cotacao', 'valorBRL', 'dividir', 'por', 'atualizadoEm', 'deletado', 'pagos'];
const CFG_PADRAO = {
  casais: [
    { id: 'c1', nome: 'Marco e Nicole' },
    { id: 'c2', nome: 'Ana e João' },
    { id: 'c3', nome: 'Norton e Má' }
  ],
  cotacao: 5.8 // R$ por 1.000 pesos chilenos (ajustável no app)
};

function planilha_() {
  return PLANILHA_ID ? SpreadsheetApp.openById(PLANILHA_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

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
  try {
    const p = (e && e.parameter) || {};
    // o Google reserva o parâmetro "c" nas URLs /exec, por isso o código vem em "k"
    const cod = p.k;
    if (!cod) return json_({ ok: false, erro: 'codigo', info: 'Servidor Gastos Chile funcionando ✅ (versão ' + VERSAO + ')' });
    if (cod !== CODIGO) return json_({ ok: false, erro: 'codigo' });
    return json_({ ok: true, versao: VERSAO, config: lerConfig_(), linhas: lerLinhas_(), agora: new Date().toISOString() });
  } catch (err) {
    return json_({ ok: false, erro: 'servidor', detalhe: String(err && err.message || err) });
  }
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, erro: 'json' }); }
  if (b.c !== CODIGO) return json_({ ok: false, erro: 'codigo' });
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    if (b.acao === 'salvar') salvarLinha_(b.linha);
    else if (b.acao === 'apagar') {
      const alvo = lerLinhas_().filter(function (o) { return o.id === b.id; })[0];
      if (alvo) { alvo.deletado = true; salvarLinha_(alvo); }
    } else if (b.acao === 'config') gravarConfig_(b.config);
    else return json_({ ok: false, erro: 'acao' });
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, erro: 'servidor', detalhe: String(err && err.message || err) });
  } finally { try { lock.releaseLock(); } catch (e2) {} }
}

/**
 * RODE ESTA FUNÇÃO UMA VEZ no editor (selecione "configurar" e clique em Executar).
 * Ela pede a autorização do Google e cria as abas Lancamentos e Config na planilha.
 */
function configurar() {
  aba_(); lerConfig_();
  Logger.log('Tudo certo! Planilha: ' + planilha_().getName() + '. Agora implante (ou crie nova versão) como App da Web.');
}
