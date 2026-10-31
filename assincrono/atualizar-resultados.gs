/**
 * Cantinho · leva as respostas do teste assíncrono para as abas da planilha de resultados.
 *
 * Onde colar: na PLANILHA DE RESPOSTAS do formulário (a que o criar-formulario.gs criou),
 * em Extensões → Apps Script. Antes, importe para dentro dela a Resultados-teste-usabilidade.xlsx
 * (Arquivo → Importar → Fazer upload → "Inserir nova(s) página(s)").
 *
 * Depois de salvar, recarregue a planilha: aparece o menu "Cantinho".
 *  - Atualizar resultados ............ lê todas as respostas e preenche as abas.
 *  - Atualizar sozinho a cada envio .. cria um gatilho que roda a cada resposta nova.
 *
 * Pode rodar quantas vezes quiser: cada resposta mantém o mesmo ID e reescreve as mesmas linhas.
 * Só mexe nas linhas de IDs do assíncrono (FA1, FA2… e CUA1, CUA2…). As linhas dos testes conduzidos ficam como estão.
 */

var ABAS = {
  partF: 'Participantes família', partC: 'Participantes cuidadoras',
  tarF: 'Tarefas família', tarC: 'Tarefas cuidadoras', fin: 'Final'
};
var PRIMEIRA_LINHA = 5;

var TELAS = {
  'Today': 'Hoje', 'Today-Maria': 'Hoje da Vó Maria', 'PersonSwitch': 'Trocar pessoa', 'PersonSwitch-Maria': 'Trocar pessoa',
  'Routine': 'Rotina', 'Routine-RO': 'Rotina (só leitura)', 'AddChoice': 'Adicionar', 'CareSingle-Novo': 'Novo remédio',
  'CantDo': 'Não consegui', 'Log': 'Registrar sintoma', 'Report': 'Relatórios', 'Measures': 'Medições', 'Agenda': 'Agenda',
  'ShiftSwap': 'Troca de turno', 'Handoff': 'Passagem do cuidado', 'RecordDetail-Handoff': 'Detalhe do registro',
  'Settings': 'Ajustes', 'Settings-Caregiver': 'Ajustes', 'HandoffNow': 'Passar o cuidado', 'Network': 'Pessoas que cuidam',
  'Network-RO': 'Pessoas que cuidam', 'NetworkInvite': 'Convidar', 'EmergencyNow': 'SOS', 'Emergency': 'Ficha de emergência',
  'FirstShift': 'Primeira vez', 'CareGuide-RO': 'Como cuidar', 'CareDone': 'Cuidado feito', 'Welcome-Caricia': 'Boas-vindas',
  'InviteJoin-1': 'Convite', 'InviteJoin-2': 'Criar conta', 'InviteJoin-3': 'Pronto'
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Cantinho')
    .addItem('Atualizar resultados', 'atualizarResultados')
    .addItem('Atualizar sozinho a cada envio', 'ligarAutomatico')
    .addSeparator()
    .addItem('Refazer tudo (reescreve todas as respostas)', 'refazerTudo')
    .addToUi();
}

function ligarAutomatico() {
  var ss = SpreadsheetApp.getActive();
  var ja = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'atualizarResultados'; });
  if (!ja) ScriptApp.newTrigger('atualizarResultados').forSpreadsheet(ss).onFormSubmit().create();
  atualizarResultados();
  aviso('Pronto: a cada resposta nova, as abas de resultados se atualizam sozinhas.');
}

function aviso(msg) {
  try { SpreadsheetApp.getActive().toast(msg, 'Cantinho', 8); } catch (e) {}
}

function refazerTudo() { atualizarResultados(true); }

function atualizarResultados(tudo) {
  tudo = tudo === true;
  var ss = SpreadsheetApp.getActive();
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return;
  try {
    var faltando = Object.keys(ABAS).map(function (k) { return ABAS[k]; }).filter(function (n) { return !ss.getSheetByName(n); });
    if (faltando.length) {
      aviso('Faltam as abas: ' + faltando.join(', ') + '. Importe a planilha de resultados para dentro desta (veja o LEIA-ME).');
      return;
    }
    var respostas = lerRespostas(ss);
    var props = PropertiesService.getDocumentProperties();
    var ids = JSON.parse(props.getProperty('ids') || '{}');
    var cont = JSON.parse(props.getProperty('cont') || '{"familia":0,"cuidadores":0}');
    var feitosAntes = JSON.parse(props.getProperty('feitos') || '{}');
    var feitas = 0, erros = 0;
    respostas.forEach(function (r) {
      var d;
      try { d = JSON.parse(r.json); } catch (e) { erros++; return; }
      if (!d || !d.grupo) { erros++; return; }
      var g = d.grupo === 'cuidadores' ? 'cuidadores' : 'familia';
      var chave = r.quando + '|' + ((d.p || {}).nome || '') + '|' + (d.inicio || '');
      if (!ids[chave]) { cont[g] = (cont[g] || 0) + 1; ids[chave] = (g === 'familia' ? 'FA' : 'CUA') + cont[g]; }
      var id = ids[chave];
      if (!tudo && feitosAntes[chave]) return;
      escreverParticipante(ss, g, id, d, r.quando);
      escreverTarefas(ss, g, id, d);
      escreverFinal(ss, g, id, d);
      feitosAntes[chave] = 1;
      feitas++;
    });
    props.setProperty('feitos', JSON.stringify(feitosAntes));
    props.setProperty('ids', JSON.stringify(ids));
    props.setProperty('cont', JSON.stringify(cont));
    aviso((feitas ? feitas + ' resposta(s) nova(s) levada(s)' : 'Nenhuma resposta nova') + ' para as abas de resultados.' + (erros ? ' ' + erros + ' sem dados legíveis.' : ''));
  } finally {
    lock.releaseLock();
  }
}

/* A aba de respostas é a que tem a coluna "Dados completos (JSON)". */
function lerRespostas(ss) {
  var out = [];
  ss.getSheets().forEach(function (sh) {
    var lastCol = sh.getLastColumn(), lastRow = sh.getLastRow();
    if (lastCol < 1 || lastRow < 2) return;
    var head = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var cj = head.indexOf('Dados completos (JSON)');
    if (cj < 0) return;
    var ct = head.indexOf('Carimbo de data/hora'); if (ct < 0) ct = 0;
    sh.getRange(2, 1, lastRow - 1, lastCol).getValues().forEach(function (row) {
      if (!row[cj]) return;
      var t = row[ct];
      out.push({ json: String(row[cj]), quando: t instanceof Date ? t.toISOString() : String(t), data: t instanceof Date ? t : null });
    });
  });
  return out;
}

/* ---------- helpers ---------- */
function resp(list, id) {
  var x = (list || []).filter(function (q) { return q.id === id; })[0];
  return x && x.a ? String(x.a) : '';
}
function junta() {
  return Array.prototype.slice.call(arguments).filter(function (s) { return s; }).join(' · ');
}
function num(v) { var n = parseInt(v, 10); return isNaN(n) ? '' : n; }
function tela(s) { return TELAS[s] || (s && s.indexOf('Onboarding') === 0 ? 'Cadastro ' + s.slice(11) : s); }
function internet(wifi, dados) {
  var w = wifi === 'Sim', m = dados === 'Sim' || dados === 'Às vezes';
  if (!wifi && !dados) return '';
  return w && m ? 'Os dois' : w ? 'Wi-fi' : m ? 'Dados móveis' : 'Nenhum';
}

/* Linhas do ID; se ainda não existem, as primeiras vazias (coluna A e B em branco). */
function linhasPara(sh, id, quantas) {
  var last = Math.max(sh.getLastRow(), PRIMEIRA_LINHA);
  var col = sh.getRange(PRIMEIRA_LINHA, 1, last - PRIMEIRA_LINHA + 1, 2).getValues();
  var achadas = [];
  col.forEach(function (v, i) { if (String(v[0]) === id) achadas.push(PRIMEIRA_LINHA + i); });
  if (achadas.length >= quantas) return achadas.slice(0, quantas);
  var livres = [];
  col.forEach(function (v, i) { if (v[0] === '' && v[1] === '' && livres.length < quantas - achadas.length) livres.push(PRIMEIRA_LINHA + i); });
  var prox = last + 1;
  while (achadas.length + livres.length < quantas) livres.push(prox++);
  return achadas.concat(livres).sort(function (a, b) { return a - b; });
}

/* Escreve só as colunas pedidas (mapa coluna → valor), sem tocar nas fórmulas.
   Colunas vizinhas vão juntas numa só escrita. */
function escreverLinha(sh, linha, mapa) {
  var cols = Object.keys(mapa).map(function (c) { return c.charCodeAt(0) - 64; }).sort(function (a, b) { return a - b; });
  var i = 0;
  while (i < cols.length) {
    var j = i;
    while (j + 1 < cols.length && cols[j + 1] === cols[j] + 1) j++;
    var vals = [];
    for (var c = cols[i]; c <= cols[j]; c++) { var v = mapa[String.fromCharCode(64 + c)]; vals.push(v === undefined || v === null ? '' : v); }
    sh.getRange(linha, cols[i], 1, vals.length).setValues([vals]);
    i = j + 1;
  }
}

/* ---------- abas ---------- */
function escreverParticipante(ss, g, id, d, quando) {
  var ini = d.ini || [], p = d.p || {};
  var data = d.inicio ? new Date(d.inicio) : new Date(quando);
  var base = { A: id, C: p.nome || resp(ini, 'nome'), D: data, E: 'Assíncrono', F: d.consentimento ? 'Sim' : '', G: 'Só anotações', H: num(p.idade || resp(ini, 'idade')) };
  var sh, mapa;
  if (g === 'familia') {
    sh = ss.getSheetByName(ABAS.partF);
    mapa = Object.assign(base, {
      B: 'Família',
      I: junta(resp(ini, 'cuida'), resp(ini, 'quem')),
      J: resp(ini, 'rotina'),
      K: junta(resp(ini, 'celular'), resp(ini, 'oculos') ? 'óculos: ' + resp(ini, 'oculos') : ''),
      L: num(resp(ini, 'letra')),
      M: internet(resp(ini, 'wifi'), resp(ini, 'dados')),
      N: num(resp(ini, 'sinal')),
      O: junta(resp(ini, 'semNet'), resp(ini, 'comodos') ? 'Todos os cômodos: ' + resp(ini, 'comodos') : '', resp(ini, 'wifiCasa') ? 'Casa da pessoa: ' + resp(ini, 'wifiCasa') : ''),
      P: resp(ini, 'acabam'),
      Q: resp(ini, 'apps'),
      R: num(resp(ini, 'vontade')),
      S: resp(ini, 'combinam'),
      T: resp(ini, 'appSaude')
    });
  } else {
    sh = ss.getSheetByName(ABAS.partC);
    mapa = Object.assign(base, {
      B: 'Cuidadora',
      I: junta(resp(ini, 'trabalha'), resp(ini, 'tempo')),
      J: resp(ini, 'quantas'),
      K: resp(ini, 'qualCel'),
      L: internet(resp(ini, 'wifi'), resp(ini, 'dados')),
      M: num(resp(ini, 'sinal')),
      N: junta(resp(ini, 'semNet'), resp(ini, 'sinalRua') ? 'Sinal no caminho: ' + resp(ini, 'sinalRua') : '', resp(ini, 'wifiCasa') ? 'Casa onde trabalha: ' + resp(ini, 'wifiCasa') : ''),
      O: resp(ini, 'acabam'),
      P: junta(resp(ini, 'celular'), resp(ini, 'oculos') ? 'óculos: ' + resp(ini, 'oculos') : ''),
      Q: num(resp(ini, 'letra')),
      R: num(resp(ini, 'vontade')),
      S: resp(ini, 'orient'),
      T: resp(ini, 'passagem')
    });
  }
  escreverLinha(sh, linhasPara(sh, id, 1)[0], mapa);
}

function escreverTarefas(ss, g, id, d) {
  var sh = ss.getSheetByName(g === 'familia' ? ABAS.tarF : ABAS.tarC);
  var total = g === 'familia' ? 13 : 9, pre = g === 'familia' ? '' : 'C';
  var linhas = linhasPara(sh, id, total);
  for (var k = 1; k <= total; k++) {
    var t = (d.tasks || {})[pre + k];
    var mapa = { A: id, B: g === 'familia' ? k : 'C' + k };
    if (t) {
      var fora = (t.fora || []).map(function (x) { return tela(x[0]) + (x[1] ? ' › ' + x[1] : ''); });
      var cinza = (t.cinza || []).map(function (x) { return tela(x[0]) + ' › ' + x[1]; });
      var caminho = (t.caminho || []).map(function (x) { return tela(x[0]) + ' ' + x[1] + 's'; }).join(' → ');
      var erros = (t.erros || []).filter(function (x, i, a) { return a.indexOf(x) === i; });
      mapa.D = Math.floor((t.seg || 0) / 60);
      mapa.E = (t.seg || 0) % 60;
      mapa.G = t.ok ? 'Sozinha' : 'Não conseguiu';
      mapa.H = num(t.nota);
      mapa.I = junta(fora.length ? fora.length + ' toque(s) fora: ' + fora.join('; ') : '', cinza.length ? cinza.length + ' toque(s) em botão cinza: ' + cinza.join('; ') : '', 'Caminho: ' + caminho);
      mapa.K = erros.length ? erros.join('; ') + ' (' + (t.erros || []).length + 'x)' : '';
      mapa.N = t.comentario || '';
      mapa.O = num(t.pos);
    } else {
      mapa.D = ''; mapa.E = ''; mapa.G = ''; mapa.H = ''; mapa.I = 'Não chegou a fazer'; mapa.K = ''; mapa.N = ''; mapa.O = '';
    }
    escreverLinha(sh, linhas[k - 1], mapa);
  }
}

function escreverFinal(ss, g, id, d) {
  var sh = ss.getSheetByName(ABAS.fin), fin = d.fin || [];
  if (!fin.length) return;
  escreverLinha(sh, linhasPara(sh, id, 1)[0], {
    A: id, B: g === 'familia' ? 'Família' : 'Cuidadora',
    C: resp(fin, 'facil'), D: resp(fin, 'dificil'), E: resp(fin, 'palavra'),
    F: junta(resp(fin, 'usaria'), resp(fin, 'usariaPq')),
    G: junta(resp(fin, 'mais'), resp(fin, 'quantoPq') ? 'Sobre o valor: ' + resp(fin, 'quantoPq') : ''),
    H: num(resp(fin, 'nps')),
    I: resp(fin, 'offline'), J: resp(fin, 'offlinePq'),
    K: resp(fin, 'pagaria'), L: resp(fin, 'quanto')
  });
}
