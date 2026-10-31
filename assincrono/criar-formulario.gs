/**
 * Cria o Formulário Google do teste assíncrono do Cantinho e uma planilha ligada a ele.
 * Como usar:
 *  1. Entre em https://script.google.com com a conta Google da equipe e toque em "Novo projeto".
 *  2. Apague o que estiver no editor, cole este arquivo inteiro e salve.
 *  3. Toque em "Executar" (função criarFormulario) e autorize o acesso.
 *  4. Em "Registro de execução" aparece um bloco "window.__ccCfg = {...}". Copie e cole no config.js
 *     das pastas "familia" e "cuidadores" (ou mande para o Claude atualizar).
 */
function criarFormulario() {
  var CAMPOS = [["grupo", "Grupo"], ["participante", "Participante"], ["ordem", "Ordem das tarefas"], ["iniciais", "Perguntas iniciais"], ["t1", "Tarefa 1"], ["t2", "Tarefa 2"], ["t3", "Tarefa 3"], ["t4", "Tarefa 4"], ["t5", "Tarefa 5"], ["t6", "Tarefa 6"], ["t7", "Tarefa 7"], ["t8", "Tarefa 8"], ["t9", "Tarefa 9"], ["t10", "Tarefa 10"], ["t11", "Tarefa 11"], ["t12", "Tarefa 12"], ["t13", "Tarefa 13"], ["finais", "Perguntas finais"], ["json", "Dados completos (JSON)"]];
  var form = FormApp.create('Teste do Cantinho · respostas');
  form.setDescription('Recebe as respostas enviadas pelo protótipo. Não precisa preencher à mão.');
  form.setCollectEmail(false);
  form.setAcceptingResponses(true);
  try { form.setRequireLogin(false); } catch (e) {}
  var itens = CAMPOS.map(function (c) { return form.addParagraphTextItem().setTitle(c[1]); });
  try { form.setPublished(true); } catch (e) {}
  var planilha = SpreadsheetApp.create('Teste do Cantinho · respostas');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, planilha.getId());
  // O link pré-preenchido revela o código de cada campo (entry.NNN).
  var r = form.createResponse();
  itens.forEach(function (it, i) { r = r.withItemResponse(it.createResponse('CAMPO' + i)); });
  var url = r.toPrefilledUrl(), entries = {};
  CAMPOS.forEach(function (c, i) {
    var m = url.match(new RegExp('(entry\\.\\d+)=CAMPO' + i + '(&|$)'));
    if (m) entries[c[0]] = m[1];
  });
  var action = form.getPublishedUrl().replace(/\/viewform.*$/, '/formResponse');
  Logger.log('Formulário (editar): ' + form.getEditUrl());
  Logger.log('Planilha de respostas: ' + planilha.getUrl());
  Logger.log('Cole no config.js:\n\nwindow.__ccCfg = ' + JSON.stringify({ formAction: action, entries: entries, whatsapp: '' }, null, 2) + ';\n');
}
