// Recebe os votos da página dos cartazes e grava na planilha "Cartazes de Natal – Votos".
// Colado em Extensões > Apps Script da planilha e publicado como App da Web (acesso: Qualquer pessoa).
var ID = '10f2wO4Qu7adgEHaoHggopAU4BP86MQGJy6A160fy6AQ';
var FILMES = ['wonderful','miracle','white','story','home','jack','grinch','love','elf','polar','klaus','chronicles'];

function aba_() { return SpreadsheetApp.openById(ID).getSheets()[0]; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents);
    var nome = String(d.nome || '').trim().slice(0, 60);
    if (!nome) return json_({ ok: false, erro: 'nome' });
    var p = d.picks || {};
    var linha = [new Date(), nome];
    for (var i = 0; i < FILMES.length; i++) linha.push(String(p[FILMES[i]] || '').slice(0, 40));
    var sh = aba_();
    var vals = sh.getDataRange().getValues();
    var achou = -1;
    for (var r = 1; r < vals.length; r++) {
      if (String(vals[r][1]).trim().toLowerCase() === nome.toLowerCase()) achou = r + 1;
    }
    if (achou > 0) sh.getRange(achou, 1, 1, linha.length).setValues([linha]);
    else sh.appendRow(linha);
    return json_({ ok: true, atualizado: achou > 0 });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  var vals = aba_().getDataRange().getValues();
  var placar = {};
  for (var r = 1; r < vals.length; r++) {
    for (var i = 0; i < FILMES.length; i++) {
      var v = vals[r][i + 2];
      if (!v) continue;
      var k = FILMES[i] + '|' + v;
      placar[k] = (placar[k] || 0) + 1;
    }
  }
  return json_({ votantes: Math.max(vals.length - 1, 0), placar: placar });
}

// Rodar uma vez: cria a aba "Resultado" com a contagem automática.
function preparar() {
  var ss = SpreadsheetApp.openById(ID);
  var sh = ss.getSheetByName('Resultado') || ss.insertSheet('Resultado');
  sh.clear();
  var opcoes = {
    grinch: ['2000 · Original (EUA)', '2000 · Brasileira', '2018 · Original (EUA)', '2018 · Brasileira']
  };
  var votos = ss.getSheets()[0];
  var nomeVotos = "'" + votos.getName() + "'";
  var linhas = [['Filme', 'Opção', 'Votos']];
  var cabecalho = votos.getRange(1, 3, 1, FILMES.length).getValues()[0];
  for (var i = 0; i < FILMES.length; i++) {
    var col = String.fromCharCode(67 + i); // C, D, E...
    var ops = opcoes[FILMES[i]] || ['Original (EUA)', 'Versão brasileira'];
    for (var j = 0; j < ops.length; j++) {
      linhas.push([j === 0 ? cabecalho[i] : '', ops[j], '=COUNTIF(' + nomeVotos + '!' + col + '2:' + col + ',"' + ops[j] + '")']);
    }
  }
  sh.getRange(1, 1, linhas.length, 3).setValues(linhas);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 3);
  votos.setFrozenRows(1);
  votos.getRange(1, 1, 1, 2 + FILMES.length).setFontWeight('bold');
}
