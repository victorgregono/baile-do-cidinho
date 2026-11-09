/**
 * ============================================================
 *  BAILE DO MC CIDINHO - Backend (Google Apps Script)
 *  Persistencia gratuita via Google Sheets (sem banco de dados)
 * ============================================================
 *
 *  COMO USAR (5 minutos, sem codar):
 *  1. Crie a planilha: https://sheets.new
 *  2. Menu > Extensoes > Apps Script
 *  3. Cole TODO este codigo em Code.gs e salve (Ctrl+S)
 *  4. Defina a senha do admin SEM colocar no codigo:
 *     Configuracoes do projeto (engrenagem) > Propriedades do script >
 *     Adicionar propriedade: nome = SENHA_ADM, valor = sua senha
 *     (e o ID_DA_PLANILHA, se usar)
 *  5. Clique em "Implantar" > "Nova implantacao" > tipo: "Web app"
 *     - Executar como: "Eu" (a propria conta)
 *     - Quem pode acessar: "Qualquer pessoa"
 *  6. Copie a URL /exec e cole na pagina HTML (CONFIG.appsScriptUrl)
 *
 *  As abas "Passageiros" e "Config" sao criadas sozinhas.
 * ============================================================
 */

/** Cole aqui o ID da planilha (aparece na URL: /spreadsheets/d/ID_AQUI/edit). */
var ID_DA_PLANILHA = '';

/** Cabecalho da aba Passageiros. */
var CABECALHO = ['nome', 'status', 'criado_em', 'id', 'pago_em'];

/* ------------------------------------------------------------------ */
/* Rotas HTTP                                                          */
/* ------------------------------------------------------------------ */

function doGet(e) {
  return responder({
    ok: true,
    acao: 'listar',
    pessoas: listarPessoas(),
    evento: pegarConfig('evento') || 'Baile do MC Cidinho',
    data: pegarConfig('data') || '05/12'
  });
}

function doPost(e) {
  var dados = lerJson(e);

  if (dados.acao === 'login') {
    return senhaValida(dados.senha)
      ? responder({ ok: true, acao: 'login' })
      : responder({ ok: false, erro: 'Senha incorreta.' });
  }

  // Acoes que alteram status exigem a senha do administrador.
  var precisaSenha = dados.acao === 'marcar_pago'
                  || dados.acao === 'mudar_status'
                  || dados.acao === 'remover'
                  || dados.acao === 'limpar';

  if (precisaSenha && !senhaValida(dados.senha)) {
    return responder({ ok: false, erro: 'Senha do administrador invalida.' });
  }

  if (dados.acao === 'inserir') {
    var nome = String(dados.nome || '').trim().slice(0, 80);
    if (!nome) {
      return responder({ ok: false, erro: 'Informe o nome.' });
    }
    var pessoa = inserirPessoa(nome);
    return responder({ ok: true, acao: 'inserir', pessoa: pessoa, pessoas: listarPessoas() });
  }

  if (dados.acao === 'marcar_pago' || dados.acao === 'mudar_status') {
    var novoStatus = dados.status === 'pago' ? 'pago' : 'pendente';
    var alvo = String(dados.id || '');
    var lista = listarPessoas();
    var encontrou = false;
    for (var i = 0; i < lista.length; i++) {
      if (lista[i].id === alvo) {
        lista[i].status = novoStatus;
        lista[i].pago_em = novoStatus === 'pago' ? new Date().toISOString() : '';
        encontrou = true;
        break;
      }
    }
    if (!encontrou) {
      return responder({ ok: false, erro: 'Pessoa nao encontrada.' });
    }
    gravarLista(lista);
    return responder({ ok: true, acao: 'mudar_status', pessoas: lista });
  }

  if (dados.acao === 'remover') {
    var removerId = String(dados.id || '');
    var restantes = listarPessoas().filter(function (p) { return p.id !== removerId; });
    gravarLista(restantes);
    return responder({ ok: true, acao: 'remover', pessoas: restantes });
  }

  if (dados.acao === 'limpar') {
    gravarLista([]);
    return responder({ ok: true, acao: 'limpar', pessoas: [] });
  }

  return responder({ ok: false, erro: 'Acao desconhecida.' });
}

/* ------------------------------------------------------------------ */
/* Regras de planilha                                                  */
/* ------------------------------------------------------------------ */

function planilha() {
  if (ID_DA_PLANILHA) {
    return SpreadsheetApp.openById(ID_DA_PLANILHA);
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Planilha nao encontrada. Defina ID_DA_PLANILHA no codigo.');
  }
  return ss;
}

function abaPassageiros() {
  var ss = planilha();
  var aba = ss.getSheetByName('Passageiros');
  if (!aba) {
    aba = ss.insertSheet('Passageiros');
    aba.appendRow(CABECALHO);
    aba.setFrozenRows(1);
  } else if (aba.getLastRow() === 0) {
    aba.appendRow(CABECALHO);
    aba.setFrozenRows(1);
  }
  return aba;
}

function abaConfig() {
  var ss = planilha();
  var aba = ss.getSheetByName('Config');
  if (!aba) {
    aba = ss.insertSheet('Config');
    aba.appendRow(['chave', 'valor']);
    aba.appendRow(['evento', 'Baile do MC Cidinho']);
    aba.appendRow(['data', '05/12']);
  }
  return aba;
}

/**
 * Compara a senha enviada com a propriedade SENHA_ADM do script.
 * Se a propriedade nao estiver definida, nenhuma senha e aceita.
 */
function senhaValida(senha) {
  var esperada = PropertiesService.getScriptProperties().getProperty('SENHA_ADM');
  return !!esperada && String(senha || '') === esperada;
}

function pegarConfig(chave) {
  var aba = abaConfig();
  var dados = aba.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === chave) return dados[i][1];
  }
  return '';
}

/* ------------------------------------------------------------------ */
/* Operacoes de dados                                                  */
/* ------------------------------------------------------------------ */

function listarPessoas() {
  var aba = abaPassageiros();
  var linhas = aba.getDataRange().getValues();
  var pessoas = [];
  for (var i = 1; i < linhas.length; i++) {
    var nome = String(linhas[i][0] || '').trim();
    if (!nome) continue;
    pessoas.push({
      id: String(linhas[i][3] || ('linha' + i)),
      nome: nome,
      status: String(linhas[i][1] || 'pendente'),
      criado_em: String(linhas[i][2] || ''),
      pago_em: String(linhas[i][4] || '')
    });
  }
  return pessoas;
}

function inserirPessoa(nome) {
  var aba = abaPassageiros();
  var agora = new Date().toISOString();
  var id = Utilities.getUuid();
  aba.appendRow([nome, 'pendente', agora, id, '']);
  return { id: id, nome: nome, status: 'pendente', criado_em: agora, pago_em: '' };
}

/** Regrava a lista inteira (usado por marcar_pago / remover / limpar). */
function gravarLista(pessoas) {
  var aba = abaPassageiros();
  aba.clearContents();
  aba.appendRow(CABECALHO);
  if (!pessoas.length) return;
  var linhas = pessoas.map(function (p) {
    return [
      p.nome,
      p.status || 'pendente',
      p.criado_em || '',
      p.id || Utilities.getUuid(),
      p.pago_em || ''
    ];
  });
  aba.getRange(2, 1, linhas.length, 5).setValues(linhas);
}

/* ------------------------------------------------------------------ */
/* Utilidades                                                          */
/* ------------------------------------------------------------------ */

function lerJson(e) {
  try {
    if (e && e.postData && e.postData.contents) {
      return JSON.parse(e.postData.contents);
    }
  } catch (err) {
    return {};
  }
  return {};
}

function responder(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
