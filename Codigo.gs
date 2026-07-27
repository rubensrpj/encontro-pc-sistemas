/**
 * Encontro PC Sistemas — backend de confirmação de presença
 * Google Apps Script vinculado a uma Planilha Google.
 *
 * Abas geradas automaticamente:
 *   Registros — uma linha por telefone (fonte da verdade, chave única)
 *   Pessoas   — uma linha por pessoa (titular + acompanhantes), para exportar
 *   Resumo    — contagens por combo e arrecadação prevista
 *
 * Implantar: Implantar > Nova implantação > App da Web
 *   Executar como: Eu
 *   Quem tem acesso: Qualquer pessoa
 * Copie a URL /exec e cole em API_URL no index.html.
 */

/* ----------------------------- configuração ----------------------------- */

const TOKEN = 'PCS-2026-TROQUE-ISTO';   // troque. Vai aparecer no código da página.

const ABA_REGISTROS = 'Registros';
const ABA_PESSOAS = 'Pessoas';
const ABA_RESUMO = 'Resumo';

const COMBOS = {
  feijoada:   { rotulo: 'Feijoada',                    valor: 50  },
  caipirinha: { rotulo: 'Feijoada + caipirinha',       valor: 80  },
  cerveja:    { rotulo: 'Feijoada + cerveja',          valor: 100 },
  crianca:    { rotulo: 'Criança — valor a combinar',  valor: 0   }
};

const CAB_REGISTROS = [
  'Telefone', 'Nome', 'Combo', 'Valor', 'Dependentes (JSON)',
  'Qtd. pessoas', 'Total (R$)', 'Recado', 'Atualizado em'
];
const CAB_PESSOAS = [
  'Telefone', 'Titular', 'Pessoa', 'Vínculo', 'Combo', 'Valor (R$)', 'Recado', 'Atualizado em'
];

/* ------------------------------- entrada ------------------------------- */

function doPost(e) { return processar_(e); }
function doGet(e)  { return processar_(e); }

function processar_(e) {
  let dados = {};
  try {
    if (e && e.postData && e.postData.contents) dados = JSON.parse(e.postData.contents);
    else if (e && e.parameter) dados = e.parameter;
  } catch (err) {
    return json_({ ok: false, erro: 'Payload inválido.' });
  }

  if (dados.token !== TOKEN) return json_({ ok: false, erro: 'Token inválido.' });

  const trava = LockService.getScriptLock();
  try {
    trava.waitLock(25000);
  } catch (err) {
    return json_({ ok: false, erro: 'Servidor ocupado. Tente de novo em instantes.' });
  }

  try {
    switch (dados.acao) {
      case 'listar':    return json_({ ok: true, mural: mural_() });
      case 'consultar': return json_({ ok: true, registro: buscar_(dados.telefone) });
      case 'confirmar': return json_(confirmar_(dados));
      case 'cancelar':  return json_(cancelar_(dados.telefone));
      default:          return json_({ ok: false, erro: 'Ação desconhecida.' });
    }
  } catch (err) {
    return json_({ ok: false, erro: String(err && err.message ? err.message : err) });
  } finally {
    trava.releaseLock();
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------ utilidades ------------------------------ */

function soDigitos_(v) { return String(v == null ? '' : v).replace(/\D/g, ''); }

function comboDe_(chave) { return COMBOS[chave] || COMBOS.feijoada; }

function aba_(nome, cabecalho) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(nome);
  if (!aba) aba = ss.insertSheet(nome);
  if (aba.getLastRow() === 0) {
    aba.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]);
    aba.getRange(1, 1, 1, cabecalho.length).setFontWeight('bold').setBackground('#123A7A').setFontColor('#FFFFFF');
    aba.setFrozenRows(1);
  }
  return aba;
}

function lerRegistros_() {
  const aba = aba_(ABA_REGISTROS, CAB_REGISTROS);
  const ultima = aba.getLastRow();
  if (ultima < 2) return [];
  const valores = aba.getRange(2, 1, ultima - 1, CAB_REGISTROS.length).getValues();
  return valores
    .filter(function (l) { return soDigitos_(l[0]); })
    .map(function (l) {
      let deps = [];
      try { deps = JSON.parse(l[4] || '[]'); } catch (e) { deps = []; }
      return {
        telefone: soDigitos_(l[0]),
        nome: String(l[1] || ''),
        combo: String(l[2] || 'feijoada'),
        dependentes: deps,
        recado: String(l[7] || ''),
        atualizadoEm: l[8] instanceof Date ? l[8] : new Date()
      };
    });
}

function totalRegistro_(r) {
  return (r.dependentes || []).reduce(function (s, d) {
    return s + comboDe_(d.combo).valor;
  }, comboDe_(r.combo).valor);
}

/* -------------------------------- ações -------------------------------- */

function buscar_(telefone) {
  const chave = soDigitos_(telefone);
  if (!chave) return null;
  const achado = lerRegistros_().filter(function (r) { return r.telefone === chave; })[0];
  return achado || null;
}

/** Mural público: sem telefone, sem recado. */
function mural_() {
  return lerRegistros_().map(function (r) {
    return {
      nome: r.nome,
      pessoas: 1 + (r.dependentes || []).length,
      acompanhantes: (r.dependentes || []).map(function (d) {
        return String(d.nome || '').split(/\s+/)[0];
      }),
      par: parseInt(r.telefone.slice(-1), 10) % 2
    };
  }).sort(function (a, b) { return a.nome.localeCompare(b.nome, 'pt-BR'); });
}

function confirmar_(dados) {
  const chave = soDigitos_(dados.telefone);
  if (chave.length !== 10 && chave.length !== 11) return { ok: false, erro: 'Telefone inválido.' };
  const nome = String(dados.nome || '').trim();
  if (nome.length < 3) return { ok: false, erro: 'Nome muito curto.' };

  const deps = (dados.dependentes || [])
    .filter(function (d) { return String(d.nome || '').trim(); })
    .map(function (d) {
      return {
        nome: String(d.nome).trim(),
        tipo: d.tipo === 'filho' ? 'filho' : 'companheiro',
        combo: COMBOS[d.combo] ? d.combo : 'feijoada'
      };
    });

  const registro = {
    telefone: chave,
    nome: nome,
    combo: COMBOS[dados.combo] ? dados.combo : 'feijoada',
    dependentes: deps,
    recado: String(dados.recado || '').trim(),
    atualizadoEm: new Date()
  };

  const aba = aba_(ABA_REGISTROS, CAB_REGISTROS);
  const linha = [
    "'" + registro.telefone,
    registro.nome,
    comboDe_(registro.combo).rotulo,
    comboDe_(registro.combo).valor,
    JSON.stringify(registro.dependentes),
    1 + registro.dependentes.length,
    totalRegistro_(registro),
    registro.recado,
    registro.atualizadoEm
  ];

  const idx = indiceDaLinha_(aba, registro.telefone);
  const atualizacao = idx > 0;
  if (atualizacao) {
    aba.getRange(idx, 1, 1, linha.length).setValues([linha]);
  } else {
    aba.appendRow(linha);
  }

  reconstruirDerivadas_();
  return { ok: true, atualizacao: atualizacao, mural: mural_(), total: totalRegistro_(registro) };
}

function cancelar_(telefone) {
  const chave = soDigitos_(telefone);
  const aba = aba_(ABA_REGISTROS, CAB_REGISTROS);
  const idx = indiceDaLinha_(aba, chave);
  if (idx > 0) aba.deleteRow(idx);
  reconstruirDerivadas_();
  return { ok: true, removido: idx > 0, mural: mural_() };
}

function indiceDaLinha_(aba, telefone) {
  const ultima = aba.getLastRow();
  if (ultima < 2) return -1;
  const col = aba.getRange(2, 1, ultima - 1, 1).getValues();
  for (let i = 0; i < col.length; i++) {
    if (soDigitos_(col[i][0]) === telefone) return i + 2;
  }
  return -1;
}

/* --------------------------- abas derivadas --------------------------- */

function reconstruirDerivadas_() {
  const registros = lerRegistros_();

  // Pessoas
  const pessoas = aba_(ABA_PESSOAS, CAB_PESSOAS);
  if (pessoas.getLastRow() > 1) {
    pessoas.getRange(2, 1, pessoas.getLastRow() - 1, CAB_PESSOAS.length).clearContent();
  }
  const linhas = [];
  registros.forEach(function (r) {
    const c = comboDe_(r.combo);
    linhas.push(["'" + r.telefone, r.nome, r.nome, 'Titular (ex-funcionário)', c.rotulo, c.valor, r.recado, r.atualizadoEm]);
    (r.dependentes || []).forEach(function (d) {
      const cd = comboDe_(d.combo);
      linhas.push(["'" + r.telefone, r.nome, d.nome, d.tipo === 'filho' ? 'Filho(a)' : 'Companheiro(a)', cd.rotulo, cd.valor, '', r.atualizadoEm]);
    });
  });
  if (linhas.length) pessoas.getRange(2, 1, linhas.length, CAB_PESSOAS.length).setValues(linhas);

  // Resumo
  const contagem = { feijoada: 0, caipirinha: 0, cerveja: 0, crianca: 0 };
  let totalPessoas = 0, totalValor = 0;
  registros.forEach(function (r) {
    [{ combo: r.combo }].concat(r.dependentes || []).forEach(function (p) {
      const k = COMBOS[p.combo] ? p.combo : 'feijoada';
      contagem[k]++; totalPessoas++; totalValor += COMBOS[k].valor;
    });
  });

  const resumo = aba_(ABA_RESUMO, ['Indicador', 'Valor']);
  if (resumo.getLastRow() > 1) resumo.getRange(2, 1, resumo.getLastRow() - 1, 2).clearContent();
  const dadosResumo = [
    ['Confirmações (telefones únicos)', registros.length],
    ['Total de pessoas', totalPessoas],
    ['Feijoada (R$ 50)', contagem.feijoada],
    ['Feijoada + caipirinha (R$ 80)', contagem.caipirinha],
    ['Feijoada + cerveja (R$ 100)', contagem.cerveja],
    ['Crianças (valor a combinar)', contagem.crianca],
    ['Arrecadação prevista (R$)', totalValor],
    ['Atualizado em', new Date()]
  ];
  resumo.getRange(2, 1, dadosResumo.length, 2).setValues(dadosResumo);
}

/* ------------- utilitário manual: rodar uma vez para criar abas ------------- */

function prepararPlanilha() {
  aba_(ABA_REGISTROS, CAB_REGISTROS);
  aba_(ABA_PESSOAS, CAB_PESSOAS);
  aba_(ABA_RESUMO, ['Indicador', 'Valor']);
  reconstruirDerivadas_();
}
