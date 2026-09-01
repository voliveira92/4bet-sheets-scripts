const TURMAS_VERIFICACAO = {
  Moneyball: '1x8DNbFMPN45SeRiYzqjIpZDA-dbX51QjmU7z6Yzhv0c',
  Black: '1YIpNeJ-aje_K3gUn5Gyh2dd6b75UPXZ7ORxzd1MELbI',
  Diamond: '1vQT3ejealrq09kFU8M0CyR0RGyd-Jz1bO3HXJf_qcFg',
};

const ANOS_VERIFICACAO = ['2021', '2022', '2023', '2024', '2025'];
const MESES_VERIFICACAO = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
function criarVerificacoesTurmaAtiva() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nomeTurma = descobrirNomeTurma_(ss);
  const alertas = gerarAlertasTurma_(ss, nomeTurma);
  escreverAbaVerificacoes_(ss, alertas);
  ss.toast(`Verificacoes atualizadas: ${Math.max(alertas.length - 1, 0)} alertas encontrados.`, 'Sucesso');
}

function criarVerificacoesTodasTurmas() {
  Object.keys(TURMAS_VERIFICACAO).forEach((nomeTurma) => {
    const ss = SpreadsheetApp.openById(TURMAS_VERIFICACAO[nomeTurma]);
    const alertas = gerarAlertasTurma_(ss, nomeTurma);
    escreverAbaVerificacoes_(ss, alertas);
  });

  SpreadsheetApp.getActiveSpreadsheet().toast('Abas Verificacoes atualizadas nas 3 turmas.', 'Sucesso');
}

function criarVerificacaoDuplicadosEntreTurmas() {
  const registros = [];

  Object.keys(TURMAS_VERIFICACAO).forEach((nomeTurma) => {
    const ss = SpreadsheetApp.openById(TURMAS_VERIFICACAO[nomeTurma]);

    ANOS_VERIFICACAO.forEach((ano) => {
      const aba = ss.getSheetByName(ano);
      if (!aba) return;

      const dados = aba.getDataRange().getValues();
      for (let linha = 1; linha < dados.length; linha++) {
        const nick = dados[linha][0];
        const info = dados[linha][1];
        if (!nick || info !== 'Profit') continue;

        const assinatura = [
          dados[linha].slice(2, 14).join('|'),
          (dados[linha + 1] || []).slice(2, 14).join('|'),
          (dados[linha + 2] || []).slice(2, 14).join('|'),
        ].join('||');

        registros.push({
          chave: `${ano}|${normalizarNick_(nick)}`,
          ano,
          nick,
          turma: nomeTurma,
          linha: linha + 1,
          assinatura,
        });
      }
    });
  });

  const porChave = {};
  registros.forEach((registro) => {
    porChave[registro.chave] = porChave[registro.chave] || [];
    porChave[registro.chave].push(registro);
  });

  const linhas = [
    ['Status', 'Tipo', 'Ano', 'Nick', 'Turmas', 'Linhas', 'Detalhe'],
  ];

  Object.keys(porChave).forEach((chave) => {
    const grupo = porChave[chave];
    const turmas = [...new Set(grupo.map((item) => item.turma))];
    if (turmas.length < 2) return;

    const assinaturas = [...new Set(grupo.map((item) => item.assinatura))];
    linhas.push([
      '',
      'Mesmo nick em mais de uma turma',
      grupo[0].ano,
      grupo[0].nick,
      turmas.join(', '),
      grupo.map((item) => `${item.turma}!${item.ano}: linha ${item.linha}`).join(' | '),
      assinaturas.length === 1 ? 'Dados de Profit/Mtt/Buy in identicos' : 'Dados diferentes entre turmas',
    ]);
  });

  const ssDestino = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ssDestino.getSheetByName('Verificacoes_Duplicados');
  if (!aba) {
    aba = ssDestino.insertSheet('Verificacoes_Duplicados');
  } else {
    aba.clear();
  }

  aba.getRange(1, 1, linhas.length, linhas[0].length).setValues(linhas);
  formatarAbaVerificacoes_(aba, linhas[0].length);
  ssDestino.toast(`Duplicidades entre turmas: ${Math.max(linhas.length - 1, 0)} alertas.`, 'Sucesso');
}

function atualizarTiposRawDataTurmaAtiva() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaRaw = ss.getSheetByName('Raw_Data');
  if (!abaRaw) {
    SpreadsheetApp.getUi().alert('Aba "Raw_Data" nao encontrada.');
    return;
  }

  garantirColunasTipoRawData_(abaRaw);
  preencherTiposRawData_(abaRaw);
  ss.toast('Colunas Tipo atualizadas na Raw_Data.', 'Sucesso');
}

function atualizarTiposRawData2021() {
  atualizarTiposRawDataAno_('2021');
}

function atualizarTiposRawData2022() {
  atualizarTiposRawDataAno_('2022');
}

function atualizarTiposRawData2023() {
  atualizarTiposRawDataAno_('2023');
}

function atualizarTiposRawData2024() {
  atualizarTiposRawDataAno_('2024');
}

function atualizarTiposRawData2025() {
  atualizarTiposRawDataAno_('2025');
}

function atualizarTiposRawDataAno_(ano) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaRaw = ss.getSheetByName('Raw_Data');
  if (!abaRaw) {
    SpreadsheetApp.getUi().alert('Aba "Raw_Data" nao encontrada.');
    return;
  }

  garantirColunasTipoRawData_(abaRaw);
  preencherTiposRawDataAno_(abaRaw, ano);
  ss.toast(`Coluna Tipo atualizada para ${ano}.`, 'Sucesso');
}

function gerarAlertasTurma_(ss, nomeTurma) {
  const linhas = [
    ['Status', 'Tipo', 'Turma', 'Ano', 'Mes', 'Nick', 'Bloco', 'Fonte Raw_Data', 'Classificacao Fonte', 'Link Fonte', 'Profit', 'Mtt', 'Buy in', 'ROI', 'ABI', 'Detalhe'],
  ];
  const fontesPorAno = carregarFontesRawData_(ss);

  ANOS_VERIFICACAO.forEach((ano) => {
    const aba = ss.getSheetByName(ano);
    if (!aba) {
      linhas.push(['', 'Aba anual ausente', nomeTurma, ano, '', '', '', '', '', '', '', '', `Aba ${ano} nao encontrada.`]);
      return;
    }

    const dados = aba.getDataRange().getValues();
    let indiceBloco = 0;
    for (let linha = 1; linha < dados.length; linha++) {
      const nick = dados[linha][0];
      const info = dados[linha][1];
      if (!nick || info !== 'Profit') continue;
      indiceBloco++;

      const profit = dados[linha];
      const mtt = dados[linha + 1] || [];
      const buyIn = dados[linha + 2] || [];
      const roi = dados[linha + 3] || [];
      const abi = dados[linha + 4] || [];
      const bloco = linha + 1;
      const fonte = (fontesPorAno[ano] || [])[indiceBloco - 1] || {};

      for (let mesIndex = 0; mesIndex < MESES_VERIFICACAO.length; mesIndex++) {
        const col = mesIndex + 2;
        const valorProfit = numero_(profit[col]);
        const valorMtt = numero_(mtt[col]);
        const valorBuyIn = numero_(buyIn[col]);
        const valorRoi = numero_(roi[col]);
        const valorAbi = numero_(abi[col]);
        const mes = MESES_VERIFICACAO[mesIndex];

        if (valorBuyIn > 0 && valorRoi < -1) {
          linhas.push([
            '',
            'ROI abaixo de -100%',
            nomeTurma,
            ano,
            mes,
            nick,
            bloco,
            fonte.nome || '',
            fonte.classificacao || '',
            fonte.link || '',
            valorProfit,
            valorMtt,
            valorBuyIn,
            valorRoi,
            valorAbi,
            `ROI calculado como ${(valorRoi * 100).toFixed(1)}%. Validar escala de Profit/Buy in.`,
          ]);
        }

        if (valorProfit !== 0 && valorMtt === 0 && valorBuyIn === 0) {
          linhas.push([
            '',
            'Profit sem volume',
            nomeTurma,
            ano,
            mes,
            nick,
            bloco,
            fonte.nome || '',
            fonte.classificacao || '',
            fonte.link || '',
            valorProfit,
            valorMtt,
            valorBuyIn,
            valorRoi,
            valorAbi,
            detalheProfitSemVolume_(fonte),
          ]);
        }

        if (valorMtt !== 0 && valorMtt % 1 !== 0) {
          linhas.push([
            '',
            'Mtt fracionario',
            nomeTurma,
            ano,
            mes,
            nick,
            bloco,
            fonte.nome || '',
            fonte.classificacao || '',
            fonte.link || '',
            valorProfit,
            valorMtt,
            valorBuyIn,
            valorRoi,
            valorAbi,
            'Quantidade de torneios possui casas decimais.',
          ]);
        }

        if (valorMtt > 0 && valorBuyIn > 0 && valorBuyIn / valorMtt < 5) {
          linhas.push([
            '',
            'ABI abaixo de 5',
            nomeTurma,
            ano,
            mes,
            nick,
            bloco,
            fonte.nome || '',
            fonte.classificacao || '',
            fonte.link || '',
            valorProfit,
            valorMtt,
            valorBuyIn,
            valorRoi,
            valorAbi,
            `ABI calculado: ${(valorBuyIn / valorMtt).toFixed(2)}.`,
          ]);
        }
      }
    }
  });

  return linhas;
}

function carregarFontesRawData_(ss) {
  const abaRaw = ss.getSheetByName('Raw_Data');
  if (!abaRaw) return {};

  const ultimaLinha = abaRaw.getLastRow();
  if (ultimaLinha < 3) return {};

  const fontesPorAno = {};
  const layouts = localizarLayoutsRawData_(abaRaw);

  Object.keys(layouts).forEach((ano) => {
    const layout = layouts[ano];
    const largura = layout.colNickUnico - layout.colNick + 1;
    const valores = abaRaw.getRange(3, layout.colNick, ultimaLinha - 2, largura).getValues();
    fontesPorAno[ano] = valores
      .filter((linha) => linha[layout.colLink - layout.colNick] && String(linha[layout.colLink - layout.colNick]).trim() !== '')
      .map((linha) => {
        const nickRaw = linha[layout.colNick - layout.colNick] || '';
        const link = linha[layout.colLink - layout.colNick] || '';
        const tipo = layout.colTipo ? linha[layout.colTipo - layout.colNick] || '' : '';
        const nickUnico = linha[layout.colNickUnico - layout.colNick] || '';
        const nome = nickUnico || nickRaw;
        return {
          nome,
          link,
          classificacao: classificarTextoTipo_(tipo) !== 'Nao identificado'
            ? classificarTextoTipo_(tipo)
            : classificarFonte_(nickRaw, link, nickUnico),
        };
      });
  });

  return fontesPorAno;
}

function garantirColunasTipoRawData_(abaRaw) {
  let alterou = false;

  ANOS_VERIFICACAO.forEach((ano) => {
    let layout = localizarLayoutAnoRawData_(abaRaw, ano);
    if (!layout || layout.colTipo) return;

    abaRaw.insertColumnBefore(layout.colNickUnico);
    abaRaw.getRange(2, layout.colNickUnico).setValue('Tipo');
    alterou = true;
  });

  if (alterou) {
    SpreadsheetApp.flush();
  }
}

function preencherTiposRawData_(abaRaw) {
  const ultimaLinha = abaRaw.getLastRow();
  if (ultimaLinha < 3) return;

  const layouts = localizarLayoutsRawData_(abaRaw);

  Object.keys(layouts).forEach((ano) => {
    preencherTiposRawDataAno_(abaRaw, ano);
  });
}

function preencherTiposRawDataAno_(abaRaw, ano) {
  const ultimaLinha = abaRaw.getLastRow();
  if (ultimaLinha < 3) return;

  const layout = localizarLayoutAnoRawData_(abaRaw, ano);
  if (!layout || !layout.colTipo) return;

  const qtdLinhas = ultimaLinha - 2;
  const links = abaRaw.getRange(3, layout.colLink, qtdLinhas, 1).getValues();
  const tiposAtuais = abaRaw.getRange(3, layout.colTipo, qtdLinhas, 1).getValues();

  const tipos = links.map((linha, index) => {
    const url = linha[0];
    if (!url || String(url).trim() === '') return [''];

    const tipoAtual = tiposAtuais[index][0];
    if (
      tipoAtual &&
      String(tipoAtual).trim() !== '' &&
      String(tipoAtual).trim() !== 'Nao informado' &&
      String(tipoAtual).trim() !== 'Nao identificado' &&
      !String(tipoAtual).startsWith('Erro:')
    ) {
      return [tipoAtual];
    }

    return [buscarTipoPlanilha_(url)];
  });

  abaRaw.getRange(3, layout.colTipo, tipos.length, 1).setValues(tipos);
}

function buscarTipoPlanilha_(url) {
  try {
    const ssJogador = SpreadsheetApp.openByUrl(url);
    const abaSumario = ssJogador.getSheetByName('Sumario');
    if (!abaSumario) return 'Erro: Sumario nao encontrada';

    const textoTipo = lerTextoTipoSumario_(abaSumario);
    return classificarTextoTipo_(textoTipo);
  } catch (erro) {
    return `Erro: ${erro.message}`;
  }
}

function lerTextoTipoSumario_(abaSumario) {
  const candidatos = ['E2', 'F2', 'D2', 'G2', 'E1', 'F1', 'E3', 'F3'];
  let primeiroValorEncontrado = '';

  for (let i = 0; i < candidatos.length; i++) {
    const range = abaSumario.getRange(candidatos[i]);
    const valor = lerValorRangeComMescla_(range);
    if (!valor) continue;
    if (!primeiroValorEncontrado) primeiroValorEncontrado = valor;
    if (classificarTextoTipo_(valor) !== 'Nao identificado') return valor;
  }

  const valoresCabecalho = abaSumario.getRange('A1:H5').getDisplayValues();
  for (let linha = 0; linha < valoresCabecalho.length; linha++) {
    for (let col = 0; col < valoresCabecalho[linha].length; col++) {
      const valor = valoresCabecalho[linha][col];
      if (classificarTextoTipo_(valor) !== 'Nao identificado') return valor;
    }
  }

  return primeiroValorEncontrado;
}

function lerValorRangeComMescla_(range) {
  const valorDireto = String(range.getDisplayValue() || '').trim();
  if (valorDireto) return valorDireto;

  const mesclas = range.getMergedRanges();
  for (let i = 0; i < mesclas.length; i++) {
    const valorMescla = String(mesclas[i].getCell(1, 1).getDisplayValue() || '').trim();
    if (valorMescla) return valorMescla;
  }

  return '';
}

function classificarTextoTipo_(texto) {
  const normalizado = String(texto || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();

  if (normalizado.includes('congel')) return 'Congelada';
  if (normalizado.includes('pool')) return 'Pool';
  if (normalizado.includes('ativa') || normalizado.includes('ativo')) return 'Ativa';
  return 'Nao identificado';
}

function localizarLayoutsRawData_(abaRaw) {
  const layouts = {};

  ANOS_VERIFICACAO.forEach((ano) => {
    const layout = localizarLayoutAnoRawData_(abaRaw, ano);
    if (layout) layouts[ano] = layout;
  });

  return layouts;
}

function localizarLayoutAnoRawData_(abaRaw, ano) {
  const ultimaColuna = abaRaw.getLastColumn();
  const titulosAno = abaRaw.getRange(1, 1, 1, ultimaColuna).getDisplayValues()[0];
  const headers = abaRaw.getRange(2, 1, 1, ultimaColuna).getDisplayValues()[0];
  const colAnoIndex = titulosAno.findIndex((valor) => String(valor).trim() === `Ano ${ano}`);
  if (colAnoIndex === -1) return null;

  const colAno = colAnoIndex + 1;
  const proximoAnoIndexRelativo = titulosAno
    .slice(colAnoIndex + 1)
    .findIndex((valor) => /^Ano \d{4}$/.test(String(valor).trim()));
  const colFim = proximoAnoIndexRelativo === -1 ? ultimaColuna : colAnoIndex + proximoAnoIndexRelativo + 1;

  let colNick = null;
  let colLink = null;
  let colTipo = null;
  let colNickUnico = null;

  for (let col = colAno; col <= colFim; col++) {
    const header = String(headers[col - 1]).trim().toLowerCase();
    if (header === 'nick') colNick = col;
    if (header === 'link') colLink = col;
    if (header === 'tipo') colTipo = col;
    if (header === 'nicks únicos' || header === 'nicks unicos') colNickUnico = col;
  }

  if (!colNick || !colLink || !colNickUnico) return null;
  return { colNick, colLink, colTipo, colNickUnico };
}

function classificarFonte_(nickRaw, link, nickUnico) {
  const texto = `${nickRaw} ${link} ${nickUnico}`.toLowerCase();
  if (texto.includes('congel')) return 'Congelada';
  if (texto.includes('pool')) return 'Pool';
  return 'Nao identificado';
}

function detalheProfitSemVolume_(fonte) {
  if (fonte && fonte.classificacao === 'Congelada') {
    return 'Esperado para planilha Congelada: transferencia de profit para abater divida congelada.';
  }

  if (fonte && fonte.classificacao === 'Pool') {
    return 'Profit sem volume em fonte marcada como Pool. Validar regra especifica do pool.';
  }

  return 'Profit diferente de zero com Mtt e Buy in zerados. Pode ser planilha Congelada se a origem nao estiver identificada no Raw_Data.';
}

function escreverAbaVerificacoes_(ss, linhas) {
  let aba = ss.getSheetByName('Verificacoes');
  if (!aba) {
    aba = ss.insertSheet('Verificacoes');
  } else {
    aba.clear();
  }

  aba.getRange(1, 1, linhas.length, linhas[0].length).setValues(linhas);
  formatarAbaVerificacoes_(aba, linhas[0].length);
}

function formatarAbaVerificacoes_(aba, colunas) {
  aba.setFrozenRows(1);
  aba.getRange(1, 1, 1, colunas).setFontWeight('bold').setBackground('#d9ead3');
  const filtroAtual = aba.getFilter();
  if (filtroAtual) filtroAtual.remove();
  aba.getRange(1, 1, Math.max(aba.getLastRow(), 1), colunas).createFilter();
  aba.autoResizeColumns(1, colunas);
}

function descobrirNomeTurma_(ss) {
  const nome = ss.getName().toLowerCase();
  if (nome.includes('moneyball') || nome.includes('elite')) return 'Moneyball';
  if (nome.includes('black')) return 'Black';
  if (nome.includes('diamond')) return 'Diamond';
  return ss.getName();
}

function normalizarNick_(nick) {
  return String(nick).trim().toLowerCase();
}

function numero_(valor) {
  if (typeof valor === 'number') return valor;
  if (valor === null || valor === '') return 0;
  const convertido = Number(String(valor).replace(',', '.'));
  return Number.isFinite(convertido) ? convertido : 0;
}
