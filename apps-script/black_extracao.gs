const CONFIG_BLACK = {
  rawSheetName: 'Raw_Data',
  summarySheetName: 'Sumario',
  summaryRange: 'E6:P8',
  years: [
    { year: 2021, rawRange: 'U3:V1000' },
    { year: 2022, rawRange: 'Q3:R1000' },
    { year: 2023, rawRange: 'M3:N1000' },
    { year: 2024, rawRange: 'I3:J1000' },
    { year: 2025, rawRange: 'E3:F1000' },
  ],
};

function compilarBlack2021a2025() {
  const resultados = [];

  CONFIG_BLACK.years.forEach(({ year, rawRange }) => {
    resultados.push(compilarAnoBlack_(year, rawRange));
  });

  criarConsolidadoBlack_(CONFIG_BLACK.years.map(({ year }) => year));

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Extração Black 2021-2025 concluída.',
    'Sucesso'
  );

  return resultados;
}

function compilarBlack2024e2025() {
  const resultados = [
    compilarBlack2024(),
    compilarBlack2025(),
  ];

  criarConsolidadoBlack2021a2025();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Extração Black 2024-2025 concluída.',
    'Sucesso'
  );

  return resultados;
}

function compilarBlack2024() {
  return compilarAnoPorConfigBlack_(2024);
}

function compilarBlack2025() {
  return compilarAnoPorConfigBlack_(2025);
}

function compilarBlack2025Parte1() {
  return compilarAnoParcialBlack_(2025, 'E3:F35', true, false);
}

function compilarBlack2025Parte2() {
  return compilarAnoParcialBlack_(2025, 'E36:F1000', false, true);
}

function criarConsolidadoBlack2021a2025() {
  return criarConsolidadoBlack_(CONFIG_BLACK.years.map(({ year }) => year));
}

function compilarAnoPorConfigBlack_(ano) {
  const configAno = CONFIG_BLACK.years.find((item) => item.year === ano);

  if (!configAno) {
    throw new Error(`Ano ${ano} não configurado em CONFIG_BLACK.years.`);
  }

  return compilarAnoBlack_(configAno.year, configAno.rawRange);
}

function compilarAnoBlack_(ano, intervaloRaw) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaRaw = ss.getSheetByName(CONFIG_BLACK.rawSheetName);

  if (!abaRaw) {
    throw new Error('Aba "Raw_Data" não encontrada.');
  }

  const dadosIniciais = abaRaw.getRange(intervaloRaw).getValues();
  const resultados = [
    ['Nick', 'Info', 'Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],
  ];

  let linksProcessados = 0;
  let erros = 0;

  dadosIniciais.forEach(([nick, url]) => {
    if (!url || url.toString().trim() === '') return;

    try {
      const ssJogador = SpreadsheetApp.openByUrl(url);
      const abaSumario = ssJogador.getSheetByName(CONFIG_BLACK.summarySheetName);

      if (!abaSumario) {
        resultados.push([nick, 'Erro: Aba "Sumario" não encontrada', '', '', '', '', '', '', '', '', '', '', '', '']);
        erros++;
        return;
      }

      const valores = abaSumario.getRange(CONFIG_BLACK.summaryRange).getValues();
      const profit = normalizarNumeros_(valores[0]);
      const mtt = normalizarNumeros_(valores[1]);
      const buyIn = normalizarNumeros_(valores[2]);
      const roi = profit.map((valor, i) => buyIn[i] ? valor / buyIn[i] : 0);
      const abi = buyIn.map((valor, i) => mtt[i] ? valor / mtt[i] : 0);

      resultados.push([nick, 'Profit', ...profit]);
      resultados.push(['', 'Mtt', ...mtt]);
      resultados.push(['', 'Buy in', ...buyIn]);
      resultados.push(['', 'ROI', ...roi]);
      resultados.push(['', 'ABI', ...abi]);

      linksProcessados++;
    } catch (erro) {
      resultados.push([nick, 'Erro de Acesso ou Link Quebrado', erro.message, '', '', '', '', '', '', '', '', '', '', '']);
      erros++;
    }
  });

  const abaDestino = prepararAba_(ss, String(ano));
  abaDestino.getRange(1, 1, resultados.length, 14).setValues(resultados);

  formatarAbaAno_(abaDestino, resultados.length);
  adicionarTotaisAno_(abaDestino, resultados.length);

  return { ano, linksProcessados, erros, linhas: resultados.length };
}

function compilarAnoParcialBlack_(ano, intervaloRaw, limparAba, finalizarTotais) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaRaw = ss.getSheetByName(CONFIG_BLACK.rawSheetName);

  if (!abaRaw) {
    throw new Error('Aba "Raw_Data" não encontrada.');
  }

  const dadosIniciais = abaRaw.getRange(intervaloRaw).getValues();
  const resultados = limparAba
    ? [['Nick', 'Info', 'Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']]
    : [];

  let linksProcessados = 0;
  let erros = 0;

  dadosIniciais.forEach(([nick, url]) => {
    if (!url || url.toString().trim() === '') return;

    try {
      const ssJogador = SpreadsheetApp.openByUrl(url);
      const abaSumario = ssJogador.getSheetByName(CONFIG_BLACK.summarySheetName);

      if (!abaSumario) {
        resultados.push([nick, 'Erro: Aba "Sumario" não encontrada', '', '', '', '', '', '', '', '', '', '', '', '']);
        erros++;
        return;
      }

      const valores = abaSumario.getRange(CONFIG_BLACK.summaryRange).getValues();
      const profit = normalizarNumeros_(valores[0]);
      const mtt = normalizarNumeros_(valores[1]);
      const buyIn = normalizarNumeros_(valores[2]);
      const roi = profit.map((valor, i) => buyIn[i] ? valor / buyIn[i] : 0);
      const abi = buyIn.map((valor, i) => mtt[i] ? valor / mtt[i] : 0);

      resultados.push([nick, 'Profit', ...profit]);
      resultados.push(['', 'Mtt', ...mtt]);
      resultados.push(['', 'Buy in', ...buyIn]);
      resultados.push(['', 'ROI', ...roi]);
      resultados.push(['', 'ABI', ...abi]);

      linksProcessados++;
    } catch (erro) {
      resultados.push([nick, 'Erro de Acesso ou Link Quebrado', erro.message, '', '', '', '', '', '', '', '', '', '', '']);
      erros++;
    }
  });

  let abaDestino = ss.getSheetByName(String(ano));

  if (!abaDestino) {
    abaDestino = ss.insertSheet(String(ano));
  }

  if (limparAba) {
    abaDestino.clear();
  } else {
    removerTotaisAnoSeExistirem_(abaDestino);
  }

  const linhaInicial = Math.max(abaDestino.getLastRow() + 1, 1);

  if (resultados.length > 0) {
    abaDestino.getRange(linhaInicial, 1, resultados.length, 14).setValues(resultados);
  }

  abaDestino.setFrozenRows(1);
  abaDestino.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#d9eaf7');

  if (finalizarTotais) {
    const ultimaLinhaDados = abaDestino.getLastRow();
    formatarLinhasMetricas_(abaDestino, 2, ultimaLinhaDados);
    adicionarTotaisAno_(abaDestino, ultimaLinhaDados);
    abaDestino.autoResizeColumns(1, 14);
  }

  return { ano, intervaloRaw, linksProcessados, erros, linhasAdicionadas: resultados.length };
}

function adicionarTotaisAno_(aba, linhasDados) {
  const linhaTotal = linhasDados + 2;
  const linhas = [
    ['Total', 'Profit'],
    ['Total', 'Mtt'],
    ['Total', 'Buy in'],
    ['Total', 'ROI'],
    ['Total', 'ABI'],
  ];

  aba.getRange(linhaTotal, 1, linhas.length, 2).setValues(linhas);

  for (let col = 3; col <= 14; col++) {
    const letra = colunaParaLetra_(col);
    aba.getRange(linhaTotal, col).setFormula(`=SUMIF($B$2:$B$${linhasDados};"Profit";${letra}$2:${letra}$${linhasDados})`);
    aba.getRange(linhaTotal + 1, col).setFormula(`=SUMIF($B$2:$B$${linhasDados};"Mtt";${letra}$2:${letra}$${linhasDados})`);
    aba.getRange(linhaTotal + 2, col).setFormula(`=SUMIF($B$2:$B$${linhasDados};"Buy in";${letra}$2:${letra}$${linhasDados})`);
    aba.getRange(linhaTotal + 3, col).setFormula(`=IFERROR(${letra}${linhaTotal}/${letra}${linhaTotal + 2};0)`);
    aba.getRange(linhaTotal + 4, col).setFormula(`=IFERROR(${letra}${linhaTotal + 2}/${letra}${linhaTotal + 1};0)`);
  }

  formatarLinhasMetricas_(aba, linhaTotal, linhaTotal + 4);
}

function removerTotaisAnoSeExistirem_(aba) {
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return;

  const inicioBusca = Math.max(1, ultimaLinha - 20);
  const valores = aba.getRange(inicioBusca, 1, ultimaLinha - inicioBusca + 1, 2).getValues();

  for (let i = 0; i < valores.length; i++) {
    const [colunaA, colunaB] = valores[i];
    if (colunaA === 'Total' && colunaB === 'Profit') {
      const linhaTotal = inicioBusca + i;
      aba.getRange(linhaTotal, 1, ultimaLinha - linhaTotal + 1, 14).clearContent();
      return;
    }
  }
}

function criarConsolidadoBlack_(anos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = prepararAba_(ss, 'Consolidado');
  const cabecalho = ['Ano', 'Info', 'Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
  const linhas = [cabecalho];

  anos.forEach((ano) => {
    const abaAno = ss.getSheetByName(String(ano));
    const ultimaLinha = abaAno.getLastRow();
    const valores = abaAno.getRange(ultimaLinha - 4, 2, 5, 13).getValues();

    valores.forEach((linha) => {
      linhas.push([ano, ...linha]);
    });

    linhas.push(new Array(14).fill(''));
  });

  aba.getRange(1, 1, linhas.length, 14).setValues(linhas);

  const linhaGrafico = 1;
  aba.getRange(linhaGrafico, 17, 1, 2).setValues([['Mtts Acumulado (Eixo X)', 'Profit Acumulado (Eixo Y)']]);

  let mttsAcumulado = 0;
  let profitAcumulado = 0;
  const pontosGrafico = [];

  anos.forEach((ano) => {
    const linhasAno = linhas.filter((linha) => linha[0] === ano);
    const linhaProfit = linhasAno.find((linha) => linha[1] === 'Profit');
    const linhaMtt = linhasAno.find((linha) => linha[1] === 'Mtt');

    if (!linhaProfit || !linhaMtt) return;

    for (let mes = 2; mes <= 13; mes++) {
      profitAcumulado += Number(linhaProfit[mes]) || 0;
      mttsAcumulado += Number(linhaMtt[mes]) || 0;
      pontosGrafico.push([mttsAcumulado, profitAcumulado]);
    }
  });

  if (pontosGrafico.length > 0) {
    aba.getRange(2, 17, pontosGrafico.length, 2).setValues(pontosGrafico);
  }
    });
  });

  aba.setFrozenRows(1);
  aba.getRange(1, 1, 1, 18).setFontWeight('bold').setBackground('#d9eaf7');
  aba.autoResizeColumns(1, 18);
}

function prepararAba_(ss, nome) {
  let aba = ss.getSheetByName(nome);

  if (!aba) {
    aba = ss.insertSheet(nome);
  } else {
    aba.clear();
  }

  return aba;
}

function formatarAbaAno_(aba, totalLinhas) {
  aba.setFrozenRows(1);
  aba.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#d9eaf7');
  formatarLinhasMetricas_(aba, 2, totalLinhas);
  aba.autoResizeColumns(1, 14);
}

function formatarLinhasMetricas_(aba, linhaInicial, linhaFinal) {
  for (let row = linhaInicial; row <= linhaFinal; row += 5) {
    const info = aba.getRange(row, 2, Math.min(5, linhaFinal - row + 1), 1).getValues().flat();

    info.forEach((nomeMetrica, offset) => {
      const linhaAtual = row + offset;

      if (nomeMetrica === 'Profit' || nomeMetrica === 'Buy in' || nomeMetrica === 'ABI') {
        aba.getRange(linhaAtual, 3, 1, 12).setNumberFormat('$#,##0.00;-$#,##0.00');
      }
      if (nomeMetrica === 'ROI') {
        aba.getRange(linhaAtual, 3, 1, 12).setNumberFormat('0.00%');
      }
      if (nomeMetrica === 'Mtt') {
        aba.getRange(linhaAtual, 3, 1, 12).setNumberFormat('0');
      }
    });
  }
}

function formatarLinhasMetricasAntigo_(aba, linhaInicial, linhaFinal) {
  for (let row = linhaInicial; row <= linhaFinal; row++) {
    const info = aba.getRange(row, 2).getValue();
    if (info === 'Profit' || info === 'Buy in' || info === 'ABI') {
      aba.getRange(row, 3, 1, 12).setNumberFormat('$#,##0.00;-$#,##0.00');
    }
    if (info === 'ROI') {
      aba.getRange(row, 3, 1, 12).setNumberFormat('0.00%');
    }
    if (info === 'Mtt') {
      aba.getRange(row, 3, 1, 12).setNumberFormat('0');
    }
  }
}

function normalizarNumeros_(valores) {
  return valores.map((valor) => {
    if (typeof valor === 'number') return valor;
    if (valor === null || valor === '') return 0;

    const texto = valor.toString().trim();
    if (!texto) return 0;

    const negativo = texto.includes('-');
    const limpo = texto
      .replace(/[^\d.,]/g, '')
      .replace(/,/g, '');

    const numero = Number(limpo);
    if (Number.isNaN(numero)) return 0;

    return negativo ? -numero : numero;
  });
}

function colunaParaLetra_(coluna) {
  let letra = '';

  while (coluna > 0) {
    const resto = (coluna - 1) % 26;
    letra = String.fromCharCode(65 + resto) + letra;
    coluna = Math.floor((coluna - 1) / 26);
  }

  return letra;
}
