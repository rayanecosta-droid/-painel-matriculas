/**
 * ============================================================================
 * PAINEL BI - CAMPANHA DE MATRÍCULAS 2026.2 - FACULDADE VIASAPIENS
 * Versão 2 - regras de negócio atualizadas em 20/07/2026
 * ============================================================================
 * Este script lê os dados diretamente da planilha "MATRÍCULAS 2026.2
 * (respostas)" e publica um JSON com todos os indicadores já calculados.
 *
 * REGRA DE OURO: este script NUNCA escreve na planilha. Ele só lê.
 * REGRA DE OURO 2: a coluna "Valor da matrícula" (coluna R) é IGNORADA por
 * completo, conforme instrução da Diretoria Comercial. Nenhum cálculo
 * financeiro em reais é feito neste painel.
 *
 * -------------------------------------------------------------------------
 * CONFIGURAÇÃO (ajuste só esta seção quando precisar mudar regras)
 * -------------------------------------------------------------------------
 */
const CONFIG = {
  SHEET_NAME: "Respostas ao formulário 1",

  META_TOTAL_GERAL: 200,
  META_PRESENCIAL: 75,
  META_POR_CONSULTOR: 50,

  CAMPANHA_INICIO: "2026-04-24",
  CAMPANHA_FIM: "2026-08-14",

  // Nenhum polo é excluído por enquanto (Ubajara e Parnaíba continuam
  // valendo normalmente nos indicadores).
  POLOS_EXCLUIDOS: [],

  // Consultores cujas matrículas não entram em nenhum indicador, gráfico,
  // ranking ou lista de consultores. Valécia Carvalho não faz parte da
  // equipe comercial, então ela e tudo que ela lançou saem do painel.
  CONSULTORES_EXCLUIDOS: ["valécia carvalho", "valecia carvalho", "valécia"],

  // Cursos que, independentemente do que estiver na coluna Turno, sempre
  // contam como bloco EAD/Semipresencial (nunca Presencial).
  CURSOS_SEMPRE_EAD: ["pedagogia"],
};

function doGet(e) {
  const payload = buildDashboardPayload();
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function buildDashboardPayload() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    return { erro: "Aba '" + CONFIG.SHEET_NAME + "' não encontrada na planilha." };
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const linhasBrutas = values.slice(1).filter(r => r[1] && String(r[1]).trim() !== "");

  const idx = mapHeaderIndexes(headers);
  const registrosTodos = [];
  const registros = [];
  const excluidos = [];

  linhasBrutas.forEach((r, i) => {
    const linhaPlanilha = i + 2;
    const registro = parseRegistro(r, idx, linhaPlanilha);
    registrosTodos.push(registro);
    if (registro.excluido) {
      excluidos.push({ linha: linhaPlanilha, aluno: registro.nome, motivoExclusao: registro.motivoExclusao });
    } else {
      registros.push(registro);
    }
  });

  return {
    geradoEm: new Date().toISOString(),
    fonte: CONFIG.SHEET_NAME,
    totalRegistrosLidos: registrosTodos.length,
    totalRegistrosExcluidos: excluidos.length,
    registrosExcluidos: excluidos,
    config: {
      metaTotalGeral: CONFIG.META_TOTAL_GERAL,
      metaPresencial: CONFIG.META_PRESENCIAL,
      metaPorConsultor: CONFIG.META_POR_CONSULTOR,
      campanhaInicio: CONFIG.CAMPANHA_INICIO,
      campanhaFim: CONFIG.CAMPANHA_FIM,
    },
    indicadores: calcularIndicadores(registros),
    presencial: calcularBloco(registros.filter(r => r.modalidadeGeral === "Presencial")),
    ead: calcularBloco(registros.filter(r => r.modalidadeGeral === "EAD/Semipresencial")),
    rankings: calcularRankings(registros),
    consultores: calcularConsultores(registros),
    insights: calcularInsights(registros),
    registros: registros,
  };
}

function mapHeaderIndexes(headers) {
  const find = (texto) => headers.findIndex(h => String(h).trim().toLowerCase().indexOf(texto.toLowerCase()) !== -1);
  return {
    carimbo: find("Carimbo"),
    nome: find("Nome completo"),
    genero: find("Gênero"),
    cidade: find("Cidade"),
    tipoEscola: find("Tipo de escola de origem"),
    polo: find("Pólo de ensino"),
    situacao: find("Situação"),
    turno: find("Turno"),
    curso: find("Curso"),
    formaPagamento: find("Forma de pagamento"),
    financiamento: find("Financiamento"),
    origem: find("Origem"),
    motivo: find("Motivo"),
    consultor: find("Consultor"),
  };
  // Nota: a coluna "Valor da matrícula" é ignorada de propósito e não é
  // mapeada aqui, conforme instrução da Diretoria Comercial.
}

/**
 * Interpreta uma linha da planilha, aplicando as regras de negócio
 * confirmadas:
 *
 *  - EXCLUSÃO DA VALÉCIA: todos os registros lançados pela consultora
 *    Valécia Carvalho são marcados como excluídos e não entram em nenhum
 *    cálculo, gráfico, ranking ou lista de consultores do dashboard,
 *    porque ela não faz parte da equipe comercial atual.
 *  - STATUS DE MATRÍCULA: "Financiamento"/"Forma de pagamento" contendo
 *    "minuta" => Minuta FIES (célula laranja na planilha). Caso contrário
 *    => Matrícula Paga (célula verde).
 *  - MODALIDADE: o curso de Pedagogia é SEMPRE Semipresencial/EAD,
 *    independentemente do que estiver na coluna Turno. Para os demais
 *    cursos, se a coluna Turno contiver "EAD" ou "semipresencial" =>
 *    bloco EAD/Semipresencial; qualquer outro valor (Manhã/Tarde/Noite)
 *    => bloco Presencial.
 *  - COLUNA "VALOR DA MATRÍCULA": totalmente ignorada, por instrução
 *    explícita. Não é lida, não aparece no JSON, não entra em nenhuma
 *    conta.
 */
function parseRegistro(r, idx, linhaPlanilha) {
  const g = (i) => (i >= 0 && r[i] !== undefined && r[i] !== null) ? String(r[i]).trim() : "";
  const normaliza = (s) => s.toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "");

  const nome = g(idx.nome);
  const polo = g(idx.polo);
  const consultor = g(idx.consultor) || "Não informado";
  const curso = g(idx.curso).replace(/^Graduação em\s*/i, "").trim();
  const turno = g(idx.turno);

  let excluido = false;
  let motivoExclusao = null;
  if (CONFIG.POLOS_EXCLUIDOS.indexOf(normaliza(polo)) !== -1) {
    excluido = true;
    motivoExclusao = "Polo excluído (" + polo + ")";
  } else if (CONFIG.CONSULTORES_EXCLUIDOS.indexOf(normaliza(consultor)) !== -1) {
    excluido = true;
    motivoExclusao = "Consultor excluído (" + consultor + ")";
  }

  const financiamentoTxt = g(idx.financiamento).toLowerCase();
  const formaPagamentoTxt = g(idx.formaPagamento).toLowerCase();
  const isMinuta = financiamentoTxt.indexOf("minuta") !== -1 || formaPagamentoTxt.indexOf("minuta") !== -1;
  const statusMatricula = isMinuta ? "Minuta FIES" : "Matrícula Paga";

  const isFIES = financiamentoTxt.indexOf("fies") !== -1;
  const isPravaler = financiamentoTxt.indexOf("pravaler") !== -1;
  const tipoFinanceiro = isFIES ? "FIES" : (isPravaler ? "Financiadora Externa" : "Particular");

  const cursoNorm = normaliza(curso);
  const turnoNorm = normaliza(turno);
  const forcaEAD = CONFIG.CURSOS_SEMPRE_EAD.some(c => cursoNorm.indexOf(c) !== -1);
  const modalidadeGeral = (forcaEAD || turnoNorm.indexOf("ead") !== -1 || turnoNorm.indexOf("semipresencial") !== -1)
    ? "EAD/Semipresencial"
    : "Presencial";

  const carimboTxt = g(idx.carimbo);
  const dataCarimbo = parseDataBR(carimboTxt);

  return {
    linha: linhaPlanilha,
    nome: nome,
    genero: g(idx.genero),
    cidade: g(idx.cidade),
    polo: polo,
    tipoEscola: g(idx.tipoEscola),
    situacao: g(idx.situacao),
    turno: turno,
    modalidadeGeral: modalidadeGeral,
    curso: curso,
    statusMatricula: statusMatricula,
    tipoFinanceiro: tipoFinanceiro,
    origem: g(idx.origem) || "Não informado",
    motivo: g(idx.motivo) || "Não informado",
    consultor: consultor,
    dataCarimbo: dataCarimbo ? dataCarimbo.toISOString() : null,
    diaSemana: dataCarimbo ? dataCarimbo.getDay() : null,
    excluido: excluido,
    motivoExclusao: motivoExclusao,
  };
}

function parseDataBR(txt) {
  if (!txt) return null;
  const m = txt.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (!m) return null;
  const dia = parseInt(m[1], 10), mes = parseInt(m[2], 10) - 1, ano = parseInt(m[3], 10);
  const data = new Date(ano, mes, dia);
  return isNaN(data.getTime()) ? null : data;
}

function calcularIndicadores(registros) {
  const total = registros.length;
  const pagas = registros.filter(r => r.statusMatricula === "Matrícula Paga").length;
  const minutas = registros.filter(r => r.statusMatricula === "Minuta FIES").length;
  const fies = registros.filter(r => r.tipoFinanceiro === "FIES").length;
  const particular = registros.filter(r => r.tipoFinanceiro === "Particular").length;
  const presencial = registros.filter(r => r.modalidadeGeral === "Presencial").length;
  const ead = registros.filter(r => r.modalidadeGeral === "EAD/Semipresencial").length;

  const inicio = new Date(CONFIG.CAMPANHA_INICIO);
  const hoje = new Date();
  const diasCorridos = Math.max(1, Math.round((hoje - inicio) / 86400000));
  const semanasCorridas = Math.max(1, diasCorridos / 7);

  const porCurso = agruparContagem(registros, "curso");
  const porModalidade = agruparContagem(registros, "modalidadeGeral");

  return {
    totalMatriculas: total,
    totalPagas: pagas,
    totalMinutas: minutas,
    totalFIES: fies,
    totalParticular: particular,
    conversaoPercentual: total > 0 ? round2((pagas / total) * 100) : 0,
    metaGeral: CONFIG.META_TOTAL_GERAL,
    percentualMeta: round2((total / CONFIG.META_TOTAL_GERAL) * 100),
    totalPresencial: presencial,
    totalEAD: ead,
    mediaDiaria: round2(total / diasCorridos),
    mediaSemanal: round2(total / semanasCorridas),
    diasCorridos: diasCorridos,
    cursoCampeao: porCurso[0] || null,
    consultorCampeao: agruparContagem(registros, "consultor")[0] || null,
    cidadeCampea: agruparContagem(registros, "cidade")[0] || null,
    participacaoPorCurso: porCurso.map(x => ({ nome: x.nome, total: x.total, percentual: round2((x.total / total) * 100) })),
    participacaoPorModalidade: porModalidade.map(x => ({ nome: x.nome, total: x.total, percentual: round2((x.total / total) * 100) })),
  };
}

function calcularBloco(registrosDoBloco) {
  const total = registrosDoBloco.length;
  const pagas = registrosDoBloco.filter(r => r.statusMatricula === "Matrícula Paga").length;
  const minutas = registrosDoBloco.filter(r => r.statusMatricula === "Minuta FIES").length;
  return {
    total: total,
    pagas: pagas,
    minutas: minutas,
    conversaoPercentual: total > 0 ? round2((pagas / total) * 100) : 0,
    porConsultor: agruparContagem(registrosDoBloco, "consultor"),
    porCurso: agruparContagem(registrosDoBloco, "curso"),
    porCidade: agruparContagem(registrosDoBloco, "cidade"),
  };
}

function calcularRankings(registros) {
  return {
    consultores: agruparContagem(registros, "consultor"),
    cursos: agruparContagem(registros, "curso"),
    cidades: agruparContagem(registros, "cidade"),
    origemEscolar: agruparContagem(registros, "tipoEscola"),
    tipoFinanceiro: agruparContagem(registros, "tipoFinanceiro"),
    canalDivulgacao: agruparContagem(registros, "origem"),
    motivoEscolha: agruparContagem(registros, "motivo"),
  };
}

function calcularConsultores(registros) {
  const nomes = [...new Set(registros.map(r => r.consultor).filter(Boolean))];
  const hoje = new Date();
  return nomes.map(nome => {
    const doConsultor = registros.filter(r => r.consultor === nome);
    const pagas = doConsultor.filter(r => r.statusMatricula === "Matrícula Paga").length;
    const minutas = doConsultor.filter(r => r.statusMatricula === "Minuta FIES").length;
    const fies = doConsultor.filter(r => r.tipoFinanceiro === "FIES").length;
    const particular = doConsultor.filter(r => r.tipoFinanceiro === "Particular").length;

    const datasValidas = doConsultor.map(r => r.dataCarimbo).filter(Boolean).map(d => new Date(d));
    const ultimaData = datasValidas.length > 0 ? new Date(Math.max.apply(null, datasValidas)) : null;
    const diasSemMatricula = ultimaData ? Math.floor((hoje - ultimaData) / 86400000) : null;

    return {
      nome: nome,
      totalMatriculas: doConsultor.length,
      pagas: pagas,
      minutas: minutas,
      fies: fies,
      particular: particular,
      conversaoPercentual: doConsultor.length > 0 ? round2((pagas / doConsultor.length) * 100) : 0,
      metaIndividual: CONFIG.META_POR_CONSULTOR,
      percentualMetaGeral: round2((doConsultor.length / CONFIG.META_POR_CONSULTOR) * 100),
      ultimaMatriculaEm: ultimaData ? ultimaData.toISOString() : null,
      diasSemMatricula: diasSemMatricula,
    };
  }).sort((a, b) => b.totalMatriculas - a.totalMatriculas);
}

function calcularInsights(registros) {
  const DIAS = ["Domingo", "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"];
  const porDiaSemana = {};
  registros.forEach(r => {
    if (r.diaSemana === null || r.diaSemana === undefined) return;
    porDiaSemana[r.diaSemana] = (porDiaSemana[r.diaSemana] || 0) + 1;
  });
  const diaSemanaOrdenado = Object.entries(porDiaSemana).sort((a, b) => b[1] - a[1]);
  const melhorDiaSemana = diaSemanaOrdenado.length > 0
    ? { dia: DIAS[diaSemanaOrdenado[0][0]], total: diaSemanaOrdenado[0][1] }
    : null;

  const consultores = calcularConsultores(registros);
  const consultorMelhorConversao = consultores.filter(c => c.totalMatriculas >= 3).sort((a, b) => b.conversaoPercentual - a.conversaoPercentual)[0] || null;
  const consultorMaisTempoParado = consultores.filter(c => c.diasSemMatricula !== null).sort((a, b) => b.diasSemMatricula - a.diasSemMatricula)[0] || null;

  const cidades = agruparContagem(registros, "cidade");
  const canais = agruparContagem(registros, "origem");
  const motivos = agruparContagem(registros, "motivo");
  const origemEscolar = agruparContagem(registros, "tipoEscola");
  const total = registros.length;
  const publica = origemEscolar.find(o => o.nome.toLowerCase().indexOf("pública") !== -1 || o.nome.toLowerCase().indexOf("publica") !== -1);
  const privada = origemEscolar.find(o => o.nome.toLowerCase().indexOf("privada") !== -1);

  const particular = registros.filter(r => r.tipoFinanceiro === "Particular").length;
  const fies = registros.filter(r => r.tipoFinanceiro === "FIES").length;

  const hoje = new Date();
  const seteDiasAtras = new Date(hoje - 7 * 86400000);
  const quatorzeDiasAtras = new Date(hoje - 14 * 86400000);
  const comData = registros.filter(r => r.dataCarimbo).map(r => Object.assign({}, r, { d: new Date(r.dataCarimbo) }));
  const ultimos7 = comData.filter(r => r.d >= seteDiasAtras).length;
  const anteriores7 = comData.filter(r => r.d >= quatorzeDiasAtras && r.d < seteDiasAtras).length;
  const tendencia = anteriores7 === 0
    ? (ultimos7 > 0 ? "crescimento" : "estável")
    : (ultimos7 > anteriores7 ? "crescimento" : (ultimos7 < anteriores7 ? "queda" : "estável"));

  return {
    melhorDiaSemana: melhorDiaSemana,
    consultorMelhorConversao: consultorMelhorConversao ? { nome: consultorMelhorConversao.nome, conversaoPercentual: consultorMelhorConversao.conversaoPercentual } : null,
    consultorMaisTempoSemMatricula: consultorMaisTempoParado ? { nome: consultorMaisTempoParado.nome, dias: consultorMaisTempoParado.diasSemMatricula } : null,
    cidadeQueMaisGeraAlunos: cidades[0] || null,
    canalQueMaisGeraMatriculas: canais[0] || null,
    principalMotivoDeEscolha: motivos[0] || null,
    percentualEscolaPublica: publica && total > 0 ? round2((publica.total / total) * 100) : null,
    percentualEscolaPrivada: privada && total > 0 ? round2((privada.total / total) * 100) : null,
    particularVsFIES: { particular: particular, fies: fies },
    tendenciaUltimos7Dias: { matriculasUltimos7Dias: ultimos7, matriculas7DiasAnteriores: anteriores7, tendencia: tendencia },
  };
}

function agruparContagem(registros, campo) {
  const mapa = {};
  registros.forEach(r => {
    const chave = r[campo] || "Não informado";
    mapa[chave] = (mapa[chave] || 0) + 1;
  });
  return Object.entries(mapa)
    .map(([nome, total]) => ({ nome, total }))
    .sort((a, b) => b.total - a.total);
}

function round2(n) {
  return Math.round(n * 100) / 100;
}

function criarGatilhoAtualizacao() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "atualizarCacheJSON") {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger("atualizarCacheJSON")
    .timeBased()
    .everyMinutes(5)
    .create();
}

function atualizarCacheJSON() {
  const payload = buildDashboardPayload();
  CacheService.getScriptCache().put("dashboard_payload", JSON.stringify(payload), 21600);
}
