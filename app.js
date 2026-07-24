// ============================================================================
// PAINEL DE MATRÍCULAS 2026.2 - lógica principal (v2)
// ============================================================================

let DADOS = null;
let ABA_ATIVA = "Presencial"; // "Presencial" | "EAD/Semipresencial"
let GRAFICOS = {};

const el = (id) => document.getElementById(id);

async function iniciar() {
  await carregarDados();
  setInterval(carregarDados, DASHBOARD_CONFIG.REFRESH_INTERVAL_MS);
}

async function carregarDados() {
  if (!DASHBOARD_CONFIG.API_URL || DASHBOARD_CONFIG.API_URL.indexOf("COLE_AQUI") !== -1) {
    mostrarEstadoConfiguracaoPendente();
    return;
  }
  try {
    const resp = await fetch(DASHBOARD_CONFIG.API_URL, { cache: "no-store" });
    const json = await resp.json();
    if (json.erro) {
      mostrarEstadoErro(json.erro);
      return;
    }
    DADOS = json;
    montarFiltros(DADOS.registros);
    renderizarTudo();
    el("status-atualizacao").textContent = "Atualizado às " + new Date(DADOS.geradoEm).toLocaleTimeString("pt-BR");
  } catch (err) {
    mostrarEstadoErro("Não foi possível conectar à API do Apps Script. Detalhe: " + err.message);
  }
}

function mostrarEstadoConfiguracaoPendente() {
  el("app").innerHTML = `
    <div class="estado-central">
      <h2>Conecte o painel à planilha</h2>
      <p>Falta apenas colar a URL do Apps Script publicado no arquivo <b>config.js</b>.</p>
      <code>DASHBOARD_CONFIG.API_URL = "https://script.google.com/macros/s/SEU_ID/exec"</code>
    </div>`;
}

function mostrarEstadoErro(msg) {
  el("app").innerHTML = `
    <div class="estado-central">
      <h2>Não foi possível carregar os dados</h2>
      <p>${msg}</p>
    </div>`;
}

// ---------------------------------------------------------------------------
// Filtros
// ---------------------------------------------------------------------------
function montarFiltros(registros) {
  preencherSelect("filtro-consultor", unicos(registros, "consultor"));
  preencherSelect("filtro-curso", unicos(registros, "curso"));
  preencherSelect("filtro-cidade", unicos(registros, "cidade"));
  preencherSelect("filtro-origemEscolar", unicos(registros, "tipoEscola"));
  preencherSelect("filtro-financeiro", unicos(registros, "tipoFinanceiro"));
  preencherSelect("filtro-status", unicos(registros, "statusMatricula"));
  preencherSelect("filtro-genero", unicos(registros, "genero"));
  preencherSelect("filtro-canal", unicos(registros, "origem"));
  preencherSelect("filtro-motivo", unicos(registros, "motivo"));
}

function unicos(registros, campo) {
  return [...new Set(registros.map(r => r[campo]).filter(Boolean))].sort();
}

function preencherSelect(id, valores) {
  const select = el(id);
  const atual = select.value;
  select.innerHTML = '<option value="">Todos</option>' + valores.map(v => `<option value="${v}">${v}</option>`).join("");
  if (valores.includes(atual)) select.value = atual;
}

function registrosFiltrados() {
  if (!DADOS) return [];
  const consultor = el("filtro-consultor").value;
  const curso = el("filtro-curso").value;
  const cidade = el("filtro-cidade").value;
  const origemEscolar = el("filtro-origemEscolar").value;
  const financeiro = el("filtro-financeiro").value;
  const status = el("filtro-status").value;
  const genero = el("filtro-genero").value;
  const canal = el("filtro-canal").value;
  const motivo = el("filtro-motivo").value;

  return DADOS.registros.filter(r => r.modalidadeGeral === ABA_ATIVA).filter(r =>
    (!consultor || r.consultor === consultor) &&
    (!curso || r.curso === curso) &&
    (!cidade || r.cidade === cidade) &&
    (!origemEscolar || r.tipoEscola === origemEscolar) &&
    (!financeiro || r.tipoFinanceiro === financeiro) &&
    (!status || r.statusMatricula === status) &&
    (!genero || r.genero === genero) &&
    (!canal || r.origem === canal) &&
    (!motivo || r.motivo === motivo)
  );
}

// ---------------------------------------------------------------------------
// Renderização
// ---------------------------------------------------------------------------
function trocarAba(aba) {
  ABA_ATIVA = aba;
  el("tab-presencial").classList.toggle("ativa", aba === "Presencial");
  el("tab-ead").classList.toggle("ativa", aba === "EAD/Semipresencial");
  renderizarTudo();
}

function renderizarTudo() {
  if (!DADOS) return;
  const regs = registrosFiltrados();
  renderizarCards(regs);
  renderizarRanking("ranking-consultores", agrupar(regs, "consultor"));
  renderizarRanking("ranking-cursos", agrupar(regs, "curso"));
  renderizarRanking("ranking-cidades", agrupar(regs, "cidade"));
  renderizarRanking("ranking-origem", agrupar(regs, "tipoEscola"));
  renderizarRanking("ranking-canal", agrupar(regs, "origem"));
  renderizarRanking("ranking-motivo", agrupar(regs, "motivo"));
  renderizarConsultores(regs);
  renderizarInsights();
  renderizarExclusoes();
  try { renderizarFunil(regs); } catch (e) { avisarFalhaGrafico("grafico-funil", e); }
  try { renderizarGraficoFinanceiro(regs); } catch (e) { avisarFalhaGrafico("grafico-financeiro", e); }
  try { renderizarGraficoDiaSemana(regs); } catch (e) { avisarFalhaGrafico("grafico-diasemana", e); }
}

function avisarFalhaGrafico(canvasId, erro) {
  const canvas = el(canvasId);
  if (canvas && canvas.parentElement) {
    canvas.parentElement.innerHTML = `<p style="font-size:12px;color:var(--vermelho);">Gráfico indisponível no momento (${erro.message}).</p>`;
  }
}

function agrupar(regs, campo) {
  const mapa = {};
  regs.forEach(r => { const c = r[campo] || "Não informado"; mapa[c] = (mapa[c] || 0) + 1; });
  return Object.entries(mapa).map(([nome, total]) => ({ nome, total })).sort((a, b) => b.total - a.total);
}

function renderizarCards(regs) {
  const total = regs.length;
  const pagas = regs.filter(r => r.statusMatricula === "Matrícula Paga").length;
  const minutas = regs.filter(r => r.statusMatricula === "Minuta FIES").length;
  const fies = regs.filter(r => r.tipoFinanceiro === "FIES").length;
  const particular = regs.filter(r => r.tipoFinanceiro === "Particular").length;
  const conversao = total > 0 ? ((pagas / total) * 100).toFixed(1) : "0.0";

  const meta = ABA_ATIVA === "Presencial" ? DADOS.config.metaPresencial : (DADOS.config.metaTotalGeral - DADOS.config.metaPresencial);
  const pctMeta = meta > 0 ? Math.min(100, (total / meta) * 100) : 0;

  const campeao = agrupar(regs, "consultor")[0];
  const cursoCampeao = agrupar(regs, "curso")[0];
  const cidadeCampea = agrupar(regs, "cidade")[0];

  el("cards").innerHTML = `
    ${card("Total de Matrículas", total, "roxo")}
    ${card("Matrículas Pagas", pagas, "verde")}
    ${card("Minutas FIES", minutas, "laranja")}
    ${card("Particular", particular, "")}
    ${card("FIES", fies, "")}
    ${card("Conversão", conversao + "%", "roxo")}
    ${cardMeta("Meta do bloco (" + ABA_ATIVA + ")", total, meta, pctMeta)}
    ${card("Consultor campeão", campeao ? campeao.nome : "-", "", campeao ? campeao.total + " matrículas" : "")}
    ${card("Curso campeão", cursoCampeao ? cursoCampeao.nome : "-", "", cursoCampeao ? cursoCampeao.total + " matrículas" : "")}
    ${card("Cidade campeã", cidadeCampea ? cidadeCampea.nome : "-", "", cidadeCampea ? cidadeCampea.total + " matrículas" : "")}
  `;
}

function card(rotulo, valor, cor, rodape) {
  return `<div class="card">
    <div class="rotulo">${rotulo}</div>
    <div class="valor ${cor || ""}">${valor}</div>
    ${rodape ? `<div class="rodape">${rodape}</div>` : ""}
  </div>`;
}

function cardMeta(rotulo, total, meta, pct) {
  return `<div class="card">
    <div class="rotulo">${rotulo}</div>
    <div class="valor roxo">${pct.toFixed(0)}%</div>
    <div class="meta-barra-fundo"><div class="meta-barra-preenchida" style="width:${pct}%"></div></div>
    <div class="rodape">${total} de ${meta} matrículas</div>
  </div>`;
}

function renderizarFunil(regs) {
  const total = regs.length;
  const pagas = regs.filter(r => r.statusMatricula === "Matrícula Paga").length;
  const minutas = regs.filter(r => r.statusMatricula === "Minuta FIES").length;

  destruirGrafico("funil");
  const ctx = el("grafico-funil").getContext("2d");
  GRAFICOS.funil = new Chart(ctx, {
    type: "bar",
    data: {
      labels: ["Total", "Minutas (aguardando FIES)", "Matrículas pagas"],
      datasets: [{ data: [total, minutas, pagas], backgroundColor: ["#4B2E83", "#F7931E", "#1E9E6B"], borderRadius: 8, maxBarThickness: 60 }]
    },
    options: {
      indexAxis: "y",
      plugins: { legend: { display: false } },
      scales: { x: { beginAtZero: true, grid: { color: "#EDEDF3" } }, y: { grid: { display: false } } },
    }
  });
}

function renderizarGraficoFinanceiro(regs) {
  const grupos = agrupar(regs, "tipoFinanceiro");
  destruirGrafico("financeiro");
  const ctx = el("grafico-financeiro").getContext("2d");
  GRAFICOS.financeiro = new Chart(ctx, {
    type: "doughnut",
    data: { labels: grupos.map(g => g.nome), datasets: [{ data: grupos.map(g => g.total), backgroundColor: ["#4B2E83", "#F7931E", "#8A8A9C"] }] },
    options: { plugins: { legend: { position: "bottom", labels: { boxWidth: 10, font: { size: 11 } } } } }
  });
}

function renderizarGraficoDiaSemana(regs) {
  const DIAS = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"];
  const contagem = [0, 0, 0, 0, 0, 0, 0];
  regs.forEach(r => {
    if (!r.dataCarimbo) return;
    const dia = new Date(r.dataCarimbo).getDay();
    contagem[dia]++;
  });
  destruirGrafico("diasemana");
  const ctx = el("grafico-diasemana").getContext("2d");
  GRAFICOS.diasemana = new Chart(ctx, {
    type: "bar",
    data: { labels: DIAS, datasets: [{ data: contagem, backgroundColor: "#7F77DD", borderRadius: 6, maxBarThickness: 36 }] },
    options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { precision: 0 } } } }
  });
}

function renderizarRanking(id, grupos) {
  el(id).innerHTML = grupos.slice(0, 8).map((g, i) => `
    <li><span class="ranking-pos">${i + 1}</span><span class="ranking-nome">${g.nome}</span><span class="ranking-total">${g.total}</span></li>
  `).join("") || '<li>Sem dados para os filtros atuais.</li>';
}

function renderizarConsultores(regs) {
  const nomes = unicos(regs, "consultor");
  el("consultores-grid").innerHTML = nomes.map(nome => {
    const doC = regs.filter(r => r.consultor === nome);
    const pagas = doC.filter(r => r.statusMatricula === "Matrícula Paga").length;
    const minutas = doC.filter(r => r.statusMatricula === "Minuta FIES").length;
    const fies = doC.filter(r => r.tipoFinanceiro === "FIES").length;
    const particular = doC.filter(r => r.tipoFinanceiro === "Particular").length;
    const conversao = doC.length > 0 ? ((pagas / doC.length) * 100).toFixed(1) : "0.0";
    const meta = DADOS.config.metaPorConsultor;
    const pctMeta = Math.min(100, (doC.length / meta) * 100);

    // Dias sem matrícula vem calculado do backend, considerando TODOS os
    // registros do consultor (não só os do bloco/filtro atual), por isso
    // buscamos direto na lista "consultores" do payload.
    const infoConsultor = (DADOS.consultores || []).find(c => c.nome === nome);
    const diasSemMatricula = infoConsultor ? infoConsultor.diasSemMatricula : null;
    const alertaParado = diasSemMatricula !== null && diasSemMatricula >= 5;

    return `<div class="consultor-card">
      <h4>${nome}</h4>
      ${diasSemMatricula !== null ? `<div class="${alertaParado ? 'alerta' : 'alerta-ok'}" style="margin-bottom:8px;">Está há <b>${diasSemMatricula}</b> dia(s) sem matrícula.</div>` : ""}
      <div class="consultor-metrica"><span>Matrículas (bloco atual)</span><span>${doC.length}</span></div>
      <div class="consultor-metrica"><span>Pagas</span><span>${pagas}</span></div>
      <div class="consultor-metrica"><span>Minutas FIES</span><span>${minutas}</span></div>
      <div class="consultor-metrica"><span>FIES</span><span>${fies}</span></div>
      <div class="consultor-metrica"><span>Particular</span><span>${particular}</span></div>
      <div class="consultor-metrica"><span>Conversão</span><span>${conversao}%</span></div>
      <div class="meta-barra-fundo"><div class="meta-barra-preenchida" style="width:${pctMeta}%"></div></div>
      <div class="rodape" style="margin-top:6px;">${doC.length} de ${meta} (meta individual)</div>
    </div>`;
  }).join("") || "<p>Nenhum consultor com registros para os filtros atuais.</p>";
}

function renderizarInsights() {
  const ins = DADOS.insights;
  if (!ins) { el("insights").innerHTML = "<p>Sem dados suficientes ainda.</p>"; return; }
  const linhas = [];
  linhas.push(`Os insights abaixo consideram Presencial e EAD/Semipresencial juntos (a soma dos dois blocos), diferente dos cards de consultor acima, que mostram só o bloco selecionado na aba.`);
  if (ins.melhorDiaSemana) linhas.push(`O dia da semana com mais matrículas é <b>${ins.melhorDiaSemana.dia}</b> (${ins.melhorDiaSemana.total} no total).`);
  if (ins.consultorMelhorConversao) linhas.push(`<b>${ins.consultorMelhorConversao.nome}</b> tem a melhor conversão entre os consultores com pelo menos 3 matrículas, somando Presencial e EAD: ${ins.consultorMelhorConversao.conversaoPercentual}%.`);
  if (ins.consultorMaisTempoSemMatricula) linhas.push(`<b>${ins.consultorMaisTempoSemMatricula.nome}</b> é quem está há mais tempo sem fechar matrícula: ${ins.consultorMaisTempoSemMatricula.dias} dia(s).`);
  if (ins.cidadeQueMaisGeraAlunos) linhas.push(`A cidade que mais gera alunos é <b>${ins.cidadeQueMaisGeraAlunos.nome}</b> (${ins.cidadeQueMaisGeraAlunos.total}).`);
  if (ins.canalQueMaisGeraMatriculas) linhas.push(`O canal que mais gera matrículas é <b>${ins.canalQueMaisGeraMatriculas.nome}</b> (${ins.canalQueMaisGeraMatriculas.total}).`);
  if (ins.principalMotivoDeEscolha) linhas.push(`O principal motivo de escolha relatado é <b>${ins.principalMotivoDeEscolha.nome}</b> (${ins.principalMotivoDeEscolha.total} respostas).`);
  if (ins.percentualEscolaPublica !== null) linhas.push(`${ins.percentualEscolaPublica}% dos alunos vêm de escola pública${ins.percentualEscolaPrivada !== null ? ` e ${ins.percentualEscolaPrivada}% de escola privada` : ""}.`);
  if (ins.particularVsFIES) linhas.push(`Particular: ${ins.particularVsFIES.particular} matrículas · FIES: ${ins.particularVsFIES.fies} matrículas.`);
  if (ins.tendenciaUltimos7Dias) linhas.push(`Nos últimos 7 dias: ${ins.tendenciaUltimos7Dias.matriculasUltimos7Dias} matrículas, contra ${ins.tendenciaUltimos7Dias.matriculas7DiasAnteriores} nos 7 dias anteriores. Tendência: <b>${ins.tendenciaUltimos7Dias.tendencia}</b>.`);
  el("insights").innerHTML = linhas.map(l => `<div class="alerta-ok" style="margin-bottom:8px;">${l}</div>`).join("") || "<p>Sem dados suficientes ainda.</p>";
}

function renderizarExclusoes() {
  // Mensagem de exclusão removida a pedido: os registros de consultores
  // fora da equipe (ex: Valécia) continuam sendo excluídos dos cálculos
  // normalmente, só não aparece mais um aviso fixo no topo do painel.
  el("exclusoes").innerHTML = "";
}

function destruirGrafico(chave) {
  if (GRAFICOS[chave]) { GRAFICOS[chave].destroy(); }
}

document.addEventListener("DOMContentLoaded", iniciar);
