/*
========================================================
INTRANET EXECUTIVA KPI CCO • SLU
========================================================
*/

const VALORES_FIXOS = {
  "P1": 296.00,
  "P2.1": 1027.42,
  "P2.2": 1027.42,
  "P3": 41992.93,
  "P4": 68.80,
  "P5": 160.94,
  "P6": 76.24,
  "P7": 49811.72,
  "P8": 81001.04,
  "P9": 122039.23,
  "P10": 346660.01,
  "P11": 272459.08,
  "P12": 0.83
};

const SERVICOS_FIXOS = ["P3", "P7", "P8", "P9", "P10"];

let painelExecutivo = [];
let painelExecutivoOriginal = [];
let operacoes = [];
let operacoesOriginal = [];
let sheetsOriginais = {};
let todasAsAbas = [];
let graficoExecucao = null;
let graficoPizza = null;

/* INICIALIZAÇÃO */
document.addEventListener("DOMContentLoaded", () => {
  atualizarData();

  const input = document.getElementById("arquivoExcel");
  if (input) input.addEventListener("change", importarPlanilhas);

  mostrarKPI("geral");
});

// Controle de acesso por perfil
document.addEventListener("DOMContentLoaded", () => {
  const usuarioLogado = JSON.parse(localStorage.getItem("usuarioLogado"));

  if (!usuarioLogado) {
    window.location.href = "login.html";
    return;
  }

  const perfil = usuarioLogado.perfil;

  // Diretoria: somente visualização
  if (perfil === "Diretoria") {
    // Esconde botões de ação
    esconderElemento("btnImportar");
    esconderElemento("btnLimpar");
    esconderElemento("btnExportar");
    esconderElemento("btnImplantar");

    // Esconde página/menu KPI
    esconderElemento("menuKpi");
    esconderElemento("paginaKpi");

    // Bloqueia qualquer botão perigoso por segurança
    bloquearBotoesDiretoria();
  }
});

function esconderElemento(id) {
  const elemento = document.getElementById(id);
  if (elemento) {
    elemento.style.display = "none";
  }
}

function bloquearBotoesDiretoria() {
  const botoesBloqueados = [
    "btnImportar",
    "btnLimpar",
    "btnExportar",
    "btnImplantar"
  ];

  botoesBloqueados.forEach(id => {
    const botao = document.getElementById(id);
    if (botao) {
      botao.disabled = true;
      botao.style.display = "none";
    }
  });
}





/* DATA */
function atualizarData() {
  const dataAtual = document.getElementById("dataAtual");
  if (!dataAtual) return;

  dataAtual.innerText = new Date().toLocaleDateString("pt-BR", {
    weekday: "long",
    day: "2-digit",
    month: "long",
    year: "numeric"
  });
}

/* IMPORTAÇÃO */
async function importarPlanilhas(evento) {
  const arquivos = Array.from(evento.target.files || []);
  if (!arquivos.length) return;

  mostrarLoading(true);

  try {
    painelExecutivo = [];
    painelExecutivoOriginal = [];
    operacoes = [];
    operacoesOriginal = [];
    sheetsOriginais = {};
    todasAsAbas = [];

    for (const arquivo of arquivos) {
      const buffer = await arquivo.arrayBuffer();

      const workbook = XLSX.read(buffer, {
        type: "array",
        cellDates: true
      });

      workbook.SheetNames.forEach(nomeAba => {
        const sheet = workbook.Sheets[nomeAba];

        const dados = XLSX.utils.sheet_to_json(sheet, {
          defval: "",
          raw: false
        });

        const normalizados = dados.map(item => normalizarObjeto(item));
        const nomeNormalizado = normalizar(nomeAba);

        sheetsOriginais[nomeNormalizado] = {
          nomeOriginal: nomeAba,
          codigoServico: extrairCodigo(nomeAba),
          dadosOriginais: dados,
          dadosNormalizados: normalizados
        };

        todasAsAbas.push({
          arquivo: arquivo.name,
          aba: nomeAba
        });
      });
    }

    gerarPainelExecutivo();
    gerarOperacoes();

    painelExecutivoOriginal = [...painelExecutivo];
    operacoesOriginal = [...operacoes];

    atualizarDashboard();

    document.getElementById("nomeArquivo").innerText =
      `${arquivos.length} arquivo(s) importado(s) | ${todasAsAbas.length} aba(s) lida(s)`;

    alert("Planilhas importadas com sucesso!");
  } catch (erro) {
    console.error(erro);
    alert("Erro ao importar planilha.");
  } finally {
    mostrarLoading(false);
    evento.target.value = "";
  }
}

/* GERA PAINEL */
function gerarPainelExecutivo() {
  const painel = sheetsOriginais["painel executivo"];

  if (!painel) {
    alert("Aba Painel Executivo não encontrada.");
    return;
  }

  painelExecutivo = painel.dadosNormalizados.map(item => {
    const servico = String(item.servico || item.programa || item.codigo || "").toUpperCase();
    const abaServico = buscarAbaServico(servico);

    const acumulado = abaServico
      ? calcularAcumulado(item.medicao, abaServico.dadosNormalizados)
      : numero(item.acumulado_mes || item.acumulado_no_mes || item.executado);

    const previsto = numero(
      item.previsto_mes ||
      item.meta_mes ||
      item.previsto ||
      item.meta
    );

    const valorUnitario = VALORES_FIXOS[servico] || 0;

    const valorFinal = SERVICOS_FIXOS.includes(servico)
      ? valorUnitario
      : valorUnitario * acumulado;

    return {
      servico,
      nome_servico: item.nome_servico || item.nome_do_servico || item.descricao || "",
      acumulado_mes: acumulado,
      semana_1: numero(item.semana_1 || item["1a_semana_8d"] || item["1o_semana_8d"]),
      semana_2: numero(item.semana_2 || item["2a_semana_6d"] || item["2o_semana_6d"]),
      semana_3: numero(item.semana_3 || item["3a_semana_6d"] || item["3o_semana_6d"]),
      semana_4: numero(item.semana_4 || item["4a_semana_6d"] || item["4o_semana_6d"]),
      medicao: item.medicao || "",
      previsto_mes: previsto,
      porcentagem_execucao: calcularPercentual(acumulado, previsto),
      valor: valorFinal,
      status: acumulado > 0 ? "Com dados" : "Sem dados"
    };
  });
}

/* GERA OPERAÇÕES */
function gerarOperacoes() {
  operacoes = [];

  Object.keys(sheetsOriginais).forEach(nome => {
    if (nome === "painel executivo") return;

    const sheet = sheetsOriginais[nome];
    if (!sheet.codigoServico) return;

    sheet.dadosNormalizados.forEach(item => {
      operacoes.push({
        servico: sheet.codigoServico,
        origem: sheet.nomeOriginal,
        data: item.data || item.data_operacao || "",
        data_normalizada: normalizarData(item.data || item.data_operacao || ""),
        turno: item.turno || "",
        ra: item.ra || item.regiao_administrativa || "Por demanda",
        peso: numero(item.peso || item.peso_total || item.peso_t),
        viagens: numero(item.viagens || item.qtd_viagem || item.qtd_viagens),
        km: numero(item.km || item.km_total || item.km_executado),
        equipe: numero(item.equipe || item.qtd_equipe || item.qdt_equipe),
        status: "Com dados"
      });
    });
  });
}

/* ATUALIZA DASHBOARD */
function atualizarDashboard() {
  renderCards();
  renderTabelaExecutiva();
  renderTabelaContratual();
  renderResumo();
  renderAlertas();
  renderPerguntas();
  renderTabelaDados();
  renderFiltros();
  renderGraficos();
}

/* CARDS */
function renderCards() {
  const servicosComDados = painelExecutivo.filter(item => item.status === "Com dados").length;

  const mediaExecucao = painelExecutivo.length
    ? painelExecutivo.reduce((soma, item) => soma + numero(item.porcentagem_execucao), 0) / painelExecutivo.length
    : 0;

  preencherTexto("kpiServicosDados", servicosComDados);
  preencherTexto("kpiExecucaoMedia", `${formatarNumero(mediaExecucao)}%`);
  preencherTexto("kpiAbas", todasAsAbas.length);
}

/* TABELA EXECUTIVA */
function renderTabelaExecutiva() {
  const tabela = document.getElementById("tabelaPainelExecutivo");
  if (!tabela) return;

  if (!painelExecutivo.length) {
    tabela.innerHTML = `<tr><td colspan="12">Importe uma planilha para visualizar os dados.</td></tr>`;
    return;
  }

  tabela.innerHTML = painelExecutivo.map(item => `
    <tr>
      <td><strong>${item.servico}</strong></td>
      <td>${item.nome_servico}</td>
      <td>${formatarNumero(item.acumulado_mes)}</td>
      <td>${formatarNumero(item.semana_1)}</td>
      <td>${formatarNumero(item.semana_2)}</td>
      <td>${formatarNumero(item.semana_3)}</td>
      <td>${formatarNumero(item.semana_4)}</td>
      <td>${item.medicao}</td>
      <td>${formatarNumero(item.previsto_mes)}</td>
      <td>${formatarNumero(item.porcentagem_execucao)}%</td>
      <td>${formatarMoeda(item.valor)}</td>
      <td><span class="badge ${item.status === "Com dados" ? "ok" : "info"}">${item.status}</span></td>
    </tr>
  `).join("");
}

/* TABELA CONTRATUAL */
function renderTabelaContratual() {
  const tabela = document.getElementById("tabelaContratual");
  if (!tabela) return;

  if (!painelExecutivo.length) {
    tabela.innerHTML = `<tr><td colspan="6">Nenhum dado contratual importado.</td></tr>`;
    return;
  }

  tabela.innerHTML = painelExecutivo.map(item => `
    <tr>
      <td>${item.servico} - ${item.nome_servico}</td>
      <td>${item.medicao}</td>
      <td>${formatarNumero(item.previsto_mes)}</td>
      <td>${formatarNumero(item.acumulado_mes)}</td>
      <td>${formatarNumero(item.porcentagem_execucao)}%</td>
      <td><span class="badge ${item.status === "Com dados" ? "ok" : "info"}">${item.status}</span></td>
    </tr>
  `).join("");
}

/* RESUMO */
function renderResumo() {
  const resumo = document.getElementById("resumoExecutivo");
  if (!resumo) return;

  const top = [...painelExecutivo]
    .filter(item => item.acumulado_mes > 0)
    .sort((a, b) => b.acumulado_mes - a.acumulado_mes)
    .slice(0, 6);

  if (!top.length) {
    resumo.innerHTML = "<p>Importe dados para gerar o resumo executivo.</p>";
    return;
  }

  const maior = Math.max(...top.map(item => item.acumulado_mes), 1);

  resumo.innerHTML = top.map(item => `
    <div class="metric-row">
      <strong>${item.servico}</strong>
      <div class="bar-bg">
        <div class="bar" style="width:${Math.min(100, item.acumulado_mes / maior * 100)}%"></div>
      </div>
      <b>${formatarNumero(item.acumulado_mes)}</b>
    </div>
  `).join("");
}

/* ALERTAS */
function renderAlertas() {
  const alertas = document.getElementById("alertasExecutivos");
  if (!alertas) return;

  if (!painelExecutivo.length) {
    alertas.innerHTML = `<p><span class="badge info">Aguardando</span> Importe uma planilha.</p>`;
    return;
  }

  const semDados = painelExecutivo.filter(item => item.status === "Sem dados");

  if (!semDados.length) {
    alertas.innerHTML = `<p><span class="badge ok">Normal</span> Sem alertas críticos.</p>`;
    return;
  }

  alertas.innerHTML = `<p><span class="badge critico">Atenção</span> ${semDados.length} serviço(s) sem dados.</p>`;
}

/* RESPOSTAS RÁPIDAS */
function renderPerguntas() {
  const tabela = document.getElementById("tabelaPerguntas");
  if (!tabela) return;

  tabela.innerHTML = `
    <tr>
      <td>Quantos serviços possuem dados?</td>
      <td>${painelExecutivo.filter(item => item.status === "Com dados").length}</td>
      <td><span class="badge ok">Operacional</span></td>
    </tr>
    <tr>
      <td>Quantas abas foram importadas?</td>
      <td>${todasAsAbas.length}</td>
      <td><span class="badge info">Sistema</span></td>
    </tr>
  `;
}

/* TABELA DADOS */
function renderTabelaDados() {
  const tabela = document.getElementById("tabelaDados");
  if (!tabela) return;

  const busca = normalizar(document.getElementById("busca")?.value || "");
  const filtroPrograma = document.getElementById("filtroPrograma")?.value || "Todos";
  const filtroStatus = document.getElementById("filtroStatus")?.value || "Todos";

  const filtrados = operacoes.filter(item => {
    const texto = normalizar(Object.values(item).join(" "));

    return texto.includes(busca) &&
      (filtroPrograma === "Todos" || item.servico === filtroPrograma) &&
      (filtroStatus === "Todos" || item.status === filtroStatus);
  });

  tabela.innerHTML = filtrados.map(item => `
    <tr>
      <td>${item.servico}</td>
      <td>${item.origem}</td>
      <td>${item.data}</td>
      <td>${item.turno}</td>
      <td>${item.ra}</td>
      <td>${formatarNumero(item.peso)}</td>
      <td>${formatarNumero(item.viagens)}</td>
      <td>${formatarNumero(item.km)}</td>
      <td>${formatarNumero(item.equipe)}</td>
    </tr>
  `).join("") || `<tr><td colspan="9">Nenhum dado encontrado.</td></tr>`;
}

/* FILTROS */
function renderFiltros() {
  const filtroPrograma = document.getElementById("filtroPrograma");
  const filtroStatus = document.getElementById("filtroStatus");

  if (!filtroPrograma || !filtroStatus) return;

  const programas = ["Todos", ...new Set(operacoes.map(item => item.servico).filter(Boolean))];

  filtroPrograma.innerHTML = programas.map(item => `<option>${item}</option>`).join("");

  filtroStatus.innerHTML = `
    <option>Todos</option>
    <option>Com dados</option>
    <option>Sem dados</option>
  `;
}

/* GRÁFICOS */
function renderGraficos() {
  const ctxExecucao = document.getElementById("graficoExecucao");
  const ctxPizza = document.getElementById("graficoAcumulado");

  if (!ctxExecucao || !ctxPizza) return;

  if (graficoExecucao) graficoExecucao.destroy();
  if (graficoPizza) graficoPizza.destroy();

  const labels = painelExecutivo.map(item => item.servico);
  const dadosExecucao = painelExecutivo.map(item => item.porcentagem_execucao);
  const dadosAcumulado = painelExecutivo.map(item => item.acumulado_mes);

  graficoExecucao = new Chart(ctxExecucao, {
    type: "bar",
    data: {
      labels,
      datasets: [{
        label: "% Execução",
        data: dadosExecucao,
        borderRadius: 10,
        backgroundColor: "rgba(12,107,63,.7)"
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false
    }
  });

  graficoPizza = new Chart(ctxPizza, {
    type: "doughnut",
    data: {
      labels,
      datasets: [{
        data: dadosAcumulado
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false
    }
  });
}

/* FILTRO POR DATA */
function aplicarFiltroPeriodo() {
  const data = document.getElementById("filtroDia").value;

  if (!data) {
    limparFiltroPeriodo();
    return;
  }

  operacoes = operacoesOriginal.filter(item => item.data_normalizada === data);
  atualizarDashboard();
}

/* LIMPAR FILTRO */
function limparFiltroPeriodo() {
  const filtro = document.getElementById("filtroDia");
  if (filtro) filtro.value = "";

  painelExecutivo = [...painelExecutivoOriginal];
  operacoes = [...operacoesOriginal];

  atualizarDashboard();
}

/* TELAS */
function mostrarTela(nome, botao) {
  document.querySelectorAll(".tela").forEach(item => item.classList.remove("ativa"));

  const tela = document.getElementById(`tela-${nome}`);
  if (tela) tela.classList.add("ativa");

  document.querySelectorAll("nav button").forEach(btn => btn.classList.remove("active"));
  if (botao) botao.classList.add("active");
}

/* SERVIÇOS */
function mostrarServico(codigo, botao) {
  document.querySelectorAll("#tela-contrato .servico-btn").forEach(btn => btn.classList.remove("active"));
  if (botao) botao.classList.add("active");

  const geral = document.getElementById("servico-geral");
  const detalhe = document.getElementById("servico-detalhe");

  if (codigo === "geral") {
    geral.classList.add("ativa");
    detalhe.classList.remove("ativa");
    return;
  }

  geral.classList.remove("ativa");
  detalhe.classList.add("ativa");
  renderDetalheServico(codigo);
}

/* DETALHE DO SERVIÇO */
function renderDetalheServico(codigo) {
  const detalhe = document.getElementById("detalheServico");
  const aba = buscarAbaServico(codigo);

  if (!aba) {
    detalhe.innerHTML = `<div class="not-found">Serviço não encontrado.</div>`;
    return;
  }

  const colunas = aba.dadosOriginais.length ? Object.keys(aba.dadosOriginais[0]) : [];

  const thead = colunas.map(col => `<th>${col}</th>`).join("");

  const tbody = aba.dadosOriginais.map(linha => `
    <tr>
      ${colunas.map(col => `<td>${linha[col] || ""}</td>`).join("")}
    </tr>
  `).join("");

/* CALCULA INFORMAÇÕES DOS CARDS */
const totalPeso = aba.dadosNormalizados.reduce(
  (soma, item) => soma + numero(item.peso || item.peso_total || item.peso_t),
  0
);

const totalViagens = aba.dadosNormalizados.reduce(
  (soma, item) => soma + numero(item.viagens || item.qtd_viagem || item.qtd_viagens),
  0
);

const totalKm = aba.dadosNormalizados.reduce(
  (soma, item) => soma + numero(item.km || item.km_total || item.km_executado),
  0
);

const totalEquipes = aba.dadosNormalizados.reduce(
  (soma, item) => soma + numero(item.equipe || item.qtd_equipe || item.qdt_equipe),
  0
);

detalhe.innerHTML = `

  <!-- ======================================== -->
  <!-- CARDS EXECUÇÃO -->
  <!-- ======================================== -->
  <section class="cards">

    <div class="card">
      <span>KM Executado</span>
      <strong>${formatarNumero(totalKm)}</strong>
      <small>quilometragem</small>
    </div>

    <div class="card">
      <span>Equipes</span>
      <strong>${formatarNumero(totalEquipes)}</strong>
      <small>equipes operacionais</small>
    </div>

  </section>

  <!-- ======================================== -->
  <!-- TABELA -->
  <!-- ======================================== -->
  <section class="section">

      <div class="section-title">
        <span>Espelho da planilha</span>
        <h2>${codigo}</h2>
      </div>

      <div class="table-wrap">
        <table>
          <thead>
            <tr>${thead}</tr>
          </thead>

          <tbody>
            ${tbody || `
              <tr>
                <td colspan="20">
                  Nenhuma informação encontrada.
                </td>
              </tr>
            `}
          </tbody>
        </table>
      </div>

  </section>
`;
}

/* KPI'S */
const KPIS_POR_SERVICO = {
  geral: {
    titulo: "Indicadores Gerais de Operação",
    descricao: "Indicadores estratégicos para acompanhamento de volume, frota, distância, produtividade, tempo, velocidade e eficiência operacional.",
    blocos: [
      {
        titulo: "Volume e Demanda",
        itens: [
          "Dias de operação efetivos por mês ou por ano.",
          "Quantidade coletada por turno.",
          "Quantidade total coletada por dia.",
          "Quantidade coletada por veículo por turno."
        ]
      },
      {
        titulo: "Frota e Viagens",
        itens: [
          "Viagens por veículo por turno.",
          "Viagens por veículo por dia ou mês.",
          "Capacidade real média dos veículos.",
          "Taxa de ocupação dos veículos.",
          "Frota ativa comparada com frota dimensionada."
        ],
        formula: "Taxa de ocupação = toneladas coletadas por viagem ÷ capacidade média × 100"
      },
      {
        titulo: "Distância e Percurso",
        itens: [
          "Distância mensal total percorrida.",
          "Distância semanal e diária média.",
          "Distância média por viagem.",
          "Distância média por rota ou RA."
        ]
      },
      {
        titulo: "Tempo e Produtividade",
        itens: [
          "Horas produtivas por veículo.",
          "Horas de veículo parado à disposição.",
          "Tempo médio por viagem.",
          "Tempo médio por km rodado.",
          "Duração de refeição comparada ao tempo de operação."
        ]
      }
    ]
  },

  "P1": {
    titulo: "P1 - Coleta Orgânica",
    descricao: "Controle do peso coletado e produtividade da coleta orgânica.",
    blocos: [{
      titulo: "Respostas do Serviço",
      itens: [
        "Peso coletado.",
        "Toneladas coletadas por dia.",
        "Toneladas coletadas por turno.",
        "Toneladas por veículo por turno.",
        "Toneladas por viagem."
      ],
      formula: "Peso coletado = soma das toneladas registradas na aba P1"
    }]
  },

  "P2.1": {
    titulo: "P2.1 - Coleta Seletiva",
    descricao: "Controle das viagens realizadas e viagens recebidas como orgânica.",
    blocos: [{
      titulo: "Respostas do Serviço",
      itens: [
        "Viagens realizadas.",
        "Viagens recebidas como orgânica.",
        "Controle de estouro do peso previsto.",
        "Viagens por veículo por dia.",
        "Produtividade por rota ou RA."
      ],
      formula: "Resultado = viagens realizadas + viagens recebidas como orgânica"
    }]
  },

  "P2.2": {
    titulo: "P2.2 - Rejeito Seletivo das IRR",
    descricao: "Controle das viagens do rejeito seletivo.",
    blocos: [{
      titulo: "Respostas do Serviço",
      itens: [
        "Viagens realizadas.",
        "Viagens recebidas como orgânica.",
        "Controle de peso previsto.",
        "Viagens por veículo.",
        "Viagens por turno."
      ],
      formula: "Resultado = viagens realizadas + viagens recebidas como orgânica"
    }]
  },

  "P3": {
    titulo: "P3 - Remoção Manual",
    descricao: "Controle da quantidade de equipes por dia.",
    blocos: [{
      titulo: "Respostas do Serviço",
      itens: [
        "Quantidade de equipes dia.",
        "Equipes previstas.",
        "Equipes em operação.",
        "Percentual de atendimento.",
        "Produtividade por equipe."
      ],
      formula: "Quantidade de equipes dia = total de equipes registradas"
    }]
  },

  "P4": {
    titulo: "P4 - Remoção Mecanizada",
    descricao: "Controle do peso coletado e produtividade mecanizada.",
    blocos: [{
      titulo: "Respostas do Serviço",
      itens: [
        "Peso coletado.",
        "Toneladas por viagem.",
        "Toneladas por km rodado.",
        "Viagens por veículo.",
        "Distância média por viagem."
      ],
      formula: "Peso coletado = soma do peso registrado na aba P4"
    }]
  },

  "P5": {
    titulo: "P5 - Varrição Manual",
    descricao: "Controle de quilometragem prevista, executada e índice de execução.",
    blocos: [{
      titulo: "Respostas do Serviço",
      itens: [
        "Quilometragem prevista.",
        "Quilometragem executada.",
        "Índice de execução.",
        "Km varridos por hora.",
        "Equipes ou varredores por km varrido."
      ],
      formula: "Índice de execução = quilometragem executada ÷ quilometragem prevista × 100"
    }]
  },

  "P6": {
    titulo: "P6 - Varrição Mecanizada",
    descricao: "Controle da quilometragem mecanizada executada.",
    blocos: [{
      titulo: "Respostas do Serviço",
      itens: [
        "Quilometragem prevista.",
        "Quilometragem executada.",
        "Índice de execução.",
        "Km por hora.",
        "Velocidade média operacional."
      ],
      formula: "Índice de execução = km executado ÷ km previsto × 100"
    }]
  },

  "P7": {
    titulo: "P7 - Lavagem de Vias e Logradouros",
    descricao: "Controle das equipes previstas e equipes em operação.",
    blocos: [{
      titulo: "Respostas do Serviço",
      itens: [
        "Equipes previstas.",
        "Equipes em operação.",
        "Percentual de execução.",
        "Horas produtivas.",
        "Produtividade por equipe."
      ],
      formula: "Execução = equipes em operação ÷ equipes previstas × 100"
    }]
  },

  "P8": {
    titulo: "P8 - Limpeza de Equipamentos e Bens",
    descricao: "Controle das equipes de limpeza de equipamentos e bens.",
    blocos: [{
      titulo: "Respostas do Serviço",
      itens: [
        "Equipes previstas.",
        "Equipes em operação.",
        "Locais atendidos.",
        "Produtividade por equipe.",
        "Percentual de atendimento."
      ],
      formula: "Execução = equipes em operação ÷ equipes previstas × 100"
    }]
  },

  "P9": {
    titulo: "P9 - Catação",
    descricao: "Controle das equipes de catação.",
    blocos: [{
      titulo: "Respostas do Serviço",
      itens: [
        "Equipes previstas.",
        "Equipes em operação.",
        "Áreas atendidas.",
        "Produtividade por equipe.",
        "Percentual de execução."
      ],
      formula: "Execução = equipes em operação ÷ equipes previstas × 100"
    }]
  },

  "P10": {
    titulo: "P10 - Pintura Mecanizada",
    descricao: "Controle das equipes de pintura mecanizada.",
    blocos: [{
      titulo: "Respostas do Serviço",
      itens: [
        "Equipes previstas.",
        "Equipes em operação.",
        "Índice de execução.",
        "Produtividade por equipe.",
        "Atendimento por rota ou área."
      ],
      formula: "Execução = equipes em operação ÷ equipes previstas × 100"
    }]
  },

  "P11": {
    titulo: "P11 - Limpeza Pós-Eventos e Coleta de Gordura",
    descricao: "Controle de equipes em serviços especiais.",
    blocos: [
      {
        titulo: "Limpeza Pós-Eventos",
        itens: [
          "Equipes previstas.",
          "Equipes em operação.",
          "Eventos atendidos.",
          "Tempo médio de atendimento.",
          "Produtividade por equipe."
        ]
      },
      {
        titulo: "Coleta de Gordura",
        itens: [
          "Equipes previstas.",
          "Equipes em operação.",
          "Demandas atendidas.",
          "Tempo médio de operação.",
          "Percentual de execução."
        ],
        formula: "Execução = equipes em operação ÷ equipes previstas × 100"
      }
    ]
  },

  "P12": {
    titulo: "P12 - Transbordo",
    descricao: "Controle de peso coletado por distrito e km percorrido.",
    blocos: [{
      titulo: "Respostas do Serviço",
      itens: [
        "Peso coletado por distrito.",
        "Km percorrido por distrito.",
        "Toneladas por km.",
        "Viagens por distrito.",
        "Distância média por viagem."
      ],
      formula: "Resultado = peso coletado por distrito × km percorrido por distrito"
    }]
  }
};

/* MOSTRAR KPI */
function mostrarKPI(codigo, botao) {
  document.querySelectorAll("#tela-kpi .servico-btn").forEach(btn => btn.classList.remove("active"));
  if (botao) botao.classList.add("active");

  const area = document.getElementById("conteudoKPI");
  if (!area) return;

  const dados = KPIS_POR_SERVICO[codigo];

  if (!dados) {
    area.innerHTML = `<div class="not-found">KPI não encontrado para este serviço.</div>`;
    return;
  }

  area.innerHTML = `
    <div class="kpi-header">
      <span>KPI Operacional</span>
      <h2>${dados.titulo}</h2>
      <p>${dados.descricao}</p>
    </div>

    <div class="kpi-grid">
      ${dados.blocos.map(bloco => `
        <div class="kpi-box">
          <h3>${bloco.titulo}</h3>
          <ul>
            ${bloco.itens.map(item => `<li>${item}</li>`).join("")}
          </ul>
          ${bloco.formula ? `<div class="kpi-formula">${bloco.formula}</div>` : ""}
        </div>
      `).join("")}
    </div>
  `;
}

/* FUNÇÕES AUXILIARES */
function buscarAbaServico(codigo) {
  for (const nome in sheetsOriginais) {
    const aba = sheetsOriginais[nome];
    if (aba.codigoServico === codigo) return aba;
  }
  return null;
}

function calcularAcumulado(medicao, dados) {
  const tipo = normalizar(medicao);

  return dados.reduce((soma, item) => {
    if (tipo.includes("km")) return soma + numero(item.km || item.km_total || item.km_executado);
    if (tipo.includes("t")) return soma + numero(item.peso || item.peso_total || item.peso_t);
    if (tipo.includes("vg") || tipo.includes("viagem")) return soma + numero(item.viagens || item.qtd_viagem || item.qtd_viagens);
    if (tipo.includes("equipe")) return soma + numero(item.equipe || item.qtd_equipe || item.qdt_equipe);
    return soma + numero(item.total || item.quantidade || item.valor);
  }, 0);
}

function calcularPercentual(realizado, previsto) {
  if (!previsto) return 0;
  return Math.min((realizado / previsto) * 100, 100);
}

function mostrarLoading(ativo) {
  const loading = document.getElementById("loadingOverlay");
  if (!loading) return;
  loading.classList.toggle("ativo", ativo);
}

function sair() {
  window.location.href = "login.html";
}

function limparBanco() {
  if (!confirm("Deseja apagar todos os dados?")) return;

  painelExecutivo = [];
  painelExecutivoOriginal = [];
  operacoes = [];
  operacoesOriginal = [];
  sheetsOriginais = {};
  todasAsAbas = [];

  atualizarDashboard();
  preencherTexto("nomeArquivo", "Nenhuma planilha importada");
}

function extrairCodigo(texto) {
  const match = String(texto).toUpperCase().match(/P\d+(\.\d+)?/);
  return match ? match[0] : "";
}

function normalizarObjeto(obj) {
  const novo = {};

  Object.keys(obj).forEach(chave => {
    const novaChave = normalizar(chave)
      .replace(/ª/g, "a")
      .replace(/º/g, "o")
      .replace(/%/g, "porcentagem")
      .replace(/\$/g, "valor")
      .replace(/\./g, "")
      .replace(/\s+/g, "_")
      .replace(/[^\w]/g, "");

    novo[novaChave] = obj[chave];
  });

  return novo;
}

function normalizar(texto) {
  return String(texto || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

function normalizarData(valor) {
  if (!valor) return "";

  const texto = String(valor).trim();

  if (/^\d{4}-\d{2}-\d{2}/.test(texto)) return texto.slice(0, 10);

  const br = texto.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);

  if (br) {
    const dia = br[1].padStart(2, "0");
    const mes = br[2].padStart(2, "0");
    let ano = br[3];
    if (ano.length === 2) ano = "20" + ano;
    return `${ano}-${mes}-${dia}`;
  }

  const data = new Date(texto);
  if (isNaN(data)) return "";

  return data.toISOString().slice(0, 10);
}

function numero(valor) {
  if (valor === null || valor === undefined || valor === "") return 0;
  if (typeof valor === "number") return valor;

  let texto = String(valor).trim().replace(/[^\d,.-]/g, "");

  const temVirgula = texto.includes(",");
  const temPonto = texto.includes(".");

  if (temVirgula && temPonto) {
    texto = texto.replace(/\./g, "").replace(",", ".");
  } else if (temVirgula && !temPonto) {
    texto = texto.replace(",", ".");
  }

  const convertido = Number(texto);
  return isNaN(convertido) ? 0 : convertido;
}

function formatarNumero(valor) {
  return numero(valor).toLocaleString("pt-BR", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2
  });
}

function formatarMoeda(valor) {
  return numero(valor).toLocaleString("pt-BR", {
    style: "currency",
    currency: "BRL"
  });
}

function preencherTexto(id, texto) {
  const elemento = document.getElementById(id);
  if (elemento) elemento.innerText = texto;
}
/* ========================================================
   DETALHE DO SERVIÇO COM CARDS EXECUTIVOS
   P1 A P12
======================================================== */
function renderDetalheServico(codigo) {

  const detalhe =
    document.getElementById("detalheServico");

  const aba =
    buscarAbaServico(codigo);

  if (!aba) {

    detalhe.innerHTML = `
      <div class="not-found">
        Serviço não encontrado.
      </div>
    `;

    return;
  }

  /* ==========================================
     LOCALIZA DADOS DO PAINEL EXECUTIVO
  ========================================== */
  const dadosPainel =
    painelExecutivo.find(
      item => item.servico === codigo
    );

  /* ==========================================
     PREVISTO
  ========================================== */
  const previsto =
    dadosPainel
      ? numero(dadosPainel.previsto_mes)
      : 0;

  /* ==========================================
     EXECUTADO
  ========================================== */
  const executado =
    dadosPainel
      ? numero(dadosPainel.acumulado_mes)
      : 0;

  /* ==========================================
     EXECUÇÃO %
  ========================================== */
  const percentual =
    dadosPainel
      ? numero(dadosPainel.porcentagem_execucao)
      : 0;

  /* ==========================================
     VALOR TOTAL
  ========================================== */
  const valor =
    dadosPainel
      ? numero(dadosPainel.valor)
      : 0;

  /* ==========================================
     KM TOTAL
  ========================================== */
  const totalKm =
    aba.dadosNormalizados.reduce(
      (soma, item) =>
        soma + numero(
          item.km ||
          item.km_total ||
          item.km_executado
        ),
      0
    );

  /* ==========================================
     TOTAL EQUIPES
  ========================================== */
  const totalEquipes =
    aba.dadosNormalizados.reduce(
      (soma, item) =>
        soma + numero(
          item.equipe ||
          item.qtd_equipe ||
          item.qdt_equipe
        ),
      0
    );

  /* ==========================================
     TOTAL VIAGENS
  ========================================== */
  const totalViagens =
    aba.dadosNormalizados.reduce(
      (soma, item) =>
        soma + numero(
          item.viagens ||
          item.qtd_viagem ||
          item.qtd_viagens
        ),
      0
    );

  /* ==========================================
     TOTAL PESO
  ========================================== */
  const totalPeso =
    aba.dadosNormalizados.reduce(
      (soma, item) =>
        soma + numero(
          item.peso ||
          item.peso_total ||
          item.peso_t
        ),
      0
    );

  /* ==========================================
     COLUNAS TABELA
  ========================================== */
  const colunas =
    aba.dadosOriginais.length
      ? Object.keys(
          aba.dadosOriginais[0]
        )
      : [];

  const thead =
    colunas.map(col =>
      `<th>${col}</th>`
    ).join("");

  const tbody =
    aba.dadosOriginais.map(linha => `
      <tr>
        ${colunas.map(col => `
          <td>${linha[col] || ""}</td>
        `).join("")}
      </tr>
    `).join("");

  /* ==========================================
     HTML FINAL
  ========================================== */
  detalhe.innerHTML = `

    <!-- ===================================== -->
    <!-- CARDS EXECUTIVOS -->
    <!-- ===================================== -->
    <section class="cards">

      <!-- PREVISTO -->
      <div class="card">
        <span>Previsto</span>

        <strong>
          ${formatarNumero(previsto)}
        </strong>

        <small>meta operacional</small>
      </div>

      <!-- EXECUTADO -->
      <div class="card">
        <span>Executado</span>

        <strong>
          ${formatarNumero(executado)}
        </strong>

        <small>acumulado realizado</small>
      </div>

      <!-- EXECUÇÃO -->
      <div class="card">
        <span>% Execução</span>

        <strong>
          ${formatarNumero(percentual)}%
        </strong>

        <small>índice operacional</small>
      </div>

      <div class="card">
        <strong>
          ${formatarNumero(totalPeso)}
        </strong>

        <small>toneladas</small>
      </div>

      <!-- VIAGENS -->
      <div class="card">
        <span>Viagens</span>

        <strong>
          ${formatarNumero(totalViagens)}
        </strong>

        <small>operações</small>
      </div>

      <!-- KM -->
      <div class="card">
        <span>KM Executado</span>

        <strong>
          ${formatarNumero(totalKm)}
        </strong>

        <small>quilometragem</small>
      </div>

      <!-- EQUIPES -->
      <div class="card">
        <span>Equipes</span>

        <strong>
          ${formatarNumero(totalEquipes)}
        </strong>

        <small>equipes operacionais</small>
      </div>

    </section>

    <!-- ===================================== -->
    <!-- TABELA -->
    <!-- ===================================== -->
    <section class="section">

      <div class="section-title">
        <span>Espelho da planilha</span>
        <h2>${codigo}</h2>
      </div>

      <div class="table-wrap">

        <table>

          <thead>
            <tr>
              ${thead}
            </tr>
          </thead>

          <tbody>

            ${tbody || `
              <tr>
                <td colspan="20">
                  Nenhuma informação encontrada.
                </td>
              </tr>
            `}

          </tbody>

        </table>

      </div>

    </section>
  `;
}