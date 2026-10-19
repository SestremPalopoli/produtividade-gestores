const INDICADORES = [
    { nome: "RIT", aba: "RIT_ACUM", celulaMeta: "I7" },
    { nome: "Falha de M.O", aba: "FMO_ACUM", celulaMeta: "I9" },
    { nome: "Checklist", aba: "CHECK_ACUM", celulaMeta: "I8" },
    { nome: "Degradação Rede", aba: "DEG_ACUM", celulaMeta: "I10" },
    { nome: "Lançamento SD", aba: "SD_ACUM", celulaMeta: "I11" },
    { nome: "Material Indevido", aba: "FISC_ACUM", celulaMeta: "I7" },
    { nome: "Clean-Up", aba: "CLEAN-UP_ACUM", celulaMeta: "I12" },
    { nome: "Liberado Comercial", aba: "LIBERADO COMERCIAL", tipo: "tres-meses" }
];

const GESTORES = [
    "Alexandre Morais Medeiros Da Silva",
    "Andre Mendes De Sousa",
    "Carmine Avarese Filho",
    "Denis Reis",
    "Deyvison Bruno Cutrim Ferreira",
    "Diego Mendonca Santos",
    "Priscila De Melo Sevilha Zandeli"
];

let workbookGlobal = null;

function normalizarTexto(texto) {
    return String(texto || "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .trim()
        .toUpperCase();
}

function escaparHtml(texto) {
    return String(texto)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function abreviarNomeGestor(nome) {
    if (
        normalizarTexto(nome) ===
        "DEYVISON BRUNO CUTRIM FERREIRA"
    ) {
        return "Bruno Cutrim";
    }

    const partes = String(nome || "").trim().split(/\s+/);

    if (partes.length <= 3) {
        return nome;
    }

    const conectores = new Set(["DA", "DE", "DO", "DAS", "DOS", "E"]);
    const partesPrincipais = partes.filter(
        parte => !conectores.has(normalizarTexto(parte))
    );

    const primeiroNome = partesPrincipais[0];
    const ultimoNome = partesPrincipais[partesPrincipais.length - 1];
    const inicialMeio = partesPrincipais.length > 2
        ? `${partesPrincipais[1].charAt(0).toUpperCase()}.`
        : "";

    return [primeiroNome, inicialMeio, ultimoNome]
        .filter(Boolean)
        .join(" ");
}

function formatarDataExcel(valor) {
    if (valor === null || valor === undefined || valor === "") {
        return "";
    }

    if (typeof valor === "number") {
        return XLSX.SSF.format("dd/mm/yyyy", valor);
    }

    const texto = String(valor).trim();

    if (/^\d{2}\/\d{2}\/\d{4}$/.test(texto)) {
        return texto;
    }

    const data = new Date(texto);

    if (!Number.isNaN(data.getTime())) {
        return data.toLocaleDateString("pt-BR");
    }

    return texto;
}

function obterMesLinha(linha) {
    const mesInformado = Number(linha["MÊS"]);

    if (mesInformado >= 1 && mesInformado <= 12) {
        return mesInformado;
    }

    const dataFormatada = formatarDataExcel(linha["DATA"]);
    const partes = dataFormatada.split("/");

    return partes.length === 3 ? Number(partes[1]) : 0;
}

function obterFarol(percentual) {
    if (percentual >= 100) {
        return { simbolo: "🟢", classe: "verde" };
    }

    if (percentual >= 75) {
        return { simbolo: "🟡", classe: "amarelo" };
    }

    return { simbolo: "🔴", classe: "vermelho" };
}

function nomeMesCurto(numeroMes) {
    const nomes = [
        "Todos", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
        "Jul", "Ago", "Set", "Out", "Nov", "Dez"
    ];

    return nomes[numeroMes] || "Mês";
}

function obterChaveMesAno(linha, colunaData = "DATA", colunaMes = "MÊS") {
    const dataFormatada = formatarDataExcel(linha[colunaData]);
    const partes = dataFormatada.split("/").map(Number);

    if (partes.length === 3 && partes.every(Number.isFinite)) {
        return {
            chave: partes[2] * 100 + partes[1],
            mes: partes[1],
            ano: partes[2]
        };
    }

    const mes = Number(linha[colunaMes]) || obterMesLinha(linha);
    const ano = new Date().getFullYear();

    return mes >= 1 && mes <= 12
        ? { chave: ano * 100 + mes, mes, ano }
        : null;
}

function obterUltimosTresMeses(dados, colunaData = "DATA", colunaMes = "MÊS") {
    const mesesEncontrados = new Map();

    dados.forEach(linha => {
        const periodo = obterChaveMesAno(linha, colunaData, colunaMes);

        if (periodo) {
            mesesEncontrados.set(periodo.chave, periodo);
        }
    });

    return [...mesesEncontrados.values()]
        .sort((a, b) => a.chave - b.chave)
        .slice(-3);
}

function contarRegistroNoPeriodo(linha, periodo, colunaData = "DATA", colunaMes = "MÊS") {
    const periodoLinha = obterChaveMesAno(linha, colunaData, colunaMes);
    return Boolean(periodoLinha && periodoLinha.chave === periodo.chave);
}

function mostrarErro(mensagem) {
    const elemento = document.getElementById("mensagemErro");
    elemento.hidden = false;
    elemento.textContent = mensagem;
}

function ocultarErro() {
    document.getElementById("mensagemErro").hidden = true;
}

function carregarFiltroGestores() {
    const select = document.getElementById("filtroGestor");

    if (!select) {
        return;
    }

    select.innerHTML = '<option value="">Selecione...</option>';

    GESTORES.forEach(gestor => {
        const opcao = document.createElement("option");
        opcao.value = gestor;
        opcao.textContent = gestor;
        select.appendChild(opcao);
    });
}

function abrirPainelGestor() {
    const selectGestor = document.getElementById("filtroGestor");
    const gestor = selectGestor.value;

    if (!gestor) {
        return;
    }

    const mes = document.getElementById("mesSelecionado").value;
    const data = document.getElementById("filtroData").value;
    const parametros = new URLSearchParams({ gestor, mes, data });

    window.open(
        `gestor.html?${parametros.toString()}`,
        "_blank",
        "noopener,noreferrer"
    );

    selectGestor.value = "";
}

async function carregarArquivoExcel() {
    const resposta = await fetch("dados.xlsx", { cache: "no-store" });

    if (!resposta.ok) {
        throw new Error(`Não foi possível carregar dados.xlsx. Código ${resposta.status}.`);
    }

    const ultimaAtualizacao = resposta.headers.get("Last-Modified");

    document.getElementById("data-atualizacao").textContent = ultimaAtualizacao
        ? new Date(ultimaAtualizacao).toLocaleString("pt-BR")
        : new Date().toLocaleString("pt-BR");

    const arquivo = await resposta.arrayBuffer();
    workbookGlobal = XLSX.read(arquivo, { type: "array" });
}

function carregarFiltroDatas() {
    const select = document.getElementById("filtroData");
    const datas = new Set();

    INDICADORES.forEach(indicador => {
        const aba = workbookGlobal.Sheets[indicador.aba];

        if (!aba) {
            return;
        }

        XLSX.utils.sheet_to_json(aba).forEach(linha => {
            const gestor = normalizarTexto(linha["VISTORIADOR"]);
            const gestorPermitido = GESTORES.some(
                nome => normalizarTexto(nome) === gestor
            );

            if (!gestorPermitido) {
                return;
            }

            const data = formatarDataExcel(linha["DATA"]);

            if (data) {
                datas.add(data);
            }
        });
    });

    const datasOrdenadas = [...datas].sort((a, b) => {
        const [diaA, mesA, anoA] = a.split("/").map(Number);
        const [diaB, mesB, anoB] = b.split("/").map(Number);
        return new Date(anoB, mesB - 1, diaB) - new Date(anoA, mesA - 1, diaA);
    });

    select.innerHTML = "";

    if (datasOrdenadas.length === 0) {
        select.innerHTML = '<option value="">Sem datas disponíveis</option>';
        return;
    }

    datasOrdenadas.slice(0, 30).forEach(data => {
        const opcao = document.createElement("option");
        opcao.value = data;
        opcao.textContent = data;
        select.appendChild(opcao);
    });
}

function obterDadosIndicador(indicador) {
    const aba = workbookGlobal.Sheets[indicador.aba];

    if (!aba) {
        return [];
    }

    return XLSX.utils.sheet_to_json(aba);
}

function obterMeta(indicador) {
    const calendario = workbookGlobal.Sheets["CALENDÁRIO"];

    if (!calendario) {
        return 0;
    }

    return Number(calendario[indicador.celulaMeta]?.v || 0);
}

function contarRegistros(dados, gestor, mesSelecionado, dataSelecionada) {
    const gestorNormalizado = normalizarTexto(gestor);
    let totalMes = 0;
    let totalDia = 0;

    dados.forEach(linha => {
        if (normalizarTexto(linha["VISTORIADOR"]) !== gestorNormalizado) {
            return;
        }

        const mesLinha = obterMesLinha(linha);
        const dataLinha = formatarDataExcel(linha["DATA"]);

        if (mesSelecionado === 0 || mesLinha === mesSelecionado) {
            totalMes++;
        }

        if (dataSelecionada && dataLinha === dataSelecionada) {
            totalDia++;
        }
    });

    return { totalMes, totalDia };
}

function criarCard(indicador, total) {
    return `
        <article class="card">
            <h3>${escaparHtml(indicador.nome)}</h3>
            <p>${total}</p>
        </article>
    `;
}

function criarPainel(indicador, dados, meta, mesSelecionado, dataSelecionada) {
    let totalMesGeral = 0;
    let totalDiaGeral = 0;
    let totalMeta = 0;
    let linhas = "";

    GESTORES.forEach(gestor => {
        const resultado = contarRegistros(
            dados,
            gestor,
            mesSelecionado,
            dataSelecionada
        );

        const percentual = meta > 0 ? (resultado.totalMes / meta) * 100 : 0;
        const farol = obterFarol(percentual);

        totalMesGeral += resultado.totalMes;
        totalDiaGeral += resultado.totalDia;
        totalMeta += meta;

        const metaExibida = indicador.nome === "Falha de M.O"
            ? Math.round(meta)
            : meta;

        linhas += `
            <tr>
                <td class="nome-gestor" title="${escaparHtml(gestor)}">
                    ${escaparHtml(abreviarNomeGestor(gestor))}
                </td>
                <td>${resultado.totalMes}</td>
                <td>${resultado.totalDia}</td>
                <td class="meta">${metaExibida}</td>
                <td class="${farol.classe}">${farol.simbolo} ${percentual.toFixed(1)}%</td>
            </tr>
        `;
    });

    const percentualTotal = totalMeta > 0
        ? (totalMesGeral / totalMeta) * 100
        : 0;

    const farolTotal = obterFarol(percentualTotal);
    const totalMetaExibida = indicador.nome === "Falha de M.O"
        ? Math.round(totalMeta)
        : totalMeta;
    const tituloMes = mesSelecionado === 0 ? "Total" : nomeMesCurto(mesSelecionado);

    return `
        <article class="painel-indicador">
            <header class="painel-indicador-cabecalho">
                <h2>${escaparHtml(indicador.nome)}</h2>
            </header>

            <div class="tabela-scroll">
                <table class="tabela-indicador">
                    <thead>
                        <tr>
                            <th>Gestor</th>
                            <th>${tituloMes}</th>
                            <th title="${escaparHtml(dataSelecionada || "Data não selecionada")}">
                                ${dataSelecionada ? escaparHtml(dataSelecionada.substring(0, 5)) : "Dia"}
                            </th>
                            <th class="meta">Meta</th>
                            <th>%</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${linhas}
                        <tr class="linha-total">
                            <td>TOTAL</td>
                            <td>${totalMesGeral}</td>
                            <td>${totalDiaGeral}</td>
                            <td class="meta">${totalMetaExibida}</td>
                            <td class="${farolTotal.classe}">${farolTotal.simbolo} ${percentualTotal.toFixed(1)}%</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </article>
    `;
}

function criarPainelLiberadoComercial(indicador, dados, dataSelecionada) {
    const periodos = obterUltimosTresMeses(dados, "Data visita Gestor", "MÊS");
    const periodosExibidos = [...periodos];

    while (periodosExibidos.length < 3) {
        periodosExibidos.unshift(null);
    }

    const totaisMeses = [0, 0, 0];
    let totalDia = 0;
    let linhas = "";

    GESTORES.forEach(gestor => {
        const gestorNormalizado = normalizarTexto(gestor);
        const dadosGestor = dados.filter(
            linha => normalizarTexto(linha["GESTOR"]) === gestorNormalizado
        );

        const valoresMeses = periodosExibidos.map((periodo, indice) => {
            const total = periodo
                ? dadosGestor.filter(linha => contarRegistroNoPeriodo(linha, periodo, "Data visita Gestor", "MÊS")).length
                : 0;
            totaisMeses[indice] += total;
            return total;
        });

        const valorDia = dadosGestor.filter(
            linha => dataSelecionada && formatarDataExcel(linha["Data visita Gestor"]) === dataSelecionada
        ).length;
        totalDia += valorDia;

        linhas += `
            <tr>
                <td class="nome-gestor" title="${escaparHtml(gestor)}">
                    ${escaparHtml(abreviarNomeGestor(gestor))}
                </td>
                <td>${valoresMeses[0]}</td>
                <td>${valoresMeses[1]}</td>
                <td>${valoresMeses[2]}</td>
                <td>${valorDia}</td>
            </tr>
        `;
    });

    const titulosMeses = periodosExibidos.map(
        periodo => periodo ? nomeMesCurto(periodo.mes) : "-"
    );

    return `
        <article class="painel-indicador painel-liberado">
            <header class="painel-indicador-cabecalho">
                <h2>${escaparHtml(indicador.nome)}</h2>
            </header>

            <div class="tabela-scroll">
                <table class="tabela-indicador tabela-liberado">
                    <thead>
                        <tr>
                            <th>Gestor</th>
                            <th>${titulosMeses[0]}</th>
                            <th>${titulosMeses[1]}</th>
                            <th>${titulosMeses[2]}</th>
                            <th title="${escaparHtml(dataSelecionada || "Data não selecionada")}">
                                ${dataSelecionada ? escaparHtml(dataSelecionada.substring(0, 5)) : "Dia"}
                            </th>
                        </tr>
                    </thead>
                    <tbody>
                        ${linhas}
                        <tr class="linha-total">
                            <td>TOTAL</td>
                            <td>${totaisMeses[0]}</td>
                            <td>${totaisMeses[1]}</td>
                            <td>${totaisMeses[2]}</td>
                            <td>${totalDia}</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </article>
    `;
}

function criarPainelReservado() {
    return `
        <article class="painel-indicador painel-reservado">
            <div>
                <span class="icone-reservado">＋</span>
                <h2>Futuro indicador</h2>
                <p>Espaço reservado para o oitavo indicador.</p>
            </div>
        </article>
    `;
}

function atualizarPainel() {
    if (!workbookGlobal) {
        return;
    }

    ocultarErro();

    const mesSelecionado = Number(
        document.getElementById("mesSelecionado").value
    );

    const dataSelecionada =
        document.getElementById("filtroData").value;

    const cards = [];
    const paineis = [];

    INDICADORES.forEach(indicador => {
        const dados = obterDadosIndicador(indicador);
        const colunaGestor = indicador.tipo === "tres-meses"
            ? "GESTOR"
            : "VISTORIADOR";

        const dadosPermitidos = dados.filter(linha => {
            const gestor = normalizarTexto(linha[colunaGestor]);
            return GESTORES.some(nome => normalizarTexto(nome) === gestor);
        });

        if (indicador.tipo === "tres-meses") {
            const periodos = obterUltimosTresMeses(
                dadosPermitidos,
                "Data visita Gestor",
                "MÊS"
            );
            const periodoMaisRecente = periodos.at(-1);
            const totalCard = periodoMaisRecente
                ? dadosPermitidos.filter(
                    linha => contarRegistroNoPeriodo(
                        linha,
                        periodoMaisRecente,
                        "Data visita Gestor",
                        "MÊS"
                    )
                ).length
                : 0;

            cards.push(criarCard(indicador, totalCard));
            paineis.push(
                criarPainelLiberadoComercial(indicador, dadosPermitidos, dataSelecionada)
            );
            return;
        }

        const meta = obterMeta(indicador);
        const totalCard = dadosPermitidos.filter(
            linha => mesSelecionado === 0 || obterMesLinha(linha) === mesSelecionado
        ).length;

        cards.push(criarCard(indicador, totalCard));
        paineis.push(
            criarPainel(indicador, dadosPermitidos, meta, mesSelecionado, dataSelecionada)
        );
    });

    document.getElementById("cardsIndicadores").innerHTML = cards.join("");
    document.getElementById("gradePaineis").innerHTML = paineis.join("");
}

async function iniciarPainel() {
    try {
        await carregarArquivoExcel();
        carregarFiltroDatas();
        atualizarPainel();
    } catch (erro) {
        console.error(erro);
        mostrarErro(
            "Não foi possível carregar o painel. Verifique o arquivo dados.xlsx e os nomes das abas."
        );
    }
}

document.addEventListener("DOMContentLoaded", () => {
    carregarFiltroGestores();

    const filtroGestor = document.getElementById("filtroGestor");

    if (filtroGestor) {
        filtroGestor.addEventListener("change", abrirPainelGestor);
    }

    document
        .getElementById("mesSelecionado")
        .addEventListener("change", atualizarPainel);

    document
        .getElementById("filtroData")
        .addEventListener("change", atualizarPainel);

    document
        .getElementById("btnFullscreen")
        .addEventListener("click", async () => {
            try {
                if (!document.fullscreenElement) {
                    await document.documentElement.requestFullscreen();
                } else {
                    await document.exitFullscreen();
                }
            } catch (erro) {
                console.error("Não foi possível alternar a tela cheia.", erro);
            }
        });

    iniciarPainel();
});
