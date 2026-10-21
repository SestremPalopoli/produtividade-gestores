const CONFIGURACAO = {
    arquivo: "dados.xlsx",
    abaRit: "RIT_ICOMON",
    abaEstrutura: "ESTRUTURA_ICOMON",
    colunasRit: {
        id: "ID",
        data: "DATA",
        vistoriador: "VISTORIADOR",
        ano: "ANO",
        mes: "MÊS",
        semana: "SEMANA"
    },
    colunasEstrutura: {
        supervisor: "SUPERVISOR",
        coordenador: "COORDENADOR"
    }
};

let dadosRit = [];
let estruturaCompleta = [];
let equipesRecolhidas = false;
let dataReferenciaChave = 0;

function normalizarTexto(texto) {
    return String(texto || "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .trim()
        .replace(/\s+/g, " ")
        .toUpperCase();
}

function escaparHtml(texto) {
    return String(texto || "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function formatarDataExcel(valor) {
    if (valor === null || valor === undefined || valor === "") return "";

    if (typeof valor === "number") {
        return XLSX.SSF.format("dd/mm/yyyy", valor);
    }

    const texto = String(valor).trim();
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(texto)) return texto;

    const data = new Date(texto);
    return Number.isNaN(data.getTime())
        ? ""
        : data.toLocaleDateString("pt-BR");
}

function decomporData(dataFormatada) {
    const partes = String(dataFormatada).split("/").map(Number);
    if (partes.length !== 3 || partes.some(Number.isNaN)) return null;

    return {
        dia: partes[0],
        mes: partes[1],
        ano: partes[2],
        chave: partes[2] * 10000 + partes[1] * 100 + partes[0]
    };
}

function nomeMesCurto(mes) {
    return [
        "", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
        "Jul", "Ago", "Set", "Out", "Nov", "Dez"
    ][mes] || "";
}

function prepararDadosRit(linhas) {
    return linhas.map(linha => {
        const dataFormatada = formatarDataExcel(
            linha[CONFIGURACAO.colunasRit.data]
        );
        const data = decomporData(dataFormatada);
        const ano = Number(linha[CONFIGURACAO.colunasRit.ano]) || data?.ano || 0;
        const mes = Number(linha[CONFIGURACAO.colunasRit.mes]) || data?.mes || 0;
        const semana = Number(linha[CONFIGURACAO.colunasRit.semana]) || 0;

        return {
            supervisor: normalizarTexto(
                linha[CONFIGURACAO.colunasRit.vistoriador]
            ),
            dataFormatada,
            dataChave: data?.chave || 0,
            ano,
            mes,
            semana,
            mesChave: ano * 100 + mes,
            semanaChave: ano * 100 + semana
        };
    }).filter(item => item.supervisor);
}

function prepararEstrutura(linhas) {
    const pares = new Map();

    linhas.forEach(linha => {
        const supervisorOriginal = String(
            linha[CONFIGURACAO.colunasEstrutura.supervisor] || ""
        ).trim();
        const coordenadorOriginal = String(
            linha[CONFIGURACAO.colunasEstrutura.coordenador] || ""
        ).trim();
        const supervisor = normalizarTexto(supervisorOriginal);
        const coordenador = normalizarTexto(coordenadorOriginal);

        if (!supervisor || !coordenador) return;

        const chave = `${coordenador}|${supervisor}`;
        if (!pares.has(chave)) {
            pares.set(chave, {
                coordenador,
                coordenadorOriginal,
                supervisor,
                supervisorOriginal
            });
        }
    });

    return [...pares.values()].sort((a, b) =>
        a.coordenadorOriginal.localeCompare(b.coordenadorOriginal, "pt-BR") ||
        a.supervisorOriginal.localeCompare(b.supervisorOriginal, "pt-BR")
    );
}

function dataLocalParaIso(data) {
    const ano = data.getFullYear();
    const mes = String(data.getMonth() + 1).padStart(2, "0");
    const dia = String(data.getDate()).padStart(2, "0");
    return `${ano}-${mes}-${dia}`;
}

function isoParaChave(valor) {
    return Number(String(valor || "").replaceAll("-", "")) || 0;
}

function chaveParaData(chave) {
    const texto = String(chave).padStart(8, "0");
    return new Date(Number(texto.slice(0, 4)), Number(texto.slice(4, 6)) - 1, Number(texto.slice(6, 8)));
}

function chaveParaTitulo(chave) {
    const data = chaveParaData(chave);
    return `${String(data.getDate()).padStart(2, "0")}/${String(data.getMonth() + 1).padStart(2, "0")}/${data.getFullYear()}`;
}

function obterDmenosUm() {
    const data = new Date();
    data.setHours(0, 0, 0, 0);
    data.setDate(data.getDate() - 1);
    return data;
}

function configurarFiltroData() {
    const input = document.getElementById("filtroData");
    const dMenosUm = obterDmenosUm();
    input.max = dataLocalParaIso(dMenosUm);
    input.value = input.max;
    dataReferenciaChave = isoParaChave(input.value);
}

function obterPeriodos() {
    const limite = dataReferenciaChave || isoParaChave(dataLocalParaIso(obterDmenosUm()));
    const meses = new Map();
    const semanas = new Map();

    dadosRit.filter(item => item.dataChave && item.dataChave <= limite).forEach(item => {
        if (item.ano && item.mes) {
            meses.set(item.mesChave, {
                chave: item.mesChave,
                ano: item.ano,
                mes: item.mes,
                titulo: `${nomeMesCurto(item.mes)}/${String(item.ano).slice(-2)}`
            });
        }
        if (item.ano && item.semana) {
            semanas.set(item.semanaChave, {
                chave: item.semanaChave,
                ano: item.ano,
                semana: item.semana,
                titulo: `Sem. ${item.semana}`
            });
        }
    });

    const dataFinal = chaveParaData(limite);
    const datas = [];
    for (let deslocamento = 4; deslocamento >= 0; deslocamento--) {
        const data = new Date(dataFinal);
        data.setDate(data.getDate() - deslocamento);
        const chave = Number(dataLocalParaIso(data).replaceAll("-", ""));
        datas.push({ chave, data: chaveParaTitulo(chave), titulo: chaveParaTitulo(chave) });
    }

    return {
        meses: [...meses.values()].sort((a, b) => a.chave - b.chave).slice(-3),
        semanas: [...semanas.values()].sort((a, b) => a.chave - b.chave).slice(-5),
        datas
    };
}

function contar(supervisores, tipo, chave) {
    const conjunto = new Set(supervisores);

    return dadosRit.filter(item => {
        if (!conjunto.has(item.supervisor)) return false;
        if (tipo === "mes") return item.mesChave === chave;
        if (tipo === "semana") return item.semanaChave === chave;
        return item.dataChave === chave;
    }).length;
}

function obterEstruturaFiltrada() {
    const coordenador = normalizarTexto(
        document.getElementById("filtroCoordenador").value
    );
    const supervisor = normalizarTexto(
        document.getElementById("filtroSupervisor").value
    );

    return estruturaCompleta.filter(item =>
        (!coordenador || item.coordenador === coordenador) &&
        (!supervisor || item.supervisor === supervisor)
    );
}

function agruparPorCoordenador(estrutura) {
    const grupos = new Map();

    estrutura.forEach(item => {
        if (!grupos.has(item.coordenador)) {
            grupos.set(item.coordenador, {
                nome: item.coordenadorOriginal,
                chave: item.coordenador,
                supervisores: []
            });
        }

        grupos.get(item.coordenador).supervisores.push({
            nome: item.supervisorOriginal,
            chave: item.supervisor
        });
    });

    return [...grupos.values()];
}

function criarCelula(valor, classes = "") {
    const classeZero = valor === 0 ? "valor-zero" : "";
    return `<td class="${classes} ${classeZero}">${valor}</td>`;
}

function criarCelulas(supervisores, periodos) {
    const meses = periodos.meses.map((periodo, indice) =>
        criarCelula(contar(supervisores, "mes", periodo.chave), indice === 0 ? "periodo-separador" : "")
    ).join("");
    const semanas = periodos.semanas.map((periodo, indice) =>
        criarCelula(contar(supervisores, "semana", periodo.chave), indice === 0 ? "periodo-separador" : "")
    ).join("");
    const datas = periodos.datas.map((periodo, indice) =>
        criarCelula(contar(supervisores, "data", periodo.chave), indice === 0 ? "periodo-separador" : "")
    ).join("");
    return meses + semanas + datas;
}

function montarCabecalho(periodos) {
    const preencher = (lista, quantidade, titulo) => {
        const vazios = Math.max(quantidade - lista.length, 0);
        return [
            ...Array(vazios).fill({ titulo: "-" }),
            ...lista
        ].map(periodo =>
            `<th>${escaparHtml(periodo.titulo)}</th>`
        ).join("");
    };

    document.querySelector("#tabelaIcomon thead").innerHTML = `
        <tr>
            <th class="coluna-nome">ESTRUTURA</th>
            ${preencher(periodos.meses, 3, "mes")}
            ${preencher(periodos.semanas, 5, "semana")}
            ${preencher(periodos.datas, 5, "data")}
        </tr>
    `;
}

function montarTabela() {
    const estrutura = obterEstruturaFiltrada();
    const grupos = agruparPorCoordenador(estrutura);
    const periodos = obterPeriodos();
    const tbody = document.querySelector("#tabelaIcomon tbody");

    montarCabecalho(periodos);
    tbody.innerHTML = "";

    const supervisoresVisiveis = estrutura.map(item => item.supervisor);

    grupos.forEach(grupo => {
        const supervisores = grupo.supervisores.map(item => item.chave);

        tbody.innerHTML += `
            <tr class="linha-coordenador" data-grupo="${escaparHtml(grupo.chave)}">
                <td class="coluna-nome">
                    <button class="btn-grupo" type="button" data-coordenador="${escaparHtml(grupo.chave)}">
                        ${equipesRecolhidas ? "▶" : "▼"}
                    </button>
                    ${escaparHtml(grupo.nome)}
                </td>
                ${criarCelulas(supervisores, periodos)}
            </tr>
        `;

        grupo.supervisores.forEach(supervisor => {
            tbody.innerHTML += `
                <tr class="linha-supervisor ${equipesRecolhidas ? "oculta" : ""}"
                    data-equipe="${escaparHtml(grupo.chave)}">
                    <td class="coluna-nome">${escaparHtml(supervisor.nome)}</td>
                    ${criarCelulas([supervisor.chave], periodos)}
                </tr>
            `;
        });
    });

    tbody.innerHTML += `
        <tr class="linha-total">
            <td class="coluna-nome">TOTAL</td>
            ${criarCelulas(supervisoresVisiveis, periodos)}
        </tr>
    `;

    const mesAtual = periodos.meses.at(-1);
    const mesAnterior = periodos.meses.at(-2);
    const semanaAtual = periodos.semanas.at(-1);
    const semanaAnterior = periodos.semanas.at(-2);
    const ultimoPeriodo = periodos.datas.at(-1);
    const totalMes = mesAtual ? contar(supervisoresVisiveis, "mes", mesAtual.chave) : 0;
    const totalMesAnterior = mesAnterior ? contar(supervisoresVisiveis, "mes", mesAnterior.chave) : 0;
    const totalSemana = semanaAtual ? contar(supervisoresVisiveis, "semana", semanaAtual.chave) : 0;
    const totalSemanaAnterior = semanaAnterior ? contar(supervisoresVisiveis, "semana", semanaAnterior.chave) : 0;
    const zerados = semanaAtual
        ? supervisoresVisiveis.filter(item => contar([item], "semana", semanaAtual.chave) === 0).length
        : 0;
    document.getElementById("totalMesAtual").textContent = totalMes;
    document.getElementById("variacaoMes").textContent = calcularVariacao(totalMes, totalMesAnterior);
    document.getElementById("totalSemanaAtual").textContent = totalSemana;
    document.getElementById("variacaoSemana").textContent = calcularVariacao(totalSemana, totalSemanaAnterior);
    document.getElementById("totalPeriodoRecente").textContent = ultimoPeriodo
        ? contar(supervisoresVisiveis, "data", ultimoPeriodo.chave)
        : 0;
    document.getElementById("dataPeriodoRecente").textContent = ultimoPeriodo
        ? `Referência: ${ultimoPeriodo.titulo}`
        : "Sem data";
    document.getElementById("totalZerados").textContent = zerados;

    document.querySelectorAll(".btn-grupo").forEach(botao => {
        botao.addEventListener("click", alternarGrupo);
    });
}

function alternarGrupo(evento) {
    const coordenador = evento.currentTarget.dataset.coordenador;
    const linhas = document.querySelectorAll(
        `.linha-supervisor[data-equipe="${CSS.escape(coordenador)}"]`
    );
    const ocultar = [...linhas].some(linha => !linha.classList.contains("oculta"));

    linhas.forEach(linha => linha.classList.toggle("oculta", ocultar));
    evento.currentTarget.textContent = ocultar ? "▶" : "▼";
}

function atualizarFiltroSupervisor() {
    const coordenador = normalizarTexto(document.getElementById("filtroCoordenador").value);
    const supervisores = new Map();
    estruturaCompleta
        .filter(item => !coordenador || item.coordenador === coordenador)
        .forEach(item => supervisores.set(item.supervisor, item.supervisorOriginal));
    preencherSelect("filtroSupervisor", supervisores);
}

function calcularVariacao(atual, anterior) {
    if (!anterior) return "Sem base anterior";
    const valor = ((atual - anterior) / anterior) * 100;
    return `${valor >= 0 ? "+" : ""}${valor.toFixed(1).replace(".", ",")}% vs. período anterior`;
}

function carregarFiltros() {
    const coordenadores = new Map();
    const supervisores = new Map();

    estruturaCompleta.forEach(item => {
        coordenadores.set(item.coordenador, item.coordenadorOriginal);
        supervisores.set(item.supervisor, item.supervisorOriginal);
    });

    preencherSelect("filtroCoordenador", coordenadores);
    preencherSelect("filtroSupervisor", supervisores);
}

function preencherSelect(id, valores) {
    const select = document.getElementById(id);
    select.innerHTML = '<option value="">Todos</option>';

    [...valores.entries()]
        .sort((a, b) => a[1].localeCompare(b[1], "pt-BR"))
        .forEach(([valor, texto]) => {
            const option = document.createElement("option");
            option.value = valor;
            option.textContent = texto;
            select.appendChild(option);
        });
}

function mostrarErro(mensagem) {
    const elemento = document.getElementById("mensagemErro");
    elemento.hidden = false;
    elemento.textContent = mensagem;
}

async function iniciar() {
    try {
        const resposta = await fetch(CONFIGURACAO.arquivo, { cache: "no-store" });
        if (!resposta.ok) throw new Error(`Erro ${resposta.status} ao carregar o Excel.`);

        const ultimaAtualizacao = resposta.headers.get("Last-Modified");
        document.getElementById("dataAtualizacao").textContent = ultimaAtualizacao
            ? new Date(ultimaAtualizacao).toLocaleString("pt-BR")
            : new Date().toLocaleString("pt-BR");

        const workbook = XLSX.read(await resposta.arrayBuffer(), { type: "array" });
        const abaRit = workbook.Sheets[CONFIGURACAO.abaRit];
        const abaEstrutura = workbook.Sheets[CONFIGURACAO.abaEstrutura];

        if (!abaRit) throw new Error(`A aba ${CONFIGURACAO.abaRit} não foi encontrada.`);
        if (!abaEstrutura) throw new Error(`A aba ${CONFIGURACAO.abaEstrutura} não foi encontrada.`);

        dadosRit = prepararDadosRit(XLSX.utils.sheet_to_json(abaRit));
        estruturaCompleta = prepararEstrutura(XLSX.utils.sheet_to_json(abaEstrutura));

        carregarFiltros();
        montarTabela();
    } catch (erro) {
        console.error(erro);
        mostrarErro("Não foi possível carregar o painel. Verifique o arquivo dados.xlsx e as abas RIT_ICOMON e ESTRUTURA_ICOMON.");
    }
}

document.addEventListener("DOMContentLoaded", () => {
    document.getElementById("filtroCoordenador").addEventListener("change", () => {
        atualizarFiltroSupervisor();
        montarTabela();
    });
    document.getElementById("filtroSupervisor").addEventListener("change", montarTabela);
    document.getElementById("filtroData").addEventListener("change", event => {
        dataReferenciaChave = isoParaChave(event.currentTarget.value);
        montarTabela();
    });
    document.getElementById("btnLimpar").addEventListener("click", () => {
        document.getElementById("filtroCoordenador").value = "";
        atualizarFiltroSupervisor();
        document.getElementById("filtroSupervisor").value = "";
        configurarFiltroData();
        montarTabela();
    });

    document.getElementById("btnExpandir").addEventListener("click", event => {
        equipesRecolhidas = !equipesRecolhidas;
        event.currentTarget.textContent = equipesRecolhidas
            ? "Expandir equipes"
            : "Recolher equipes";
        montarTabela();
    });

    document.getElementById("btnFullscreen").addEventListener("click", async () => {
        if (!document.fullscreenElement) {
            await document.documentElement.requestFullscreen();
        } else {
            await document.exitFullscreen();
        }
    });

    configurarFiltroData();
    iniciar();
});
