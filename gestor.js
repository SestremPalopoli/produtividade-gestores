const INDICADORES_GESTOR = [
    { nome: "RIT", aba: "RIT_ACUM", celulaMeta: "I7" },
    { nome: "Falha de M.O", aba: "FMO_ACUM", celulaMeta: "I9" },
    { nome: "Checklist", aba: "CHECK_ACUM", celulaMeta: "I8" },
    { nome: "Degradação Rede", aba: "DEG_ACUM", celulaMeta: "I10" },
    { nome: "Lançamento SD", aba: "SD_ACUM", celulaMeta: "I11" },
    { nome: "Material Indevido", aba: "FISC_ACUM", celulaMeta: "I7" },
    { nome: "Clean-Up", aba: "CLEAN-UP_ACUM", celulaMeta: "I12" },
    { nome: "Liberado Comercial", aba: "LIBERADO COMERCIAL", tipo: "tres-meses" }
];

function normalizarTexto(texto) {
    return String(texto || "").normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "").trim().toUpperCase();
}

function escaparHtml(texto) {
    return String(texto).replaceAll("&", "&amp;").replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#039;");
}

function formatarDataExcel(valor) {
    if (valor === null || valor === undefined || valor === "") return "";
    if (typeof valor === "number") return XLSX.SSF.format("dd/mm/yyyy", valor);
    const texto = String(valor).trim();
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(texto)) return texto;
    const data = new Date(texto);
    return Number.isNaN(data.getTime()) ? texto : data.toLocaleDateString("pt-BR");
}

function obterMesLinha(linha) {
    const mes = Number(linha["MÊS"]);
    if (mes >= 1 && mes <= 12) return mes;
    const partes = formatarDataExcel(linha["DATA"]).split("/");
    return partes.length === 3 ? Number(partes[1]) : 0;
}

function obterChaveMesAno(linha, colunaData = "DATA", colunaMes = "MÊS") {
    const partes = formatarDataExcel(linha[colunaData]).split("/").map(Number);
    if (partes.length === 3 && partes.every(Number.isFinite)) {
        return { chave: partes[2] * 100 + partes[1], mes: partes[1], ano: partes[2] };
    }
    const mes = Number(linha[colunaMes]) || obterMesLinha(linha);
    const ano = new Date().getFullYear();
    return mes >= 1 && mes <= 12 ? { chave: ano * 100 + mes, mes, ano } : null;
}

function obterUltimosTresMeses(dados, colunaData = "DATA", colunaMes = "MÊS") {
    const mapa = new Map();
    dados.forEach(linha => {
        const periodo = obterChaveMesAno(linha, colunaData, colunaMes);
        if (periodo) mapa.set(periodo.chave, periodo);
    });
    return [...mapa.values()].sort((a, b) => a.chave - b.chave).slice(-3);
}

function nomeMesCurto(numero) {
    return ["-", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"][numero] || "-";
}

function obterFarol(percentual) {
    if (percentual >= 100) return { simbolo: "🟢", classe: "verde" };
    if (percentual >= 75) return { simbolo: "🟡", classe: "amarelo" };
    return { simbolo: "🔴", classe: "vermelho" };
}

function nomeMes(numero) {
    return ["Todos", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"][numero] || "Todos";
}

function obterParametros() {
    const p = new URLSearchParams(window.location.search);
    return { gestor: p.get("gestor") || "", mes: Number(p.get("mes") || 0), data: p.get("data") || "" };
}

function contarGestor(dados, gestor, mes, data) {
    let totalMes = 0;
    let totalDia = 0;
    const nome = normalizarTexto(gestor);
    dados.forEach(linha => {
        if (normalizarTexto(linha["VISTORIADOR"]) !== nome) return;
        if (mes === 0 || obterMesLinha(linha) === mes) totalMes++;
        if (data && formatarDataExcel(linha["DATA"]) === data) totalDia++;
    });
    return { totalMes, totalDia };
}

function criarCardGestor(indicador, resultado, meta) {
    const percentual = meta > 0 ? (resultado.totalMes / meta) * 100 : 0;
    const farol = obterFarol(percentual);
    const metaExibida = ["Falha de M.O", "Clean-Up"].includes(indicador.nome)
        ? Math.round(meta)
        : meta;
    return `
        <article class="card-gestor-indicador">
            <header><h2>${escaparHtml(indicador.nome)}</h2></header>
            <div class="valor-principal">${resultado.totalMes}</div>
            <div class="metricas-gestor">
                <div><span>Na data</span><strong>${resultado.totalDia}</strong></div>
                <div><span>Meta</span><strong>${metaExibida}</strong></div>
            </div>
            <div class="desempenho-gestor ${farol.classe}">${farol.simbolo} ${percentual.toFixed(1)}%</div>
        </article>`;
}

function criarCardLiberadoGestor(indicador, dados, gestor, dataSelecionada) {
    const dadosGestor = dados.filter(
        linha => normalizarTexto(linha["GESTOR"]) === normalizarTexto(gestor)
    );
    const periodos = obterUltimosTresMeses(dados, "Data visita Gestor", "MÊS");
    while (periodos.length < 3) periodos.unshift(null);
    const valores = periodos.map(periodo => periodo
        ? dadosGestor.filter(
            linha => obterChaveMesAno(linha, "Data visita Gestor", "MÊS")?.chave === periodo.chave
        ).length
        : 0
    );
    const totalDia = dadosGestor.filter(
        linha => dataSelecionada && formatarDataExcel(linha["Data visita Gestor"]) === dataSelecionada
    ).length;

    return `
        <article class="card-gestor-indicador card-liberado-gestor">
            <header><h2>${escaparHtml(indicador.nome)}</h2></header>
            <div class="meses-liberado-gestor">
                ${periodos.map((periodo, i) => `
                    <div><span>${periodo ? nomeMesCurto(periodo.mes) : "-"}</span><strong>${valores[i]}</strong></div>
                `).join("")}
            </div>
            <div class="data-liberado-gestor">
                <span>${dataSelecionada ? dataSelecionada.substring(0, 5) : "Na data"}</span>
                <strong>${totalDia}</strong>
            </div>
        </article>`;
}

function mostrarErroGestor(mensagem) {
    const erro = document.getElementById("mensagemErroGestor");
    erro.hidden = false;
    erro.textContent = mensagem;
}

async function carregarPainelGestor() {
    const { gestor, mes, data } = obterParametros();
    if (!gestor) return mostrarErroGestor("Nenhum gestor foi informado.");

    document.getElementById("nomeGestor").textContent = gestor;
    document.getElementById("mesResumo").textContent = nomeMes(mes);
    document.getElementById("dataResumo").textContent = data || "Não selecionada";
    document.getElementById("periodoGestor").textContent = `Resultados individuais | ${nomeMes(mes)} | ${data || "sem data selecionada"}`;
    document.title = `${gestor} | Indicadores`;

    try {
        const resposta = await fetch("dados.xlsx", { cache: "no-store" });
        if (!resposta.ok) throw new Error(`Erro ${resposta.status}.`);
        const workbook = XLSX.read(await resposta.arrayBuffer(), { type: "array" });
        const calendario = workbook.Sheets["CALENDÁRIO"];
        if (!calendario) throw new Error('A aba "CALENDÁRIO" não foi encontrada.');

        const cards = INDICADORES_GESTOR.map(indicador => {
            const aba = workbook.Sheets[indicador.aba];
            const dados = aba ? XLSX.utils.sheet_to_json(aba) : [];
            if (indicador.tipo === "tres-meses") {
                return criarCardLiberadoGestor(indicador, dados, gestor, data);
            }
            const meta = Number(calendario[indicador.celulaMeta]?.v || 0);
            return criarCardGestor(indicador, contarGestor(dados, gestor, mes, data), meta);
        });
        document.getElementById("gradeGestor").innerHTML = cards.join("");
    } catch (erro) {
        console.error(erro);
        mostrarErroGestor("Não foi possível carregar os indicadores do gestor.");
    }
}

document.addEventListener("DOMContentLoaded", () => {
    document.getElementById("btnFechar").addEventListener("click", () => window.close());
    carregarPainelGestor();
});
