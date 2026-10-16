const INDICADORES_GESTOR = [
    { nome: "RIT", aba: "RIT_ACUM", celulaMeta: "I7" },
    { nome: "Falha de M.O", aba: "FMO_ACUM", celulaMeta: "I9" },
    { nome: "Checklist", aba: "CHECK_ACUM", celulaMeta: "I8" },
    { nome: "Degradação Rede", aba: "DEG_ACUM", celulaMeta: "I10" },
    { nome: "Lançamento SD", aba: "SD_ACUM", celulaMeta: "I11" },
    { nome: "Material Indevido", aba: "FISC_ACUM", celulaMeta: "I7" },
    { nome: "Clean-Up", aba: "CLEAN-UP_ACUM", celulaMeta: "I12" }
];

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

function obterFarol(percentual) {
    if (percentual >= 100) return { simbolo: "🟢", classe: "verde" };
    if (percentual >= 75) return { simbolo: "🟡", classe: "amarelo" };
    return { simbolo: "🔴", classe: "vermelho" };
}

function nomeMes(numero) {
    const meses = [
        "Todos", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ];
    return meses[numero] || "Todos";
}

function obterParametros() {
    const parametros = new URLSearchParams(window.location.search);
    return {
        gestor: parametros.get("gestor") || "",
        mes: Number(parametros.get("mes") || 0),
        data: parametros.get("data") || ""
    };
}

function contarGestor(dados, gestor, mesSelecionado, dataSelecionada) {
    let totalMes = 0;
    let totalDia = 0;
    const gestorNormalizado = normalizarTexto(gestor);

    dados.forEach(linha => {
        if (normalizarTexto(linha["VISTORIADOR"]) !== gestorNormalizado) return;

        if (mesSelecionado === 0 || obterMesLinha(linha) === mesSelecionado) totalMes++;
        if (dataSelecionada && formatarDataExcel(linha["DATA"]) === dataSelecionada) totalDia++;
    });

    return { totalMes, totalDia };
}

function criarCardGestor(indicador, resultado, meta) {
    const percentual = meta > 0 ? (resultado.totalMes / meta) * 100 : 0;
    const farol = obterFarol(percentual);
    const metaExibida = indicador.nome === "Falha de M.O" ? Math.round(meta) : meta;

    return `
        <article class="card-gestor-indicador">
            <header><h2>${escaparHtml(indicador.nome)}</h2></header>
            <div class="valor-principal">${resultado.totalMes}</div>
            <div class="metricas-gestor">
                <div><span>Na data</span><strong>${resultado.totalDia}</strong></div>
                <div><span>Meta</span><strong>${metaExibida}</strong></div>
            </div>
            <div class="desempenho-gestor ${farol.classe}">
                ${farol.simbolo} ${percentual.toFixed(1)}%
            </div>
        </article>
    `;
}

function criarCardReservadoGestor() {
    return `
        <article class="card-gestor-indicador card-gestor-reservado">
            <div>
                <span class="icone-reservado">＋</span>
                <h2>Futuro indicador</h2>
                <p>Espaço reservado para o oitavo indicador.</p>
            </div>
        </article>
    `;
}

function mostrarErroGestor(mensagem) {
    const erro = document.getElementById("mensagemErroGestor");
    erro.hidden = false;
    erro.textContent = mensagem;
}

async function carregarPainelGestor() {
    const { gestor, mes, data } = obterParametros();

    if (!gestor) {
        mostrarErroGestor("Nenhum gestor foi informado.");
        return;
    }

    document.getElementById("nomeGestor").textContent = gestor;
    document.getElementById("mesResumo").textContent = nomeMes(mes);
    document.getElementById("dataResumo").textContent = data || "Não selecionada";
    document.getElementById("periodoGestor").textContent =
        `Resultados individuais | ${nomeMes(mes)} | ${data || "sem data selecionada"}`;
    document.title = `${gestor} | Indicadores`;

    try {
        const resposta = await fetch("dados.xlsx", { cache: "no-store" });
        if (!resposta.ok) throw new Error(`Erro ${resposta.status} ao carregar dados.xlsx.`);

        const arquivo = await resposta.arrayBuffer();
        const workbook = XLSX.read(arquivo, { type: "array" });
        const calendario = workbook.Sheets["CALENDÁRIO"];
        if (!calendario) throw new Error('A aba "CALENDÁRIO" não foi encontrada.');

        const cards = INDICADORES_GESTOR.map(indicador => {
            const aba = workbook.Sheets[indicador.aba];
            const dados = aba ? XLSX.utils.sheet_to_json(aba) : [];
            const meta = Number(calendario[indicador.celulaMeta]?.v || 0);
            const resultado = contarGestor(dados, gestor, mes, data);
            return criarCardGestor(indicador, resultado, meta);
        });

        document.getElementById("gradeGestor").innerHTML =
            cards.join("") + criarCardReservadoGestor();
    } catch (erro) {
        console.error(erro);
        mostrarErroGestor("Não foi possível carregar os indicadores do gestor.");
    }
}

document.addEventListener("DOMContentLoaded", () => {
    document.getElementById("btnFechar").addEventListener("click", () => window.close());
    carregarPainelGestor();
});
