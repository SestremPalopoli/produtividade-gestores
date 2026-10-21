const CONFIG = {
    arquivo: "dados.xlsx",
    aba: "Plano Preventivo",
    frentesEsperadas: ["Rede Limpa", "Modernização", "Ações Plano Preventivo"]
};

let frentes = [];
let mesesGlobais = [];
let graficoEvolucao = null;
let graficoBacklog = null;
let detalhesAbertos = true;

function normalizarTexto(valor) {
    return String(valor ?? "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .trim()
        .replace(/\s+/g, " ")
        .toUpperCase();
}

function numero(valor) {
    if (typeof valor === "number") return valor;
    const texto = String(valor ?? "").trim().replace(/\./g, "").replace(",", ".");
    const convertido = Number(texto);
    return Number.isFinite(convertido) ? convertido : 0;
}

function escaparHtml(valor) {
    return String(valor ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function formatarNumero(valor) {
    return Number(valor || 0).toLocaleString("pt-BR", { maximumFractionDigits: 1 });
}

function identificarMes(cabecalho) {
    const texto = normalizarTexto(cabecalho);
    const mapa = {
        JAN: 1, FEV: 2, MAR: 3, ABR: 4, MAI: 5, JUN: 6,
        JUL: 7, AGO: 8, SET: 9, OUT: 10, NOV: 11, DEZ: 12
    };
    const sigla = Object.keys(mapa).find(item => texto.startsWith(item));
    if (!sigla) return null;
    const anoEncontrado = texto.match(/(20)?\d{2}/)?.[0] || "";
    const ano = anoEncontrado.length === 2 ? 2000 + Number(anoEncontrado) : Number(anoEncontrado || 0);
    return { chave: `${ano || 0}-${String(mapa[sigla]).padStart(2, "0")}`, mes: mapa[sigla], ano, titulo: String(cabecalho).trim() };
}

function encontrarLinhasFrentes(matriz) {
    const encontrados = [];
    matriz.forEach((linha, indice) => {
        const primeira = normalizarTexto(linha[0]);
        const frente = CONFIG.frentesEsperadas.find(nome => normalizarTexto(nome) === primeira);
        if (frente) encontrados.push({ nome: frente, indice });
    });
    return encontrados;
}

function interpretarAba(matriz) {
    const marcadores = encontrarLinhasFrentes(matriz);
    if (!marcadores.length) throw new Error("As três frentes não foram encontradas na aba Plano Preventivo.");

    return marcadores.map((marcador, posicao) => {
        const fim = posicao + 1 < marcadores.length ? marcadores[posicao + 1].indice : matriz.length;
        const cabecalho = matriz[marcador.indice];
        const colunas = cabecalho.map((valor, indice) => ({ valor: String(valor ?? "").trim(), indice }));
        const backlog = colunas.find(item => normalizarTexto(item.valor) === "BACKLOG")?.indice;
        const anual = colunas.find(item => /^20\d{2}$/.test(item.valor))?.indice;
        const meses = colunas.map(item => ({ ...item, periodo: identificarMes(item.valor) })).filter(item => item.periodo);
        const atividades = [];

        for (let i = marcador.indice + 1; i < fim; i++) {
            const linha = matriz[i] || [];
            const nome = String(linha[0] ?? "").trim();
            if (!nome || normalizarTexto(nome) === "TOTAL") continue;
            const valoresMensais = {};
            meses.forEach(item => valoresMensais[item.periodo.chave] = numero(linha[item.indice]));
            atividades.push({
                nome,
                backlog: backlog === undefined ? 0 : numero(linha[backlog]),
                anual: anual === undefined ? 0 : numero(linha[anual]),
                meses: valoresMensais
            });
        }

        return { nome: marcador.nome, meses: meses.map(item => item.periodo), atividades };
    });
}

function atualizarMesesGlobais() {
    const mapa = new Map();
    frentes.forEach(frente => frente.meses.forEach(mes => mapa.set(mes.chave, mes)));
    mesesGlobais = [...mapa.values()].sort((a, b) => a.chave.localeCompare(b.chave));
}

function soma(frente, campo) {
    return frente.atividades.reduce((total, atividade) => total + Number(atividade[campo] || 0), 0);
}

function somaMes(frente, chave) {
    return frente.atividades.reduce((total, atividade) => total + Number(atividade.meses[chave] || 0), 0);
}

function atividadeDestaque(frente, campo, chaveMes) {
    return [...frente.atividades].sort((a, b) => {
        const valorA = campo === "mes" ? a.meses[chaveMes] || 0 : a[campo] || 0;
        const valorB = campo === "mes" ? b.meses[chaveMes] || 0 : b[campo] || 0;
        return valorB - valorA;
    })[0];
}

function frentesVisiveis() {
    const filtro = document.getElementById("filtroFrente").value;
    return filtro === "TODAS" ? frentes : frentes.filter(frente => frente.nome === filtro);
}

function criarCards() {
    const chaveMes = document.getElementById("filtroMes").value;
    const container = document.getElementById("cardsFrentes");
    container.innerHTML = frentesVisiveis().map((frente, indice) => {
        const backlog = soma(frente, "backlog");
        const anual = soma(frente, "anual");
        const atual = somaMes(frente, chaveMes);
        const destaque = atividadeDestaque(frente, "mes", chaveMes);
        const classes = ["frente-rede", "frente-modernizacao", "frente-acoes"];
        return `
            <article class="card-frente ${classes[frentes.indexOf(frente)]}">
                <header><h2>${escaparHtml(frente.nome)}</h2></header>
                <div class="card-numeros">
                    <div><span>Backlog</span><strong>${formatarNumero(backlog)}</strong></div>
                    <div><span>2026</span><strong>${formatarNumero(anual)}</strong></div>
                    <div><span>Mês</span><strong>${formatarNumero(atual)}</strong></div>
                </div>
                <p class="card-destaque">
                    Maior atividade no mês: <strong>${escaparHtml(destaque?.nome || "Sem movimento")}</strong>
                </p>
            </article>`;
    }).join("");
}

function montarGraficoEvolucao() {
    const selecionadas = frentesVisiveis();
    const ctx = document.getElementById("graficoEvolucao");
    graficoEvolucao?.destroy();
    const cores = ["#5b1fa6", "#00a46c", "#d43a8f"];
    graficoEvolucao = new Chart(ctx, {
        type: "line",
        data: {
            labels: mesesGlobais.map(mes => mes.titulo),
            datasets: selecionadas.map(frente => ({
                label: frente.nome,
                data: mesesGlobais.map(mes => somaMes(frente, mes.chave)),
                borderColor: cores[frentes.indexOf(frente)],
                backgroundColor: `${cores[frentes.indexOf(frente)]}22`,
                borderWidth: 3,
                pointRadius: 4,
                pointHoverRadius: 6,
                tension: 0.3,
                fill: true
            }))
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: "bottom" } },
            scales: { y: { beginAtZero: true, ticks: { precision: 0 } } }
        }
    });
}

function montarGraficoBacklog() {
    const atividades = frentesVisiveis()
        .flatMap(frente => frente.atividades.map(atividade => ({ ...atividade, frente: frente.nome })))
        .filter(item => item.backlog > 0)
        .sort((a, b) => b.backlog - a.backlog)
        .slice(0, 8);
    const ctx = document.getElementById("graficoBacklog");
    graficoBacklog?.destroy();
    graficoBacklog = new Chart(ctx, {
        type: "bar",
        data: {
            labels: atividades.map(item => item.nome),
            datasets: [{
                label: "Backlog",
                data: atividades.map(item => item.backlog),
                backgroundColor: ["#d43a8f", "#b83280", "#8a2ca8", "#6d28b5", "#5b1fa6", "#7439ba", "#8f59c7", "#aa7bd3"],
                borderRadius: 7
            }]
        },
        options: {
            indexAxis: "y",
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: { x: { beginAtZero: true, ticks: { precision: 0 } } }
        }
    });
}

function criarDetalhes() {
    const chaveMes = document.getElementById("filtroMes").value;
    document.getElementById("detalhesFrentes").innerHTML = frentesVisiveis().map(frente => `
        <article class="tabela-frente">
            <header><h3>${escaparHtml(frente.nome)}</h3></header>
            <div class="tabela-scroll">
                <table>
                    <thead><tr><th>Atividade</th><th>Backlog</th><th>2026</th><th>Mês</th></tr></thead>
                    <tbody>
                        ${frente.atividades.map(atividade => `
                            <tr>
                                <td title="${escaparHtml(atividade.nome)}">${escaparHtml(atividade.nome)}</td>
                                <td class="backlog">${formatarNumero(atividade.backlog)}</td>
                                <td>${formatarNumero(atividade.anual)}</td>
                                <td class="mes-atual">${formatarNumero(atividade.meses[chaveMes] || 0)}</td>
                            </tr>`).join("")}
                    </tbody>
                </table>
            </div>
        </article>`).join("");
}

function preencherFiltros() {
    const frente = document.getElementById("filtroFrente");
    frentes.forEach(item => frente.add(new Option(item.nome, item.nome)));
    const mes = document.getElementById("filtroMes");
    mesesGlobais.forEach(item => mes.add(new Option(item.titulo, item.chave)));
    if (mesesGlobais.length) mes.value = mesesGlobais.at(-1).chave;
}

function atualizarPainel() {
    criarCards();
    montarGraficoEvolucao();
    montarGraficoBacklog();
    criarDetalhes();
}

function mostrarErro(mensagem) {
    const elemento = document.getElementById("mensagemErro");
    elemento.hidden = false;
    elemento.textContent = mensagem;
}

async function iniciar() {
    try {
        const resposta = await fetch(CONFIG.arquivo, { cache: "no-store" });
        if (!resposta.ok) throw new Error(`Erro ${resposta.status} ao carregar dados.xlsx.`);
        const ultimaAtualizacao = resposta.headers.get("Last-Modified");
        document.getElementById("dataAtualizacao").textContent = ultimaAtualizacao
            ? new Date(ultimaAtualizacao).toLocaleString("pt-BR")
            : new Date().toLocaleString("pt-BR");
        const workbook = XLSX.read(await resposta.arrayBuffer(), { type: "array" });
        const aba = workbook.Sheets[CONFIG.aba];
        if (!aba) throw new Error(`A aba ${CONFIG.aba} não foi encontrada.`);
        const matriz = XLSX.utils.sheet_to_json(aba, { header: 1, defval: "", raw: true });
        frentes = interpretarAba(matriz);
        atualizarMesesGlobais();
        preencherFiltros();
        atualizarPainel();
    } catch (erro) {
        console.error(erro);
        mostrarErro("Não foi possível carregar o Plano Preventivo. Verifique a aba Plano Preventivo no arquivo dados.xlsx.");
    }
}

document.addEventListener("DOMContentLoaded", () => {
    document.getElementById("filtroFrente").addEventListener("change", atualizarPainel);
    document.getElementById("filtroMes").addEventListener("change", atualizarPainel);
    document.getElementById("btnTelaCheia").addEventListener("click", async () => {
        if (!document.fullscreenElement) await document.documentElement.requestFullscreen();
        else await document.exitFullscreen();
    });
    document.getElementById("btnAlternarDetalhes").addEventListener("click", event => {
        detalhesAbertos = !detalhesAbertos;
        document.getElementById("detalhesFrentes").classList.toggle("recolhido", !detalhesAbertos);
        event.currentTarget.textContent = detalhesAbertos ? "Recolher" : "Expandir";
    });
    iniciar();
});
