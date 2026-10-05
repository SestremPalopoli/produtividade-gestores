function carregarFiltroDatas() {

    const select = document.getElementById("filtroData");

    select.innerHTML = "";

    const hoje = new Date();

    for (let i = 1; i <= 5; i++) {

        const data = new Date();

        data.setDate(hoje.getDate() - i);

        const dataFormatada =
            data.toLocaleDateString("pt-BR");

        select.innerHTML += `
            <option value="${dataFormatada}">
                ${dataFormatada}
            </option>
        `;
    }
}

window.onload = function () {

    carregarFiltroDatas();

    carregarIndicadores();

    document
        .getElementById("mesSelecionado")
        .addEventListener("change", carregarIndicadores);

    document
        .getElementById("filtroData")
        .addEventListener("change", carregarIndicadores);

    document
        .getElementById("filtroIndicador")
        .addEventListener("change", carregarIndicadores);

};

function carregarIndicadores() {

    fetch("dados.xlsx")

        .then(response => {

    const ultimaAtualizacao =
        response.headers.get("Last-Modified");

    if (ultimaAtualizacao) {

        const dataArquivo =
            new Date(ultimaAtualizacao);

        document.getElementById("data-atualizacao")
            .innerHTML =
            dataArquivo.toLocaleString("pt-BR");
    }

    return response.arrayBuffer();

})

        .then(data => {

            const workbook = XLSX.read(data, {
                type: "array"
            });
            const indicador =
    document.getElementById("filtroIndicador").value;

            const mesSelecionado =
                Number(document.getElementById("mesSelecionado").value);
                const gestoresPermitidos = [
    "AFONSO HENRIQUES MIGUEL MOREIRA",
    "ALEXANDRE MORAIS MEDEIROS DA SILVA",
    "ANDRE MENDES DE SOUSA",
    "CARMINE AVARESE FILHO",
    "DENIS REIS",
    "DEYVISON BRUNO CUTRIM FERREIRA",
    "DIEGO MENDONCA SANTOS",
    "HUMBERTO DE FARIA COSTA",
    "NILTON CAMPOS DE BARROS",
    "PRISCILA DE MELO SEVILHA ZANDELI"
];

            /* =========================
               RIT
            ========================== */

let nomeAba = "";
let tituloTabela = "";

switch(indicador){

    case "RIT":
        nomeAba = "RIT_ACUM";
        tituloTabela = "Performance RIT por Gestor";
        break;

    case "Falha de M.O":
        nomeAba = "FMO_ACUM";
        tituloTabela = "Performance Falha de M.O por Gestor";
        break;

    case "Checklist":
        nomeAba = "CHECK_ACUM";
        tituloTabela = "Performance Checklist por Gestor";
        break;

    case "Degradação Rede":
        nomeAba = "DEG_ACUM";
        tituloTabela = "Performance Degradação Rede por Gestor";
        break;

    case "Lançamento SD":
        nomeAba = "SD_ACUM";
        tituloTabela = "Performance Lançamento SD por Gestor";
        break;

    case "Material Indevido":
        nomeAba = "FISC_ACUM";
        tituloTabela = "Performance Material Indevido por Gestor";
        break;

    case "Clean-Up":
        nomeAba = "CLEAN-UP_ACUM";
        tituloTabela = "Performance Clean-Up por Gestor";
        break;
}


document.getElementById("tituloTabela").innerHTML =
    tituloTabela;

const abaSelecionada =
    workbook.Sheets[nomeAba];

const dadosTabela =
    XLSX.utils.sheet_to_json(abaSelecionada);

const abaCalendario =
    workbook.Sheets["CALENDÁRIO"];

let meta = 0;

switch (indicador) {

    case "RIT":
        meta = abaCalendario["I7"]?.v || 0;
        break;

    case "Checklist":
        meta = abaCalendario["I8"]?.v || 0;
        break;

    case "Falha de M.O":
        meta = abaCalendario["I9"]?.v || 0;
        break;

    case "Degradação Rede":
        meta = abaCalendario["I10"]?.v || 0;
        break;

    case "Lançamento SD":
        meta = abaCalendario["I11"]?.v || 0;
        break;

    case "Clean-Up":
        meta = abaCalendario["I12"]?.v || 0;
        break;

    case "Material Indevido":
        meta = abaCalendario["I7"]?.v || 0;
        break;
}


montarTabelaGestores(dadosTabela, meta);

const abaRit = workbook.Sheets["RIT_ACUM"];
const dadosRit = XLSX.utils.sheet_to_json(abaRit);

let totalRit = 0;

dadosRit.forEach(linha => {

    const mesLinha = Number(linha["MÊS"]);

    const vistoriador =
        String(linha["VISTORIADOR"] || "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .trim()
        .toUpperCase();

    if (
        gestoresPermitidos.includes(vistoriador) &&
        (
            mesSelecionado === 0 ||
            mesLinha === mesSelecionado
        )
    ) {
        totalRit++;
    }

});

document.getElementById("rit-total").innerHTML = totalRit;


/* FMO */

const abaFMO = workbook.Sheets["FMO_ACUM"];
const dadosFMO = XLSX.utils.sheet_to_json(abaFMO);

let totalFMO = 0;

dadosFMO.forEach(linha => {

    const mesLinha = Number(linha["MÊS"]);

    const vistoriador =
        String(linha["VISTORIADOR"] || "")
        .trim()
        .toUpperCase();

    if (
        gestoresPermitidos.includes(vistoriador) &&
        (
            mesSelecionado === 0 ||
            mesLinha === mesSelecionado
        )
    ) {
        totalFMO++;
    }

});

document.getElementById("fmo-total").innerHTML = totalFMO;


/* CHECKLIST */

const abaCheck = workbook.Sheets["CHECK_ACUM"];
const dadosCheck = XLSX.utils.sheet_to_json(abaCheck);

let totalCheck = 0;

dadosCheck.forEach(linha => {

    const mesLinha = Number(linha["MÊS"]);

    const vistoriador =
        String(linha["VISTORIADOR"] || "")
        .trim()
        .toUpperCase();

    if (
        gestoresPermitidos.includes(vistoriador) &&
        (
            mesSelecionado === 0 ||
            mesLinha === mesSelecionado
        )
    ) {
        totalCheck++;
    }

});

document.getElementById("checklist-total").innerHTML = totalCheck;


/* DEGRADAÇÃO */

const abaDeg = workbook.Sheets["DEG_ACUM"];
const dadosDeg = XLSX.utils.sheet_to_json(abaDeg);


let totalDeg = 0;

dadosDeg.forEach(linha => {

    const mesLinha = Number(linha["MÊS"]);

    const vistoriador =
        String(linha["VISTORIADOR"] || "")
        .trim()
        .toUpperCase();

    if (
        gestoresPermitidos.includes(vistoriador) &&
        (
            mesSelecionado === 0 ||
            mesLinha === mesSelecionado
        )
    ) {
        totalDeg++;
    }

});

document.getElementById("deg-total").innerHTML = totalDeg;


/* SD */

const abaSD = workbook.Sheets["SD_ACUM"];
const dadosSD = XLSX.utils.sheet_to_json(abaSD);

let totalSD = 0;

dadosSD.forEach(linha => {

    const mesLinha = Number(linha["MÊS"]);

    const vistoriador =
        String(linha["VISTORIADOR"] || "")
        .trim()
        .toUpperCase();

    if (
        gestoresPermitidos.includes(vistoriador) &&
        (
            mesSelecionado === 0 ||
            mesLinha === mesSelecionado
        )
    ) {
        totalSD++;
    }

});

document.getElementById("sd-total").innerHTML = totalSD;


/* MATERIAL INDEVIDO */

const abaFisc = workbook.Sheets["FISC_ACUM"];
const dadosFisc = XLSX.utils.sheet_to_json(abaFisc);

let totalFisc = 0;

dadosFisc.forEach(linha => {

    const mesLinha = Number(linha["MÊS"]);

    const vistoriador =
        String(linha["VISTORIADOR"] || "")
        .trim()
        .toUpperCase();

    if (
        gestoresPermitidos.includes(vistoriador) &&
        (
            mesSelecionado === 0 ||
            mesLinha === mesSelecionado
        )
    ) {
        totalFisc++;
    }

});

document.getElementById("fisc-total").innerHTML = totalFisc;


/* CLEAN-UP */

const abaCleanup = workbook.Sheets["CLEAN-UP_ACUM"];
const dadosCleanup = XLSX.utils.sheet_to_json(abaCleanup);

let totalCleanup = 0;

dadosCleanup.forEach(linha => {

    const mesLinha = Number(linha["MÊS"]);

    const vistoriador =
        String(linha["VISTORIADOR"] || "")
        .trim()
        .toUpperCase();

    if (
        gestoresPermitidos.includes(vistoriador) &&
        (
            mesSelecionado === 0 ||
            mesLinha === mesSelecionado
        )
    ) {
        totalCleanup++;
    }

});

document.getElementById("cleanup-total").innerHTML = totalCleanup;

        })

        .catch(error => {

            console.error(error);

            document.getElementById("rit-total").innerHTML = "ERRO";
            document.getElementById("fmo-total").innerHTML = "ERRO";
            document.getElementById("checklist-total").innerHTML = "ERRO";
            document.getElementById("deg-total").innerHTML = "ERRO";
            document.getElementById("sd-total").innerHTML = "ERRO";
            document.getElementById("fisc-total").innerHTML = "ERRO";
            document.getElementById("cleanup-total").innerHTML = "ERRO";

        });

}

function montarTabelaGestores(dados, meta) {

    const gestores = [
        "Afonso Henriques Miguel Moreira",
        "Alexandre Morais Medeiros Da Silva",
        "Andre Mendes De Sousa",
        "Carmine Avarese Filho",
        "Denis Reis",
        "Deyvison Bruno Cutrim Ferreira",
        "Diego Mendonca Santos",
        "Humberto De Faria Costa",
        "Nilton Campos De Barros",
        "Priscila De Melo Sevilha Zandeli"
    ];

    const dataSelecionada =
    document.getElementById("filtroData").value;

const thead =
    document.querySelector("#tabela-rit thead");

const tbody =
    document.querySelector("#tabela-rit tbody");


thead.innerHTML = `
    <tr>
        <th>Gestor</th>
        <th>Ago/26</th>
        <th>Set/26</th>
        <th>Out/26</th>
        <th>${dataSelecionada}</th>
        <th class="meta">Meta</th>
        <th>%</th>
    </tr>
`;

    tbody.innerHTML = "";

    let totalAgo = 0;
    let totalSet = 0;
    let totalOut = 0;
    let totalDia = 0;
    let totalMeta = 0;

    gestores.forEach(nome => {

        let ago = 0;
let set = 0;
let out = 0;
let dia = 0;

        dados.forEach(linha => {

            const vistoriador =
                String(linha["VISTORIADOR"] || "")
                .trim()
                .toUpperCase();

            const gestor =
                nome.toUpperCase();

            const mes =
                Number(linha["MÊS"]);

            if (vistoriador === gestor) {

                if (mes === 8) ago++;
                if (mes === 9) set++;
                if (mes === 10) out++;

                const valorData = linha["DATA"];

                if (valorData) {

                    let dataExcel = "";

                    if (typeof valorData === "number") {

                        dataExcel = XLSX.SSF.format(
                            "dd/mm/yyyy",
                            valorData
                        );

                    } else {

                        dataExcel = String(valorData).trim();

                    }

                    if (dataExcel === dataSelecionada) {
                        dia++;
                    }
                }
            }
        });

        totalAgo += ago;
totalSet += set;
totalOut += out;
        totalDia += dia;
        totalMeta += Number(meta);

const percentual =
    meta > 0
        ? ((out / meta) * 100)
        : 0;

        let farol = "";

if (percentual >= 100) {
    farol = "🟢";
}
else if (percentual >= 75) {
    farol = "🟡";
}
else {
    farol = "🔴";
}

        tbody.innerHTML += `
    <tr>
        <td>${nome}</td>
        <td>${ago}</td>
<td>${set}</td>
<td>${out}</td>
        <td>${dia}</td>
        <td class="meta">${meta}</td>
<td>${farol} ${percentual.toFixed(1)}%</td>
    </tr>
`;
    });

    const percentualTotal =
    totalMeta > 0
        ? ((totalOut / totalMeta) * 100).toFixed(1)
        : 0;

        let farolTotal = "";

if (percentualTotal >= 100) {
    farolTotal = "🟢";
}
else if (percentualTotal >= 75) {
    farolTotal = "🟡";
}
else {
    farolTotal = "🔴";
}

    tbody.innerHTML += `
        <tr class="linha-total">
    <td>TOTAL</td>
    <td>${totalAgo}</td>
<td>${totalSet}</td>
<td>${totalOut}</td>
    <td>${totalDia}</td>
    <td class="meta">${Number(totalMeta.toFixed(1))}</td>
    <td>${farolTotal} ${percentualTotal}%</td>
</tr>
    `;
}
document
    .getElementById("btnFullscreen")
    .addEventListener("click", () => {

        if (!document.fullscreenElement) {
            document.documentElement.requestFullscreen();
        } else {
            document.exitFullscreen();
        }

    });
