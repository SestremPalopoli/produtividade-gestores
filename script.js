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
    alert("SCRIPT NOVO");

    fetch("dados.xlsx")

        .then(response => {

            document.getElementById("data-atualizacao")
                .innerHTML = new Date().toLocaleString("pt-BR");

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

montarTabelaGestores(dadosTabela);

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

const gestor =
    nome
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
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

console.log("DEG PRIMEIRA LINHA:", dadosDeg[0]);

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
function montarTabelaGestores(dados) {

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
            <th>Jul/26</th>
            <th>Ago/26</th>
            <th>Set/26</th>
            <th>${dataSelecionada}</th>
        </tr>
    `;

    tbody.innerHTML = "";

    let totalJul = 0;
    let totalAgo = 0;
    let totalSet = 0;
    let totalDia = 0;

    gestores.forEach(nome => {

        let jul = 0;
        let ago = 0;
        let set = 0;
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

                if (mes === 7) jul++;
                if (mes === 8) ago++;
                if (mes === 9) set++;

                const valorData = linha["DATA"];

if (valorData) {

    let dataExcel = "";

    if (typeof valorData === "number") {

        dataExcel = XLSX.SSF.format(
            "dd/mm/yyyy",
            valorData
        );

    } else {

        dataExcel = String(valorData)
            .trim();

    }

    if (
        vistoriador.includes("ANDRE") ||
        vistoriador.includes("PRISCILA")
    ) {

        console.log({
            gestor,
            valorData,
            dataExcel,
            dataSelecionada
        });

    }

    if (dataExcel === dataSelecionada) {
        dia++;
    }

}

        });

        totalJul += jul;
        totalAgo += ago;
        totalSet += set;
        totalDia += dia;

        tbody.innerHTML += `
            <tr>
                <td>${nome}</td>
                <td>${jul}</td>
                <td>${ago}</td>
                <td>${set}</td>
                <td>${dia}</td>
            </tr>
        `;

    });

    tbody.innerHTML += `
        <tr class="linha-total">
            <td>TOTAL</td>
            <td>${totalJul}</td>
            <td>${totalAgo}</td>
            <td>${totalSet}</td>
            <td>${totalDia}</td>
        </tr>
    `;
}
