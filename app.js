/*==================================================
    GERADOR CPTM x TRIVIA
==================================================*/

pdfjsLib.GlobalWorkerOptions.workerSrc =
"https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js";

/*==================================================
    ELEMENTOS
==================================================*/

const pdfInput = document.getElementById("pdf");
const excelInput = document.getElementById("excel");

const btnPDF = document.getElementById("btnPDF");
const btnExcel = document.getElementById("btnExcel");
const btnCruzar = document.getElementById("btnCruzar");

const btnCopiar = document.getElementById("copiar");
const btnLimpar = document.getElementById("limpar");

const resultado = document.getElementById("resultado");

const bancoInput = document.getElementById("banco");
const btnBanco = document.getElementById("btnBanco");


/*==================================================
    BASES
==================================================*/

let textoPDF = "";

let maquinistas = [];

let operadores = [];

let resultadoFinal = [];

let bancoCPTM = [];

let bancoTrivia = [];

/*==================================================
    EVENTOS
==================================================*/

btnPDF.addEventListener("click", gerarListaPDF);

btnExcel.addEventListener("click", carregarGestao);

btnCruzar.addEventListener("click", cruzarDados);

btnCopiar.addEventListener("click", copiarResultado);

btnLimpar.addEventListener("click", limparTudo);

btnBanco.addEventListener("click", importarBanco);

bancoCPTM = JSON.parse(localStorage.getItem("bancoCPTM")) || [];
bancoTrivia = JSON.parse(localStorage.getItem("bancoTrivia")) || [];

/*==================================================
    GERAR LISTA PDF
==================================================*/

async function gerarListaPDF(){

    if(!pdfInput.files.length){

        alert("Selecione o PDF.");

        return;

    }

    resultado.value = "Lendo PDF...\n";

    textoPDF = "";

    maquinistas = [];

    await lerPDF(pdfInput.files[0]);
console.log(textoPDF.substring(0,1500));

mostrarListaPDF();

window.scrollTo({
    top: document.body.scrollHeight,
    behavior: "smooth"
});

/*==================================================
    CARREGAR GESTÃO
==================================================*/

async function carregarGestao(){

    if(!excelInput.files.length){

        alert("Selecione a Gestão de Escala.");

        return;

    }

    operadores = [];

    await lerExcel(excelInput.files[0]);

    resultado.value =
`Gestão carregada.

Operadores encontrados: ${operadores.length}`;

    console.table(operadores);

}

/*==================================================
    CRUZAR DADOS
==================================================*/
function cruzarDados(){

    resultado.value = "";
    resultadoFinal = [];

    for(const operador of operadores){

        const maquinistaPDF = maquinistas.find(m =>
            normalizarNome(m.nome) === normalizarNome(operador.maquinista)
        );

        resultadoFinal.push({

            cptm: buscarNomeGuerraCPTM(operador.maquinista),

            hora: maquinistaPDF
                ? maquinistaPDF.entrada
                : operador.horaMaquinista,

            trivia: buscarNomeGuerraTrivia(operador.nome)

        });

    }

    resultadoFinal.forEach(item=>{

        resultado.value +=
`${item.cptm} ${item.hora} / ${item.trivia}\n`;

    });

}
/*==================================================
    LEITURA PDF
==================================================*/

async function lerPDF(file){

    const bytes = await file.arrayBuffer();

    const pdf = await pdfjsLib.getDocument({
        data: bytes
    }).promise;

    maquinistas = [];

    for(let pagina=1; pagina<=pdf.numPages; pagina++){

        const page = await pdf.getPage(pagina);

        const content = await page.getTextContent();

        const linhas = {};

        content.items.forEach(item=>{

            const y = Math.round(item.transform[5]);

            if(!linhas[y]){
                linhas[y] = [];
            }

            linhas[y].push(item);

        });

        Object.keys(linhas)
            .sort((a,b)=>b-a)
            .forEach(y=>{

                const linha = linhas[y]
                    .sort((a,b)=>a.transform[4]-b.transform[4])
                    .map(i=>i.str)
                    .join(" ")
                    .replace(/\s+/g," ")
                    .trim();

                processarLinhaPDF(linha);

            });

    }

    console.table(maquinistas);

}


function processarLinhaPDF(linha){

    const match = linha.match(
        /^(.+?)\s+([A-Z]{2}\d{3})\s+(.+?)\s+(\d{2}:\d{2})/
    );

    if(!match) return;

    const posto = match[1].trim();
    const escala = match[2].trim();
    const nome = match[3].trim();
    const entrada = match[4].replace(":","");

    if(
        nome === "" ||
        nome.startsWith("APOIO") ||
        nome.startsWith("TREIN.")
    ){
        return;
    }

    maquinistas.push({

        posto,
        escala,
        nome,
        entrada

    });

}
/*==================================================
    LEITURA EXCEL
==================================================*/

async function lerExcel(file){

    const bytes = await file.arrayBuffer();

    const workbook = XLSX.read(bytes);

    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    const dados = XLSX.utils.sheet_to_json(sheet,{
        range: 9,
        defval: ""
    });

    operadores = [];

    dados.forEach(linha=>{

        const textoMaquinista = String(
            linha["Maquinista"] ||
            linha["MAQUINISTA"] ||
            ""
        ).trim();

// Ignora linhas que não possuem um maquinista CPTM válido
if(
    !/\d{4}$/.test(textoMaquinista) ||
    textoMaquinista.startsWith("APOIO") ||
    textoMaquinista.startsWith("LOCOMOTIVA") ||
    textoMaquinista.startsWith("MQT ")
){
    return;
}

        let nomeMaquinista = textoMaquinista;
        let horaMaquinista = "";

        const hora = textoMaquinista.match(/(\d{4})$/);

        if(hora){

            horaMaquinista = hora[1];

            nomeMaquinista = textoMaquinista
                .replace(/\d{4}$/,"")
                .trim();

        }

        operadores.push({

            nome:
                linha["Nome"] ||
                linha["NOME"] ||
                linha["NOME COMPLETO"] ||
                "",

            maquinista: nomeMaquinista,

            horaMaquinista: horaMaquinista,

            entrada:
                String(
                    linha["Entrada"] ||
                    linha["ENTRADA"] ||
                    linha["ENTRADA HORA"] ||
                    ""
                ).replace(":",""),

            local:
                linha["Local"] ||
                linha["LOCAL"] ||
                ""

        });

    });

    console.table(operadores.slice(0,10));

}
/*==================================================
    COPIAR
==================================================*/

function copiarResultado(){

    if(!resultado.value.trim()){

        alert("Nenhum resultado.");

        return;

    }

    navigator.clipboard.writeText(resultado.value);

    alert("Resultado copiado.");

}

/*==================================================
    LIMPAR
==================================================*/

function limparTudo(){

    pdfInput.value = "";

    excelInput.value = "";

    resultado.value = "";

    textoPDF = "";

    maquinistas = [];

    operadores = [];

    resultadoFinal = [];

}

/*==================================================
    EXTRAIR MAQUINISTAS DO PDF
==================================================*/

function extrairMaquinistas(texto){

    maquinistas = [];

    texto = texto
        .replace(/\r/g," ")
        .replace(/\n/g," ")
        .replace(/\s+/g," ");

    const regex = /([0-9]{2}-[A-Z0-9 ]+(?:-[A-Z0-9 ]+)*)\s+([A-Z]{2}\d{3})\s+([A-ZÁÀÂÃÉÊÍÓÔÕÚÇ.'\- ]+?)\s+(\d{2}:\d{2})/gi;

    let item;

    while((item = regex.exec(texto)) !== null){

        const nome = item[3].trim();

        if(
            !nome ||
            nome.startsWith("APOIO ESC") ||
            nome.startsWith("TREIN.") ||
            /^\d+$/.test(nome)
        ){
            continue;
        }

        maquinistas.push({

            posto: item[1].trim(),
            escala: item[2].trim(),
            nome,
            entrada: item[4].replace(":","")

        });

    }

    console.table(maquinistas);

}
/*==================================================
    MOSTRAR LISTA DO PDF
==================================================*/

function mostrarListaPDF(){

    resultado.value = "";

    maquinistas.forEach(m=>{

        resultado.value += `${m.nome} ${m.entrada}\n`;

    });

}

/*==================================================
    IMPORTAR BANCO DE DADOS
==================================================*/

async function importarBanco(){

    if(!bancoInput.files.length){

        alert("Selecione o Banco de Dados.");

        return;

    }

    const bytes = await bancoInput.files[0].arrayBuffer();

    const workbook = XLSX.read(bytes);
console.log("CPTM:", bancoCPTM.length);
console.log("TRIVIA:", bancoTrivia.length);

console.table(bancoCPTM.slice(0,5));
    /*======================
        MQTS CPTM
    ======================*/

    bancoCPTM = [];

    const abaCPTM = workbook.Sheets["MQTS CPTM"];

    const dadosCPTM = XLSX.utils.sheet_to_json(abaCPTM,{
        header:1,
        defval:""
    });

    for(let i=1;i<dadosCPTM.length;i++){

        if(!dadosCPTM[i][0]) continue;

        bancoCPTM.push({

            nome:dadosCPTM[i][0].toString().trim().toUpperCase(),

            guerra:dadosCPTM[i][1].toString().trim()

        });

    }

    /*======================
        OPT TRIVIA
    ======================*/

    bancoTrivia = [];

    const abaTrivia = workbook.Sheets["OPT TRIVIA"];

    const dadosTrivia = XLSX.utils.sheet_to_json(abaTrivia,{
        header:1,
        defval:""
    });

    for(let i=1;i<dadosTrivia.length;i++){

        if(!dadosTrivia[i][0]) continue;

        bancoTrivia.push({

            nome:dadosTrivia[i][0].toString().trim().toUpperCase(),

            guerra:dadosTrivia[i][1].toString().trim()

        });

    }

    localStorage.setItem("bancoCPTM",JSON.stringify(bancoCPTM));
    localStorage.setItem("bancoTrivia",JSON.stringify(bancoTrivia));

    resultado.value =
`BANCO IMPORTADO COM SUCESSO

MQTS CPTM : ${bancoCPTM.length}

OPT TRIVIA : ${bancoTrivia.length}`;

}

/*==================================================
    NORMALIZAR NOME
==================================================*/

function normalizarNome(nome){

    return String(nome)
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/\s+/g, " ")
        .trim()
        .toUpperCase();

}

/*==================================================
    BUSCAR NOME DE GUERRA CPTM
==================================================*/
function buscarNomeGuerraCPTM(nome){

    const procurado = normalizarNome(nome);

    for(const operador of bancoCPTM){

        if(normalizarNome(operador.nome) === procurado){

            return operador.guerra;

        }

    }

    // Se não encontrou, retorna o nome original do PDF
    return nome;

}
/*==================================================
    BUSCAR NOME DE GUERRA TRIVIA
==================================================*/

function buscarNomeGuerraTrivia(nome){

    const procurado = normalizarNome(nome);

    for(const operador of bancoTrivia){

        if(normalizarNome(operador.nome) === procurado){

            return operador.guerra;

        }

    }

    // Se não encontrar, retorna o nome da Gestão
    return nome;

}
