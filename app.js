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

const btnPostos = document.getElementById("btnPostos");


/*==================================================
    BASES
==================================================*/

let textoPDF = "";

let maquinistas = [];

let operadores = [];

let resultadoFinal = [];

let bancoCPTM = [];

let bancoTrivia = [];

let operadoresPostos = [];

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
btnPostos.addEventListener("click", gerarPostos);

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

    mostrarListaPDF();

    setTimeout(() => {

        window.scrollTo({
            top: document.body.scrollHeight,
            behavior: "smooth"
        });

    }, 100);

}
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

console.table(
    operadores.map(op => ({
        nome: op.nome,
        local: op.local,
        maquinista: op.maquinista,
        hora: op.horaMaquinista
    }))
);

}  
/*==================================================
    CRUZAR DADOS
==================================================*/
/*==================================================
    CRUZAR DADOS
==================================================*/

function cruzarDados(){

    if(!operadores.length){

        alert("Carregue a Gestão de Escala.");

        return;

    }

    resultado.value = "";

    resultadoFinal = [];

    operadores.forEach(operador=>{

        resultadoFinal.push({

            cptm: buscarNomeGuerraCPTM(operador.maquinista),

            hora: operador.horaMaquinista || operador.entrada,

            trivia: buscarNomeGuerraTrivia(operador.nome),

            local: operador.local,

            situacao: operador.situacao,

            observacoes: operador.observacoes

        });

    });

    resultadoFinal.sort((a,b)=>{

        if(a.hora !== b.hora){
            return a.hora.localeCompare(b.hora);
        }

        return a.cptm.localeCompare(b.cptm);

    });

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
    LEITURA EXCEL (UNIVERSAL)
==================================================*/

async function lerExcel(file){

    const bytes = await file.arrayBuffer();

    const workbook = XLSX.read(bytes,{type:"array"});

    const aba =
        workbook.SheetNames.find(nome =>
            ["MANHÃ","MANHA","TARDE","NOITE","GERAL"]
                .includes(nome.toUpperCase().trim())
        ) || workbook.SheetNames[0];

    const sheet = workbook.Sheets[aba];

    const dados = XLSX.utils.sheet_to_json(sheet,{
        header:1,
        defval:""
    });

    /*---------------------------------------
      Localiza automaticamente o cabeçalho
    ---------------------------------------*/

    let linhaCabecalho = -1;

    for(let i=0;i<dados.length;i++){

        const linha = dados[i].map(v=>String(v||"").trim().toUpperCase());

        if(
            linha.includes("NOME COMPLETO") &&
            linha.includes("LOCAL")
        ){
            linhaCabecalho = i;
            break;
        }

    }

    if(linhaCabecalho === -1){

        alert("Cabeçalho da Gestão não encontrado.");

        return;

    }

    const cab = dados[linhaCabecalho]
        .map(v=>String(v||"").trim().toUpperCase());

    const localizarColuna = (...nomes) =>
        cab.findIndex(col => nomes.includes(col));

    const idxNome = localizarColuna("NOME COMPLETO");

    const idxLocal = localizarColuna("LOCAL");

    const idxEntradaHora = localizarColuna(
        "ENTRADA HORA",
        "HORA ENTRADA"
    );

    const idxEntrada = localizarColuna("ENTRADA");

    const idxMaquinista = localizarColuna("MAQUINISTA");

    const idxObs = localizarColuna(
        "OBSERVAÇÕES",
        "OBSERVACOES"
    );

    operadores = [];
    operadoresPostos = [];

    for(let i=linhaCabecalho+1;i<dados.length;i++){

        const linha = dados[i];

        if(!linha.length) continue;

        const nomeCompleto = String(linha[idxNome] || "").trim();

        if(!nomeCompleto) continue;

        const local = idxLocal >= 0
            ? String(linha[idxLocal] || "").trim()
            : "";

        const entrada = idxEntradaHora >= 0
            ? formatarHora(linha[idxEntradaHora])
            : "";

        const situacao = idxEntrada >= 0
            ? String(linha[idxEntrada] || "").trim()
            : "";

        const observacoes = idxObs >= 0
            ? String(linha[idxObs] || "").trim()
            : "";

        const texto = idxMaquinista >= 0
            ? String(linha[idxMaquinista] || "").trim()
            : "";

        operadoresPostos.push({

            nome: nomeCompleto,

            local,

            hora: entrada

        });

        if(!texto) continue;

        if(
            /^APOIO/i.test(texto) ||
            /^LOCOMOTIVA/i.test(texto) ||
            /^EQUIPE LOCOMOTIVA/i.test(texto) ||
            /^MQT/i.test(texto)
        ){
            continue;
        }

        const regexNovo = /^(.*?)(?:\s+(\d{2}:\d{2}|\d{4}))?$/;

        const regexAntigo = /([A-ZÀ-Ú'. ]+?)\s+(\d{4})/gi;

        if(texto.includes("/")){

            let item;

            while((item = regexAntigo.exec(texto)) !== null){

                operadores.push({

                    nome: nomeCompleto,

                    nomeCompleto,

                    local,

                    entrada,

                    situacao,

                    maquinista: item[1].trim(),

                    horaMaquinista: item[2],

                    observacoes

                });

            }

        }else{

            let maquinista = texto;
            let hora = "";

            const partes = texto.match(regexNovo);

            if(partes){

                maquinista = partes[1].trim();

                if(partes[2]){
                    hora = partes[2].replace(":","");
                }

            }

            operadores.push({

                nome: nomeCompleto,

                nomeCompleto,

                local,

                entrada,

                situacao,

                maquinista,

                horaMaquinista: hora,

                observacoes

            });

        }

    }

    resultado.value =
`Gestão carregada.

Aba: ${aba}

Operadores: ${operadores.length}

Postos: ${operadoresPostos.length}`;

    console.table(operadoresPostos);
    console.table(operadores);

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

/*==================================================
    GERAR POSTOS
==================================================*/

function gerarPostos(){

    if(!operadoresPostos.length){
        alert("Carregue a Gestão de Escala.");
        return;
    }

    const grupos = {};

    operadoresPostos.forEach(op=>{

        if(!grupos[op.local]){
            grupos[op.local] = [];
        }

        grupos[op.local].push(op);

    });

    resultado.value = "";

    Object.keys(grupos)
        .sort()
        .forEach(local=>{

            resultado.value += local + "\n";

            grupos[local]
                .sort((a,b)=>a.hora.localeCompare(b.hora))
                .forEach(op=>{

                    resultado.value +=
`${buscarNomeGuerraTrivia(op.nome)} ${op.hora}\n`;

                });

            resultado.value += "\n";

        });

}


function formatarHora(valor){

    if(valor == null || valor === "")
        return "";

    // Hora vinda do Excel (0.25, 0.5...)
    if(typeof valor === "number"){

        const total = Math.round(valor * 86400);

        const h = Math.floor(total / 3600);
        const m = Math.floor((total % 3600) / 60);

        return String(h).padStart(2,"0") +
               String(m).padStart(2,"0");
    }

    // Date do JavaScript
    if(valor instanceof Date){

        return String(valor.getHours()).padStart(2,"0") +
               String(valor.getMinutes()).padStart(2,"0");
    }

    // Objeto Time lido pelo SheetJS
    if(typeof valor === "object"){

        if("h" in valor && "m" in valor){

            return String(valor.h).padStart(2,"0") +
                   String(valor.m).padStart(2,"0");
        }

    }

    const texto = String(valor).trim();

    const hhmm = texto.match(/^(\d{1,2}):(\d{2})/);

    if(hhmm){

        return hhmm[1].padStart(2,"0") + hhmm[2];
    }

    return texto.replace(":","");
}