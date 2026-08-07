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
const btnRendicoes = document.getElementById("btnRendicoes");

const btnMonitoria = document.getElementById("btnMonitoria");

btnMonitoria.addEventListener("click", gerarMonitoria);

btnExcelMonitoria.onclick = exportarExcel;


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

let turnoAtual = "";

let operadoresMonitoria = [];

let operadoresApoio = [];

let operadoresIgnorados = [];

let operadoresSemMonitoria = [];

let maquinistasSemOperador = [];

let dadosExcel = [];

let vagasCPTM = [];

let ocorrenciasPDF = [];

let turnoCA = "";
/*==================================================
    EVENTOS
==================================================*/

//======================
// BANCO
//======================

btnBanco.addEventListener("click",()=>{

    bancoInput.click();

});

bancoInput.addEventListener("change",()=>{

    if(bancoInput.files.length){

        importarBanco();

        document.getElementById("statusBanco").textContent =
        "Carregado";

        document.getElementById("statusBanco").className =
        "fw-bold text-success";

    }

});

//======================
// PDF
//======================

btnPDF.addEventListener("click",()=>{

    pdfInput.click();

});

pdfInput.addEventListener("change",()=>{

    if(pdfInput.files.length){

        gerarListaPDF();

    }

});

//======================
// GESTÃO
//======================

btnExcel.addEventListener("click",()=>{

    excelInput.click();

});

excelInput.addEventListener("change",()=>{

    if(excelInput.files.length){

        carregarGestao();

        document.getElementById("statusGestao").textContent =
        "Carregada";

        document.getElementById("statusGestao").className =
        "fw-bold text-success";

    }

});

//======================
// PROCESSAMENTO
//======================

btnCruzar.addEventListener("click",cruzarDados);

btnPostos.addEventListener("click",gerarPostos);

btnRendicoes.addEventListener("click",gerarRendicoes);

btnMonitoria.addEventListener("click",gerarMonitoria);

//======================
// RESULTADO
//======================

btnCopiar.addEventListener("click",copiarResultado);

btnLimpar.addEventListener("click",limparTudo);

//======================
// BANCO LOCAL
//======================

bancoCPTM =
JSON.parse(localStorage.getItem("bancoCPTM")) || [];

bancoTrivia =
JSON.parse(localStorage.getItem("bancoTrivia")) || [];

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

    document.getElementById("statusPDF").textContent =
`${maquinistas.length} registros`;

document.getElementById("statusPDF").className =
"fw-bold text-success";

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

    operadoresPostos = [];

    operadoresMonitoria = [];

    operadoresApoio = [];

    operadoresIgnorados = [];

    await lerExcel(excelInput.files[0]);

    console.table(
        operadores.map(op=>({
            nome: op.nome,
            local: op.local,
            maquinista: op.maquinista,
            hora: op.horaMaquinista
        }))
    );

    document.getElementById("statusGestao").textContent =
        `${operadoresPostos.length} operadores`;

    document.getElementById("statusGestao").className =
        "fw-bold text-success";

    document.getElementById("statusTurno").textContent =
        turnoAtual;

}
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
    vagasCPTM = [];

    for(let pagina=1; pagina<=pdf.numPages; pagina++){

        const page = await pdf.getPage(pagina);

        const content = await page.getTextContent();
      //======================================
// IDENTIFICA O TURNO PELO HORÁRIO
//======================================

const horarios = content.items
    .map(item => item.str.trim())
    .filter(txt => /^\d{2}:\d{2}$/.test(txt));

const primeiraEntrada = horarios[0] || "";

if(primeiraEntrada){

    const hora = Number(primeiraEntrada.replace(":",""));

    if(hora >= 400 && hora < 1200){

        turnoCA = "MANHÃ";

    }
    else if(hora >= 1200 && hora < 1800){

        turnoCA = "TARDE";

    }
    else{

        turnoCA = "NOITE";

    }

}

console.log("Primeira entrada:", primeiraEntrada);
console.log("Turno identificado:", turnoCA);
        const linhas = {};

        content.items.forEach(item=>{

            const y = Math.round(item.transform[5]);

            if(!linhas[y]){

                linhas[y] = [];

            }

            linhas[y].push(item);

        });

const listaLinhas = Object.keys(linhas)
    .sort((a,b)=>b-a)
    .map(y=>{

        return linhas[y]
            .sort((a,b)=>a.transform[4]-b.transform[4])
            .map(i=>i.str)
            .join(" ")
            .replace(/\s+/g," ")
            .trim();

    });

for(let i=0;i<listaLinhas.length;i++){

    let linha = listaLinhas[i];

    if(!linha) continue;

    // Junta linhas quebradas
    while(

        i+1 < listaLinhas.length &&

        /-\s*$/.test(linha)

    ){

        linha = linha.replace(/-\s*$/,"");

        linha += " " + listaLinhas[++i];

    }

    processarLinhaPDF(linha);

}

    }

    console.table(maquinistas);

    console.table(vagasCPTM);
    console.table(ocorrenciasPDF);

    console.log(
        `Maquinistas: ${maquinistas.length}`
    );

    console.log(
        `Vagas: ${vagasCPTM.length}`
    );
console.log("Turno do CA:", turnoCA);

document.getElementById("statusTurno").textContent = turnoCA || "--";
}

function processarLinhaPDF(linha){

    linha = linha
        .replace(/\s+/g," ")
        .trim();

    if(!linha) return;

    const match = linha.match(
        /^(.+?)\s+([A-Z]{2}\d{3})\s+(.+?)\s+(\d{2}:\d{2})(.*)$/
    );

    if(!match) return;

    const posto = match[1].trim();

    const escala = match[2].trim();

    const nome = match[3].trim().toUpperCase();

    const entrada = match[4].replace(":","");

    const restante = match[5]
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g,"")
        .toUpperCase();

    //=========================================
    // LINHAS ADMINISTRATIVAS
    //=========================================

    if(

        nome.startsWith("TREIN.") ||

        nome.startsWith("TREINAMENTO")

    ){

        return;

    }

    //=========================================
    // VAGA CPTM
    //=========================================

    if(

        /^(\d{2}:\d{2}|\d{4})$/.test(nome)

    ){

        vagasCPTM.push({

            posto,

            escala,

            hora:entrada

        });

        return;

    }

    //=========================================
    // OCORRÊNCIAS
    //=========================================

    const palavras = [

        "AUSENCIA",

        "FISCAL",

        "CCM",

        "FOLGA",

        "APOIO"

    ];

    const encontrou = palavras.some(p=>

        restante.includes(p)

    );

    if(encontrou){

        ocorrenciasPDF.push({

            posto,

            escala,

            nome,

            entrada,

            observacao:restante

        });

        return;

    }

    //=========================================
    // MAQUINISTA
    //=========================================

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

  //======================================
// LOCALIZA AS ABAS DE TURNO
//======================================

const abasTurno = workbook.SheetNames.filter(nome=>{

    const nomeAba = nome
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g,"")
        .toUpperCase()
        .trim();

    return ["MANHA","MANHÃ","TARDE","NOITE"].includes(nomeAba);

});

if(!abasTurno.length){

    alert("Nenhuma aba de turno foi encontrada na Gestão de Escala.");

    return;

}

let aba;

// Apenas uma aba
if(abasTurno.length===1){

    aba = abasTurno[0];

}
// Mais de uma aba
else{

    const opcoes = abasTurno.join("\n");

    const resposta = prompt(
`Selecione o turno da Gestão:

${opcoes}`
    );

    if(!resposta) return;

    const turnoSelecionado = resposta
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g,"")
        .toUpperCase()
        .trim();

    aba = abasTurno.find(nome=>

        nome
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g,"")
            .toUpperCase()
            .trim()===turnoSelecionado

    );

    if(!aba){

        alert("Turno inválido.");

        return;

    }

}

turnoAtual = aba;

const sheet = workbook.Sheets[aba];
    const dados = XLSX.utils.sheet_to_json(sheet,{
        header:1,
        defval:""
    });

    let linhaCabecalho = -1;

    for(let i=0;i<dados.length;i++){

        const linha = dados[i].map(v=>String(v||"").trim().toUpperCase());

        if(
            linha.includes("NOME COMPLETO") &&
            linha.includes("LOCAL")
        ){
            linhaCabecalho=i;
            break;
        }

    }

    if(linhaCabecalho==-1){

        alert("Cabeçalho da Gestão não encontrado.");

        return;

    }

    const cab=dados[linhaCabecalho]
        .map(v=>String(v||"").trim().toUpperCase());

    const localizarColuna=(...nomes)=>
        cab.findIndex(col=>nomes.includes(col));

    const idxNome=localizarColuna("NOME COMPLETO");
    const idxLocal=localizarColuna("LOCAL");
    const idxEntradaHora=localizarColuna("ENTRADA HORA","HORA ENTRADA");
    const idxEntrada=localizarColuna("ENTRADA");
    const idxMaquinista=localizarColuna("MAQUINISTA CPTM","MAQUINISTA");
    const idxObs=localizarColuna("OBSERVAÇÕES","OBSERVACOES");

    operadores=[];
    operadoresPostos=[];
    operadoresMonitoria=[];
    operadoresApoio=[];
    operadoresIgnorados=[];

    for(let i=linhaCabecalho+1;i<dados.length;i++){

        const linha=dados[i];

        if(!linha.length) continue;

        const nomeCompleto=String(linha[idxNome]||"").trim();

        if(!nomeCompleto) continue;

        const local=idxLocal>=0
            ? String(linha[idxLocal]||"").trim()
            : "";

        const entrada=idxEntradaHora>=0
            ? formatarHora(linha[idxEntradaHora])
            : "";

        const situacao=idxEntrada>=0
            ? String(linha[idxEntrada]||"").trim()
            : "";

        const observacoes=idxObs>=0
            ? String(linha[idxObs]||"").trim()
            : "";

        const texto=idxMaquinista>=0
            ? String(linha[idxMaquinista]||"").trim()
            : "";

        const localMaiusculo=local.toUpperCase();

        let grupo="";

        if(localMaiusculo==="SUZ") grupo="SUZ";
        else if(localMaiusculo==="BAS") grupo="BAS";
        else if(localMaiusculo==="EGO") grupo="EGO";

        const operador={

    nome: nomeCompleto,

    local,

    grupo,

    posto: local,

    hora: entrada,

    entrada,

    situacao,

    observacoes

};

        if(

            localMaiusculo.includes("CCM") ||

            localMaiusculo.includes("AUS") ||

            localMaiusculo.includes("RETORNO") ||

            localMaiusculo.includes("PSO") ||

            localMaiusculo.includes("FISCAL")

        ){

            operadoresIgnorados.push(operador);

        }
        else if(localMaiusculo.includes("APOIO")){

            operadoresApoio.push(operador);

            operadoresPostos.push(operador);

        }
        else{

            operadoresMonitoria.push(operador);

            operadoresPostos.push(operador);

        }

        if(!texto) continue;

        if(

            /^LOCOMOTIVA/i.test(texto) ||

            /^EQUIPE LOCOMOTIVA/i.test(texto) ||

            /^MQT/i.test(texto)

        ){

            continue;

        }

        const regexNovo=/^(.*?)(?:\s+(\d{2}:\d{2}|\d{4}))?$/;

        const regexAntigo=/([A-ZÀ-Ú'. ]+?)\s+(\d{4})/gi;

        if(texto.includes("/")){

            let item;

            while((item=regexAntigo.exec(texto))!==null){

                operadores.push({

                    nome:nomeCompleto,

                    nomeCompleto,

                    local,

                    grupo,

                    entrada,

                    situacao,

                    maquinista:item[1].trim(),

                    horaMaquinista:item[2],

                    observacoes

                });

            }

        }else{

            let maquinista=texto;

            let hora="";

            const partes=texto.match(regexNovo);

            if(partes){

                maquinista=partes[1].trim();

                if(partes[2]){

                    hora=partes[2].replace(":","");

                }

            }

            operadores.push({

                nome:nomeCompleto,

                nomeCompleto,

                local,

                grupo,

                entrada,

                situacao,

                maquinista,

                horaMaquinista:hora,

                observacoes

            });

        }

    }

 if(maquinistas.length){

    resultado.value =
`MONITORIA CPTM x TRIVIA

CONTROLE DE APRESENTAÇÃO

Turno..................... ${turnoCA}

Maquinistas CPTM.......... ${maquinistas.length}

Vagas CPTM................ ${vagasCPTM.length}

Ocorrências............... ${ocorrenciasPDF.length}

Controle de apresentação carregado com sucesso.`;

}else{

    resultado.value =
`GESTÃO DE ESCALA

Turno..................... ${turnoAtual}

Operadores................. ${operadoresPostos.length}

Gestão de Escala carregada com sucesso.`;

}

    console.table(operadoresPostos);
    console.table(operadores);

document.getElementById("statusGestao").textContent =
    `${operadoresPostos.length} operadores`;

document.getElementById("statusTurno").textContent =
    aba;

console.log("Turno CA:", turnoCA);
console.log("Aba carregada:", aba);
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
    vagasCPTM = [];
    ocorrenciasPDF = [];

    texto = texto
        .replace(/\r/g," ")
        .replace(/\n/g," ")
        .replace(/\s+/g," ");

    const regex = /([0-9]{2}-[A-Z0-9 ]+(?:-[A-Z0-9 ]+)*)\s+([A-Z]{2}\d{3})\s+(.+?)\s+(\d{2}:\d{2})(.*?)(?=(?:[0-9]{2}-[A-Z0-9 ]+(?:-[A-Z0-9 ]+)*)\s+[A-Z]{2}\d{3}|$)/gis;

    let item;

    while((item = regex.exec(texto)) !== null){

        const posto = item[1].trim();

        const escala = item[2].trim();

        const nome = item[3].trim().toUpperCase();

        const entrada = item[4].replace(":","");

        const observacao = item[5]
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g,"")
            .toUpperCase()
            .trim();

        //=========================
        // TREINAMENTO
        //=========================

        if(nome.startsWith("TREIN.")){

            continue;

        }

        //=========================
        // VAGA CPTM
        //=========================

        if(/^(\d{2}:\d{2}|\d{4})$/.test(nome)){

            vagasCPTM.push({

                posto,

                escala,

                hora:entrada

            });

            continue;

        }

        //=========================
        // OCORRÊNCIAS
        //=========================

        if(

            observacao.includes("AUSENCIA") ||

            observacao.includes("FISCAL") ||

            observacao.includes("CCM") ||

            observacao.includes("FOLGA") ||

            observacao.includes("APOIO")

        ){

            ocorrenciasPDF.push({

                posto,

                escala,

                nome,

                entrada,

                observacao

            });

            continue;

        }

        //=========================
        // MAQUINISTA
        //=========================

        maquinistas.push({

            posto,

            escala,

            nome,

            entrada

        });

    }

    console.table(maquinistas);
    console.table(vagasCPTM);
    console.table(ocorrenciasPDF);

    console.log("Maquinistas:", maquinistas.length);
    console.log("Vagas:", vagasCPTM.length);
    console.log("Ocorrências:", ocorrenciasPDF.length);

}
/*==================================================
    MOSTRAR LISTA DO PDF
==================================================*/
function mostrarListaPDF(){

    resultado.value =
`MONITORIA CPTM x TRIVIA

CONTROLE DE APRESENTAÇÃO

Turno..................... ${turnoCA}

Maquinistas CPTM.......... ${maquinistas.length}

Vagas CPTM................ ${vagasCPTM.length}

Ocorrências............... ${ocorrenciasPDF.length}

========================================
MAQUINISTAS CPTM
========================================

`;

    //========================================
    // MAQUINISTAS
    //========================================

    maquinistas.forEach(m=>{

        resultado.value +=
`${m.posto}  ${m.escala}  ${m.nome}  ${m.entrada}
`;

    });

    //========================================
    // VAGAS
    //========================================

    resultado.value +=
`
========================================
VAGAS CPTM
========================================

`;

    vagasCPTM.forEach(v=>{

        resultado.value +=
`${v.posto}  ${v.escala}  ${v.hora}
`;

    });

    //========================================
    // OCORRÊNCIAS
    //========================================

    resultado.value +=
`
========================================
OCORRÊNCIAS PDF
========================================

`;

    if(ocorrenciasPDF.length){

        ocorrenciasPDF.forEach(o=>{

            resultado.value +=
`${o.posto}  ${o.escala}  ${o.nome}  ${o.entrada}  ${o.observacao}
`;

        });

    }else{

        resultado.value +=
"Nenhuma ocorrência encontrada.\n";

    }

}/*==================================================
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

/*==================================================
    GERAR RENDIÇÕES
==================================================*/


function gerarRendicoes(){

    if(!operadoresPostos.length){

        alert("Carregue a Gestão de Escala.");

        return;

    }

    // Agrupa operadores por LOCAL
    const postos = {};

    operadoresPostos.forEach(op=>{

        const local = String(op.local)
            .trim()
            .toUpperCase();

        if(!local) return;

        if(!postos[local]){

            postos[local] = [];

        }

        postos[local].push(op);

    });

    resultado.value = "";

    // Processa cada posto
    Object.keys(postos)
        .sort()
        .forEach(local=>{

            const lista = postos[local]
                .sort((a,b)=>a.hora.localeCompare(b.hora));

            // Agrupa por horário
            const horarios = {};

            lista.forEach(op=>{

                if(!horarios[op.hora]){

                    horarios[op.hora] = [];

                }

                horarios[op.hora].push(op);

            });

            const listaHorarios = Object.keys(horarios).sort();

            let totalEquipes;

switch(local){

    case "SUZ":
        totalEquipes = 5;
        break;

    case "BAS":
        totalEquipes = 3;
        break;

    default:
        totalEquipes = Math.max(1, Math.ceil(lista.length / 10));

}

//=====================================
// DISTRIBUIÇÃO POR HORÁRIO
//=====================================

const equipes = Array.from(
    { length: totalEquipes },
    () => []
);

let equipeAtual = 0;

for(const hora of listaHorarios){

    const operadoresHora = [...horarios[hora]];

    while(operadoresHora.length){

        equipes[equipeAtual].push(
            operadoresHora.shift()
        );

        equipeAtual++;

        if(equipeAtual >= totalEquipes){

            equipeAtual = 0;

        }

    }

}
            // Escreve no Resultado

            equipes.forEach((equipe,index)=>{

                resultado.value +=
`${local}-${String(index+1).padStart(2,"0")}
`;

                equipe.forEach(op=>{

                    resultado.value +=
`${buscarNomeGuerraTrivia(op.nome)} ${op.hora}
`;

                });

                resultado.value += "\n";

            });

        });



}
function exportarRendicoesExcel(postos){

    const wb = XLSX.utils.book_new();

    Object.keys(postos)
        .sort()
        .forEach(local=>{

            const lista = postos[local]
                .sort((a,b)=>a.hora.localeCompare(b.hora));

            // Agrupa por horário
            const horarios = {};

            lista.forEach(op=>{

                if(!horarios[op.hora])
                    horarios[op.hora]=[];

                horarios[op.hora].push(op);

            });

            const listaHorarios = Object.keys(horarios).sort();

            const totalEquipes = Math.ceil(lista.length/10);

            const equipes=[];

            for(let i=0;i<totalEquipes;i++)
                equipes.push([]);

            let equipeAtual=0;

            while(true){

                let adicionou=false;

                for(const hora of listaHorarios){

                    if(horarios[hora].length){

                        equipes[equipeAtual].push(
                            horarios[hora].shift()
                        );

                        equipeAtual++;

                        if(equipeAtual>=equipes.length)
                            equipeAtual=0;

                        adicionou=true;

                    }

                }

                if(!adicionou)
                    break;

            }

            const dados=[];

            equipes.forEach((equipe,index)=>{

                dados.push([`${local}-${String(index+1).padStart(2,"0")}`]);

                equipe.forEach(op=>{

                    dados.push([
                        buscarNomeGuerraTrivia(op.nome),
                        op.hora
                    ]);

                });

                dados.push([]);

            });

            const ws = XLSX.utils.aoa_to_sheet(dados);

            ws["!cols"]=[
                {wch:35},
                {wch:10}
            ];

            XLSX.utils.book_append_sheet(
                wb,
                ws,
                local.substring(0,31)
            );

        });

//========================
// Nome do arquivo
//========================

let nomeArquivo = "Rendição";

switch(turnoAtual){

    case "MANHÃ":
    case "MANHA":
        nomeArquivo = "Rendição Manhã";
        break;

    case "TARDE":
        nomeArquivo = "Rendição Tarde";
        break;

    case "NOITE":
        nomeArquivo = "Rendição Noite";
        break;

    default:
        nomeArquivo = "Rendição";

}

XLSX.writeFile(
    wb,
    `${nomeArquivo}.xlsx`
);

}


/*==================================================
    GERAR MONITORIA
==================================================*/

function gerarMonitoria(){

    if(!maquinistas.length){

        alert("Carregue o PDF.");

        return;

    }

    if(operadoresPostos.length===0){

        alert("Carregue a Gestão de Escala.");

        return;

    }

    const listaOperadores = gerarListaOperadores();

    if(listaOperadores.length===0){

        alert("Nenhum operador disponível.");

        return;

    }

    operadoresSemMonitoria = [];

    maquinistasSemOperador = [];

    resultado.value = "";

    dadosExcel = [];

    let listaSimples = "";

    let relatorioPostos = "";

    let monitorados = 0;

    let postoAtual = "";

    maquinistas.sort((a,b)=>{

        if(a.posto!==b.posto){

            return a.posto.localeCompare(b.posto);

        }

        if(a.escala!==b.escala){

            return a.escala.localeCompare(b.escala);

        }

        return converterHora(a.entrada)-converterHora(b.entrada);

    });

    maquinistas.forEach(m=>{

        if(m.posto!==postoAtual){

            postoAtual = m.posto;

            relatorioPostos +=
`
==================================================
POSTO ${postoAtual}
==================================================

`;

        }

        const operador = localizarOperador(

            m,

            listaOperadores

        );
if(operador){

    monitorados++;

    dadosExcel.push({

        tipo:"MONITORIA",

        posto:m.posto,

        escala:m.escala,

        maquinista:m.nome,

        hora:m.entrada,

        operador:operador.nomeCompleto,

        nomeCompletoOperador:operador.nomeCompleto,

        entradaOperador:operador.hora,

        local:operador.local,

        status:"MONITORADO",

        observacao:"",

        monitoria:`${m.nome} ${m.entrada} / ${operador.operador}`

    });

    listaSimples +=
`${m.nome} ${m.entrada} / ${operador.operador}
`;

            relatorioPostos +=
`Escala.....: ${m.escala}
Maquinista.: ${m.nome} ${m.entrada}
Operador...: ${operador.operador}
Grupo......: ${operador.grupo}
Local......: ${operador.local}
Monitoria..: ${m.nome} ${m.entrada} / ${operador.operador}

--------------------------------------------------

`;

        }else{

            maquinistasSemOperador.push({

                posto:m.posto,

                escala:m.escala,

                nome:m.nome,

                hora:m.entrada

            });

            dadosExcel.push({

                posto:m.posto,

                escala:m.escala,

                maquinista:m.nome,

                hora:m.entrada,

                operador:"",

                local:"",

                status:"SEM OPERADOR",

                monitoria:"",
                observacao: ""

            });

            relatorioPostos +=
`Escala.....: ${m.escala}
Maquinista.: ${m.nome} ${m.entrada}
Operador...: SEM OPERADOR TRIVIA

--------------------------------------------------

`;

        }

    });

    operadoresSemMonitoria =
        listaOperadores.filter(op=>!op.utilizado);

    document.getElementById("totalTrivia").textContent =
        listaOperadores.length;

    document.getElementById("totalCPTM").textContent =
        maquinistas.length;

    document.getElementById("totalMonitorias").textContent =
        monitorados;

    document.getElementById("semMonitoria").textContent =
        operadoresSemMonitoria.length;

    document.getElementById("semOperador").textContent =
        maquinistasSemOperador.length;
            //==================================================
    // RESUMO
    //==================================================

    resultado.value =
`
==================================================
RESUMO
==================================================

Operadores TRIVIA................. ${listaOperadores.length}

Maquinistas CPTM.................. ${maquinistas.length}

Monitorias CPTM x TRIVIA.......... ${monitorados}

Operadores TRIVIA sem Monitoria... ${operadoresSemMonitoria.length}

Maquinistas CPTM sem Operador..... ${maquinistasSemOperador.length}
`;

    //==================================================
    // OPERADORES SEM MONITORIA
    //==================================================

    if(operadoresSemMonitoria.length){

        resultado.value +=
`
==================================================
OPERADORES TRIVIA SEM MONITORIA CPTM
==================================================

`;

        operadoresSemMonitoria
            .sort((a,b)=>converterHora(a.hora)-converterHora(b.hora))
            .forEach(op=>{

                resultado.value +=
`${op.operador} ${op.hora} ${op.local}
`;

            });

    }

    //==================================================
    // MAQUINISTAS SEM OPERADOR
    //==================================================

    if(maquinistasSemOperador.length){

        resultado.value +=
`

==================================================
MAQUINISTAS CPTM SEM OPERADOR
==================================================

`;

        maquinistasSemOperador.forEach(m=>{

            resultado.value +=
`${m.posto} ${m.escala} ${m.nome} ${m.hora}
`;

        });

    }
//==============================================
// OCORRÊNCIAS DO PDF
//==============================================

if(ocorrenciasPDF.length){

    resultado.value +=
`
==================================================
OCORRÊNCIAS DO PDF
==================================================

`;

    ocorrenciasPDF.forEach(o=>{

        resultado.value +=
`${o.posto}
Escala......: ${o.escala}
Nome........: ${o.nome}
Entrada.....: ${o.entrada}
Ocorrência..: ${o.observacao}

--------------------------------------------------

`;

        dadosExcel.push({

            posto: o.posto,

            escala: o.escala,

            maquinista: o.nome,

            hora: o.entrada,

            operador: "",

            local: "",

            status: "OCORRÊNCIA PDF",

            monitoria: "",

            observacao: o.observacao

        });

    });

}
    //==================================================
    // VAGAS CPTM
    //==================================================

    if(vagasCPTM.length){

        resultado.value +=
`
==================================================
VAGAS CPTM
==================================================

`;

        vagasCPTM
            .sort((a,b)=>converterHora(a.hora)-converterHora(b.hora))
            .forEach(v=>{

                resultado.value +=
`${v.posto} ${v.escala} ${v.hora}
`;

                dadosExcel.push({

                    posto:v.posto,

                    escala:v.escala,

                    maquinista:"",

                    hora:v.hora,

                    operador:"",

                    local:"",

                    status:"VAGA CPTM",

                    monitoria:"",
                    observacao: ""

                });

            });

    }

    //==================================================
    // LISTA SIMPLES
    //==================================================

    resultado.value +=
`
==================================================
LISTA SIMPLES PARA ESCALA
MONITORIA CPTM x TRIVIA
==================================================

`;

    resultado.value += listaSimples;

        //==================================================
    // RELATÓRIO DETALHADO POR POSTO
    //==================================================

    resultado.value +=
`
==================================================
RELATÓRIO DETALHADO POR POSTO
==================================================

`;

    resultado.value += relatorioPostos;

    console.table(listaOperadores);

}


/*==================================================
    GERAR LISTA OPERADORES
==================================================*/

function gerarListaOperadores(){

    const lista = [];

    const origem = [

        ...operadoresMonitoria,

        ...operadoresApoio

    ];

    origem.forEach(op=>{

        if(!op.hora) return;

        const local = op.local.toUpperCase();

        // Ignorar operadores que nunca entram na monitoria
        if(

            local.includes("CCM") ||

            local.includes("AUS") ||

            local.includes("RETORNO") ||

            local.includes("PSO") ||

            local.includes("FISCAL")

        ){

            return;

        }

lista.push({

    operador: buscarNomeGuerraTrivia(op.nome),

    nomeCompleto: op.nome,

    local: op.local,

    posto: op.posto,

    grupo: op.grupo,

    hora: String(op.hora)
        .replace(":","")
        .padStart(4,"0"),

    utilizado: false

});

    });

    lista.sort((a,b)=>a.hora.localeCompare(b.hora));
console.table(

    lista.map(op=>({

        operador:op.operador,

        grupo:op.grupo,

        local:op.local,

        posto:op.posto

    }))

);
    return lista;

}

const CORRELACAO_POSTOS = {

    "BRÁS":[

        "06-BAS-02",
        "07-BAS-04",
        "09-BAS-12",
        "10-BAS-06 APOIO",

        // EGO passa a ser atendido por BRÁS
        "20-LINHA 13 EGO-03",
        "16-LINHA 12 EGO-02",
        "17-LINHA 12 EGO-04"

    ],

    "SUZANO":[

        "SUZ-01",
        "SUZ-02",
        "SUZ-03",
        "SUZ-04",
        "SUZ-05",
        "SUZ-06 APOIO",
        "SUZ-11"

    ]

};
function localizarOperador(maquinista, lista){

    const posto = maquinista.posto.toUpperCase().trim();

    //====================================
    // POSTOS SEM MONITORIA
    //====================================

    if(

        posto.includes("10-BAS-06") ||

        posto.includes("SUZ-06")

    ){

        return null;

    }

    //====================================
    // MAQUINISTAS DA 00:10
    //====================================

    if(converterHora(maquinista.entrada) === converterHora("0010")){

        return null;

    }

    const grupo = obterGrupoCPTM(posto);

    if(!grupo){

        return null;

    }

    const horaBase = converterHora(maquinista.entrada);

    const candidatos = lista.filter(op=>{

        if(op.utilizado) return false;

        if(op.grupo !== grupo) return false;

        //====================================
        // EGO SOMENTE PARA 20-LINHA 13
        //====================================

        if(

            op.grupo === "EGO" &&

            !posto.includes("20-LINHA 13")

        ){

            return false;

        }

        const horaOperador = converterHora(op.hora);

        //====================================
        // OPERADOR NÃO PODE ENTRAR
        // DEPOIS DO MAQUINISTA
        //====================================

        return horaOperador <= horaBase;

    });

    if(!candidatos.length){

        return null;

    }

    candidatos.sort((a,b)=>{

        const horaA = converterHora(a.hora);

        const horaB = converterHora(b.hora);

        //====================================
        // QUANTO MAIS PRÓXIMO DO HORÁRIO
        // DO MAQUINISTA, MELHOR
        //====================================

        const diffA = horaBase - horaA;

        const diffB = horaBase - horaB;

        if(diffA !== diffB){

            return diffA - diffB;

        }

        //====================================
        // DESEMPATE
        //====================================

        return horaA - horaB;

    });

    const operador = candidatos[0];

    operador.utilizado = true;

    return operador;

}

function converterHora(hora){

    hora = String(hora)
        .replace(":","")
        .trim();

    if(hora.length===3){

        hora="0"+hora;

    }

    const h = parseInt(hora.substring(0,2));

    const m = parseInt(hora.substring(2,4));

    return h*60+m;

}
function obterGrupoCPTM(posto){

    posto = posto
        .toUpperCase()
        .replace(/\s+/g," ")
        .replace(/￾/g,"-")
        .trim();

    console.log("POSTO PDF:", posto);

    //==========================
    // SUZ
    //==========================

    if(posto.includes("SUZ")){

        return "SUZ";

    }

    //==========================
    // BAS
    //==========================

    if(posto.includes("BAS")){

        return "BAS";

    }

    //==========================
    // EGO
    //==========================

    if(

        posto.includes("EGO-02") ||

        posto.includes("EGO-03") ||

        posto.includes("EGO-04") ||

        posto.includes("ENGENHEIRO GOULART") ||

        posto.includes("LINHA 12") ||

        posto.includes("LINHA 13")

    ){

        return "EGO";

    }

    return null;

}

function exportarExcel(){

    if(!dadosExcel.length){

        alert("Gere a monitoria primeiro.");

        return;

    }

    const monitorados =
        dadosExcel.filter(x=>x.status==="MONITORADO");

    const semOperador =
        dadosExcel.filter(x=>x.status==="SEM OPERADOR");

    const vagas =
        dadosExcel.filter(x=>x.status==="VAGA CPTM");

    const linhas = [];

    linhas.push([

        "MONITORIA / PENDÊNCIAS CPTM","","","","","","","","","",

        "",

        "OPERADORES SEM MAQUINISTA","","",""

    ]);

    linhas.push([

        "Tipo",
        "Posto",
        "Escala",
        "Maquinista",
        "Hora",
        "Operador",
        "Entrada",
        "Local",
        "Status",
        "Observação",

        "",

        "Operador",
        "Entrada",
        "Local",
        "Status"

    ]);

    const blocoEsquerdo = [];

    //==================================================
    // MONITORADOS
    //==================================================

    monitorados.forEach(item=>{

        blocoEsquerdo.push({

            tipo:"MONITORIA",

            posto:item.posto,

            escala:item.escala,

            maquinista:item.maquinista,

            hora:item.hora,

            operador:item.nomeCompletoOperador || item.operador,

            entradaOperador:item.entradaOperador || "",

            local:item.local,

            status:item.status,

            observacao:""

        });

    });

    //==================================================
    // OCORRÊNCIAS PDF
    //==================================================

    ocorrenciasPDF.forEach(item=>{

        blocoEsquerdo.push({

            tipo:"OCORRÊNCIA PDF",

            posto:item.posto,

            escala:item.escala,

            maquinista:item.nome,

            hora:item.entrada,

            operador:"",

            entradaOperador:"",

            local:"",

            status:"OCORRÊNCIA",

            observacao:item.observacao

        });

    });

    //==================================================
    // SEM OPERADOR
    //==================================================

    semOperador.forEach(item=>{

        blocoEsquerdo.push({

            tipo:"SEM OPERADOR",

            posto:item.posto,

            escala:item.escala,

            maquinista:item.maquinista,

            hora:item.hora,

            operador:"",

            entradaOperador:"",

            local:"",

            status:item.status,

            observacao:""

        });

    });

    

    //==================================================
    // VAGAS CPTM
    //==================================================

    vagas.forEach(item=>{

        blocoEsquerdo.push({

            tipo:"VAGA CPTM",

            posto:item.posto,

            escala:item.escala,

            maquinista:"",

            hora:item.hora,

            operador:"",

            entradaOperador:"",

            local:"",

            status:item.status,

            observacao:""

        });

    });

    const total = Math.max(

        blocoEsquerdo.length,

        operadoresSemMonitoria.length

    );

    for(let i=0;i<total;i++){

        const e = blocoEsquerdo[i] || {};

        const o = operadoresSemMonitoria[i] || {};

        linhas.push([

            e.tipo || "",
            e.posto || "",
            e.escala || "",
            e.maquinista || "",
            e.hora || "",
            e.operador || "",
            e.entradaOperador || "",
            e.local || "",
            e.status || "",
            e.observacao || "",

            "",

            o.operador || "",
            o.hora || "",
            o.local || "",
            o.operador ? "SEM MAQUINISTA" : ""

        ]);

    }

    const ws = XLSX.utils.aoa_to_sheet(linhas);

    ws["!cols"]=[

        {wch:18}, // Tipo
        {wch:18}, // Posto
        {wch:10}, // Escala
        {wch:35}, // Maquinista
        {wch:8},  // Hora
        {wch:35}, // Operador
        {wch:10}, // Entrada
        {wch:12}, // Local
        {wch:18}, // Status
        {wch:40}, // Observação

        {wch:3},

        {wch:25}, // Operador
        {wch:10}, // Entrada
        {wch:12}, // Local
        {wch:18}  // Status

    ];

    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(

        wb,

        ws,

        "Monitoria"

    );
//==================================================
// ABA LISTA SIMPLES
//==================================================

const listaSimples = [];

listaSimples.push([

    "MAQUINISTA",

    "OPERADOR",

    "MONITORIA CPTM x TRIVIA",

    "",

    "MAQUINISTA CPTM SEM MONITORIA"

]);

listaSimples.push([

    "Nome Completo + Hora",

    "Nome Completo",

    "Lista Simples",

    "",

    "Nome Completo + Hora"

]);

const monitoradosLista = [...monitorados].sort((a,b)=>{

    if((a.entradaOperador || "") !== (b.entradaOperador || "")){
        return (a.entradaOperador || "").localeCompare(b.entradaOperador || "");
    }

    return (a.nomeCompletoOperador || "")
        .localeCompare(b.nomeCompletoOperador || "");

});

const totalLista = Math.max(
    monitoradosLista.length,
    maquinistasSemOperador.length
);

for(let i=0;i<totalLista;i++){

    const item = monitoradosLista[i] || {};

    const sem = maquinistasSemOperador[i] || {};

    let nomeGuerraMaquinista = "";

    let nomeGuerraOperador = "";

    if(item.maquinista){

        nomeGuerraMaquinista =
            buscarNomeGuerraCPTM(item.maquinista);

        nomeGuerraOperador =
            buscarNomeGuerraTrivia(
                item.nomeCompletoOperador || item.operador
            );

    }

listaSimples.push([

    item.maquinista
        ? `${item.maquinista} ${item.hora}`
        : "",

    item.nomeCompletoOperador || item.operador || "",

    item.maquinista
        ? `${nomeGuerraMaquinista} ${item.hora} / ${nomeGuerraOperador}`
        : "",

    "",

    sem.nome
        ? `${sem.nome} ${sem.hora}`
        : ""

]);
}

const wsLista = XLSX.utils.aoa_to_sheet(listaSimples);

wsLista["!cols"] = [

    {wch:40},

    {wch:35},

    {wch:45},

    {wch:3},

    {wch:40}

];

XLSX.utils.book_append_sheet(

    wb,

    wsLista,

    "LISTA SIMPLES"

);

    XLSX.writeFile(

        wb,

        `Monitoria_${turnoAtual}.xlsx`

    );

}

