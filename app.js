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
const btnExcelMonitoria = document.getElementById("btnExcelMonitoria");

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

    const nome = limparNomeMaquinistaPDF(match[3]);

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

        // A classificação ESCALA/ESCALANTE pode aparecer na coluna
        // MAQUINISTA SAÍDA, em OBSERVAÇÕES ou em outro campo da mesma linha.
        // Nesses casos o operador permanece na Gestão, mas fica bloqueado
        // para o algoritmo de monitoria.
        const linhaGestaoTexto = linha
            .map(v => String(v ?? "").trim().toUpperCase())
            .join(" ");
        const bloqueadoEscala = /\bESCALA\b|\bESCALANTE\b/.test(linhaGestaoTexto);

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

    observacoes,

    bloqueadoEscala

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

                    observacoes,
                    bloqueadoEscala

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

                observacoes,
                bloqueadoEscala

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

        const nome = limparNomeMaquinistaPDF(item[3]);

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


let duplasRealizadas = [];

const MAX_DIFERENCA_HORARIO_MONITORIA = 15;

const MAQUINISTAS_NAO_ESCALAR = new Set([
    "MARCOS ROBERTO LEMES",
    "FANI JULIANA PIMENTA",
    "ABRAAO MOURA DE HOLANDA",
    "SIMASCER DE SOUSA SANTOS",
    "PAOLA DE LIMA SILVA",
    "FABRICIO GERMANO DA SILVA"
]);

function normalizarNomeRestricao(nome){
    return String(nome || "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/\s+/g, " ")
        .trim()
        .toUpperCase();
}

function maquinistaTemRestricao(nome){
    return MAQUINISTAS_NAO_ESCALAR.has(normalizarNomeRestricao(nome));
}

function formatarHora4(hora){
    if(hora === undefined || hora === null || String(hora).trim() === "") return "";
    let h=String(hora).replace(":","").trim();
    if(h.length===3) h="0"+h;
    return h.padStart(4,"0");
}

function limparNomeMaquinistaPDF(nome){
    return String(nome || "")
        .replace(/^\s*\d+\s+/, "")
        .replace(/^\s*[-–—:]+\s*/, "")
        .replace(/\s+/g, " ")
        .trim()
        .toUpperCase();
}

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
    duplasRealizadas = [];

    let listaSimples = "";

    let relatorioPostos = "";

    let monitorados = 0;
    const monitoriasNormais = [];

    let postoAtual = "";

    // Classificação principal dos maquinistas por horário, mantendo as regras
    // de escala/compatibilidade já existentes. Em caso de empate, mantém
    // posto e escala como critérios secundários.
    maquinistas.sort((a,b)=>{

        const difHora = converterHora(a.entrada) - converterHora(b.entrada);
        if(difHora !== 0) return difHora;

        if(a.posto!==b.posto){
            return String(a.posto || '').localeCompare(String(b.posto || ''));
        }

        return String(a.escala || '').localeCompare(String(b.escala || ''));

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

        let operador = null;

        if(!maquinistaTemRestricao(m.nome)){
            operador = localizarOperador(
                m,
                listaOperadores
            );
        }
if(maquinistaTemRestricao(m.nome)){

            maquinistasSemOperador.push({
                posto:m.posto, escala:m.escala, nome:m.nome,
                hora:formatarHora4(m.entrada), restricao:true
            });

            dadosExcel.push({
                tipo:"NÃO ESCALAR", posto:m.posto, escala:m.escala,
                maquinista:m.nome, hora:formatarHora4(m.entrada),
                operador:"", nomeCompletoOperador:"", entradaOperador:"",
                local:"", status:"NÃO ESCALAR", observacao:"NÃO ESCALAR", monitoria:""
            });

            relatorioPostos +=
`Escala.....: ${m.escala}
Maquinista.: ${m.nome} ${formatarHora4(m.entrada)}
Operador...: NÃO ESCALAR

--------------------------------------------------

`;

        }else if(operador){

    monitorados++;

    dadosExcel.push({

        tipo:"MONITORIA",

        posto:m.posto,

        escala:m.escala,

        maquinista:m.nome,

        hora:formatarHora4(m.entrada),

        operador:operador.nomeCompleto,

        nomeCompletoOperador:operador.nomeCompleto,

        entradaOperador:formatarHora4(operador.hora),

        local:operador.local,

        status:"MONITORADO",

        observacao:"",

        monitoria:`${m.nome} ${m.entrada} / ${operador.operador}`

    });

    monitoriasNormais.push({
        maquinista:m,
        operador:operador
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

                hora:formatarHora4(m.entrada),
                restricao:false

            });

            dadosExcel.push({

                posto:m.posto,

                escala:m.escala,

                maquinista:m.nome,

                hora:formatarHora4(m.entrada),

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

    // ==================================================
    // DUPLAS AUTOMÁTICAS POR HORÁRIO
    // ==================================================
    // Não existem mais "sugestões de dupla".
    // Após a escalação normal, procura-se um segundo operador
    // que tenha horário compatível simultaneamente com:
    // 1) o maquinista; e
    // 2) o operador já escalado.
    //
    // Regra de horário:
    // - operador no mesmo horário ou ANTES do maquinista;
    // - no máximo 20 minutos antes do maquinista;
    // - o segundo operador também deve estar a no máximo 20 min
    //   do operador que já foi escalado.
    // - operador 03h não entra.
    // ==================================================
    const MAX_DIFERENCA_DUPLA = 20;

    function operadorPodeFormarDupla(maquinista, operadorEscalado, candidato){

        const horaMaq = converterHora(maquinista.entrada);
        const horaEscalado = converterHora(operadorEscalado.hora);
        const horaCandidato = converterHora(candidato.hora);

        const postoMaq = String(maquinista.posto || "")
            .toUpperCase()
            .trim();

        if(formatarHora4(candidato.hora).startsWith("03")) return false;
        if(candidato.utilizado) return false;
        if(candidato.duplaSugerida) return false;

        const grupoMaq = obterGrupoCPTM(maquinista.posto);
        if(!grupoMaq || candidato.grupo !== grupoMaq) return false;

        if(
            candidato.grupo === "EGO" &&
            !postoMaq.includes("20-LINHA 13")
        ){
            return false;
        }

        // Nunca aceitar operador depois do maquinista.
        if(horaCandidato > horaMaq) return false;

        // O candidato precisa estar até 20 min antes do maquinista.
        if((horaMaq - horaCandidato) > MAX_DIFERENCA_DUPLA) return false;

        // O candidato também precisa ser compatível com o operador
        // que já está escalado.
        if(Math.abs(horaCandidato - horaEscalado) > MAX_DIFERENCA_DUPLA){
            return false;
        }

        return true;
    }

    monitoriasNormais.forEach(mon=>{

        const m = mon.maquinista;
        const operadorEscalado = mon.operador;

        const candidatos = listaOperadores
            .filter(op => operadorPodeFormarDupla(m, operadorEscalado, op))
            .sort((a,b)=>{

                const horaA = converterHora(a.hora);
                const horaB = converterHora(b.hora);
                const horaMaq = converterHora(m.entrada);
                const horaEsc = converterHora(operadorEscalado.hora);

                const scoreA =
                    Math.abs(horaA - horaMaq) +
                    Math.abs(horaA - horaEsc);

                const scoreB =
                    Math.abs(horaB - horaMaq) +
                    Math.abs(horaB - horaEsc);

                if(scoreA !== scoreB) return scoreA - scoreB;

                return horaA - horaB;
            });

        const segundoOperador = candidatos[0];

        if(!segundoOperador) return;

        // O segundo operador deixa de ser sobra e passa a fazer parte
        // da dupla automaticamente.
        segundoOperador.utilizado = true;
        segundoOperador.duplaSugerida = false;

        const registro = dadosExcel.find(d =>
            d.status === "MONITORADO" &&
            d.maquinista === m.nome &&
            d.posto === m.posto &&
            d.escala === m.escala
        );

        if(registro){
            registro.status = "MONITORADO - DUPLA";
            registro.segundoOperador = segundoOperador.nomeCompleto;
            registro.horaSegundoOperador = formatarHora4(segundoOperador.hora);
            registro.localSegundoOperador = segundoOperador.local;
            registro.monitoria =
                `${m.nome} ${formatarHora4(m.entrada)} / ${operadorEscalado.operador} ${formatarHora4(operadorEscalado.hora)} + ${segundoOperador.operador} ${formatarHora4(segundoOperador.hora)}`;
        }

        duplasRealizadas.push({
            posto:m.posto,
            escala:m.escala,
            maquinista:m.nome,
            horaMaquinista:formatarHora4(m.entrada),
            operador1:operadorEscalado.nomeCompleto,
            hora1:formatarHora4(operadorEscalado.hora),
            operador2:segundoOperador.nomeCompleto,
            hora2:formatarHora4(segundoOperador.hora)
        });
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

    if(duplasRealizadas.length){
        resultado.value +=
`
==================================================
DUPLAS REALIZADAS POR COMPATIBILIDADE DE HORÁRIO
==================================================

`;

        duplasRealizadas.forEach(d=>{
            resultado.value +=
`${d.maquinista} ${d.horaMaquinista} / ${d.operador1} ${d.hora1} + ${d.operador2} ${d.hora2} - DUPLA
`;
        });
    }

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

        const local = String(op.local || '').toUpperCase();
        const nomeOperador = String(op.nome || '').toUpperCase().trim();

        // ESCALA/ESCALANTE pode estar na classificação da linha da Gestão,
        // e não necessariamente no nome do operador.
        // O operador continua listado na planilha, mas nunca entra no algoritmo.
        if(
            op.bloqueadoEscala ||
            /\bESCALA\b|\bESCALANTE\b/.test(nomeOperador)
        ){
            return;
        }

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

    hora: formatarHora4(op.hora),

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
        // REGRA DE HORÁRIO:
        // O operador deve ter o mesmo horário ou entrar ANTES
        // do maquinista. NUNCA depois.
        // Limite máximo: 15 minutos antes.
        //====================================

        return horaOperador <= horaBase && (horaBase - horaOperador) <= MAX_DIFERENCA_HORARIO_MONITORIA;

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
    if(!dadosExcel.length){ alert("Gere a monitoria primeiro."); return; }

    // Todos os maquinistas processados aparecem na MONITORIA, inclusive bloqueados.
    const registrosMaquinistas = dadosExcel.filter(x => x && x.maquinista);
    const operadoresSobraram = (operadoresSemMonitoria || [])
        .map(o => ({nome:o.nomeCompleto || o.operador || "", hora:formatarHora4(o.hora)}))
        .sort((a,b)=>converterHora(a.hora)-converterHora(b.hora) || a.nome.localeCompare(b.nome));

    const linhas=[[
        "POSTO DO MAQUINISTA",
        "NOME COMPLETO MAQUINISTA + HORÁRIO",
        "NOME COMPLETO DO OPERADOR",
        "HORÁRIO DE ENTRADA DO OPERADOR",
        "NOME COMPLETO DO OPERADOR (EM CASO DE DUPLA)",
        "HORÁRIO DE ENTRADA DO OPERADOR (EM CASO DE DUPLA)",
        "OPERADORES QUE SOBRARAM",
        "HORÁRIO DOS OPERADORES QUE SOBRARAM"
    ]];

    const total=Math.max(registrosMaquinistas.length, operadoresSobraram.length);
    for(let i=0;i<total;i++){
        const m=registrosMaquinistas[i] || {};
        const os=operadoresSobraram[i] || {};
        linhas.push([
            m.posto || "",
            m.maquinista ? `${m.maquinista} ${formatarHora4(m.hora)}`.trim() : "",
            m.nomeCompletoOperador || m.operador || "",
            formatarHora4(m.entradaOperador || m.horaOperador || ""),
            m.segundoOperador || "",
            formatarHora4(m.horaSegundoOperador || ""),
            os.nome || "",
            os.hora || ""
        ]);
    }

    const ws=XLSX.utils.aoa_to_sheet(linhas);
    ws["!cols"]=[{wch:24},{wch:48},{wch:42},{wch:28},{wch:42},{wch:34},{wch:42},{wch:30}];

    // GESTAO: operadores e horários fixos; somente o maquinista é fórmula vinculada à MONITORIA.
    const gestao=[["NOME DO OPERADOR","HORÁRIO DO OPERADOR","MAQUINISTA CPTM + HORÁRIO"]];
    const registrosGestao=[];

    registrosMaquinistas.forEach((linha,index)=>{
        const row=index+2;
        const op1=linha.nomeCompletoOperador || linha.operador || "";
        const h1=formatarHora4(linha.entradaOperador || linha.horaOperador || "");
        const op2=linha.segundoOperador || "";
        const h2=formatarHora4(linha.horaSegundoOperador || "");
        if(op1) registrosGestao.push({hora:converterHora(h1),operador:op1,horaOriginal:h1,linha:row,colunaOperador:"C"});
        if(op2) registrosGestao.push({hora:converterHora(h2),operador:op2,horaOriginal:h2,linha:row,colunaOperador:"E"});
    });

    registrosGestao.sort((a,b)=>a.hora-b.hora || a.operador.localeCompare(b.operador));
    registrosGestao.forEach(reg=>gestao.push([
        reg.operador,
        reg.horaOriginal,
        {f:`IF(Monitoria!${reg.colunaOperador}${reg.linha}="","",Monitoria!B${reg.linha})`}
    ]));

    const wsGestao=XLSX.utils.aoa_to_sheet(gestao);
    wsGestao["!cols"]=[{wch:42},{wch:28},{wch:48}];

    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Monitoria");
    XLSX.utils.book_append_sheet(wb,wsGestao,"GESTAO");

    const agora=new Date();
    const dataArquivo=[String(agora.getDate()).padStart(2,"0"),String(agora.getMonth()+1).padStart(2,"0"),agora.getFullYear()].join("-");
    XLSX.writeFile(wb,`monitoria (${dataArquivo}).xlsm`);
}
