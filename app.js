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

    turnoAtual = aba.toUpperCase().trim();

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

    resultado.value=
`MONITORIA CPTM x TRIVIA

GESTÃO CARREGADA

Turno..................... ${turnoAtual}

Operadores TRIVIA......... ${operadoresPostos.length}

Operadores Aptos.......... ${operadoresMonitoria.length}

Operadores Apoio.......... ${operadoresApoio.length}

Operadores Ignorados...... ${operadoresIgnorados.length}`;

    console.table(operadoresPostos);
    console.table(operadores);

    document.getElementById("statusGestao").textContent =
        `${operadoresPostos.length} operadores`;

    document.getElementById("statusTurno").textContent =
        turnoAtual;

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

            const totalEquipes = Math.ceil(lista.length / 10);

            const equipes = [];

            for(let i=0;i<totalEquipes;i++){

                equipes.push([]);

            }

            let equipeAtual = 0;

            while(true){

                let adicionou = false;

                for(const hora of listaHorarios){

                    if(horarios[hora].length){

                        equipes[equipeAtual].push(
                            horarios[hora].shift()
                        );

                        equipeAtual++;

                        if(equipeAtual >= equipes.length){

                            equipeAtual = 0;

                        }

                        adicionou = true;

                    }

                }

                if(!adicionou){

                    break;

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

    //=========================
    // GERAR EXCEL
    //=========================

    const linhas = resultado.value
        .trim()
        .split("\n")
        .map(linha=>[linha]);

    const wb = XLSX.utils.book_new();

    const ws = XLSX.utils.aoa_to_sheet(linhas);

    ws["!cols"] = [
        { wch: 50 }
    ];

    XLSX.utils.book_append_sheet(
        wb,
        ws,
        "Rendições"
    );

    XLSX.writeFile(
        wb,
        "Rendicoes.xlsx"
    );

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

    const listaOperadores=gerarListaOperadores();

    if(listaOperadores.length===0){

        alert("Nenhum operador disponível.");

        return;

    }

    operadoresSemMonitoria=[];

    maquinistasSemOperador=[];

    resultado.value="";

    let listaSimples="";
    dadosExcel = [];

    let monitorados=0;

    let postoAtual="";

    maquinistas.sort((a,b)=>{

        if(a.posto!==b.posto){

            return a.posto.localeCompare(b.posto);

        }

        if(a.escala!==b.escala){

            return a.escala.localeCompare(b.escala);

        }

        return converterHora(a.entrada)-converterHora(b.entrada);

    });

    resultado.value+=
`==================================================
MONITORIA CPTM x TRIVIA
==================================================

`;

    maquinistas.forEach(m=>{

        if(m.posto!==postoAtual){

            postoAtual=m.posto;

            resultado.value+=
`
==================================================
POSTO ${postoAtual}
==================================================

`;

        }

        const operador=localizarOperador(

            m,

            listaOperadores

        );

        if(operador){

            monitorados++;
            dadosExcel.push({

    posto: m.posto,

    escala: m.escala,

    maquinista: m.nome,

    hora: m.entrada,

    operador: operador.operador,

    local: operador.local,

    status: "MONITORADO",

    monitoria: `${m.nome} ${m.entrada} / ${operador.operador}`

});

            listaSimples +=
`${m.nome} ${m.entrada} / ${operador.operador}
`;

            resultado.value+=
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

    posto: m.posto,

    escala: m.escala,

    maquinista: m.nome,

    hora: m.entrada,

    operador: "",

    local: "",

    status: "SEM OPERADOR",

    monitoria: ""

});

            resultado.value+=
`Escala.....: ${m.escala}
Maquinista.: ${m.nome} ${m.entrada}
Operador...: SEM OPERADOR

--------------------------------------------------

`;

        }

    });

    operadoresSemMonitoria=
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

    resultado.value+=
`
==================================================
RESUMO
==================================================

Operadores TRIVIA............. ${listaOperadores.length}

Maquinistas CPTM.............. ${maquinistas.length}

Monitorias.................... ${monitorados}

Operadores sem Monitoria...... ${operadoresSemMonitoria.length}

Maquinistas sem Operador...... ${maquinistasSemOperador.length}

`;

    if(operadoresSemMonitoria.length){

        resultado.value+=
`
==================================================
OPERADORES SEM MONITORIA
==================================================

`;

        operadoresSemMonitoria
            .sort((a,b)=>converterHora(a.hora)-converterHora(b.hora))
            .forEach(op=>{

                resultado.value+=
`${op.operador} ${op.hora} ${op.local}
`;

            });

    }

    if(maquinistasSemOperador.length){

        resultado.value+=
`
==================================================
MAQUINISTAS SEM OPERADOR
==================================================

`;

        maquinistasSemOperador.forEach(m=>{

            resultado.value+=
`${m.posto} ${m.escala} ${m.nome} ${m.hora}
`;

        });

    }

    resultado.value+=
`
==================================================
LISTA SIMPLES
==================================================

${listaSimples}
`;

    console.table(listaOperadores);

}

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

    "BOA VISTA":[

        "06-BAS-02",
        "07-BAS-04",
        "09-BAS-12",
        "10-BAS-06 APOIO"

    ],

    "BRÁS":[

        "SUZ-01",
        "SUZ-02",
        "SUZ-03",
        "SUZ-04",
        "SUZ-05",
        "SUZ-06 APOIO",
        "SUZ-11"

    ],

    "ENGENHEIRO GOULART":[

        "20-LINHA 13 EGO-03",
        "16-LINHA 12 EGO-02",
        "17-LINHA 12 EGO-04"

    ]

};
function localizarOperador(maquinista, lista){

    const grupo = obterGrupoCPTM(maquinista.posto);

    if(!grupo){

        return null;

    }

    const horaBase = converterHora(maquinista.entrada);

    const candidatos = lista.filter(op=>{

        if(op.utilizado) return false;

        if(op.grupo !== grupo) return false;

        const diferenca = Math.abs(

            converterHora(op.hora) -

            horaBase

        );

        return diferenca <= 15;

    });

    console.log(

        maquinista.nome,

        grupo,

        candidatos.length

    );

    if(!candidatos.length){

        return null;

    }

    candidatos.sort((a,b)=>{

        const diffA = Math.abs(

            converterHora(a.hora) -

            horaBase

        );

        const diffB = Math.abs(

            converterHora(b.hora) -

            horaBase

        );

        return diffA - diffB;

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

    posto = posto.toUpperCase().trim();

    // SUZ
    if(
        posto === "SUZ-01" ||
        posto === "SUZ-02" ||
        posto === "SUZ-03" ||
        posto === "SUZ-04" ||
        posto === "SUZ-05" ||
        posto === "SUZ-11"
    ){
        return "SUZ";
    }

    // BAS
    if(
        posto === "06-BAS-02" ||
        posto === "07-BAS-04" ||
        posto === "09-BAS-12"
    ){
        return "BAS";
    }

    // EGO
    if(
        posto === "20-LINHA 13 EGO-03" ||
        posto === "16-LINHA 12 EGO-02" ||
        posto === "17-LINHA 12 EGO-04"
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

    const ws = XLSX.utils.json_to_sheet(dadosExcel);

    const wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(

        wb,

        ws,

        "Monitoria"

    );

    XLSX.writeFile(

        wb,

        `Monitoria_${turnoAtual}.xlsx`

    );

}

btnExcelMonitoria.onclick = exportarExcel;