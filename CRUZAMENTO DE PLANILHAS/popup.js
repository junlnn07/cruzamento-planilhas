const ordem = [
    "nrcc",
    "nome",
    "cpf",
    "Cooperativa",
    "FLG_PREJUIZO",
    "Pagamento_Total",
    "Recuperação_Efetiva",
    "contrato",
    "DS_GARANTIA",
    "Acordo",
    "SituaÃ§Ã£o",
    "Saldo (todos contratos do CPF)",
    "Dias Atraso (mais antigo do CPF)"
];

const mapaRetaca = {
    "1": "01. RETENÇÃO EFETIVADA COM SUCESSO",
    "2": "02. CLIENTE RECUSOU PROPOSTA",
    "3": "03. SEM CONTATO COM CLIENTE",
    "4": "04. AGUARDANDO RETORNO DO CLIENTE"
};

const mapaSitcod = {
    "1": "Ativo / Em Dia",
    "2": "Em Atraso",
    "3": "Quitado / Acordo Realizado",
    "4": "Cancelado"
};

function atualizarCamposFormatados(nomeOuCodigo) {
    const textarea = document.getElementById('_RETACA');
    if (textarea) {
        const textoRetaca = mapaRetaca[nomeOuCodigo] || nomeOuCodigo;
        textarea.value = textoRetaca;
        textarea.dispatchEvent(new Event('change', { bubbles: true }));
    }

    const select2Container = document.getElementById('select2-_SITCOD-container');
    if (select2Container) {
        const textoSitcod = mapaSitcod[nomeOuCodigo] || nomeOuCodigo;
        select2Container.textContent = textoSitcod;
        select2Container.title = textoSitcod;
    }
}

let resultadoFinal = null;

// Elementos HTML
const toggleAcordos = document.getElementById("toggleAcordos");
const acordosContainer = document.getElementById("acordosContainer");

const biFile = document.getElementById("biFile");
const siscobraFile = document.getElementById("siscobraFile");
const acordosFile = document.getElementById("acordosFile");

const biName = document.getElementById("biName");
const siscobraName = document.getElementById("siscobraName");
const acordosName = document.getElementById("acordosName");

const columns = document.getElementById("columns");
const status = document.getElementById("status");

const processBtn = document.getElementById("process");
const downloadBtn = document.getElementById("download");

const rowsElement = document.getElementById("rows");
const colsElement = document.getElementById("cols");

// Controle de exibição da 3ª planilha
toggleAcordos.addEventListener("change", () => {
    if (toggleAcordos.checked) {
        acordosContainer.style.display = "block";
    } else {
        acordosContainer.style.display = "none";
        acordosFile.value = "";
        acordosName.textContent = "Nenhum arquivo selecionado";
    }
});

// Nomes de Arquivos
biFile.addEventListener("change", () => {
    biName.textContent = biFile.files.length > 0 ? biFile.files[0].name : "Nenhum arquivo selecionado";
});

siscobraFile.addEventListener("change", () => {
    siscobraName.textContent = siscobraFile.files.length > 0 ? siscobraFile.files[0].name : "Nenhum arquivo selecionado";
});

acordosFile.addEventListener("change", () => {
    acordosName.textContent = acordosFile.files.length > 0 ? acordosFile.files[0].name : "Nenhum arquivo selecionado";
});

// Adicionar Coluna Extra
document.getElementById("addColumn").addEventListener("click", () => {
    if (columns.children.length >= 5) {
        status.textContent = "⚠ Máximo de 5 colunas extras.";
        return;
    }

    const div = document.createElement("div");
    div.className = "column";
    div.innerHTML = `
        <input type="text" placeholder="Nome da coluna">
        <button class="remove">×</button>
    `;

    div.querySelector(".remove").addEventListener("click", () => div.remove());
    columns.appendChild(div);
});

// Normalizações
function normalizarColuna(valor) {
    return String(valor ?? "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .trim()
        .toLowerCase()
        .replace(/[^a-z0-9]/g, "");
}

function normalizarContrato(valor) {
    if (valor === null || valor === undefined) return "";
    let texto = String(valor).trim().toLowerCase();
    texto = texto.replace(/\s+/g, "");
    texto = texto.replace(/\.0$/, "");
    return texto;
}

function normalizarCPF(valor) {
    if (valor === null || valor === undefined) return "";
    // Mantém apenas dígitos numéricos
    let cpf = String(valor).replace(/\D/g, "");
    // Preenche com zeros à esquerda se necessário (garante 11 dígitos)
    if (cpf.length > 0 && cpf.length < 11) {
        cpf = cpf.padStart(11, "0");
    }
    return cpf;
}

function localizarColuna(dados, nomesPossiveis) {
    if (!dados || !dados.length) return null;
    const colunas = Object.keys(dados[0]);
    const nomesNormalizados = nomesPossiveis.map(nome => normalizarColuna(nome));

    for (const coluna of colunas) {
        if (nomesNormalizados.includes(normalizarColuna(coluna))) {
            return coluna;
        }
    }
    return null;
}

function obterValorNormalizado(objeto, nomeColunaBuscada) {
    if (!objeto) return undefined;
    const chaveNormalizadaBuscada = normalizarColuna(nomeColunaBuscada);
    const chavesObjeto = Object.keys(objeto);

    for (const chave of chavesObjeto) {
        if (normalizarColuna(chave) === chaveNormalizadaBuscada) {
            return objeto[chave];
        }
    }
    return undefined;
}

// Ler Excel
function lerArquivo(file, colunaProcurada = null) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = function(event) {
            try {
                const workbook = XLSX.read(event.target.result, { type: "array", cellDates: false });
                let abaEncontrada = null;
                let dadosEncontrados = null;

                for (const nomeAba of workbook.SheetNames) {
                    const sheet = workbook.Sheets[nomeAba];
                    const dados = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });

                    if (!dados.length) continue;

                    if (!colunaProcurada) {
                        abaEncontrada = nomeAba;
                        dadosEncontrados = dados;
                        break;
                    }

                    const cabecalhos = Object.keys(dados[0]);
                    const encontrou = cabecalhos.some(c => normalizarColuna(c) === normalizarColuna(colunaProcurada));

                    if (encontrou) {
                        abaEncontrada = nomeAba;
                        dadosEncontrados = dados;
                        break;
                    }
                }

                if (!dadosEncontrados) {
                    const todasColunas = [];
                    workbook.SheetNames.forEach(nomeAba => {
                        const sheet = workbook.Sheets[nomeAba];
                        const dados = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
                        if (dados.length) {
                            todasColunas.push(`Aba "${nomeAba}": ` + Object.keys(dados[0]).join(", "));
                        }
                    });

                    throw new Error(
                        `Não encontrei a coluna "${colunaProcurada}". ` +
                        `Colunas encontradas: ${todasColunas.join(" | ")}`
                    );
                }

                console.log(`✓ Coluna "${colunaProcurada}" encontrada na aba: ${abaEncontrada}`);
                resolve(dadosEncontrados);

            } catch (erro) {
                reject(erro);
            }
        };

        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Processamento
processBtn.addEventListener("click", async () => {

    if (!biFile.files[0]) {
        status.textContent = "⚠ Selecione a planilha BI.";
        return;
    }

    if (!siscobraFile.files[0]) {
        status.textContent = "⚠ Selecione a planilha Siscobra.";
        return;
    }

    const usarAcordos = toggleAcordos.checked;
    if (usarAcordos && !acordosFile.files[0]) {
        status.textContent = "⚠ Selecione a planilha de Acordos ou desative a opção.";
        return;
    }

    processBtn.disabled = true;
    processBtn.textContent = "PROCESSANDO...";
    downloadBtn.disabled = true;
    status.textContent = "⏳ Carregando planilhas...";

    try {
        let bi = await lerArquivo(biFile.files[0], "nrcontrato_Cyber");
        let siscobra = await lerArquivo(siscobraFile.files[0], "Contrato Cliente");
        let acordos = [];

        if (!bi.length) throw new Error("A planilha BI está vazia.");
        if (!siscobra.length) throw new Error("A planilha Siscobra está vazia.");

        if (usarAcordos) {
            acordos = await lerArquivo(acordosFile.files[0], "CPF");
            if (!acordos.length) throw new Error("A planilha de Acordos está vazia.");
        }

        const colunaContratoBI = localizarColuna(bi, [
            "nrcontrato_Cyber", "nrcontrato cyber", "nr contrato cyber", "nrcontrato", "numero contrato cyber"
        ]);

        const colunaCPFBI = localizarColuna(bi, ["cpf", "nr_cpf", "num_cpf"]);

        const colunaContratoSiscobra = localizarColuna(siscobra, [
            "Contrato Cliente", "contrato cliente", "contrato_cliente", "contrato-cliente"
        ]);

        if (!colunaContratoBI) {
            throw new Error('Não encontrei a coluna "nrcontrato_Cyber" na BI. Colunas: ' + Object.keys(bi[0]).join(", "));
        }

        if (!colunaContratoSiscobra) {
            throw new Error('Não encontrei a coluna "Contrato Cliente" na Siscobra. Colunas: ' + Object.keys(siscobra[0]).join(", "));
        }

        let colunaCPFAcordos = null;
        let colunaValor1Pagamento = null;
        let colunaStatusAcordos = null;

        if (usarAcordos) {
            colunaCPFAcordos = localizarColuna(acordos, ["cpf", "nr_cpf", "num_cpf", "cpf_cnpj"]);
            colunaValor1Pagamento = localizarColuna(acordos, ["VALOR 1 PAGAMENTO", "valor 1 pagamento", "valor1pagamento", "valor_1_pagamento"]);
            colunaStatusAcordos = localizarColuna(acordos, ["STATUS", "status"]);

            if (!colunaCPFAcordos) {
                throw new Error('Não encontrei a coluna "CPF" na planilha de Acordos. Colunas: ' + Object.keys(acordos[0]).join(", "));
            }
        }

        status.textContent = "⏳ Colunas encontradas. Cruzando dados...";

        // Mapa Siscobra (Chave: Contrato)
        const mapaSiscobra = new Map();
        siscobra.forEach(linha => {
            const contrato = normalizarContrato(linha[colunaContratoSiscobra]);
            if (contrato) mapaSiscobra.set(contrato, linha);
        });

        // Mapa Acordos (Chave: CPF - PROCV)
        const mapaAcordos = new Map();
        if (usarAcordos) {
            acordos.forEach(linha => {
                const cpf = normalizarCPF(linha[colunaCPFAcordos]);
                if (cpf && !mapaAcordos.has(cpf)) {
                    mapaAcordos.set(cpf, linha);
                }
            });
        }

        // Colunas Extras adicionadas pelo usuário
        const extras = [];
        document.querySelectorAll("#columns input").forEach(input => {
            const nome = input.value.trim();
            if (nome && !extras.includes(nome)) extras.push(nome);
        });

        // Montagem da estrutura de colunas do resultado
        let ordemFinal = [...ordem];
        if (usarAcordos) {
            ordemFinal.push("VALOR 1 PAGAMENTO", "STATUS");
        }
        ordemFinal = [...ordemFinal, ...extras];

        resultadoFinal = bi.map(linhaBI => {
            const contrato = normalizarContrato(linhaBI[colunaContratoBI]);
            const cpfBI = normalizarCPF(linhaBI[colunaCPFBI]);

            const linhaSiscobra = mapaSiscobra.get(contrato) || {};
            const linhaAcordos = usarAcordos ? (mapaAcordos.get(cpfBI) || {}) : {};

            const resultado = {};

            ordemFinal.forEach(coluna => {
                // Tratamento específico da coluna Contrato
                if (normalizarColuna(coluna) === normalizarColuna("contrato")) {
                    resultado[coluna] = linhaBI[colunaContratoBI] ?? "";
                    return;
                }

                // Tratamento das novas colunas vindas do PROCV da 3ª planilha
                if (usarAcordos && normalizarColuna(coluna) === normalizarColuna("VALOR 1 PAGAMENTO")) {
                    resultado[coluna] = colunaValor1Pagamento ? (linhaAcordos[colunaValor1Pagamento] ?? "") : "";
                    return;
                }

                if (usarAcordos && normalizarColuna(coluna) === normalizarColuna("STATUS")) {
                    resultado[coluna] = colunaStatusAcordos ? (linhaAcordos[colunaStatusAcordos] ?? "") : "";
                    return;
                }

                // Busca nos dados da BI
                const valorBI = obterValorNormalizado(linhaBI, coluna);
                if (valorBI !== undefined) {
                    resultado[coluna] = valorBI;
                    return;
                }

                // Busca nos dados da Siscobra
                const valorSiscobra = obterValorNormalizado(linhaSiscobra, coluna);
                if (valorSiscobra !== undefined) {
                    resultado[coluna] = valorSiscobra;
                    return;
                }

                resultado[coluna] = "";
            });

            return resultado;
        });

        let encontrados = 0;
        let naoEncontrados = 0;

        bi.forEach(linhaBI => {
            const contrato = normalizarContrato(linhaBI[colunaContratoBI]);
            if (contrato && mapaSiscobra.has(contrato)) {
                encontrados++;
            } else {
                naoEncontrados++;
            }
        });

        rowsElement.textContent = resultadoFinal.length.toLocaleString("pt-BR");
        colsElement.textContent = ordemFinal.length;

        status.textContent =
            `✓ Concluído! ` +
            `${encontrados.toLocaleString("pt-BR")} contratos cruzados. ` +
            (usarAcordos ? `(Planilha de Acordos incluída via CPF)` : ``);

        downloadBtn.disabled = false;

    } catch (erro) {
        console.error("Erro completo:", erro);
        status.textContent = "✕ Erro: " + erro.message;
        resultadoFinal = null;
    } finally {
        processBtn.disabled = false;
        processBtn.textContent = "⚡ PROCESSAR EXCEL";
    }
});

// Download
downloadBtn.addEventListener("click", () => {
    if (!resultadoFinal) return;

    const worksheet = XLSX.utils.json_to_sheet(resultadoFinal);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, "Resultado");
    XLSX.writeFile(workbook, "resultado.xlsx");

    status.textContent = "✓ resultado.xlsx baixado com sucesso.";
});