let bufferReservas = [];
const camposDeFiltro = ["Data", "Reserva", "Historico", "Ficha", "UO", "NaturezaDespesa", "UE", "Processo", "ValorReserva", "Fonte", "SaldoReserva", "SaldoAtual", "ValorEmpenhado"];
let filtrosAplicados = {};
let limiteLinhasExibidas = 50;

const mapaColunas = {
    'B': 1, 'Q': 16, 'R': 17, 'W': 22, 'Y': 24, 
    'AJ': 35, 'AS': 44, 'BC': 54, 'BD': 55, 
    'BL': 63, 'BM': 64, 'BP': 67, 'BT': 71
};

document.addEventListener("DOMContentLoaded", () => {
    document.getElementById('uploadExcel').addEventListener('change', processarUploadExcel);
    document.getElementById('btnNovaReserva').addEventListener('click', abrirModalNovaReserva);
    document.getElementById('btnLimparFiltros').addEventListener('click', limparFiltros);
    document.getElementById('btnLimparDados').addEventListener('click', limparDadosCarregados);
    document.getElementById('btnExportar').addEventListener('click', () => exportarExcel(true));
    document.getElementById('btnImprimir').addEventListener('click', () => window.print());
    document.getElementById('btnFecharModal').addEventListener('click', fecharModal);
    document.getElementById('btnCancelarModal').addEventListener('click', fecharModal);
    document.getElementById('btnDataHoje').addEventListener('click', preencherDataHoje);
    document.getElementById('btnSalvarReserva').addEventListener('click', salvarReserva);
    
    document.getElementById('btnCarregarMais').addEventListener('click', () => {
        limiteLinhasExibidas += 50;
        renderizarTabela();
    });
    
    document.querySelectorAll('.input-money').forEach(input => {
        input.addEventListener('input', (e) => aplicarMascaraMoeda(e.target));
    });

    document.addEventListener('click', (event) => {
        if (!event.target.closest('.multi-select')) {
            document.querySelectorAll('.ms-dropdown').forEach(d => d.classList.remove('show'));
        }
    });

    window.addEventListener('beforeunload', (e) => {
        if (bufferReservas && bufferReservas.length > 0) {
            e.preventDefault();
            e.returnValue = ''; 
        }
    });
});

function aplicarMascaraMoeda(input) {
    let valor = input.value.replace(/\D/g, "");
    if (valor === "") {
        input.value = "";
        return;
    }
    valor = (parseFloat(valor) / 100).toFixed(2);
    input.value = valor.replace(".", ",").replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
}

function formatarDataExcel(val) {
    if (!val) return "";
    if (typeof val === 'number') {
        const dataObj = XLSX.SSF.parse_date_code(val);
        if (dataObj) {
            return `${String(dataObj.d).padStart(2, '0')}/${String(dataObj.m).padStart(2, '0')}/${dataObj.y}`;
        }
    }
    let str = String(val).trim();
    if (val instanceof Date && !isNaN(val)) {
        return `${String(val.getDate()).padStart(2, '0')}/${String(val.getMonth() + 1).padStart(2, '0')}/${val.getFullYear()}`;
    }
    if (str.includes('-') && str.length >= 10) {
        const partes = str.split('T')[0].split('-');
        if (partes.length === 3) return `${partes[2]}/${partes[1]}/${partes[0]}`;
    }
    if (str.includes('/')) {
        const partes = str.split('/');
        if (partes.length === 3) {
            let ano = partes[2];
            if (ano.length === 2) ano = '20' + ano;
            return `${partes[0].padStart(2, '0')}/${partes[1].padStart(2, '0')}/${ano}`;
        }
    }
    return str;
}

function processarUploadExcel(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {
                type: 'array', 
                cellDates: true,
                raw: false
            });

            if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
                alert("O arquivo selecionado está vazio ou em um formato incompatível.");
                return;
            }

            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            let jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1, defval: ""});
            
            bufferReservas = [];

            if (jsonData.length === 0) {
                alert("Nenhum dado foi encontrado nas linhas do arquivo.");
                return;
            }

            for (let i = 0; i < jsonData.length; i++) {
                let linha = jsonData[i];
                if (!linha || linha.length === 0) continue;

                let regData = linha[mapaColunas['B']] !== undefined ? linha[mapaColunas['B']] : (linha[1] || "");
                let regReserva = linha[mapaColunas['BT']] !== undefined ? linha[mapaColunas['BT']] : (linha[0] || "");
                
                let dataFormatada = formatarDataExcel(regData);
                let reservaTxt = String(regReserva).trim();

                if (dataIsoEInvalida(dataFormatada, reservaTxt) && i < 3) continue;

                bufferReservas.push({
                    Data: dataFormatada, 
                    Reserva: reservaTxt,
                    Historico: String(linha[mapaColunas['Q']] || linha[2] || "").trim(),
                    Ficha: String(linha[mapaColunas['R']] || linha[3] || "").trim(), 
                    UO: String(linha[mapaColunas['W']] || linha[4] || "").trim(),
                    NaturezaDespesa: String(linha[mapaColunas['Y']] || linha[5] || "").trim(), 
                    UE: String(linha[mapaColunas['AJ']] || linha[6] || "").trim(),
                    Processo: String(linha[mapaColunas['AS']] || linha[7] || "").trim(), 
                    ValorReserva: linha[mapaColunas['BP']] || linha[8] || "",
                    Fonte: String(linha[mapaColunas['BD']] || linha[9] || "").trim(), 
                    SaldoReserva: linha[mapaColunas['BL']] || linha[10] || "",
                    SaldoAtual: linha[mapaColunas['BM']] || linha[11] || "",
                    ValorEmpenhado: linha[mapaColunas['BC']] || linha[12] || ""
                });
            }

            filtrosAplicados = {};
            limiteLinhasExibidas = 50; 
            gerarFiltrosMultiplos();
            renderizarTabela();

            if (bufferReservas.length === 0) {
                alert("Arquivo lido, porém nenhuma linha de reserva válida foi identificada.");
            }
        } catch (err) {
            console.error("Erro ao processar arquivo:", err);
            alert("Erro ao ler o arquivo selecionado. Verifique se o formato é válido.");
        }
    };

    reader.readAsArrayBuffer(file);
}

function dataIsoEInvalida(data, reserva) {
    return (data.toLowerCase().includes("data") || reserva.toLowerCase().includes("reserva"));
}

function converterParaNumero(val) {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    let str = String(val).replace(/R\$\s?/g, '').trim();
    if (!str) return 0;
    if (str.includes('.') && str.includes(',')) str = str.replace(/\./g, '').replace(',', '.');
    else if (str.includes(',')) str = str.replace(',', '.');
    const num = parseFloat(str);
    return isNaN(num) ? 0 : num;
}

function formatarMoeda(val) {
    if (val === "" || val === null || val === undefined) return "R$ 0,00";
    return converterParaNumero(val).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function gerarFiltrosMultiplos() {
    camposDeFiltro.forEach(campo => {
        const container = document.getElementById('filtro_' + campo);
        if(!container) return;

        const valoresUnicos = [...new Set(bufferReservas.map(item => String(item[campo] || '').trim()))]
                              .filter(val => val !== '').sort();

        let html = `
            <button type="button" class="ms-btn" id="btn_drop_${campo}">
                <span id="txt_${campo}">Todos</span> <span>▾</span>
            </button>
            <div class="ms-dropdown" id="drop_${campo}">
                <div class="ms-search">
                    <input type="text" placeholder="Pesquisar...">
                </div>
                <div class="ms-options-container">
                    <label class="ms-select-all-label">
                        <input type="checkbox" class="ms-select-all" data-campo="${campo}"> 
                        (Selecionar Tudo)
                    </label>
        `;

        valoresUnicos.forEach(val => {
            const valEscaped = val.replace(/"/g, '&quot;');
            html += `
                <label class="ms-item-label">
                    <input type="checkbox" class="ms-item ms-item-${campo}" value="${valEscaped}"> 
                    ${val}
                </label>
            `;
        });

        html += `
                </div>
                <div class="ms-footer">
                    <button type="button" class="btn btn-success btn-ok">OK</button>
                    <button type="button" class="btn btn-danger btn-cancelar">Cancelar</button>
                </div>
            </div>`;
        container.innerHTML = html;
        
        const btnDrop = document.getElementById(`btn_drop_${campo}`);
        btnDrop.addEventListener('click', (e) => toggleDropdown(campo, e));
        
        const inputBusca = container.querySelector('.ms-search input');
        inputBusca.addEventListener('keyup', (e) => filtrarDropdownPesquisa(campo, e.target));
        
        const masterCb = container.querySelector('.ms-select-all');
        masterCb.addEventListener('change', (e) => toggleSelectAll(campo, e.target));
        
        container.querySelectorAll(`.ms-item-${campo}`).forEach(cb => {
            cb.addEventListener('change', () => verificarSelectAll(campo));
        });
        
        container.querySelector('.btn-ok').addEventListener('click', () => aplicarFiltro(campo));
        container.querySelector('.btn-cancelar').addEventListener('click', () => fecharDropdown(campo));

        atualizarTextoBotaoFiltro(campo);
    });
}

function toggleDropdown(campo, event) {
    event.stopPropagation();
    const drop = document.getElementById('drop_' + campo);
    const estaAberto = drop.classList.contains('show');
    document.querySelectorAll('.ms-dropdown').forEach(d => d.classList.remove('show'));
    if (!estaAberto) {
        sincronizarCheckboxesComFiltroAplicado(campo);
        drop.classList.add('show');
    }
}

function fecharDropdown(campo) {
    const drop = document.getElementById('drop_' + campo);
    if(drop) drop.classList.remove('show');
}

function sincronizarCheckboxesComFiltroAplicado(campo) {
    const checkboxes = document.querySelectorAll(`.ms-item-${campo}`);
    const master = document.querySelector(`.ms-select-all[data-campo="${campo}"]`);
    const valoresPermitidos = filtrosAplicados[campo];

    if (!valoresPermitidos) {
        if(master) master.checked = false;
        checkboxes.forEach(cb => cb.checked = false);
    } else {
        let todos = true;
        checkboxes.forEach(cb => {
            if (valoresPermitidos.includes(cb.value)) {
                cb.checked = true;
            } else {
                cb.checked = false;
                todos = false;
            }
        });
        if(master) master.checked = todos;
    }
}

function toggleSelectAll(campo, masterCheckbox) {
    const checkboxes = document.querySelectorAll(`.ms-item-${campo}`);
    checkboxes.forEach(cb => {
        if(cb.closest('label').style.display !== 'none') cb.checked = masterCheckbox.checked;
    });
}

function verificarSelectAll(campo) {
    const checkboxes = document.querySelectorAll(`.ms-item-${campo}`);
    const master = document.querySelector(`.ms-select-all[data-campo="${campo}"]`);
    let todosMarcados = true;
    checkboxes.forEach(cb => { if(!cb.checked) todosMarcados = false; });
    if(master) master.checked = todosMarcados;
}

function filtrarDropdownPesquisa(campo, input) {
    const termo = input.value.toLowerCase();
    document.querySelectorAll(`#drop_${campo} .ms-item-label`).forEach(label => {
        label.style.display = label.textContent.toLowerCase().includes(termo) ? '' : 'none';
    });
}

function atualizarTextoBotaoFiltro(campo) {
    const checkboxes = document.querySelectorAll(`.ms-item-${campo}`);
    const txtBox = document.getElementById(`txt_${campo}`);
    const btnBox = document.getElementById(`btn_drop_${campo}`);
    const valoresPermitidos = filtrosAplicados[campo];

    if (!valoresPermitidos || valoresPermitidos.length === 0) {
        txtBox.innerText = "Todos";
        if(btnBox) btnBox.classList.remove('active-filter');
    } else if (valoresPermitidos.length === checkboxes.length) {
        txtBox.innerText = "Todos";
        if(btnBox) btnBox.classList.remove('active-filter');
    } else if (valoresPermitidos.length === 1) {
        txtBox.innerText = valoresPermitidos[0];
        if(btnBox) btnBox.classList.add('active-filter');
    } else {
        txtBox.innerText = valoresPermitidos.length + " selecionados";
        if(btnBox) btnBox.classList.add('active-filter');
    }
}

function aplicarFiltro(campo) {
    const checkboxes = document.querySelectorAll(`.ms-item-${campo}`);
    const valoresMarcados = [];
    checkboxes.forEach(cb => { if(cb.checked) valoresMarcados.push(cb.value); });

    if (valoresMarcados.length === checkboxes.length || valoresMarcados.length === 0) {
        delete filtrosAplicados[campo];
    } else {
        filtrosAplicados[campo] = valoresMarcados;
    }

    limiteLinhasExibidas = 50; 
    atualizarTextoBotaoFiltro(campo);
    fecharDropdown(campo);
    renderizarTabela();
}

function obterDadosFiltrados() {
    return bufferReservas.filter(reg => {
        for (let campo in filtrosAplicados) {
            const valoresPermitidos = filtrosAplicados[campo];
            if (!valoresPermitidos || valoresPermitidos.length === 0) continue;
            const valorLinha = String(reg[campo] || '').trim().toLowerCase();
            let encontrou = false;
            for (let i = 0; i < valoresPermitidos.length; i++) {
                const valorPermitidoSanitizado = valoresPermitidos[i].replace(/&quot;/g, '"').trim().toLowerCase();
                if (valorLinha === valorPermitidoSanitizado) {
                    encontrou = true;
                    break;
                }
            }
            if (!encontrou) return false;
        }
        return true;
    });
}

function limparFiltros() { 
    filtrosAplicados = {};
    limiteLinhasExibidas = 50;
    
    camposDeFiltro.forEach(campo => {
        const master = document.querySelector(`.ms-select-all[data-campo="${campo}"]`);
        if(master) master.checked = false;
        
        document.querySelectorAll(`.ms-item-${campo}`).forEach(cb => {
            cb.checked = false;
        });

        const inputBusca = document.querySelector(`#drop_${campo} .ms-search input`);
        if (inputBusca) inputBusca.value = "";

        document.querySelectorAll(`#drop_${campo} .ms-item-label`).forEach(label => {
            label.style.display = '';
        });
        
        atualizarTextoBotaoFiltro(campo);
    });
    
    renderizarTabela(); 
}

function limparDadosCarregados() {
    if (bufferReservas.length === 0) {
        alert("Não há dados carregados para limpar.");
        return;
    }

    if (confirm("Tem certeza que deseja apagar TODOS os dados carregados na tabela? Esta ação não pode ser desfeita.")) {
        bufferReservas = []; 
        limiteLinhasExibidas = 50;
        
        const uploadInput = document.getElementById('uploadExcel');
        if (uploadInput) uploadInput.value = ""; 

        filtrosAplicados = {};

        camposDeFiltro.forEach(campo => {
            const container = document.getElementById('filtro_' + campo);
            if (container) container.innerHTML = '';
        });
        
        const infoEl = document.getElementById('infoRegistros');
        if (infoEl) infoEl.innerText = "";

        renderizarTabela();
        alert("Todos os dados e seleções de filtros foram limpos com sucesso.");
    }
}

// FUNÇÃO DE PROCESSAMENTO DO PAINEL DASHBOARD (APENAS CARDS DE UO + FONTE DESTACADA)
function atualizarPainelDashboard(dadosFiltrados) {
    const elPainel = document.getElementById('painelDashboard');
    const containerCards = document.getElementById('containerCardsDinamicos');
    if (!elPainel || !containerCards) return;

    if (!dadosFiltrados || dadosFiltrados.length === 0 || bufferReservas.length === 0) {
        elPainel.style.display = 'none';
        return;
    }

    elPainel.style.display = 'block';

    const combinacoesUOFonte = {};

    dadosFiltrados.forEach(reg => {
        const vReserva = converterParaNumero(reg.ValorReserva);

        let uoOriginal = String(reg.UO || "").trim();
        let secChave = uoOriginal.length >= 4 ? uoOriginal.substring(0, 4) : (uoOriginal || "N/A");
        let fonteChave = String(reg.Fonte || "N/A").trim();

        let chaveComposta = `${secChave}___${fonteChave}`;
        if (!combinacoesUOFonte[chaveComposta]) {
            combinacoesUOFonte[chaveComposta] = {
                sec: secChave,
                fonte: fonteChave,
                valor: 0
            };
        }
        combinacoesUOFonte[chaveComposta].valor += vReserva;
    });

    let htmlCards = '';
    const listaCombinacoes = Object.values(combinacoesUOFonte).sort((a, b) => b.valor - a.valor);

    listaCombinacoes.forEach(item => {
        htmlCards += `
            <div class="kpi-card">
                <div class="kpi-header">
                    <span class="kpi-secretaria">UO ${item.sec}</span>
                    <span class="kpi-fonte-tag">Fonte ${item.fonte}</span>
                </div>
                <span class="kpi-title">Valor Reserva</span>
                <span class="kpi-value">${formatarMoeda(item.valor)}</span>
            </div>
        `;
    });

    containerCards.innerHTML = htmlCards;
}

function renderizarTabela() {
    const tbody = document.getElementById('tabelaCorpo');
    const tfoot = document.getElementById('tabelaRodapeSubtotal');
    const infoEl = document.getElementById('infoRegistros');
    const btnCarregarMais = document.getElementById('btnCarregarMais');
    
    if(!tbody) return;

    if(tfoot) {
        tfoot.innerHTML = '';
    }

    const dadosFiltrados = obterDadosFiltrados();

    let somaValorReserva = 0;
    dadosFiltrados.forEach(reg => {
        somaValorReserva += converterParaNumero(reg.ValorReserva);
    });
    
    const bannerVal = document.getElementById('bannerValorReserva');
    if (bannerVal) bannerVal.innerText = formatarMoeda(somaValorReserva);

    atualizarPainelDashboard(dadosFiltrados);

    tbody.innerHTML = '';
    
    if (bufferReservas.length === 0) {
        tbody.innerHTML = '<tr><td colspan="14" class="empty-table-message">Nenhum arquivo carregado. Selecione um arquivo Excel acima para começar.</td></tr>';
        if (infoEl) infoEl.innerText = "";
        if (btnCarregarMais) btnCarregarMais.style.display = 'none';
        return;
    }

    if(dadosFiltrados.length === 0) {
        tbody.innerHTML = '<tr><td colspan="14" class="empty-table-message">Nenhum registro encontrado para os filtros atuais.</td></tr>';
        if (infoEl) infoEl.innerText = "Mostrando 0 de 0 registros filtrados";
        if (btnCarregarMais) btnCarregarMais.style.display = 'none';
        return;
    }

    const totalRegistrosVisiveis = Math.min(dadosFiltrados.length, limiteLinhasExibidas);
    
    if (infoEl) {
        infoEl.innerText = `Exibindo ${totalRegistrosVisiveis} de ${dadosFiltrados.length} registros encontrados`;
        infoEl.style.color = "#64748b";
    }

    for(let i = 0; i < totalRegistrosVisiveis; i++) {
        const reg = dadosFiltrados[i];
        const indexReal = bufferReservas.indexOf(reg);
        
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="col-acoes">
                <button type="button" class="btn btn-warning btn-sm btn-editar">Editar</button>
                <button type="button" class="btn btn-danger btn-sm btn-excluir">Excluir</button>
            </td>
            <td class="col-nao-saldo col-print-data" style="text-align: center;">${reg.Data || ''}</td>
            <td class="col-nao-saldo col-print-reserva" style="text-align: center;"><b>${reg.Reserva || ''}</b></td>
            <td class="col-nao-saldo col-print-historico col-texto-longo">${reg.Historico || ''}</td>
            <td class="col-nao-saldo col-print-ficha" style="text-align: center;">${reg.Ficha || ''}</td>
            <td class="col-nao-saldo col-print-uo" style="text-align: center;">${reg.UO || ''}</td>
            <td class="col-nao-saldo col-print-natureza">${reg.NaturezaDespesa || ''}</td>
            <td class="col-nao-saldo col-print-ue" style="text-align: center;">${reg.UE || ''}</td>
            <td class="col-nao-saldo col-print-processo col-texto-longo">${reg.Processo || ''}</td>
            <td class="col-nao-saldo col-print-vlr-reserva col-valores" style="color: #6d28d9;">${formatarMoeda(reg.ValorReserva)}</td>
            <td class="col-nao-saldo col-print-fonte" style="text-align: center;">${reg.Fonte || ''}</td>
            <td class="col-nao-saldo col-valores">${formatarMoeda(reg.SaldoReserva)}</td>
            <td class="col-saldo-atual col-print-sld-atual col-valores" style="color: #047857;">${formatarMoeda(reg.SaldoAtual)}</td>
            <td class="col-nao-saldo col-valores">${formatarMoeda(reg.ValorEmpenhado)}</td>
        `;
        
        const btnEdit = tr.querySelector('.btn-editar');
        const btnDel = tr.querySelector('.btn-excluir');
        if(btnEdit) btnEdit.addEventListener('click', () => editarReserva(indexReal));
        if(btnDel) btnDel.addEventListener('click', () => excluirReserva(indexReal));
        
        tbody.appendChild(tr);
    }

    if (btnCarregarMais) {
        if (dadosFiltrados.length > totalRegistrosVisiveis) {
            const restantes = dadosFiltrados.length - totalRegistrosVisiveis;
            btnCarregarMais.innerText = `Carregar Mais (${Math.min(restantes, 50)} de ${restantes} restantes)`;
            btnCarregarMais.style.display = "inline-block";
        } else {
            btnCarregarMais.style.display = "none";
        }
    }
}

function abrirModalNovaReserva() {
    document.getElementById('modalTitulo').innerText = "Nova Reserva";
    document.getElementById('editIndex').value = "-1";
    document.getElementById('formReserva').reset();
    document.getElementById('modalFormulario').style.display = "flex";
}

function preencherDataHoje() {
    const hoje = new Date();
    const ano = hoje.getFullYear();
    const mes = String(hoje.getMonth() + 1).padStart(2, '0');
    const dia = String(hoje.getDate()).padStart(2, '0');
    document.getElementById('form_Data').value = `${ano}-${mes}-${dia}`;
}

function editarReserva(index) {
    const reg = bufferReservas[index];
    document.getElementById('modalTitulo').innerText = "Editar Reserva";
    document.getElementById('editIndex').value = index;
    
    Object.keys(reg).forEach(key => {
        const input = document.getElementById('form_' + key);
        if(input && key !== 'Data') {
            input.value = reg[key];
            if(['ValorEmpenhado', 'ValorReserva', 'SaldoReserva', 'SaldoAtual'].includes(key)){
                aplicarMascaraMoeda(input);
            }
        }
    });

    const inputData = document.getElementById('form_Data');
    if (reg.Data && reg.Data.includes('/')) {
        const partes = reg.Data.split('/');
        if (partes.length === 3) {
            inputData.value = `${partes[2]}-${partes[1]}-${partes[0]}`;
        }
    }

    window.scrollTo({ top: 0, behavior: 'smooth' });
    document.getElementById('modalFormulario').style.display = "flex";
}

function fecharModal() { 
    document.getElementById('modalFormulario').style.display = "none"; 
}

function salvarReserva() {
    const index = parseInt(document.getElementById('editIndex').value);
    const dataIso = document.getElementById('form_Data').value;
    
    let dataFormatada = "";
    if (dataIso) {
        const partes = dataIso.split('-');
        if (partes.length === 3) dataFormatada = `${partes[2]}/${partes[1]}/${partes[0]}`;
    }

    const novoRegistro = {
        Data: dataFormatada, 
        Reserva: document.getElementById('form_Reserva').value,
        Historico: document.getElementById('form_Historico').value,
        Ficha: document.getElementById('form_Ficha').value, 
        UO: document.getElementById('form_UO').value,
        NaturezaDespesa: document.getElementById('form_NaturezaDespesa').value, 
        UE: document.getElementById('form_UE').value,
        Processo: document.getElementById('form_Processo').value, 
        ValorReserva: document.getElementById('form_ValorReserva').value,
        Fonte: document.getElementById('form_Fonte').value, 
        SaldoReserva: document.getElementById('form_SaldoReserva').value,
        SaldoAtual: document.getElementById('form_SaldoAtual').value,
        ValorEmpenhado: document.getElementById('form_ValorEmpenhado').value
    };

    if(index >= 0) bufferReservas[index] = novoRegistro; else bufferReservas.push(novoRegistro);
    fecharModal(); gerarFiltrosMultiplos(); renderizarTabela();
}

function excluirReserva(index) {
    if (confirm("Deseja excluir esta reserva?")) {
        bufferReservas.splice(index, 1);
        gerarFiltrosMultiplos(); renderizarTabela();
    }
}

function exportarExcel(apenasVisiveis) {
    if (!bufferReservas.length) return alert("Não há dados para exportar.");
    const dadosParaExportar = apenasVisiveis ? obterDadosFiltrados() : bufferReservas;
    
    const cabecalho = [
        "Data", "Reserva", "Histórico", "Ficha", "UO", "Natureza Desp.", "UE", "Processo", 
        "Valor Reserva", "Fonte", "Saldo Reserva", "Saldo Atual", "Valor Empenhado"
    ];
    
    const dadosFinais = [cabecalho];
    dadosParaExportar.forEach(reg => {
        dadosFinais.push([
            reg.Data || "", reg.Reserva || "", reg.Historico || "", reg.Ficha || "",
            reg.UO || "", reg.NaturezaDespesa || "", reg.UE || "", reg.Processo || "",
            converterParaNumero(reg.ValorReserva), reg.Fonte || "", converterParaNumero(reg.SaldoReserva),
            converterParaNumero(reg.SaldoAtual), converterParaNumero(reg.ValorEmpenhado)
        ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(dadosFinais);
    
    for (let R = 1; R < dadosFinais.length; ++R) {
        const cellData = XLSX.utils.encode_cell({c: 0, r: R});
        if (ws[cellData]) {
            ws[cellData].z = 'dd/mm/yyyy';
            const partes = String(ws[cellData].v).split('/');
            if (partes.length === 3) {
                const d = new Date(partes[2], partes[1] - 1, partes[0]);
                if (!isNaN(d)) {
                    ws[cellData].v = d;
                    ws[cellData].t = 'd';
                }
            }
        }

        const colunasMonetarias = [8, 10, 11, 12];
        colunasMonetarias.forEach(C => {
            const cell = XLSX.utils.encode_cell({c: C, r: R});
            if (ws[cell]) ws[cell].z = '"R$"#,##0.00;("R$"#,##0.00);"-"';
        });
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Reservas");
    XLSX.writeFile(wb, "Relatorio_Reservas.xlsx");
}