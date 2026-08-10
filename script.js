let bufferReservas = [];
const camposDeFiltro = ["Data", "Reserva", "Historico", "Ficha", "UO", "NaturezaDespesa", "UE", "Processo", "ValorReserva", "Fonte", "SaldoReserva", "SaldoAtual", "ValorEmpenhado"];
let filtrosAplicados = {};
let limiteLinhasExibidas = 50;

let graficoInstancia = null; 

const dicionarioUO = {
    "0201": "GABINETE",
    "0202": "HABITAÇÃO",
    "0204": "EDUCAÇÃO",
    "0206": "SAÚDE",
    "0207": "ESPORTES",
    "0208": "SEGURANÇA",
    "0209": "ASSIST. SOCIAL",
    "0210": "FUNDO ASSIST. SOCIAL",
    "0211": "CULTURA",
    "0212": "INFRAESTRUTURA",
    "0221": "ENCARGOS GERAIS",
    "0232": "PROCURADORIA",
    "0234": "DESENVOLVIMENTO",
    "0235": "ZELADORIA",
    "0236": "GOVERNO",
    "0237": "ADMINISTRAÇÃO",
    "0238": "FAZENDA",
    "0239": "COMUNICAÇÃO",
    "0240": "TURISMO",
    "0241": "MEIO AMBIENTE",
    "0242": "AGRICULTURA"
};

function obterNomeSecretaria(codigoUO) {
    let cod = String(codigoUO).replace(/\D/g, '').substring(0, 4);
    return dicionarioUO[cod] ? `${cod}-${dicionarioUO[cod]}` : `UO ${codigoUO}`;
}

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
    document.getElementById('btnImprimir').addEventListener('click', executarImpressaoOtimizada);
    
    document.getElementById('btnFecharModal').addEventListener('click', fecharModal);
    document.getElementById('btnCancelarModal').addEventListener('click', fecharModal);
    document.getElementById('btnFecharDetalhes').addEventListener('click', fecharModalDetalhes);
    document.getElementById('btnFecharDetalhesBottom').addEventListener('click', fecharModalDetalhes);
    
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

    const btnDark = document.getElementById('btnDarkMode');
    btnDark.addEventListener('click', () => {
        document.body.classList.toggle('dark-theme');
        const isDark = document.body.classList.contains('dark-theme');
        btnDark.innerText = isDark ? '☀️ Claro' : '🌙 Escuro';
        
        if (graficoInstancia) {
            graficoInstancia.options.plugins.title.color = isDark ? '#f8fafc' : '#1e293b';
            graficoInstancia.options.scales.x.ticks.color = isDark ? '#cbd5e1' : '#64748b';
            graficoInstancia.options.scales.y.ticks.color = isDark ? '#f8fafc' : '#334155';
            if (graficoInstancia.options.plugins.datalabels) {
                graficoInstancia.options.plugins.datalabels.color = isDark ? '#cbd5e1' : '#475569';
            }
            graficoInstancia.update();
        }
    });
});

function executarImpressaoOtimizada() {
    const dadosFiltrados = obterDadosFiltrados();
    if (!dadosFiltrados || dadosFiltrados.length === 0) return alert("Não há dados para imprimir.");

    const limiteAnterior = limiteLinhasExibidas;
    limiteLinhasExibidas = dadosFiltrados.length;
    renderizarTabela();

    setTimeout(() => {
        window.print();
        limiteLinhasExibidas = limiteAnterior;
        renderizarTabela();
    }, 50);
}

function aplicarMascaraMoeda(input) {
    let valor = input.value.replace(/\D/g, "");
    if (valor === "") { input.value = ""; return; }
    valor = (parseFloat(valor) / 100).toFixed(2);
    input.value = valor.replace(".", ",").replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1.");
}

function formatarDataExcel(val) {
    if (!val) return "";
    if (typeof val === 'number') {
        const dataObj = XLSX.SSF.parse_date_code(val);
        if (dataObj) return `${String(dataObj.d).padStart(2, '0')}/${String(dataObj.m).padStart(2, '0')}/${dataObj.y}`;
    }
    let str = String(val).trim();
    if (val instanceof Date && !isNaN(val)) return `${String(val.getDate()).padStart(2, '0')}/${String(val.getMonth() + 1).padStart(2, '0')}/${val.getFullYear()}`;
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
            const workbook = XLSX.read(data, { type: 'array', cellDates: true, raw: false });

            if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) return alert("Arquivo incompatível.");

            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            let jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1, defval: ""});
            
            bufferReservas = [];
            if (jsonData.length === 0) return alert("Nenhum dado encontrado.");

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

        } catch (err) {
            console.error(err);
            alert("Erro ao ler o arquivo.");
        }
    };
    reader.readAsArrayBuffer(file);
}

function dataIsoEInvalida(data, reserva) { return (data.toLowerCase().includes("data") || reserva.toLowerCase().includes("reserva")); }

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

        const valoresUnicos = [...new Set(bufferReservas.map(item => String(item[campo] || '').trim()))].filter(val => val !== '').sort();

        let html = `
            <button type="button" class="ms-btn" id="btn_drop_${campo}">
                <span id="txt_${campo}">Todos</span> <span>▾</span>
            </button>
            <div class="ms-dropdown" id="drop_${campo}">
                <div class="ms-search"><input type="text" placeholder="Pesquisar..."></div>
                <div class="ms-options-container">
                    <label class="ms-select-all-label"><input type="checkbox" class="ms-select-all" data-campo="${campo}"> (Selecionar Tudo)</label>
        `;

        valoresUnicos.forEach(val => {
            const valEscaped = val.replace(/"/g, '&quot;');
            html += `<label class="ms-item-label"><input type="checkbox" class="ms-item ms-item-${campo}" value="${valEscaped}"> ${val}</label>`;
        });

        html += `</div>
                <div class="ms-footer">
                    <button type="button" class="btn btn-success btn-ok">OK</button>
                    <button type="button" class="btn btn-danger btn-cancelar">Cancelar</button>
                </div>
            </div>`;
        container.innerHTML = html;
        
        container.querySelector(`#btn_drop_${campo}`).addEventListener('click', (e) => toggleDropdown(campo, e));
        container.querySelector('.ms-search input').addEventListener('keyup', (e) => filtrarDropdownPesquisa(campo, e.target));
        container.querySelector('.ms-select-all').addEventListener('change', (e) => toggleSelectAll(campo, e.target));
        container.querySelectorAll(`.ms-item-${campo}`).forEach(cb => cb.addEventListener('change', () => verificarSelectAll(campo)));
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
    if (!estaAberto) { sincronizarCheckboxesComFiltroAplicado(campo); drop.classList.add('show'); }
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
            if (valoresPermitidos.includes(cb.value)) cb.checked = true;
            else { cb.checked = false; todos = false; }
        });
        if(master) master.checked = todos;
    }
}

function toggleSelectAll(campo, masterCheckbox) {
    document.querySelectorAll(`.ms-item-${campo}`).forEach(cb => {
        if(cb.closest('label').style.display !== 'none') cb.checked = masterCheckbox.checked;
    });
}

function verificarSelectAll(campo) {
    const master = document.querySelector(`.ms-select-all[data-campo="${campo}"]`);
    let todos = true;
    document.querySelectorAll(`.ms-item-${campo}`).forEach(cb => { if(!cb.checked) todos = false; });
    if(master) master.checked = todos;
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
    const val = filtrosAplicados[campo];

    if (!val || val.length === 0 || val.length === checkboxes.length) {
        txtBox.innerText = "Todos";
        if(btnBox) btnBox.classList.remove('active-filter');
    } else if (val.length === 1) {
        txtBox.innerText = val[0];
        if(btnBox) btnBox.classList.add('active-filter');
    } else {
        txtBox.innerText = val.length + " selecionados";
        if(btnBox) btnBox.classList.add('active-filter');
    }
}

function aplicarFiltro(campo) {
    const checkboxes = document.querySelectorAll(`.ms-item-${campo}`);
    const valoresMarcados = [];
    checkboxes.forEach(cb => { if(cb.checked) valoresMarcados.push(cb.value); });

    if (valoresMarcados.length === checkboxes.length || valoresMarcados.length === 0) delete filtrosAplicados[campo];
    else filtrosAplicados[campo] = valoresMarcados;

    limiteLinhasExibidas = 50; 
    atualizarTextoBotaoFiltro(campo);
    fecharDropdown(campo);
    renderizarTabela();
}

function obterDadosFiltrados() {
    return bufferReservas.filter(reg => {
        for (let campo in filtrosAplicados) {
            const valPermitidos = filtrosAplicados[campo];
            if (!valPermitidos || valPermitidos.length === 0) continue;
            const valorLinha = String(reg[campo] || '').trim().toLowerCase();
            let encontrou = false;
            for (let i = 0; i < valPermitidos.length; i++) {
                if (valorLinha === valPermitidos[i].replace(/&quot;/g, '"').trim().toLowerCase()) {
                    encontrou = true; break;
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
        document.querySelectorAll(`.ms-item-${campo}`).forEach(cb => cb.checked = false);
        const inputBusca = document.querySelector(`#drop_${campo} .ms-search input`);
        if (inputBusca) inputBusca.value = "";
        document.querySelectorAll(`#drop_${campo} .ms-item-label`).forEach(l => l.style.display = '');
        atualizarTextoBotaoFiltro(campo);
    });
    renderizarTabela(); 
}

function limparDadosCarregados() {
    if (bufferReservas.length === 0) return alert("Não há dados carregados para limpar.");
    if (confirm("Apagar TODOS os dados carregados?")) {
        bufferReservas = []; limiteLinhasExibidas = 50;
        document.getElementById('uploadExcel').value = ""; 
        filtrosAplicados = {};
        camposDeFiltro.forEach(campo => document.getElementById('filtro_' + campo).innerHTML = '');
        document.getElementById('infoRegistros').innerText = "";
        renderizarTabela();
    }
}

// ============== RENDERIZAÇÃO DO GRÁFICO (BARRAS SIMPLES COM AUTO-SKIP FALSE E DATALABELS EXTERNOS) ==============
function renderizarGrafico(dadosFiltrados) {
    const ctx = document.getElementById('graficoReservas');
    if (!ctx) return;
    if (graficoInstancia) graficoInstancia.destroy();
    if (!dadosFiltrados || dadosFiltrados.length === 0) return;

    const uoTotais = {}; 
    let totalFiltrado = 0;

    // Processamento dos dados agrupados apenas por Secretaria (UO)
    dadosFiltrados.forEach(reg => {
        let uo = String(reg.UO || "").replace(/\D/g, '').substring(0, 4);
        if (!uo) uo = "N/A";
        let valor = converterParaNumero(reg.ValorReserva);

        if (!uoTotais[uo]) uoTotais[uo] = 0;
        uoTotais[uo] += valor;
        totalFiltrado += valor;
    });

    // Filtra apenas as UOs que possuem valor > 0
    const uosValidas = Object.keys(uoTotais).filter(uo => uoTotais[uo] > 0);

    // Ordenar UOs da maior para a menor
    const uosOrdenadas = uosValidas.sort((a, b) => uoTotais[b] - uoTotais[a]);
    const labels = uosOrdenadas.map(uo => obterNomeSecretaria(uo));
    const valores = uosOrdenadas.map(uo => uoTotais[uo]);
    const isDark = document.body.classList.contains('dark-theme');
    
    // Paleta de cores moderna para as barras
    const paletaCores = [
        '#4361ee', '#3a0ca3', '#7209b7', '#f72585', '#4cc9f0', 
        '#2ec4b6', '#ff9f1c', '#e71d36', '#fb8500', '#06d6a0', 
        '#118ab2', '#073b4c', '#8338ec', '#ff006e', '#8ac926', 
        '#1982c4', '#6a4c93', '#ff595e', '#ffca3a', '#10002b'
    ];

    graficoInstancia = new Chart(ctx, {
        type: 'bar',
        plugins: [ChartDataLabels],
        data: {
            labels: labels,
            datasets: [{
                label: 'Valor Reserva',
                data: valores,
                backgroundColor: paletaCores,
                borderRadius: 6,
                barThickness: 16
            }]
        },
        options: {
            indexAxis: 'y', // Define o gráfico como BARRAS HORIZONTAIS
            responsive: true,
            maintainAspectRatio: false,
            // Aumentado o padding direito para dar espaço para o texto do rótulo
            layout: { padding: { top: 10, right: 65, bottom: 10, left: 10 } },
            plugins: {
                legend: { display: false }, // Sem legenda inferior (não está empilhado)
                datalabels: {
                    anchor: 'end',
                    align: 'end',
                    color: isDark ? '#cbd5e1' : '#475569',
                    font: { size: 10, weight: '700' },
                    formatter: function(value) {
                        if (value >= 1000000) {
                            return 'R$ ' + (value / 1000000).toFixed(1).replace('.', ',') + 'M';
                        } else if (value >= 1000) {
                            return 'R$ ' + (value / 1000).toFixed(1).replace('.', ',') + 'k';
                        }
                        return 'R$ ' + value.toLocaleString('pt-BR');
                    }
                },
                tooltip: {
                    callbacks: {
                        label: function(c) {
                            let pct = totalFiltrado > 0 ? ((c.raw / totalFiltrado) * 100).toFixed(1) : "0.0";
                            let val = c.raw.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
                            return ` ${val} (${pct}%)`;
                        }
                    }
                },
                title: { 
                    display: true, 
                    text: 'Distribuição do Valor de Reserva por Secretaria', 
                    color: isDark ? '#f8fafc' : '#1e293b',
                    font: { size: 13, weight: '600' }, 
                    padding: { bottom: 15 } 
                }
            },
            scales: {
                x: {
                    grid: { color: isDark ? '#334155' : '#f1f5f9' },
                    ticks: {
                        color: isDark ? '#cbd5e1' : '#64748b',
                        font: { size: 10 },
                        callback: function(v) {
                            return 'R$ ' + (v / 1000).toLocaleString('pt-BR') + 'k';
                        }
                    }
                },
                y: {
                    grid: { display: false },
                    ticks: {
                        autoSkip: false, // <-- GARANTE QUE NENHUMA SECRETARIA SEJA ESCONDIDA
                        color: isDark ? '#f8fafc' : '#334155',
                        font: { size: 10, weight: '600' }
                    }
                }
            }
        }
    });
}

function atualizarPainelDashboard(dadosFiltrados) {
    const elPainel = document.getElementById('painelDashboard');
    const elGrafico = document.getElementById('painelGrafico');
    const containerCards = document.getElementById('containerCardsDinamicos');
    
    if (!elPainel || !containerCards) return;

    if (!dadosFiltrados || dadosFiltrados.length === 0 || bufferReservas.length === 0) {
        elPainel.style.display = 'none';
        if(elGrafico) elGrafico.style.display = 'none';
        return;
    }

    elPainel.style.display = 'block';
    if(elGrafico) elGrafico.style.display = 'flex';

    const combinacoesUOFonte = {};

    dadosFiltrados.forEach(reg => {
        const vReserva = converterParaNumero(reg.ValorReserva);
        let secChave = String(reg.UO || "").replace(/\D/g, '').substring(0, 4); 
        if (!secChave) secChave = "N/A";
        let fonteChave = String(reg.Fonte || "N/A").trim();

        let chaveComposta = `${secChave}___${fonteChave}`;
        if (!combinacoesUOFonte[chaveComposta]) combinacoesUOFonte[chaveComposta] = { sec: secChave, fonte: fonteChave, valor: 0 };
        combinacoesUOFonte[chaveComposta].valor += vReserva;
    });

    const listaCombinacoes = Object.values(combinacoesUOFonte).sort((a, b) => {
        const compSec = a.sec.localeCompare(b.sec, undefined, { numeric: true, sensitivity: 'base' });
        if (compSec !== 0) return compSec;
        return a.fonte.localeCompare(b.fonte, undefined, { numeric: true, sensitivity: 'base' });
    });

    let htmlCards = '';
    listaCombinacoes.forEach(item => {
        let nomeSecretariaCard = obterNomeSecretaria(item.sec); 
        htmlCards += `
            <div class="kpi-card">
                <div class="kpi-header">
                    <span class="kpi-secretaria" title="${nomeSecretariaCard}">${nomeSecretariaCard}</span>
                    <span class="kpi-fonte-tag">Fonte ${item.fonte}</span>
                </div>
                <span class="kpi-title">Valor Reserva</span>
                <span class="kpi-value">${formatarMoeda(item.valor)}</span>
            </div>`;
    });

    containerCards.innerHTML = htmlCards;
    renderizarGrafico(dadosFiltrados); 
}

function renderizarTabela() {
    const tbody = document.getElementById('tabelaCorpo');
    const infoEl = document.getElementById('infoRegistros');
    const btnCarregarMais = document.getElementById('btnCarregarMais');
    
    if(!tbody) return;

    const dadosFiltrados = obterDadosFiltrados();

    let somaValorReserva = 0;
    dadosFiltrados.forEach(reg => { somaValorReserva += converterParaNumero(reg.ValorReserva); });
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
    if (infoEl) infoEl.innerText = `Exibindo ${totalRegistrosVisiveis} de ${dadosFiltrados.length} registros encontrados`;

    for(let i = 0; i < totalRegistrosVisiveis; i++) {
        const reg = dadosFiltrados[i];
        const indexReal = bufferReservas.indexOf(reg);
        
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="col-acoes" style="width: 100px; min-width: 100px;">
                <button type="button" class="btn btn-info btn-sm btn-ver">Ficha</button>
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
            <td class="col-saldo-atual col-print-sld-atual col-valores cor-saldo-atual">${formatarMoeda(reg.SaldoAtual)}</td>
            <td class="col-nao-saldo col-valores">${formatarMoeda(reg.ValorEmpenhado)}</td>
        `;
        
        tr.querySelector('.btn-ver').addEventListener('click', () => abrirModalDetalhes(indexReal));
        tr.querySelector('.btn-editar').addEventListener('click', () => editarReserva(indexReal));
        tr.querySelector('.btn-excluir').addEventListener('click', () => excluirReserva(indexReal));
        tbody.appendChild(tr);
    }

    if (btnCarregarMais) {
        if (dadosFiltrados.length > totalRegistrosVisiveis) {
            const restantes = dadosFiltrados.length - totalRegistrosVisiveis;
            btnCarregarMais.innerText = `Carregar Mais (${Math.min(restantes, 50)} de ${restantes} restantes)`;
            btnCarregarMais.style.display = "inline-block";
        } else btnCarregarMais.style.display = "none";
    }
}

function abrirModalDetalhes(index) {
    const reg = bufferReservas[index];
    const container = document.getElementById('conteudoDetalhes');
    const uoNome = obterNomeSecretaria(reg.UO);
    
    document.body.classList.add('imprimindo-ficha');
    
    container.innerHTML = `
        <div class="detalhes-grid">
            <div class="detalhe-item">
                <span class="detalhe-label">Reserva N.º</span>
                <span class="detalhe-valor" style="font-size: 18px; font-weight: 700; color: var(--primary);">${reg.Reserva || 'N/A'}</span>
            </div>
            <div class="detalhe-item">
                <span class="detalhe-label">Data</span>
                <span class="detalhe-valor">${reg.Data || 'N/A'}</span>
            </div>
            <div class="detalhe-item full-width">
                <span class="detalhe-label">Secretaria / Unidade Orçamentária</span>
                <span class="detalhe-valor">${uoNome}</span>
            </div>
            <div class="detalhe-item full-width">
                <span class="detalhe-label">Histórico</span>
                <span class="detalhe-valor">${reg.Historico || 'N/A'}</span>
            </div>
            
            <div class="detalhe-item">
                <span class="detalhe-label">Ficha</span>
                <span class="detalhe-valor">${reg.Ficha || 'N/A'}</span>
            </div>
            <div class="detalhe-item">
                <span class="detalhe-label">Fonte</span>
                <span class="detalhe-valor">${reg.Fonte || 'N/A'}</span>
            </div>
            <div class="detalhe-item full-width">
                <span class="detalhe-label">Processo</span>
                <span class="detalhe-valor">${reg.Processo || 'N/A'}</span>
            </div>
            <div class="detalhe-item full-width">
                <span class="detalhe-label">Natureza de Despesa</span>
                <span class="detalhe-valor">${reg.NaturezaDespesa || 'N/A'}</span>
            </div>
            
            <div class="detalhe-item">
                <span class="detalhe-label">Valor Reserva</span>
                <span class="detalhe-valor">${formatarMoeda(reg.ValorReserva)}</span>
            </div>
            <div class="detalhe-item">
                <span class="detalhe-label">Valor Empenhado</span>
                <span class="detalhe-valor">${formatarMoeda(reg.ValorEmpenhado)}</span>
            </div>
            <div class="detalhe-item">
                <span class="detalhe-label">Saldo Reserva</span>
                <span class="detalhe-valor">${formatarMoeda(reg.SaldoReserva)}</span>
            </div>
            <div class="detalhe-item">
                <span class="detalhe-label">Saldo Atual da Ficha</span>
                <span class="detalhe-valor destaque">${formatarMoeda(reg.SaldoAtual)}</span>
            </div>
        </div>
    `;

    document.getElementById('modalDetalhes').style.display = "flex";
}

function fecharModalDetalhes() {
    document.body.classList.remove('imprimindo-ficha');
    document.getElementById('modalDetalhes').style.display = "none";
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
            if(['ValorEmpenhado', 'ValorReserva', 'SaldoReserva', 'SaldoAtual'].includes(key)) aplicarMascaraMoeda(input);
        }
    });
    const inputData = document.getElementById('form_Data');
    if (reg.Data && reg.Data.includes('/')) {
        const partes = reg.Data.split('/');
        if (partes.length === 3) inputData.value = `${partes[2]}-${partes[1]}-${partes[0]}`;
    }
    window.scrollTo({ top: 0, behavior: 'smooth' });
    document.getElementById('modalFormulario').style.display = "flex";
}

function fecharModal() { document.getElementById('modalFormulario').style.display = "none"; }

function salvarReserva() {
    const index = parseInt(document.getElementById('editIndex').value);
    const dataIso = document.getElementById('form_Data').value;
    let dataFormatada = "";
    if (dataIso) {
        const partes = dataIso.split('-');
        if (partes.length === 3) dataFormatada = `${partes[2]}/${partes[1]}/${partes[0]}`;
    }

    const novoRegistro = {
        Data: dataFormatada, Reserva: document.getElementById('form_Reserva').value,
        Historico: document.getElementById('form_Historico').value, Ficha: document.getElementById('form_Ficha').value, 
        UO: document.getElementById('form_UO').value, NaturezaDespesa: document.getElementById('form_NaturezaDespesa').value, 
        UE: document.getElementById('form_UE').value, Processo: document.getElementById('form_Processo').value, 
        ValorReserva: document.getElementById('form_ValorReserva').value, Fonte: document.getElementById('form_Fonte').value, 
        SaldoReserva: document.getElementById('form_SaldoReserva').value, SaldoAtual: document.getElementById('form_SaldoAtual').value,
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
    const cabecalho = ["Data", "Reserva", "Histórico", "Ficha", "UO", "Natureza Desp.", "UE", "Processo", "Valor Reserva", "Fonte", "Saldo Reserva", "Saldo Atual", "Valor Empenhado"];
    const dadosFinais = [cabecalho];
    dadosParaExportar.forEach(reg => {
        dadosFinais.push([reg.Data || "", reg.Reserva || "", reg.Historico || "", reg.Ficha || "", reg.UO || "", reg.NaturezaDespesa || "", reg.UE || "", reg.Processo || "", converterParaNumero(reg.ValorReserva), reg.Fonte || "", converterParaNumero(reg.SaldoReserva), converterParaNumero(reg.SaldoAtual), converterParaNumero(reg.ValorEmpenhado)]);
    });
    const ws = XLSX.utils.aoa_to_sheet(dadosFinais);
    for (let R = 1; R < dadosFinais.length; ++R) {
        const cellData = XLSX.utils.encode_cell({c: 0, r: R});
        if (ws[cellData]) {
            ws[cellData].z = 'dd/mm/yyyy';
            const partes = String(ws[cellData].v).split('/');
            if (partes.length === 3) {
                const d = new Date(partes[2], partes[1] - 1, partes[0]);
                if (!isNaN(d)) { ws[cellData].v = d; ws[cellData].t = 'd'; }
            }
        }
        [8, 10, 11, 12].forEach(C => {
            const cell = XLSX.utils.encode_cell({c: C, r: R});
            if (ws[cell]) ws[cell].z = '"R$"#,##0.00;("R$"#,##0.00);"-"';
        });
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Reservas");
    XLSX.writeFile(wb, "Relatorio_Reservas.xlsx");
}