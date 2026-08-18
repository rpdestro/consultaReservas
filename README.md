# Sistema de Consulta de Reservas Orçamentárias

O **Sistema de Consulta de Reservas Orçamentárias** é uma aplicação web focada na extração, visualização e análise de dados financeiros. Seu principal objetivo é importar planilhas de sistemas de gestão pública e permitir a consulta dinâmica das reservas, aplicando filtros múltiplos e exibindo os valores sumarizados por Secretaria (Unidade Orçamentária - UO), Fonte de Recurso, Processo, entre outros.

## 

## 🚀 Funcionalidades

O sistema foi projetado para exibir as reservas de forma clara e organizada, transformando planilhas densas em painéis executivos interativos. As principais funcionalidades incluem:

* **Leitura Dinâmica:** Upload de relatórios `.xls` ou `.xlsx` gerados pelo sistema ERP, com mapeamento automático de colunas.
* **Filtros Inteligentes:** Capacidade de isolar informações cruzando dados como Data, Reserva, Histórico, Ficha, UO, Natureza da Despesa, Processo e Fonte.
* **Dashboard Visual:** Renderização de um gráfico de barras horizontais e cards totalizadores (KPIs) com o somatório dos valores em tempo real, organizados hierarquicamente.
* **Modo de Impressão:** Geração de uma "Ficha de Reserva" estruturada e otimizada visualmente para impressão física ou arquivamento em PDF.
* **Exportação:** Exportação dos dados filtrados de volta para o formato Excel.

## 

## 💻 Tecnologias Utilizadas

A aplicação mantém uma arquitetura intencionalmente "Client-Side" (processamento local), visando alta velocidade e dispensando a necessidade de infraestrutura de servidores de retaguarda (*backend*). O código é simples, objetivo e fundamentado em:

* **HTML5:** Estrutura e semântica da página.
* **CSS3:** Estilização visual, responsividade, modo noturno e regras específicas de paginação (`@media print`).
* **JavaScript (Vanilla):** Lógica central de negócio, abrangendo o processamento do array de objetos (`bufferReservas`), tratamentos de strings (Regex), formatação de moedas e controle de estado do DOM.
* **Bibliotecas Externas via CDN:**

  * *SheetJS (`xlsx.full.min.js`):* Para manipulação e conversão dos binários do Excel.
  * *Chart.js \& ChartDataLabels:* Para a construção do gráfico analítico.

## 

## ⚙️ Pré-requisitos e Instalação

A grande vantagem arquitetural deste projeto é que ele roda nativamente em qualquer *browser* moderno. Não há necessidade de configurar ambientes (como Node.js, Apache ou bancos de dados).

**Para executar a aplicação:**

1. Realize o download ou clone este repositório para sua máquina local.
2. Certifique-se de que os três arquivos principais estejam na mesma pasta.
3. Dê um duplo clique no arquivo `index.html` para abri-lo imediatamente no navegador web de sua preferência (Google Chrome, Edge, Firefox, etc.).

## 

## 📁 Estrutura de Pastas

Para manter a simplicidade e facilitar a manutenção, os recursos do sistema estão centralizados em três arquivos raiz, não havendo subdiretórios complexos.

##

```text
/consultaReservas
  ├── index.html    # Ponto de entrada da aplicação, contendo o esqueleto da interface e chamadas das bibliotecas.
  ├── style.css     # Arquivo de estilos responsável por toda a interface gráfica e formatação de impressão.
  ├── script.js     # Motor da aplicação, contém o dicionário de Unidades Orçamentárias, a lógica de renderização e as regras de filtragem.
  ├── logo.png      # Logotipo utilizado no cabeçalho dos relatórios oficiais.
  └── README.md     # Documentação do projeto.
```

## 

## 👨‍💻 Autor:

Renato Pinheiro Destro
renato.destro@gmail.com
Auxiliar de Escritório / Prefeitura Municipal de Botucatu/SP

###### Seja LIVRE, use Linux!