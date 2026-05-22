# UnifyData

> **Ferramenta inteligente de conciliação de folha de pagamento com Python e Streamlit.**

O UnifyData automatiza o cruzamento mensal de consumos de convênios de funcionários (extraídos do Linx AutoSystem) com as listas de RH. Diferente de uma busca simples, ele usa **correspondência aproximada de nomes** para detectar erros de digitação, diferenças de acentuação e inconsistências entre sistemas — ideal para ambientes onde os cadastros não compartilham um código de matrícula em comum.

---

## ✨ Funcionalidades

- **Correspondência Aproximada de Nomes** — usa `thefuzz` (`token_sort_ratio`) para identificar nomes similares entre os sistemas, detectando automaticamente erros de digitação e diferenças de formatação
- **Captura de Dois Valores** — extrai tanto o valor *vencido* quanto o *total da dívida* de cada entrada do relatório Linx, exibindo-os em colunas separadas
- **Central de Revisão de Divergências** — nomes ambíguos ou não encontrados aparecem em um painel interativo com opções para aprovar, ignorar ou vincular manualmente por CPF
- **Suporte a Múltiplas Empresas** — processa várias planilhas de RH ao mesmo tempo, gerando uma aba por unidade no Excel
- **Interface Web Local** — roda como um app Streamlit no navegador, com design customizado em dark mode
- **Exportação Excel** — gera um relatório `.xlsx` com cabeçalhos coloridos, formatação de moeda, linhas cinzas para funcionários sem consumo e largura automática de colunas

---

## 🗂️ Estrutura do Projeto

```
unifydata-python/
├── unifydata.py             # Aplicação Streamlit principal
├── requirements.txt         # Dependências Python
├── Iniciar_UnifyData.bat    # Inicializador Windows com um clique
└── README.md
```

---

## 🚀 Como Executar

### Pré-requisitos
- Python 3.8+ instalado na máquina

### Opção A — Windows (um clique)
Dê um duplo clique em **`Iniciar_UnifyData.bat`**.

Na primeira execução, as dependências serão instaladas automaticamente e o app abrirá no navegador.

### Opção B — Manual (qualquer sistema operacional)

```bash
# Instalar dependências
pip install -r requirements.txt

# Iniciar o app
streamlit run unifydata.py
```

O app abrirá em `http://localhost:8501`.

---

## 📖 Como Usar

1. **Suba os Arquivos de RH** — arraste as planilhas de funcionários da empresa (`.xlsx`, `.xls` ou `.csv`). O nome do arquivo (ex: `Empregados Posto Rosário.xls`) é usado como nome da unidade.
2. **Suba os Relatórios do Linx** — arraste os arquivos CSV de consumo mensal exportados do Linx AutoSystem.
3. **Clique em "Processar e Unificar Dados"** — o app cruza as duas fontes e detecta divergências.
4. **Revise as Divergências** — para cada item sinalizado, escolha:
   - ✅ **Aprovar** um vínculo de nome sugerido
   - 🔗 **Vincular por CPF** para correção manual
   - 🚫 **Ignorar** para excluir o registro do relatório
5. **Baixe** o relatório consolidado em `.xlsx`.

---

## 📊 Formato do Relatório

Cada unidade ganha sua própria aba. Colunas por aba:

| Funcionário | CPF | Vencido [Origem] | Total [Origem] | ... | Total Desconto |
|---|---|---|---|---|---|
| FULANO DE TAL | 000.000.000-00 | R$ 0,00 | R$ 150,00 | ... | R$ 150,00 |

- Cabeçalho: fundo vinho escuro, texto branco em negrito
- Funcionários com consumo zero: exibidos em cinza
- Rodapé: total geral em negrito

---

## 🛠️ Tecnologias Utilizadas

| Tecnologia | Função |
|---|---|
| [Python 3](https://python.org) | Linguagem principal |
| [Streamlit](https://streamlit.io) | Framework de interface web |
| [pandas](https://pandas.pydata.org) | Processamento e pivotamento de dados |
| [thefuzz](https://github.com/seatgeek/thefuzz) | Correspondência aproximada de nomes |
| [openpyxl](https://openpyxl.readthedocs.io) | Geração do relatório Excel |

---

## 📋 Dependências

```
streamlit
pandas
openpyxl
thefuzz[speedup]
python-Levenshtein
```

Instale com:
```bash
pip install -r requirements.txt
```

---

## 🔄 Diferença em Relação à Versão Web

| | UnifyData Web | UnifyData Python |
|---|---|---|
| **Instalação** | Nenhuma (abrir HTML) | Python obrigatório |
| **Método de cruzamento** | Código de matrícula (exato) | Correspondência aproximada de nomes |
| **Persistência** | `localStorage` (permanente) | Apenas na sessão (re-upload todo mês) |
| **Ideal para** | Fechamento mensal rápido | Sistemas sem código de matrícula em comum |

---

## 🤝 Contribuições

Sinta-se à vontade para abrir issues ou pull requests. Contribuições que ampliem a compatibilidade com outros formatos de CSV ou adicionem novas estratégias de cruzamento são especialmente bem-vindas.

---

## 📄 Licença

Licença MIT — livre para usar, modificar e distribuir.

---

*Desenvolvido por AFN Systems · 2026*
