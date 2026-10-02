# UnifyData Python

![UnifyData Python: o fechamento no computador e a revisão de nomes no celular](docs/telas/capa.png)

**O UnifyData faz o fechamento mensal do consumo dos funcionários** de
um grupo de empresas. Eles compram fiado nas lojas do próprio grupo, e o
Linx AutoSystem, o sistema de gestão das lojas, registra cada compra. No
fim do mês, alguém precisa transformar esses relatórios do Linx em uma
planilha de desconto em folha para cada empresa. Quando o Linx e as
planilhas do RH não têm código em comum, a única ligação entre eles é o
**nome**, e nome se escreve de muitos jeitos: *BAPTISTA* num lugar,
*BATISTA* no outro; *DEBORA* sem acento; um sobrenome cortado.

**O UnifyData liga as pessoas pelo nome** e só pergunta a uma pessoa o
que não dá para decidir sozinho. Nomes que mudam só no acento, nas
maiúsculas ou nos espaços são ligados automaticamente; nomes parecidos,
nomes repetidos e pessoas que não estão no RH viram cartões de revisão; o
resultado é a planilha de desconto, uma aba por empresa.

Feito em Python com Streamlit, roda no próprio computador.

**Experimente** com dois comandos (abaixo) e clique em **Experimentar com
arquivos de exemplo**: arquivos fictícios, nada é salvo. Ou experimente
agora a [versão Web](https://alyssom-fernandes.github.io/UnifyData-Web/?demo=1),
no navegador, sem instalar nada.

[![Testes](https://github.com/alyssom-fernandes/UnifyData-Python/actions/workflows/testes.yml/badge.svg)](https://github.com/alyssom-fernandes/UnifyData-Python/actions/workflows/testes.yml)
![Python](https://img.shields.io/badge/Python-3.10+-3776ab?style=flat-square&logo=python&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-1.57+-ff4b4b?style=flat-square&logo=streamlit&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-150458?style=flat-square&logo=pandas&logoColor=white)
![Tema](https://img.shields.io/badge/tema-claro_e_escuro-0d6e66?style=flat-square)
![Licença](https://img.shields.io/badge/licen%C3%A7a-MIT-blue?style=flat-square)

Este README também está em [inglês](README.md).

## Em 30 segundos

1. Instale e abra (um ou dois minutos na primeira vez):

   ```bash
   pip install -r requirements.txt
   streamlit run unifydata.py
   ```

   No Windows, o `Iniciar_UnifyData.bat` faz os dois com dois cliques.
2. Clique em **Experimentar com arquivos de exemplo**. Entram três
   planilhas do RH e dois relatórios do Linx. *DEBORA* (Linx) é ligada a
   *DÉBORA* (RH) sozinha, e aparecem cinco cartões na etapa
   *Divergências*: um nome parecido (*Rafael Cunha Baptista* no Linx,
   *Batista* no RH, 98% parecido) e quatro pessoas que não estão em
   nenhuma planilha do RH.
3. Aprove o vínculo, digite CPF e empresa, ou ignore. Quando não sobra
   nenhum cartão, o fechamento fica pronto, com **Baixar Excel**, **CSV**
   e **Imprimir ou PDF**.

## Telas

Tiradas com os arquivos de exemplo.

| Revisão de nomes, tema escuro | Fechamento pronto, tema claro |
|---|---|
| ![Cartões de revisão no tema escuro: um nome parecido e pessoas fora do RH](docs/telas/divergencias-escuro.png) | ![Fechamento pronto no tema claro, com totais por empresa e loja](docs/telas/resultado-claro.png) |
| **Início, tema escuro** | **Fechamento pronto, tema escuro** |
| ![Tela inicial, com as quatro etapas na lateral](docs/telas/inicio-escuro.png) | ![Fechamento pronto no tema escuro](docs/telas/resultado-escuro.png) |

| No celular, tema claro | No celular, tema escuro |
|---|---|
| <img src="docs/telas/celular-divergencias-claro.png" alt="Revisão de nomes no celular, tema claro" width="260"> | <img src="docs/telas/celular-escuro.png" alt="Fechamento pronto no celular, tema escuro" width="260"> |

## O que sai

![O relatório A4, a planilha do Excel com uma aba por empresa e o CSV](docs/telas/saidas.png)

- **Excel**: uma aba por empresa com **todos os funcionários do RH**; quem
  não consumiu no mês aparece em cinza, então a aba também serve de lista
  da folha. Cabeçalho fixo, formato de moeda, CPF formatado, linha de
  total e nomes de aba sempre aceitos pelo Excel. Cada aba já sai
  configurada para imprimir em A4, na largura da página.
- **CSV**: separado por `;` e com vírgula decimal, como o Excel em
  português abre, com a coluna da empresa.
- **Relatório** para imprimir ou salvar em PDF: A4, em paisagem quando há
  três lojas ou mais, cada empresa inteira na página quando cabe, e páginas
  numeradas. Quem foi ignorado aparece no rodapé, com o valor.

## O que faz

### Leitura

- **Planilhas do RH**: uma por empresa, .xls, .xlsx ou .csv, com colunas
  de nome e CPF, esteja o cabeçalho na linha que estiver. O nome do
  arquivo vira o nome da empresa; uma pasta de trabalho com várias abas
  vira *Empresa - Aba*, a menos que a aba tenha o nome padrão do Excel
  (Plan1, Planilha1...).
- **Relatórios do Linx**: o CSV *Pendências por responsável*, um por loja.
  Dois valores por pessoa e loja, *vencido* e *total*, e o período de
  vencimentos.

### Cruzamento, do mais seguro ao mais cuidadoso

1. **Mesmo nome depois de tirar acentos, maiúsculas e espaços**, e só uma
   pessoa no RH com ele: ligado automaticamente (a tela diz quantos).
2. **Nome parecido** (85% ou mais com `token_sort_ratio`, que também pega
   palavras fora de ordem): um cartão sugere o vínculo. Aprove, ou diga
   que é outra pessoa digitando o CPF dela; o cartão mostra de quem é o
   CPF antes de confirmar.
3. **Nome repetido no RH** (homônimos, ou a mesma pessoa em duas
   planilhas): o cartão pergunta de qual linha descontar, para ninguém ser
   cobrado duas vezes. Nunca é escolhido sozinho.
4. **Fora do RH**: digite o CPF. Se for de alguém do RH, o consumo vai
   para essa pessoa; se não for, o registro é criado na empresa escolhida
   (uma que já existe ou uma nova). Ou ignore.

Todo cartão pode ser ignorado; quem foi ignorado aparece, com o valor, no
resultado e no relatório.

### Celular e temas

- Tema claro e escuro, seguindo o sistema, com troca pelo menu ⋮. As
  partes próprias do app acompanham o tema mesmo quando ele muda sem
  recarregar.
- Lateral com as quatro etapas e a situação de cada uma, e a assinatura
  AFN Systems; no celular, cartões e botões ocupam a largura toda.

## Privacidade e segurança

- **Roda no seu computador, só para o seu computador.** O servidor
  atende apenas em `localhost`, então outras máquinas da rede não
  conseguem abrir o app. Os arquivos ficam na memória da sua sessão do
  Streamlit e não são gravados em lugar nenhum; fechar a sessão descarta
  tudo. As estatísticas de uso do Streamlit estão desligadas.
- **O modo de exemplo** carrega os arquivos fictícios de `exemplos/`; não
  mexe em mais nada.
- **Texto vindo dos arquivos nunca é confiável**: nomes e nomes de arquivo são
  escapados antes de virar HTML, e células do CSV que começam com `=`,
  `+`, `-` ou `@` ganham um apóstrofo na frente, para a planilha não
  executar uma fórmula escondida num nome.

## Como é feito

| Parte | Tecnologia |
|---|---|
| Interface | Streamlit 1.57, com a configuração de tema e um pouco de CSS |
| Dados | pandas para ler e consolidar; xlrd para .xls antigo |
| Nomes | thefuzz (`token_sort_ratio`) para a semelhança |
| Escrita | openpyxl para o .xlsx; o CSV e o relatório são feitos pelo app |
| Fontes | Instrument Sans na interface; JetBrains Mono em rótulos e na assinatura |
| Testes | `unittest`, a cada push (GitHub Actions) |

Algumas decisões por trás:

- **As regras ficam separadas da tela.** O `nucleo.py` tem funções puras
  para ler, cruzar, consolidar e escrever o Excel e o CSV; o
  `unifydata.py` só desenha a tela e chama essas funções. Os testes rodam
  as regras com os mesmos arquivos de exemplo, sem o Streamlit.
- **Automático só quando não há ambiguidade.** O vínculo sozinho só
  acontece quando o nome é o mesmo depois de tirar acentos, maiúsculas e
  espaços, e o RH tem uma única grafia dele; o resto espera um clique.
- **Homônimos protegidos.** Quando um nome aparece mais de uma vez no RH,
  o consumo vai só para a linha escolhida, e as outras ficam zeradas. O
  consumo que um CPF já ligou a uma pessoa fica com ela, seja qual for a
  escolha feita depois para o nome.
- **Cliques seguros.** Cada ação é um callback que roda antes do próximo
  desenho da tela, então um clique duplo num botão que já fez o trabalho
  não faz nada.

## Limitações conhecidas

- Depende do layout do CSV *Pendências por responsável* do Linx.
- Cruzar pelo nome precisa de uma pessoa para confirmar os casos em
  dúvida; é proposital.
- Nada fica guardado entre um fechamento e outro: as planilhas do RH são
  lidas de novo a cada mês. (A
  [versão Web](https://github.com/alyssom-fernandes/UnifyData-Web) guarda
  uma base e cruza pela matrícula.)
- Precisa de Python no computador; não há versão hospedada.
- As fontes vêm do Google Fonts; sem internet, entram as fontes do
  sistema.
- Algumas partes são do próprio Streamlit: o menu ⋮ está em inglês, e o
  *Print* dele imprime a página inteira (o botão **Imprimir ou PDF**
  imprime só o relatório).
- Testado no Chrome e no Edge, no computador e na largura de celular.

## Rodando

Python 3.10 ou mais novo (os testes rodam no 3.12 no CI; o app foi
usado com o 3.14):

```bash
pip install -r requirements.txt
streamlit run unifydata.py
```

O app abre em `http://localhost:8501`; `http://localhost:8501/?demo=1`
já abre com os arquivos de exemplo. No Windows, o `Iniciar_UnifyData.bat`
instala o que faltar na primeira vez, cria um arquivo de credenciais vazio
do Streamlit (para ele não pedir e-mail) e abre o app.

Para rodar os testes:

```bash
python -m unittest discover -s tests
```

## Estrutura do projeto

```
unifydata.py            o app Streamlit: etapas, cartões de revisão, resultado e exportações
nucleo.py               regras puras: leitura, cruzamento de nomes, consolidação, Excel, CSV
requirements.txt        dependências
Iniciar_UnifyData.bat   atalho para Windows
.streamlit/config.toml  temas claro e escuro, fontes, barra de ferramentas, telemetria desligada
assets/favicon.png      a marca do UnifyData
tests/                  testes das regras, com os arquivos de exemplo
exemplos/               arquivos fictícios nos formatos do Linx e do RH
.github/workflows/      roda os testes a cada push
docs/telas/             imagens deste README e da prévia para redes (og.png)
```

## Comparando com a versão Web

| | UnifyData Web | UnifyData Python |
|---|---|---|
| **Roda** | No navegador, sem instalar nada | No computador, com Python |
| **Entradas** | PDF de Contas a Receber do Linx, planilhas do RH, CSVs do Linx | Planilhas do RH e CSVs do Linx |
| **Cruzamento** | Pela matrícula (exato) | Pelo nome (aproximado, com revisão) |
| **Homônimos** | Não atrapalham, a matrícula separa | Um cartão pergunta de qual linha descontar |
| **Dados dos funcionários** | Uma base salva no navegador | Planilhas do RH lidas todo mês |
| **Excel** | Quem teve consumo; totais com fórmulas `SOMA` | Todos do RH (sem consumo em cinza) |
| **Melhor para** | Linx e RH com a matrícula em comum | Sistemas sem código em comum |

## Licença

[MIT](LICENSE). Feito por [Alyssom Fernandes](https://github.com/alyssom-fernandes), AFN Systems.
