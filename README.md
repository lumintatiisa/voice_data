# 📊 Dashboard de Acompanhamento Comercial — TFP

Dashboard em Streamlit para acompanhamento semanal do percentual de clientes
com **título/fatura em aberto (TFP)** por **consultor** e **parceiro**,
segmentado por **Linha Móvel** e **Linha Fixa**, com comparativo entre
períodos.

## 🗂 Estrutura do projeto

```
.
├── app.py                 # aplicação Streamlit (front-end + orquestração)
├── data_loader.py         # ingestão, normalização e regras de negócio
├── data/                  # arquivos .xlsx de origem (um por período)
│   ├── TFP_SEMANA_1_E_2.xlsx
│   └── TFP_SEMANA_3.xlsx
├── requirements.txt
├── .streamlit/config.toml # tema visual
└── README.md
```

## ▶️ Rodando localmente

```bash
python -m venv .venv
source .venv/bin/activate       # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## ☁️ Publicando no Streamlit Community Cloud

1. Crie um repositório no GitHub e suba **todo** este conteúdo (incluindo a
   pasta `data/` — o Streamlit Cloud não tem acesso ao seu OneDrive, então os
   arquivos `.xlsx` precisam estar versionados no repositório).
2. Em https://share.streamlit.io, clique em **New app**, aponte para o
   repositório, branch `main` e arquivo principal `app.py`.
3. Pronto — o app é publicado com o tema já configurado em
   `.streamlit/config.toml` e as dependências de `requirements.txt`.

## 🔄 Atualização semanal dos dados

O fluxo pensado para você é:

1. Baixe o(s) arquivo(s) mais recente(s) da pasta compartilhada (OneDrive/TFP).
2. Salve-o em `data/`, seguindo o padrão de nome:
   - `TFP_SEMANA_<n>.xlsx` para uma semana isolada (ex.: `TFP_SEMANA_4.xlsx`)
   - `TFP_SEMANA_<n>_E_<m>.xlsx` para um arquivo que já vem consolidando mais
     de uma semana (ex.: `TFP_SEMANA_1_E_2.xlsx`)
3. O arquivo deve ter (pelo menos) duas abas — uma para **linhas móveis** e
   outra para **linhas fixas** — contendo, no mínimo, as colunas:
   `CNPJ`, `NOME CLIENTE`, `CONSULTOR`, `PARCEIRO`. O nome da aba só precisa
   conter "MOV" ou "FIX" em algum lugar (maiúsculas/minúsculas e acentos não
   importam) para ser reconhecida automaticamente.
4. Commit + push. Se o app já estiver publicado no Streamlit Cloud, ele
   reimplanta sozinho a cada push na branch conectada.
5. Não precisa mexer em nenhum código — `data_loader.py` varre a pasta
   `data/` inteira a cada carregamento e descobre novos arquivos sozinho.

> Se algum arquivo futuro vier com abas trocadas (o mesmo problema encontrado
> no arquivo da Semana 1 e 2), adicione uma entrada em
> `SHEET_TYPE_OVERRIDES`, em `data_loader.py`, apontando o nome do arquivo e
> da aba para o tipo correto.

## 🧮 Metodologia (resumo — o texto completo também está no app, na barra
lateral, em "ℹ️ Metodologia")

- **TFP** = lista de clientes com título/fatura em aberto (inadimplência),
  separada por tipo de linha.
- As planilhas de origem só trazem quem **tem** débito — não existe, nelas,
  a carteira total de clientes por consultor/parceiro. Por isso, o
  percentual mostrado é a **participação de cada consultor/parceiro dentro
  do total de inadimplentes** daquele tipo de linha e período (não é uma
  taxa de inadimplência sobre a base total de clientes).
- **Deduplicação:** um mesmo CNPJ pode se repetir na planilha original
  (múltiplos contratos). Nas métricas de "clientes", cada CNPJ é contado uma
  única vez por consultor/parceiro/tipo/período; a contagem original de
  registros fica disponível na coluna `QTD_REGISTROS` do detalhamento.
- **Correção de dados aplicada:** no arquivo `TFP_SEMANA_1_E_2.xlsx`, as
  abas "TFP MÓVEL" e "TFP FIXA" estavam com conteúdo trocado em relação ao
  nome (confirmado comparando os CNPJs com `TFP_SEMANA_3.xlsx`). A correção
  foi aplicada programaticamente — ver `SHEET_TYPE_OVERRIDES` em
  `data_loader.py`.
- **Limitação atual:** o arquivo que consolida as semanas 1 e 2 não tem
  coluna de data/semana, então não é possível separar os registros de cada
  semana dentro dele — esse período é tratado como um bloco único ("Semana
  1 e 2"). Assim que uma planilha futura trouxer essa granularidade (ex.:
  uma coluna `SEMANA` ou arquivos individuais por semana), o comparativo
  semana a semana passa a ser automático, sem qualquer alteração de código.

## 🛠 Stack

- [Streamlit](https://streamlit.io/) — front-end e app framework
- [Pandas](https://pandas.pydata.org/) — processamento de dados
- [Plotly](https://plotly.com/python/) — gráficos interativos
- CSS customizado embutido em `app.py` para os cards e o cabeçalho
