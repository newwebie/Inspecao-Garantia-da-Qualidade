# Sistema de Arquivamento de Documentos

Aplicação **Streamlit** para cadastro, consulta, movimentação e gestão de status de documentos arquivados, com persistência em um único arquivo **Excel** hospedado no **SharePoint**. Usa autenticação por **UserCredential** e leitura/gravação de múltiplas abas.

---

## 📌 Sumário

- [Visão Geral](#visão-geral)
- [Recursos Principais](#recursos-principais)
- [Arquitetura & Fluxo de Dados](#arquitetura--fluxo-de-dados)
- [Pré-requisitos](#pré-requisitos)
- [Instalação](#instalação)
- [Configuração de Segredos (Streamlit)](#configuração-de-segredos-streamlit)
- [Estrutura do Excel no SharePoint](#estrutura-do-excel-no-sharepoint)
- [Esquema de ID (PPPPNNL)](#esquema-de-id-ppppnnl)
- [Guia de Uso (Abas)](#guia-de-uso-abas)
  - [Cadastrar](#cadastrar)
  - [Status](#status)
  - [Consultar](#consultar)
  - [Movimentar](#movimentar)
  - [⚙️ Opções](#️-opções)
- [Cache, Estado de Sessão e Atualização](#cache-estado-de-sessão-e-atualização)
- [Tratamento de Erros e Concorrência](#tratamento-de-erros-e-concorrência)
- [Boas Práticas e Segurança](#boas-práticas-e-segurança)
- [Testes Locais](#testes-locais)
- [Problemas Conhecidos & Sugestões de Melhoria](#problemas-conhecidos--sugestões-de-melhoria)
- [Licença](#licença)

---

## Visão Geral

Esta aplicação centraliza o **ciclo de vida de caixas/documentos**: cadastro com geração de ID, controle de localização física (local, estante, prateleira, caixa), **retenção** e **descarte previsto**, além de **desarquivamento/rearquivamento** e **consulta**. Toda a persistência ocorre em um **arquivo Excel** com múltiplas abas dentro de uma pasta do SharePoint.

---

## Recursos Principais

- 🔐 Integração nativa com **SharePoint** via `Office365-REST-Python-Client` (leitura e upload binário).
- 📑 Leitura/Gravação de **todas as abas** do Excel em memória, preservando as existentes ao salvar.
- 🧮 **Geração de ID determinística** por prefixo (PPPP) + sufixo cíclico (NNL → 00A..99Z, 2600 combinações por prefixo).
- 🧭 **Mapeamento de Siglas** para Departamento/Tipo de Documento a partir da aba **Selectboxes**.
- 🗃️ **Estruturas (Espaços)** parametrizáveis (quantidade de estantes/prateleiras por arquivo físico).
- 🗓️ **Retenção** por origem com cálculo de **Data Prevista de Descarte**.
- 🧰 Abas funcionais: **Cadastrar**, **Status**, **Consultar**, **Movimentar**, **⚙️ Opções**.
- 🚦 Tratamento de **arquivo bloqueado** no SharePoint com retry exponencial simples.
- ⚡ Uso de `@st.cache_data` para reduzir leituras e melhorar responsividade.

---

## Arquitetura & Fluxo de Dados

1. **Configuração**: credenciais e caminhos são carregados de `st.secrets`.
2. **Leitura**: `carregar_excel()` usa `File.open_binary` para puxar o Excel e ler **todas** as abas via `pandas.read_excel(sheet_name=None)`.
3. **Camada de Negócio**:
   - Mapeamento de siglas (departamento/tipo).
   - Cálculo de IDs e retenção.
   - Operações de status e movimentação.
4. **Persistência**: `update_sharepoint_file()` recompõe o **workbook completo em memória** e faz upload do binário para a mesma pasta/arquivo no SharePoint, preservando as demais abas quando `keep_existing=True`.
5. **Cache/Estado**: limpeza de cache e `st.session_state` para consistência após salvamentos.

---

## Pré-requisitos

- Python 3.10+
- Acesso a um **site do SharePoint** com permissões de leitura e escrita na biblioteca de documentos onde o Excel reside.

### Pacotes (requirements.txt sugerido)

```
streamlit>=1.33
pandas>=2.0
python-dateutil>=2.8
XlsxWriter>=3.1
Office365-REST-Python-Client>=2.5
openpyxl>=3.1
```

> Observação: `pandas` usa `openpyxl` para leitura de .xlsx; a escrita é feita aqui via **XlsxWriter**.

---

## Instalação

```bash
# criar venv (opcional)
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate

pip install -r requirements.txt
```

## Configuração de Segredos (Streamlit)

Crie um arquivo `.streamlit/secrets.toml` na raiz do projeto:

```toml
[sharepoint]
USERNAME = "seu.usuario@org.com"
PASSWORD = "SUA_SENHA_OU_APP_PASSWORD"
SITE_BASE = "https://suaorg.sharepoint.com/sites/SeuSite"
# Caminho server-relative do arquivo Excel (ex.: /sites/SeuSite/Shared Documents/arquivos/Repositorio.xlsx)
ARQUIVO   = "/sites/SeuSite/Shared Documents/arquivos/Repositorio.xlsx"
```

> 🔎 **ARQUIVO** deve ser o **server-relative URL** do arquivo (inclui `/sites/...`). A conta usada precisa ter permissão de escrita.

---

## Estrutura do Excel no SharePoint

O aplicativo espera um workbook com as abas (nomes exatos):

- **Arquivos** → cadastro principal
- **Espaços** → lista de locais físicos, com colunas: `Arquivo`, `Estantes`, `Prateleiras`
- **Selectboxes** → opções de selects e mapeamentos, com possíveis colunas:
  - `Departamentos`, `Tipos de Documento`
  - `Sigla Departamento`, `Sigla Documento`
  - `RESPONSÁVEL ARQUIVAMENTO`
- **Retenção** → colunas: `ORIGEM DOCUMENTO SUBMISSÃO`, `Retenção` (ex.: `5 anos`)

### Colunas esperadas — Aba **Arquivos**

- `ID` (gerado pelo sistema)
- `Local`, `Estante`, `Prateleira`, `Caixa`, `Codificação`, `Tag`, `Livro`, `Lacre`
- `Tipo de Documento`, `Conteúdo da Caixa`, `Departamento Origem`, `Origem Departamento Submissão`
- `Responsável Arquivamento`, `Solicitante`
- `Data Arquivamento`, `Período Utilizado Início`, `Período Utilizado Fim`
- `Status` ("ARQUIVADO" | "DESARQUIVADO")
- `Período de Retenção`, `Data Prevista de Descarte`
- (Opcional) `Responsável Desarquivamento`, `Data Desarquivamento`, `Observação Desarquivamento`

> Se alguma aba estiver ausente, a aplicação exibe **warning** e prossegue com DataFrames vazios.

---

## Esquema de ID (PPPPNNL)

- **Prefixo (PPPP)**: concatenação de 2 letras do **Departamento** (sigla ou fallback) + 2 letras do **Tipo de Documento** (sigla ou fallback). Fallback remove não alfanuméricos e usa os **2 primeiros** caracteres, completando com `X` se necessário.
- **Sufixo (NNL)**: 2 dígitos (`00`..`99`) + 1 letra (`A`..`Z`).
  - Mapeamento: `idx 0..2599 → 00A..99Z` (`num = idx // 26`, `letra = A + (idx % 26)`).
- **Capacidade por prefixo**: 2600 IDs.
- O próximo índice é calculado com base no conteúdo atual da aba **Arquivos** (memória + cache), evitando colisões.

---

## Guia de Uso (Abas)

### Cadastrar

1. Selecione **Tipo de Documento** e **Origem do Documento** (alimentados por **Selectboxes**).
2. O sistema mostra o **ID atual** (pré-visualização) baseado no próximo índice do prefixo.
3. Preencha os campos obrigatórios (Local, Estante, Prateleira, Caixa, Conteúdo da Caixa, Origem Departamento Submissão, Solicitante, Responsável, etc.).
4. Opcional: informe **Período Utilizado**, **Tag/Lacre/Livro**, **Codificação**.
5. Ao clicar em **Cadastrar**, o registro é adicionado à aba **Arquivos** e salvo no SharePoint preservando as demais abas.

### Status

- Informe um **ID** válido para ver dados e executar:
  - **Desarquivar**: muda `Status → DESARQUIVADO`, registra responsável/data e opcionalmente **desarquivamento parcial** (observação textual).
  - **Rearquivar**: volta `Status → ARQUIVADO` e limpa campos de desarquivamento.
- A seção **“Documentos Desarquivados”** lista todos os registros com `Status = DESARQUIVADO` e destaca parciais.

### Consultar

- **Por Codificação**: filtro exato em `Codificação`.
- **Por Período**: entre `Data Arquivamento` inicial e final.

### Movimentar

- **Lote**: informe **vários IDs** separados por vírgula. O app valida existentes/não encontrados, e aplica a **mesma nova localização** (Local/Estante/Prateleira) a todos.
- **Unitário (variante)**: há uma seção alternativa para movimentação unitária, com validação de **slot ocupado**.

### ⚙️ Opções

- **Selectboxes**: editor de dados para manter listas (departamentos, tipos, responsáveis, siglas).
- **Período de Retenção**: editor da aba `Retenção`.
- **Espaços**: editor da aba `Espaços` (quantidades por arquivo físico).

> Todos os editores salvam **apenas a aba** correspondente, preservando as demais.

---

## Cache, Estado de Sessão e Atualização

- `@st.cache_data` acelera leitura do Excel e cálculo de mapas.
- Após **salvar**, o app limpa o cache (`st.cache_data.clear()`) e atualiza `st.session_state` para refletir os dados mais recentes.
- O botão **🔄 Atualizar** (sidebar) força limpeza de cache e `st.rerun()`.

---

## Tratamento de Erros e Concorrência

- **Arquivo em uso / bloqueado (423 / -2147018894 / “lock”)**: a função de salvamento exibe aviso e **tenta novamente** após 5s, em loop até concluir ou falhar por outro motivo.
- Mensagens de erro amigáveis são mostradas via `st.error`/`st.warning`.

---

## Boas Práticas e Segurança

- Prefira **App Password**/MFA/App Registration\*\* conforme a política da organização. Evite credenciais em claro.
- Restrinja permissões no SharePoint à **biblioteca/pasta** necessárias.
- Faça **backup/versionamento** do Excel (a biblioteca do SharePoint mantém versões; valide o limite).
- Considere validações adicionais (ex.: formatos de prateleira/estante, ranges permitidos por local).

---

## Testes Locais

1. Configure `secrets.toml` com um **arquivo Excel** de teste (pode ser local, caso ajuste o código para bypass do SharePoint em dev).
2. Rode a aplicação:

```bash
streamlit run app.py
```

3. Acesse `http://localhost:8501`.

---
