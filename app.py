import streamlit as st
import pandas as pd
from datetime import date, datetime
from dateutil.relativedelta import relativedelta
from dateutil.relativedelta import relativedelta
import io
import time
import re
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File


# === Configurações do SharePoint (mesmo padrão do seu trecho que funciona) ===
username = st.secrets["sharepoint"]["USERNAME"]
password = st.secrets["sharepoint"]["PASSWORD"]
site_url = st.secrets["sharepoint"]["SITE_BASE"]
file_name = st.secrets["sharepoint"]["ARQUIVO"]  # caminho server-relative do Excel com as abas
solicitacoes = st.secrets["sharepoint"]["SOLICITAÇÕES"]

@st.cache_data
def carregar_excel():
    try:
        ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
        response = File.open_binary(ctx, file_name)

        # Lê todas as abas de uma vez
        sheets = pd.read_excel(io.BytesIO(response.content), sheet_name=None)

        # Usa .get() com defaults vazios se a aba não existir
        df          = sheets.get("Arquivos",    pd.DataFrame())
        df_espacos  = sheets.get("Espaços",     pd.DataFrame())
        df_selects  = sheets.get("Selectboxes", pd.DataFrame())
        Retencao_df = sheets.get("Retenção",    pd.DataFrame())

        # Avisos úteis se alguma aba estiver faltando
        faltando = [n for n, d in [
            ("Arquivos", df),
            ("Espaços", df_espacos),
            ("Selectboxes", df_selects),
            ("Retenção", Retencao_df),
        ] if d.empty]

        if faltando:
            st.warning(f"A(s) aba(s) não encontrada(s) ou vazia(s): {', '.join(faltando)}")

        return df, df_espacos, df_selects, Retencao_df

    except Exception as e:
        st.error(f"Erro ao acessar o arquivo no SharePoint: {e}")
        # SEMPRE retorne 4 DataFrames
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    

def carregar_solicitacoes():
    try:
        ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
        response = File.open_binary(ctx, solicitacoes)

        # Lê todas as abas de uma vez
        sheets = pd.read_excel(io.BytesIO(response.content), sheet_name=None)

        # Usa .get() com defaults vazios se a aba não existir
        df_solicitacoes          = sheets.get("Arquivos",    pd.DataFrame())


        return df_solicitacoes
    
    except Exception as e:
        st.error(f"Erro ao acessar o arquivo no SharePoint: {e}")
        return pd.DataFrame()
    


def _sanitize_sheet_name(name: str) -> str:
    # Excel: máx 31 chars e proíbe []:*?/\ 
    return (
        name.replace("[", " ").replace("]", " ")
            .replace(":", " ").replace("*", " ")
            .replace("?", " ").replace("/", " ")
            .replace("\\", " ")
    )[:31]


def update_sharepoint_file(
    df: pd.DataFrame,
    file_path: str,               # ex.: "/sites/site/Shared Documents/arquivo.xlsx"
    sheet_name: str = "Sheet1",
    keep_existing: bool = False,
    index: bool = False,
):
    """
    Salva um DataFrame em uma aba do Excel no SharePoint.
    Usa credenciais/URL de st.secrets['sharepoint'].
    """
    site_url = st.secrets["sharepoint"]["SITE_BASE"]
    username = st.secrets["sharepoint"]["USERNAME"]
    password = st.secrets["sharepoint"]["PASSWORD"]

    if "/" not in file_path:
        st.error("file_path inválido (use server-relative, ex.: /sites/.../arquivo.xlsx)")
        return

    folder_path, file_name_only = file_path.rsplit("/", 1)
    safe_sheet = _sanitize_sheet_name(sheet_name)

    while True:
        try:
            # Lê abas existentes se precisar preservar
            existing_sheets = {}
            if keep_existing:
                try:
                    ctx_rd = ClientContext(site_url).with_credentials(UserCredential(username, password))
                    resp = File.open_binary(ctx_rd, file_path)
                    existing_sheets = pd.read_excel(io.BytesIO(resp.content), sheet_name=None) or {}
                except Exception:
                    existing_sheets = {}

            # Atualiza/insere a aba
            existing_sheets[safe_sheet] = df

            # Escreve tudo num buffer
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                for name, data in existing_sheets.items():
                    (data if isinstance(data, pd.DataFrame) else pd.DataFrame(data))\
                        .to_excel(writer, sheet_name=_sanitize_sheet_name(name), index=index)
            output.seek(0)

            # Upload pro SharePoint
            ctx_wr = ClientContext(site_url).with_credentials(UserCredential(username, password))
            target_folder = ctx_wr.web.get_folder_by_server_relative_url(folder_path)
            target_folder.upload_file(file_name_only, output.read()).execute_query()

            # Cache e feedback
            try:
                st.cache_data.clear()
            except Exception:
                pass
            st.success("Salvo!")
            break

        except Exception as e:
            locked = (
                getattr(e, "response_status", None) == 423
                or "-2147018894" in str(e)
                or "lock" in str(e).lower()
            )
            if locked:
                st.warning("Arquivo está em uso. Tentando novamente em 5 segundos...")
                time.sleep(5)
                continue
            else:
                st.error(f"Erro ao salvar no SharePoint: {e}")
                break



# ===== Configuração da página =====
st.set_page_config(page_title="Sistema de Arquivo", layout="wide")
st.title("📂 Sistema de Arquivamento de Documentos")


# TABs
with st.sidebar:
    aba = st.selectbox("Escolha o que deseja", ["Cadastrar", "Status","Consultar", "Movimentar", "⚙️ Opções"])

    #Botao atualizar limpando cache
    if st.button("🔄 Atualizar"):
        st.cache_data.clear()      
        st.cache_resource.clear()
        st.rerun()  


    df, df_espacos, df_selects, Retencao_df = carregar_excel()
    # Estruturas (Espaços)
    estruturas = {
        f"ARQUIVO {str(row['Arquivo']).strip().upper()}": {
            "estantes": int(row["Estantes"]),
            "prateleiras": int(row["Prateleiras"])
        }
        for _, row in df_espacos.iterrows() if pd.notna(row['Arquivo'])
    }

    # Listas dos selects
    responsaveis = ["", *sorted(df_selects["RESPONSÁVEL ARQUIVAMENTO"].dropna().unique().tolist())]
    origens_submissao = ["", *sorted(Retencao_df["ORIGEM DOCUMENTO SUBMISSÃO"].dropna().unique().tolist())]
    dpto_op = ["", *sorted(df_selects["Departamentos"].dropna().unique().tolist())]
    doc_op = ["", *sorted(df_selects["Tipos de Documento"].dropna().unique().tolist())]
    local_op = [""] + sorted(df_espacos["Arquivo"].dropna().astype(str).unique())

    # Session state seguros
    if "ja_salvou" not in st.session_state:
        st.session_state.ja_salvou = False
    # Local default baseado em 'local_op'
    if "local" not in st.session_state:
        st.session_state.local = local_op[0] if local_op else ""


# -------------------------------------------
# ABA: Cadastrar  (ID = PPPP + NNL, com siglas vindas de Selectboxes)
# -------------------------------------------
if aba == "Cadastrar":
    st.header("Cadastrar Novo Documento")

    # -----------------------------
    # Helpers de nomes de colunas
    # -----------------------------
    DEPT_NAME_CANDIDATES = [
        "Departamento Origem", "Departamento", "Depto", "Departamento/Submissão"
    ]
    TIPO_NAME_CANDIDATES = [
        "Tipo de Documento", "Tipos de Documento", "Tipo", "Documento"
    ]
    DEPT_SIGLA_COL = "Sigla Departamento"
    TIPO_SIGLA_COL = "Sigla Documento"

    def pick_first_existing(df_local: pd.DataFrame, candidates):
        for c in candidates:
            if c in df_local.columns:
                return c
        return None

    # -----------------------------
    # Mapas de sigla (cacheados)
    # -----------------------------
    def carregar_mapas_de_sigla_de_df_selects():
        """
        Monta:
          - dept_map: nome_depto_upper -> sigla_depto_upper
          - tipo_map: nome_tipo_upper  -> sigla_tipo_upper
        Usa df_selects já carregado e cacheia em session_state.
        """
        if "sigla_maps" in st.session_state:
            return st.session_state["sigla_maps"]["dept_map"], st.session_state["sigla_maps"]["tipo_map"]

        dept_map, tipo_map = {}, {}
        if not df_selects.empty:
            dept_name_col = pick_first_existing(df_selects, DEPT_NAME_CANDIDATES)
            tipo_name_col = pick_first_existing(df_selects, TIPO_NAME_CANDIDATES)

            if dept_name_col and DEPT_SIGLA_COL in df_selects.columns:
                for _, r in df_selects[[dept_name_col, DEPT_SIGLA_COL]].dropna().iterrows():
                    nome = str(r[dept_name_col]).strip().upper()
                    sigla = str(r[DEPT_SIGLA_COL]).strip().upper()
                    if nome and sigla:
                        dept_map[nome] = sigla

            if tipo_name_col and TIPO_SIGLA_COL in df_selects.columns:
                for _, r in df_selects[[tipo_name_col, TIPO_SIGLA_COL]].dropna().iterrows():
                    nome = str(r[tipo_name_col]).strip().upper()
                    sigla = str(r[TIPO_SIGLA_COL]).strip().upper()
                    if nome and sigla:
                        tipo_map[nome] = sigla

        st.session_state["sigla_maps"] = {"dept_map": dept_map, "tipo_map": tipo_map}
        return dept_map, tipo_map

    # -----------------------------
    # Abreviações a partir dos mapas
    # -----------------------------
    def abrev_depto(nome: str) -> str:
        dept_map, _ = carregar_mapas_de_sigla_de_df_selects()
        if pd.isna(nome) or not str(nome).strip():
            return "XX"
        nome = str(nome).strip().upper()
        if nome in dept_map and dept_map[nome]:
            return str(dept_map[nome]).upper()
        alnum = re.sub(r"[^A-Z0-9]", "", nome)
        return (alnum[:2] or "XX").ljust(2, "X")

    def abrev_tipo(nome: str) -> str:
        _, tipo_map = carregar_mapas_de_sigla_de_df_selects()
        if pd.isna(nome) or not str(nome).strip():
            return "XX"
        nome = str(nome).strip().upper()
        if nome in tipo_map and tipo_map[nome]:
            return str(tipo_map[nome]).upper()
        alnum = re.sub(r"[^A-Z0-9]", "", nome)
        return (alnum[:2] or "XX").ljust(2, "X")

    def montar_prefixo(origem_depto: str, tipo_doc: str) -> str:
        return f"{abrev_depto(origem_depto)}{abrev_tipo(tipo_doc)}"  # 4 letras

    # -----------------------------
    # Conversões NNL <-> índice (00A..99Z)
    # -----------------------------
    def idx_to_sufixo(idx: int) -> str:
        """
        idx 0..2599 -> 'NNL' (00A..99Z)
        num = idx // 26, letra = A + (idx % 26)
        """
        if idx < 0 or idx >= 100 * 26:
            raise ValueError("Capacidade esgotada para este prefixo (00A..99Z = 2600 IDs).")
        num = idx // 26
        letra_idx = idx % 26
        letra = chr(ord('A') + letra_idx)
        return f"{num:02d}{letra}"

    def sufixo_to_idx(nnletra: str) -> int:
        m = re.fullmatch(r"(\d{2})([A-Z])", nnletra)
        if not m:
            raise ValueError(f"Sufixo inválido: {nnletra}")
        num = int(m.group(1))
        letra = m.group(2)
        return num * 26 + (ord(letra) - ord('A'))

    def extrair_prefixo_e_idx(id_str: str):
        """
        De PPPPNNL (ex.: GQES03C) -> (prefixo='GQES', idx=int). Ignora formatos fora do padrão.
        """
        if not isinstance(id_str, str):
            return None
        s = id_str.strip().upper()
        m = re.fullmatch(r"([A-Z0-9]{4})(\d{2}[A-Z])", s)
        if not m:
            return None
        prefixo = m.group(1)
        sufixo = m.group(2)
        try:
            idx = sufixo_to_idx(sufixo)
        except Exception:
            return None
        return prefixo, idx

    @st.cache_data(show_spinner=False)
    def ler_arquivos_existentes(path):
        try:
            return pd.read_excel(path, sheet_name="Arquivos")
        except FileNotFoundError:
            return pd.DataFrame()

    def carregar_ultimo_idx_por_prefixo():
        if "ultimo_idx_por_prefixo" in st.session_state:
            return st.session_state["ultimo_idx_por_prefixo"]

        # prioriza df que você salvou em session_state dentro do update_sharepoint_file
        base_df = st.session_state.get("df_Arquivos", df)

        ultimo = {}
        if base_df is not None and not base_df.empty and "ID" in base_df.columns:
            for val in base_df["ID"].astype(str):
                parsed = extrair_prefixo_e_idx(val)
                if parsed:
                    prefixo, idx = parsed
                    if prefixo not in ultimo or idx > ultimo[prefixo]:
                        ultimo[prefixo] = idx

        st.session_state["ultimo_idx_por_prefixo"] = ultimo
        return ultimo


    def proximo_idx_para_prefixo(prefixo: str, df_mem: pd.DataFrame = None) -> int:
        """
        Calcula o próximo índice (0..2599) para um prefixo, considerando:
          1) cache de 'ultimo_idx_por_prefixo' (do DataFrame em memória)
          2) opcionalmente o df em memória adicional (para prévia)
        """
        ultimo = carregar_ultimo_idx_por_prefixo()
        base = ultimo.get(prefixo, -1)

        # também olha o df em memória (IDs desta sessão já carregados em df, antes de salvar)
        if df_mem is not None and not df_mem.empty and "ID" in df_mem.columns:
            padrao = re.compile(rf"^{re.escape(prefixo)}(\d{{2}}[A-Z])$")
            for _id in df_mem["ID"].astype(str):
                m = padrao.match(_id)
                if m:
                    try:
                        idx_local = sufixo_to_idx(m.group(1))
                        if idx_local > base:
                            base = idx_local
                    except Exception:
                        pass

        proximo = base + 1
        if proximo >= 100 * 26:
            raise ValueError(f"Capacidade esgotada para o prefixo {prefixo} (00A..99Z).")
        return proximo




    def garantir_id_definitivo_prefixado(origem_depto: str, tipo_doc: str, df_mem: pd.DataFrame):
        # zera o cache de último índice para recomputar com df atualizado
        st.session_state.pop("ultimo_idx_por_prefixo", None)
        ultimo = carregar_ultimo_idx_por_prefixo() or {}

        prefixo = montar_prefixo(origem_depto, tipo_doc)
        base = ultimo.get(prefixo, -1)

        # procura o maior índice já usado para esse prefixo no df em memória
        if df_mem is not None and not df_mem.empty and "ID" in df_mem.columns:
            padrao = re.compile(rf"^{re.escape(prefixo)}(\d{{2}}[A-Z])$")
            for _id in df_mem["ID"].astype(str):
                m = padrao.match(_id)
                if m:
                    try:
                        base = max(base, sufixo_to_idx(m.group(1)))
                    except Exception:
                        pass

        proximo = base + 1
        if proximo >= 100 * 26:
            raise ValueError(f"Capacidade esgotada para o prefixo {prefixo} (00A..99Z).")

        ultimo[prefixo] = proximo
        st.session_state["ultimo_idx_por_prefixo"] = ultimo

        return f"{prefixo}{idx_to_sufixo(proximo)}", df_mem




    # -----------------------------
    # Estado inicial seguro
    # -----------------------------
    if "ja_salvou" not in st.session_state:
        st.session_state.ja_salvou = False

    if "local" not in st.session_state or st.session_state.local not in local_op:
        st.session_state.local = local_op[0] if local_op else ""

    # -----------------------------
    # UI
    # -----------------------------

    # Campos que fazem cálculos (fora do form para recarregar quando mudarem)
    col1, col2 = st.columns(2)
    with col1:
        tipo_doc = st.selectbox("Tipo de Documento*", doc_op, key="sb_tipo_doc")
    with col2:
        origem_depto = st.selectbox(
            "Origem do Documento*",
            dpto_op,
            key="sb_origem"
        )

    # Retenção + descarte (calculado baseado em origem_depto)
    retencao_selecionada, data_prevista_descarte = None, None
    if origem_depto:
        filtro = Retencao_df[Retencao_df["ORIGEM DOCUMENTO SUBMISSÃO"] == origem_depto]
        if not filtro.empty:
            retencao_selecionada = str(filtro["Retenção"].iloc[0]).strip()
            try:
                anos_reten = int(retencao_selecionada.split()[0])
                data_prevista_descarte = datetime.now() + relativedelta(years=anos_reten)
            except Exception:
                pass

    # Form para campos que NÃO devem recarregar o sistema
    with st.form("form_campos_estaticos", clear_on_submit=True):
        conteudo = st.text_input("Conteúdo da Caixa*")

        col3, col4 = st.columns(2)
        with col3:
            origem_submissao = st.selectbox(
                "Origem Departamento Submissão*",
                origens_submissao,
                key="sb_origem_submissao"
            )
        with col4:
            local_atual = st.session_state.get("local", local_op[0] if local_op else "")
            try:
                idx_local = local_op.index(local_atual)
            except ValueError:
                idx_local = 0

            local = st.selectbox(
                "Local*",
                local_op,
                index=idx_local if local_op else 0,
                key="sb_local"
            )
            st.session_state.local = local

        col5, col6 = st.columns(2)
        with col5:
            estante = st.text_input("Estante*", key="tx_estante")
        with col6:
            prateleira = st.text_input("Prateleira*", key="tx_prateleira")

        col7, col8 = st.columns(2)
        with col7:
            caixa = st.text_input("Caixa*", key="tx_caixa")
        with col8:
            
            codificacao = st.text_input("Codificação", key="tx_codificacao") or "N/A"

        col9, col10 = st.columns(2)
        with col9:
            data_ini = st.date_input("Período Utilizado - Início", format="DD/MM/YYYY", key="dt_ini")
            
        with col10:
            
            data_fim = st.date_input("Período Utilizado - Fim", format="DD/MM/YYYY", key="dt_fim")

        col11, col12 = st.columns(2)
        with col11:
            tag = st.text_input("TAG", key="tx_tag") or "N/A"
        with col12:
            lacre = st.text_input("Lacre", key="tx_lacre") or "N/A"

        col13, col14 = st.columns(2)
        with col13:
            livro = st.text_input("Livro", key="tx_livro")
        with col14:
            solicitante = st.text_input("Solicitante*", key="tx_solic")

        # Responsável pelo Arquivamento fica sozinho na última linha
        responsavel = st.selectbox("Responsável pelo Arquivamento*", responsaveis, key="sb_resp")

        # Botão Cadastrar dentro do form
        colA, colB = st.columns([1, 3])
        with colA:
            cadastrar = st.form_submit_button("Cadastrar", type="primary", use_container_width=True)

    # Define e mostra apenas o ID atual (próximo disponível para o prefixo selecionado)
    if origem_depto and tipo_doc:
        try:
            prefixo_atual = montar_prefixo(origem_depto, tipo_doc)

            # Pega do cache/disco o último índice usado por prefixo
            ultimo_idx_por_prefixo = carregar_ultimo_idx_por_prefixo()
            base = ultimo_idx_por_prefixo.get(prefixo_atual, -1)  # -1 significa que ainda não existe

            proximo_idx = base + 1
            # 00A..99Z => 100*26 = 2600 possibilidades (índices 0..2599)
            if proximo_idx >= 100 * 26:
                raise ValueError(f"Capacidade esgotada para o prefixo {prefixo_atual} (00A..99Z).")

            id_atual = f"{prefixo_atual}{idx_to_sufixo(proximo_idx)}"

            # Guarda em sessão para manter consistente com o ID definitivo no salvar
            st.session_state.id_preview = id_atual

            # Mostra somente o ID atual
            st.caption(f"ID atual: **{id_atual}**")

        except Exception as e:
            st.error(f"Erro ao calcular o ID: {e}")

    # Fluxo: Cadastrar
    if cadastrar and not st.session_state.ja_salvou:
        obrig = [caixa, conteudo, origem_depto, solicitante, responsavel, prateleira, local, estante, tipo_doc, origem_submissao]
        if any((c is None) or (str(c).strip() == "") for c in obrig):
            st.warning("Preencha todos os campos obrigatórios e selecione as opções válidas.")
        else:
            try:
                unique_id, df_fresh = garantir_id_definitivo_prefixado(origem_depto, tipo_doc, df)
            except ValueError as e:
                st.error(str(e))
                st.stop()

            novo_doc = {
                "ID": unique_id,
                "Local": local,
                "Estante": estante,
                "Prateleira": prateleira,
                "Caixa": caixa,
                "Codificação": codificacao,
                "Tag": tag,
                "Livro": livro,
                "Lacre": lacre,
                "Tipo de Documento": tipo_doc,
                "Conteúdo da Caixa": conteudo,
                "Departamento Origem": origem_depto,
                "Origem Departamento Submissão": origem_submissao,
                "Responsável Arquivamento": responsavel,
                "Data Arquivamento": datetime.now(),
                "Período Utilizado Início": data_ini,
                "Período Utilizado Fim": data_fim,
                "Status": "ARQUIVADO",
                "Período de Retenção": retencao_selecionada,
                "Data Prevista de Descarte": data_prevista_descarte,
                "Solicitante": solicitante,
            }

            df_final = pd.concat([df_fresh, pd.DataFrame([novo_doc])], ignore_index=True)

            # Salva no SharePoint na aba "Arquivos", mantendo as outras abas
            update_sharepoint_file(df_final, file_name, sheet_name="Arquivos", keep_existing=True)



            st.session_state.ja_salvou = True
            st.cache_data.clear()

            st.info(f"O ID gerado é: {unique_id}")





#====================================#
#         MOVIMENTAR
# ===================================#
elif aba == "Movimentar":
    st.header("📦 Movimentar Documento(s) de Lugar")

    # Entrada múltipla de IDs, separados por vírgula
    ids_raw = st.text_input(
        "Informe um ou mais IDs para movimentação",
        placeholder="Ex: GQES00, GQES01, EQOT12"
    )

    # Parse dos IDs: remove espaços, força maiúsculas e deduplica mantendo ordem
    def parse_ids(s: str):
        if not s:
            return []
        vistos = set()
        ids = []
        for part in s.split(","):
            p = part.strip().upper()
            if p and p not in vistos:
                vistos.add(p)
                ids.append(p)
        return ids

    ids_list = parse_ids(ids_raw)

    # Filtra no DF
    if ids_list:
        df_ids_upper = df.assign(ID_UP=df["ID"].astype(str).str.upper())
        encontrados_df = df_ids_upper[df_ids_upper["ID_UP"].isin(ids_list)].copy()
        encontrados = encontrados_df["ID_UP"].tolist()
        faltando = [i for i in ids_list if i not in encontrados]

        # Feedback ao usuário
        if encontrados:
            st.success(f"{len(encontrados)} documento(s) localizado(s): {', '.join(encontrados)}")
        if faltando:
            st.warning(f"Não encontrado(s): {', '.join(faltando)}")

        if encontrados:
            # Seleção da NOVA localização (aplicada a todos os IDs encontrados)
            local = st.selectbox("Novo Local", list(estruturas.keys()))
            estantes_disp = [str(i + 1).zfill(3) for i in range(estruturas[local]["estantes"])]
            prateleiras_disp = [str(i + 1).zfill(3) for i in range(estruturas[local]["prateleiras"])]

            col1, col2 = st.columns(2)
            with col1:
                estante = st.selectbox("Nova Estante", estantes_disp)
            with col2:
                prateleira = st.selectbox("Nova Prateleira", prateleiras_disp)


            # Confirmar movimentação para TODOS os encontrados
            if st.button("Confirmar Movimentação"):
                idxs = df[df["ID"].astype(str).str.upper().isin(encontrados)].index
                df.loc[idxs, "Local"] = local
                df.loc[idxs, "Estante"] = estante
                df.loc[idxs, "Prateleira"] = prateleira

                update_sharepoint_file(df, file_name, sheet_name="Arquivos", keep_existing=True)


                st.success(f"Movimentação concluída para {len(idxs)} documento(s).")
    else:
        nada = None




#====================================#
#   DESARQUIVAR
# ===================================#
elif aba == "Status":
    st.header("📤 Gerenciar Status do Documento")
    
    # Passo 1: Seleção do ID
    id_input = st.text_input("Digite o ID do Documento", placeholder="Ex: GQES00A")
    
    if id_input:
        id_input = id_input.strip().upper()
        resultado = df[df["ID"] == id_input].copy()
        
        if not resultado.empty:
            # Mostra informações do documento
            st.success(f"✅ Documento encontrado: {id_input}")
            
            resultado_display = resultado[[
                "Status","ID","Conteúdo da Caixa", "Tipo de Documento", "Local", 
                "Estante", "Prateleira", "Caixa", "Solicitante",
                "Responsável Arquivamento", "Data Arquivamento"
            ]].copy()
            
            if "Data Arquivamento" in resultado_display.columns:
                resultado_display["Data Arquivamento"] = pd.to_datetime(
                    resultado_display["Data Arquivamento"]
                ).dt.strftime("%d/%m/%Y")
            
            st.dataframe(resultado_display, use_container_width=True)
            
            # Passo 2: Seleção da Operação
            st.markdown("---")
            st.subheader("🔧 Escolha a Operação")
            
            col_op1, col_op2 = st.columns(2)
            with col_op1:
                operacao_desarquivar = st.checkbox("📤 Desarquivar", value=False, key="cb_desarquivar")
            with col_op2:
                operacao_rearquivar = st.checkbox("📥 Rearquivar", value=False, key="cb_rearquivar")
            
            # Validação: apenas uma operação pode ser selecionada
            if operacao_desarquivar and operacao_rearquivar:
                st.error("❌ Selecione apenas uma operação: Desarquivar OU Rearquivar")
                st.stop()
            elif not operacao_desarquivar and not operacao_rearquivar:
                st.info("ℹ️ Selecione uma operação para continuar")
            else:
                # Passo 3: Captura de Dados
                st.markdown("---")
                st.subheader("📝 Dados da Operação")
                
                responsavel_operacao = st.selectbox(
                    "Responsável pela Operação", 
                    responsaveis,
                    key="sb_resp_operacao"
                )
                
                data_operacao = st.date_input(
                    "Data da Operação",
                    value=datetime.now().date(),
                    format="DD/MM/YYYY",
                    key="dt_operacao"
                )
                
                # Campo específico para desarquivamento parcial
                observacao_operacao = ""
                if operacao_desarquivar:
                    desarquivamento_parcial = st.checkbox("Desarquivamento Parcial", key="cb_parcial")
                    if desarquivamento_parcial:
                        observacao_operacao = st.text_area(
                            "Informe quais documentos da caixa foram desarquivados:",
                            placeholder="Ex: CRF'S DOS PP 01 AO 03, Somente Relatório Clínico",
                            key="tx_obs_parcial"
                        )
                
                # Passo 4: Execução da Operação
                if st.button("🚀 Executar Operação", type="primary", key="btn_executar"):
                    if responsavel_operacao.strip() == "":
                        st.warning("⚠️ Selecione o responsável pela operação")
                    else:
                        try:
                            idx = df[df["ID"] == id_input].index[0]
                            status_atual = str(df.at[idx, "Status"]).strip().upper()

                            if operacao_desarquivar:
                                # Bloqueia se já estiver DESARQUIVADO
                                if status_atual == "DESARQUIVADO":
                                    st.error(f"❌ O documento {id_input} já está desarquivado e não pode ser desarquivado novamente.")
                                    st.stop()

                                # Desarquivar: ARQUIVADO → DESARQUIVADO
                                df.at[idx, "Status"] = "DESARQUIVADO"
                                df.at[idx, "Responsável Desarquivamento"] = responsavel_operacao
                                df.at[idx, "Data Desarquivamento"] = data_operacao.strftime("%d/%m/%Y")

                                # Inicializa colunas se não existirem
                                if "Observação Desarquivamento" not in df.columns:
                                    df["Observação Desarquivamento"] = ""

                                df.at[idx, "Observação Desarquivamento"] = observacao_operacao.strip()
                                st.success(f"✅ Documento {id_input} desarquivado com sucesso!")

                            elif operacao_rearquivar:
                                # Rearquivar: DESARQUIVADO → ARQUIVADO
                                df.at[idx, "Status"] = "ARQUIVADO"
                                df.at[idx, "Responsável Arquivamento"] = responsavel_operacao
                                df.at[idx, "Data Arquivamento"] = data_operacao.strftime("%d/%m/%Y")

                                # Limpa campos de desarquivamento
                                if "Responsável Desarquivamento" in df.columns:
                                    df.at[idx, "Responsável Desarquivamento"] = ""
                                if "Data Desarquivamento" in df.columns:
                                    df.at[idx, "Data Desarquivamento"] = ""
                                if "Observação Desarquivamento" in df.columns:
                                    df.at[idx, "Observação Desarquivamento"] = ""

                                

                            # Salva no Excel
                            update_sharepoint_file(df, file_name, sheet_name="Arquivos", keep_existing=True)
                            st.success(f"✅ Documento {id_input} rearquivado com sucesso!")


                            # Limpa cache e recarrega
                            st.cache_data.clear()
                            st.rerun()

                        except Exception as e:
                            st.error(f"❌ Erro ao executar operação: {e}")

                            
                            # Salva no Excel
                            update_sharepoint_file(df, file_name, sheet_name="Arquivos", keep_existing=True)

                            
                            # Limpa o cache e recarrega
                            st.cache_data.clear()
                            st.rerun()
                            
                        except Exception as e:
                            st.error(f"❌ Erro ao executar operação: {e}")
        else:
            st.warning(f"⚠️ Documento com ID '{id_input}' não encontrado!")

    # Seção de documentos desarquivados
    desarquivados = df[df["Status"] == "DESARQUIVADO"].copy()
    total_desarquivados = len(desarquivados)
    with st.expander(f"📄 Ver Documentos Desarquivados ({total_desarquivados})"):
        if not desarquivados.empty:
            desarquivados["Data Desarquivamento"] = pd.to_datetime(
                desarquivados["Data Desarquivamento"]
            ).dt.strftime("%d/%m/%Y")

            # Inicializa a coluna se estiver faltando
            if "Observação Desarquivamento" not in desarquivados.columns:
                desarquivados["Observação Desarquivamento"] = ""

            st.dataframe(desarquivados[[
                "Status","ID","Conteúdo da Caixa", "Tipo de Documento", "Local", 
                "Estante", "Prateleira", "Caixa", "Solicitante",
                "Responsável Arquivamento", "Data Desarquivamento", 
                "Observação Desarquivamento"
            ]])

            # Destaque visual para desarquivamentos parciais (opcional)
            parciais = desarquivados[desarquivados["Observação Desarquivamento"].str.strip() != ""]
            if not parciais.empty:
                st.markdown("**📌 Desarquivamentos Parciais Identificados:**")
                for _, row in parciais.iterrows():
                    st.markdown(f"- **ID {row['ID']}**: {row['Observação Desarquivamento']}")

        else:
            st.info("Nenhum documento foi desarquivado ainda.")

elif aba == "⚙️ Opções":
    st.header("⚙️ Editar Lista de opções")


    # Editor de dados (tabela editável)
    edited_df = st.data_editor(
        df_selects.copy(),
        use_container_width=True,
        num_rows="dynamic",
        key="selectboxes_editor",
    )

    # Verificação de alterações
    state = st.session_state.get("selectboxes_editor", {})
    houve_alteracao = bool(state.get("edited_rows") or state.get("added_rows") or state.get("deleted_rows"))


    col_a, col_b = st.columns([1, 3])
    with col_a:
        salvar = st.button(
            "Salvar alterações",
            type="primary",
            disabled=not houve_alteracao,
        )

    if houve_alteracao:
        st.info("Foram detectadas alterações não salvas.")

    if salvar and houve_alteracao:
        # Persiste apenas a aba "Selectboxes" no Excel, preservando as demais
        update_sharepoint_file(edited_df, file_name, sheet_name="Selectboxes", keep_existing=True)


        try:
            st.rerun()
        except Exception:
            pass

    st.markdown("---")
    st.header("📅 Período de Retenção")

    # --- Período de Retenção ---
    df_editado = st.data_editor(
        Retencao_df.copy(),
        use_container_width=True,
        num_rows="dynamic", 
        key="selectboxes_retenção",
    )

    houve_alteracao_reten = not df_editado.equals(Retencao_df)

    salvar_reten = st.button(
        "Salvar Retenção",
        type="primary",
        disabled=not houve_alteracao_reten,
        key="btn_salvar_reten"
    )

    if houve_alteracao_reten:
        st.info("Foram detectadas alterações não salvas.")

    if salvar_reten and houve_alteracao_reten:
        update_sharepoint_file(df_editado, file_name, sheet_name="Retenção", keep_existing=True)

        st.success("Retenção atualizada com sucesso.")
        st.rerun()
    

    st.markdown("---")
    st.header("📍 Espaços")

    # --- Espaços ---
    df_editado_espacos = st.data_editor(
        df_espacos.copy(),
        use_container_width=True,
        num_rows="dynamic",
        key="selectboxes_espacos",
    )

    # Verifica mudanças pelo session_state do data_editor
    state_espacos = st.session_state.get("selectboxes_espacos", {})
    houve_alteracao_espacos = bool(
        state_espacos.get("edited_rows") or 
        state_espacos.get("added_rows") or 
        state_espacos.get("deleted_rows")
    )

    salvar_espacos = st.button(
        "Salvar Espaços",
        type="primary",
        disabled=not houve_alteracao_espacos,
        key="btn_salvar_espacos"
    )

    if houve_alteracao_espacos:
        st.info("Foram detectadas alterações não salvas.")

    if salvar_espacos and houve_alteracao_espacos:
        update_sharepoint_file(df_editado_espacos,file_name, sheet_name="Espaços", keep_existing=True)

        st.success("Espaços atualizados com sucesso.")
        st.rerun()




#====================================#
#   CONSULTAR
# ===================================#
elif aba == "Consultar":
    st.header("🔎 Consulta de Documentos")

    with st.expander("🔍 Buscar por Codificação"):
        opcoes_cod = sorted(df["Codificação"].dropna().unique())
        opcoes_cod = sorted(df["Codificação"].dropna().unique())
        cod_select = st.selectbox("Selecione a Codificação do Documento", [""] + list(opcoes_cod))

        if st.button("Buscar por Codificação") and cod_select:
            resultado = df[df["Codificação"] == cod_select].copy()
            if not resultado.empty:
                resultado["Data Arquivamento"] = pd.to_datetime(resultado["Data Arquivamento"]).dt.strftime("%d/%m/%Y")
                st.dataframe(resultado[["ID","Status", "Conteúdo da Caixa", "Tipo de Documento","Departamento Origem", "Local", "Estante", "Prateleira", "Caixa", "Responsável Arquivamento", "Data Arquivamento"]])
            else:
                st.warning("Nenhum documento encontrado com esta codificação.")

    st.subheader("📅 Buscar por Período")
    col1, col2 = st.columns(2)
    with col1:
        data_ini = st.date_input("Data Inicial", value=date.today(), format="DD/MM/YYYY")
    with col2:
        data_fim = st.date_input("Data Final", value=date.today(), format="DD/MM/YYYY")

    if st.button("Buscar por Período"):
        filtrado = df[
            (df["Data Arquivamento"] >= pd.to_datetime(data_ini)) &
            (df["Data Arquivamento"] <= pd.to_datetime(data_fim))
        ].copy()

        if filtrado.empty:
            st.info("Nenhum documento encontrado no período especificado.")
        else:
            filtrado["Data Arquivamento"] = pd.to_datetime(filtrado["Data Arquivamento"]).dt.strftime("%d/%m/%Y")
            st.dataframe(filtrado[["Status", "Codificação", "Conteúdo da Caixa", "Tipo de Documento", "Local", "Estante", "Prateleira", "Caixa", "Responsável Arquivamento", "Data Arquivamento"]])



elif aba == "Desarquivar":
    st.header("📤 Desarquivar Documento")
    id_input = st.text_input("Digite o ID do Documento para desarquivar")

    resultado = df[df["ID"] == id_input.strip().upper()].copy()
    if not resultado.empty:
        resultado_display = resultado[["Status","ID","Conteúdo da Caixa", "Tipo de Documento", "Local", "Estante", "Prateleira", "Caixa", "Responsável Arquivamento", "Data Arquivamento"]]
        resultado_display["Data Arquivamento"] = pd.to_datetime(resultado_display["Data Arquivamento"]).dt.strftime("%d/%m/%Y")
        st.dataframe(resultado_display)

        responsavel_saida = st.selectbox("Responsável pelo Desarquivamento", responsaveis)

        if st.button("Desarquivar"):
            if responsavel_saida.strip() != "":
                idx = df[df["ID"] == id_input.strip().upper()].index[0]
                df.at[idx, "Status"] = "DESARQUIVADO"
                df.at[idx, "Responsável Desarquivamento"] = responsavel_saida
                df.at[idx, "Data Desarquivamento"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

                update_sharepoint_file(df, file_name,sheet_name="Arquivos", keep_existing=True)


                st.success("Documento desarquivado com sucesso!")
            elif responsavel_saida.strip() == "":
                st.warning("Selecione o responsável pelo desarquivamento.")

    elif id_input:
        st.warning("Documento não encontrado.")

    desarquivados = df[df["Status"] == "DESARQUIVADO"].copy()
    total_desarquivados = len(desarquivados)
    with st.expander(f"📄 Ver Documentos Desarquivados ({total_desarquivados})"):
        desarquivados = df[df["Status"] == "DESARQUIVADO"].copy()
        if not desarquivados.empty:
            desarquivados["Data Arquivamento"] = pd.to_datetime(desarquivados["Data Arquivamento"]).dt.strftime("%d/%m/%Y")
            st.dataframe(desarquivados[["Status","ID","Conteúdo da Caixa", "Tipo de Documento", "Local", "Estante", "Prateleira", "Caixa", "Responsável Arquivamento", "Data Arquivamento"]])
        else:
            st.info("Nenhum documento foi desarquivado ainda.")