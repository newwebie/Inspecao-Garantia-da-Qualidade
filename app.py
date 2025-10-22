import streamlit as st
import pandas as pd
from datetime import date, datetime
from dateutil.relativedelta import relativedelta
from typing import Tuple
import io, time, re
import random
from sp_connector import SPConnector 
from urllib.parse import quote

# ===== Config via novo secrets =====
TENANT_ID = st.secrets["graph"]["tenant_id"]
CLIENT_ID = st.secrets["graph"]["client_id"]
CLIENT_SECRET = st.secrets["graph"]["client_secret"]
HOSTNAME = st.secrets["graph"]["hostname"]           
SITE_PATH = st.secrets["graph"]["site_path"]          
LIBRARY   = st.secrets["graph"]["library_name"]    
USER_UPN  = st.secrets.get("onedrive", {}).get("user_upn", "")
file_name = st.secrets["files"]["arquivo"]   

GRAPH = "https://graph.microsoft.com/v1.0"


# ====== Instancia o conector (um único lugar) =======
@st.cache_resource
def _sp():
    return SPConnector(
        TENANT_ID, CLIENT_ID, CLIENT_SECRET,
        hostname=HOSTNAME, site_path=SITE_PATH, library_name=LIBRARY,
        user_upn=USER_UPN,  # se preencher, entra em modo OneDrive
    )

# ===== (mantido) saneamento de nome de aba =====
def _sanitize_sheet_name(name: str) -> str:
    invalid = ['\\', '/', '?', '*', '[', ']']
    for ch in invalid:
        name = name.replace(ch, '_')
    return (name or "Sheet1")[:31]



# ===== Carregar Excel (todas as abas que você usa) =====
@st.cache_data
def carregar_excel():
    try:
        content = _sp().download(file_name)

        sheets = pd.read_excel(io.BytesIO(content), sheet_name=None)
        df          = sheets.get("Arquivos",    pd.DataFrame())
        df_espacos  = sheets.get("Espaços",     pd.DataFrame())
        df_selects  = sheets.get("Selectboxes", pd.DataFrame())
        Retencao_df = sheets.get("Retenção",    pd.DataFrame())
        df_hist     = sheets.get("Histórico",    pd.DataFrame())

        faltando = [n for n, d in [
            ("Arquivos", df),
            ("Espaços", df_espacos),
            ("Selectboxes", df_selects),
            ("Retenção", Retencao_df),
        ] if d.empty]
        if faltando:
            st.warning(f"A(s) aba(s) não encontrada(s) ou vazia(s): {', '.join(faltando)}")

        return df, df_espacos, df_selects, Retencao_df, df_hist
    except Exception as e:
        st.error(f"Erro ao acessar o arquivo (Graph): {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()


def update_sharepoint_file(file_path: str,
                           updates: dict[str, pd.DataFrame] | None = None,
                           *,
                           # retrocompat:
                           df: pd.DataFrame | None = None,
                           sheet_name: str = "Sheet1",
                           df_hist: pd.DataFrame | None = None,
                           history_sheet_name: str | None = None,
                           keep_existing: bool = True,
                           index: bool = False):
    """
    Escreve várias abas de uma vez.
    Use EITHER `updates={"Aba1": df1, "Aba2": df2}` OR o par (df,sheet) + (df_hist,history_sheet_name).
    """
    # validação mínima
    if not isinstance(file_path, str) or not file_path:
        st.error("file_path inválido")
        return

    # monta o pacote de atualizações
    write_map: dict[str, pd.DataFrame] = {}

    if updates is not None:
        # normaliza/sanitiza nomes de abas do dict
        for k, v in (updates or {}).items():
            if v is None:
                continue
            write_map[_sanitize_sheet_name(k)] = v
    else:
        # modo retrocompatível
        if df is not None:
            write_map[_sanitize_sheet_name(sheet_name)] = df
        if df_hist is not None and history_sheet_name:
            write_map[_sanitize_sheet_name(history_sheet_name)] = df_hist

    if not write_map:
        st.warning("Nada para salvar: nenhum dataframe fornecido.")
        return

    attempts = 0
    while True:
        try:
            # lê workbook atual (se existir) para preservar abas
            existing_sheets = {}
            if keep_existing:
                try:
                    content = _sp().download(file_path)
                    existing_sheets = pd.read_excel(io.BytesIO(content), sheet_name=None) or {}
                except Exception:
                    existing_sheets = {}

            # aplica apenas as abas pedidas
            def _append_frames(existing, new):
                if existing is None or (isinstance(existing, pd.DataFrame) and existing.empty):
                    return new
                # une colunas; o que faltar vira NaN
                all_cols = list(dict.fromkeys(
                    (list(existing.columns) if isinstance(existing, pd.DataFrame) else []) + list(new.columns)
                ))
                if isinstance(existing, pd.DataFrame):
                    existing = existing.reindex(columns=all_cols)
                else:
                    existing = pd.DataFrame(existing).reindex(columns=all_cols)
                new = new.reindex(columns=all_cols)
                return pd.concat([existing, new], ignore_index=True)

            # depois (aplica append automático só na aba "Historico")
            for sheet, data in write_map.items():
                if _sanitize_sheet_name(sheet).lower() == "historico":
                    prev = existing_sheets.get(sheet)
                    existing_sheets[sheet] = _append_frames(prev if isinstance(prev, pd.DataFrame) else None, data)
                else:
                    existing_sheets[sheet] = data  # overwrite normal


            # escreve de volta
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                for name, data in existing_sheets.items():
                    (data if isinstance(data, pd.DataFrame) else pd.DataFrame(data)) \
                        .to_excel(writer, sheet_name=_sanitize_sheet_name(name), index=index)
            output.seek(0)
            _sp().upload_small(file_path, output.read(), overwrite=True)

            try:
                st.cache_data.clear()
            except Exception:
                pass
            st.success("Salvo!")
            break

        except Exception as e:
            attempts += 1
            msg = str(e)
            if any(x in msg for x in ["409", "412", "429"]) and attempts < 5:
                st.warning("Arquivo já em uso. Tentando novamente em 5s...")
                time.sleep(5)
                continue
            st.error(f"Erro ao salvar (Graph): {msg}")
            break


# ===== Utilitários de Histórico =====
HISTORY_SHEET_PREFERRED = "Histórico"
HISTORY_SHEET_ALIASES = ("Histórico", "Historico")
HISTORY_COLUMNS = [
    "Evento", "Data", "ID", "Tipo de Documento", "Origem Documento Submissão",
    "Codificação", "Tag", "Conteúdo da Caixa", "Local", "Prateleira", "Estante",
    "Solicitante", "Responsável", "Observação"
]


def _normalize_history_df(df_hist: pd.DataFrame) -> pd.DataFrame:
    """Garante que o DataFrame do histórico possua exatamente as colunas esperadas."""
    df_hist = (df_hist if isinstance(df_hist, pd.DataFrame) else pd.DataFrame()).copy()
    for coluna in HISTORY_COLUMNS:
        if coluna not in df_hist.columns:
            df_hist[coluna] = ""
    return df_hist[HISTORY_COLUMNS]


def get_history_df() -> Tuple[pd.DataFrame, str]:
    """Lê a planilha de histórico garantindo colunas padrão.

    Retorna o DataFrame normalizado e o nome da aba utilizada (existente ou preferida).
    """
    sheet_name = HISTORY_SHEET_PREFERRED
    try:
        content = _sp().download(file_name)
        sheets = pd.read_excel(io.BytesIO(content), sheet_name=None) or {}
        for possible in HISTORY_SHEET_ALIASES:
            hist = sheets.get(possible)
            if isinstance(hist, pd.DataFrame):
                sheet_name = possible
                return _normalize_history_df(hist), sheet_name
    except Exception:
        pass
    return pd.DataFrame(columns=[
        "Mudança", "Data", "ID", "Conteúdo da Caixa", "Local", "Prateleira", "Estante",
        "Solicitante", "Responsável", "Observação"
    ])


def log_history(evento: str, id_val: str, solicitante_val: str, responsavel_val: str,
                data_val: datetime, observacao_val: str = "",
                conteudo_val: str = "", local_val: str = "", prateleira_val: str = "",
                estante_val: str = ""):
    """Acrescenta uma linha no histórico e salva na sheet Historico mantendo abas existentes."""
    try:
        hist_df, hist_sheet = get_history_df()
        nova_linha = {
            "Mudança": str(evento).upper(),
            "Data da Operação": pd.to_datetime(data_val),
            "ID": id_val,
            "Conteúdo da Caixa": conteudo_val,
            "Local": local_val,
            "Prateleira": prateleira_val,
            "Estante": estante_val,
            "Solicitante": solicitante_val,
            "Responsável": responsavel_val,
            "Observação": observacao_val or ""
        }
        novo_hist = pd.concat([hist_df, pd.DataFrame([nova_linha])], ignore_index=True)
        novo_hist = _normalize_history_df(novo_hist)
        update_sharepoint_file(novo_hist, file_name, sheet_name=hist_sheet, keep_existing=True)
    except Exception as e:
        st.warning(f"Não foi possível registrar histórico: {e}")




# ===== Configuração da página =====
st.set_page_config(page_title="Sistema de Arquivo", layout="wide")


# TABs
with st.sidebar:
    aba = st.selectbox("Escolha o que deseja", ["Cadastrar", "Status","Consultar", "Editar", "Movimentar", "Histórico", "⚙️ Opções"])

    #Botao atualizar limpando cache
    if st.button("🔄 Atualizar"):
        st.cache_data.clear()      
        st.cache_resource.clear()
        st.rerun()  


    df, df_espacos, df_selects, Retencao_df, df_hist = carregar_excel()
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
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:         
        st.header("🆕 Cadastrar Documento")
        st.markdown("<br>", unsafe_allow_html=True)
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

    def _duas_letras_aleatorias() -> str:
        return "".join(random.choice("ABCDEFGHIJKLMNOPQRSTUVWXYZ") for _ in range(2))

    def _prefixo_aleatorio_estavel(tipo_doc: str) -> str:
        # Mantém duas letras aleatórias estáveis enquanto o usuário preenche o cadastro
        rand_tipo = st.session_state.get("rand_tipo")
        rand_pref = st.session_state.get("rand_prefix")
        if rand_tipo == tipo_doc and isinstance(rand_pref, str) and len(rand_pref) == 2:
            return rand_pref
        novo = _duas_letras_aleatorias()
        st.session_state["rand_tipo"] = tipo_doc
        st.session_state["rand_prefix"] = novo
        return novo

    def montar_prefixo(origem_depto: str, tipo_doc: str) -> str:
        # Ignora o departamento de origem; usa sigla do tipo primeiro + 2 letras aleatórias
        letras = _prefixo_aleatorio_estavel(tipo_doc)
        return f"{abrev_tipo(tipo_doc)}{letras}"  # 4 letras
    
    # --------------------------------
    # Capacidade do sufixo (AUMENTADA)
    # --------------------------------
    NUM_DIGITS = 3  # 3 -> 000..999  (se precisar mais, use 4 -> 0000..9999)
    CAP_MAX = (10 ** NUM_DIGITS) * 26  # 26.000 IDs por prefixo

    # -----------------------------
    # Conversões N..NL <-> índice (000A..999Z)
    # -----------------------------
    def idx_to_sufixo(idx: int) -> str:
        """
        idx 0..(CAP_MAX-1) -> 'NN..NL' (000A..999Z se NUM_DIGITS=3)
        num = idx // 26, letra = A + (idx % 26)
        """
        if idx < 0 or idx >= CAP_MAX:
            raise ValueError(
                f"Capacidade esgotada para este prefixo "
                f"(000A..{10**NUM_DIGITS - 1:0{NUM_DIGITS}d}Z = {CAP_MAX} IDs)."
            )
        num = idx // 26
        letra_idx = idx % 26
        letra = chr(ord('A') + letra_idx)
        return f"{num:0{NUM_DIGITS}d}{letra}"

    def sufixo_to_idx(nnletra: str) -> int:
        m = re.fullmatch(rf"(\d{{{NUM_DIGITS}}})([A-Z])", nnletra)
        if not m:
            raise ValueError(f"Sufixo inválido: {nnletra}")
        num = int(m.group(1))
        letra = m.group(2)
        return num * 26 + (ord(letra) - ord('A'))

    def extrair_prefixo_e_idx(id_str: str):
        """
        De PPPP + N..NL (ex.: ESA7 000A) -> (prefixo='ESA7', idx=int).
        Ignora formatos fora do padrão.
        """
        if not isinstance(id_str, str):
            return None
        s = id_str.strip().upper()
        m = re.fullmatch(rf"([A-Z0-9]{{4}})(\d{{{NUM_DIGITS}}}[A-Z])", s)
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
        Calcula o próximo índice (0..CAP_MAX-1) para um prefixo, considerando:
          1) cache de 'ultimo_idx_por_prefixo' (do DataFrame em memória)
          2) opcionalmente o df em memória adicional (para prévia)
        """
        ultimo = carregar_ultimo_idx_por_prefixo()
        base = ultimo.get(prefixo, -1)

        # também olha o df em memória (IDs desta sessão já carregados em df, antes de salvar)
        if df_mem is not None and not df_mem.empty and "ID" in df_mem.columns:
            padrao = re.compile(rf"^{re.escape(prefixo)}(\d{{{NUM_DIGITS}}}[A-Z])$")
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
        if proximo >= CAP_MAX:
            raise ValueError(
                f"Capacidade esgotada para o prefixo {prefixo} "
                f"(000A..{10**NUM_DIGITS - 1:0{NUM_DIGITS}d}Z)."
            )
        return proximo





    def garantir_id_definitivo_prefixado(origem_depto: str, tipo_doc: str, df_mem: pd.DataFrame):
        # zera o cache de último índice para recomputar com df atualizado
        st.session_state.pop("ultimo_idx_por_prefixo", None)
        ultimo = carregar_ultimo_idx_por_prefixo() or {}

        prefixo = montar_prefixo(origem_depto, tipo_doc)
        base = ultimo.get(prefixo, -1)

        # procura o maior índice já usado para esse prefixo no df em memória
        if df_mem is not None and not df_mem.empty and "ID" in df_mem.columns:
            padrao = re.compile(rf"^{re.escape(prefixo)}(\d{{{NUM_DIGITS}}}[A-Z])$")
            for _id in df_mem["ID"].astype(str):
                m = padrao.match(_id)
                if m:
                    try:
                        base = max(base, sufixo_to_idx(m.group(1)))
                    except Exception:
                        pass

        proximo = base + 1
        if proximo >= CAP_MAX:
            raise ValueError(
                f"Capacidade esgotada para o prefixo {prefixo} "
                f"(000A..{10**NUM_DIGITS - 1:0{NUM_DIGITS}d}Z)."
            )

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
            retencao_selecionada = str(filtro["Retenção"].iloc(0) if hasattr(filtro["Retenção"], 'iloc') else filtro["Retenção"]).split()[0]
            try:
                anos_reten = int(str(retencao_selecionada).split()[0])
                data_prevista_descarte = datetime.now() + relativedelta(years=anos_reten)
            except Exception:
                pass


    conteudo = st.text_input("Conteúdo da Caixa*")

    col3, col4 = st.columns(2)
    with col3:
        origem_submissao = st.selectbox(
            "Origem Documento Submissão*",
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

    if tipo_doc == "LOGBOOK":
        col9, col10 = st.columns(2)
        with col9:
            data_ini = st.date_input("Período Utilizado - Início", format="DD/MM/YYYY", key="dt_ini")

        with col10:

            data_fim = st.date_input("Período Utilizado - Fim", format="DD/MM/YYYY", key="dt_fim")
    else:
        data_ini = None
        data_fim = None
        for key in ("dt_ini", "dt_fim"):
            st.session_state.pop(key, None)

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

    colA, colB = st.columns([1, 3])
    with colA:
        cadastrar = st.button("Cadastrar", type="primary", key="btn_cadastrar")

    # Define e mostra apenas o ID atual (próximo disponível para o prefixo selecionado)
    if origem_depto and tipo_doc:
        try:
            prefixo_atual = montar_prefixo(origem_depto, tipo_doc)

            # Pega do cache/disco o último índice usado por prefixo
            ultimo_idx_por_prefixo = carregar_ultimo_idx_por_prefixo()
            base = ultimo_idx_por_prefixo.get(prefixo_atual, -1)  # -1 significa que ainda não existe

            proximo_idx = base + 1
            # 000A..999Z => (10**NUM_DIGITS)*26 possibilidades
            if proximo_idx >= CAP_MAX:
                raise ValueError(
                    f"Capacidade esgotada para o prefixo {prefixo_atual} "
                    f"(000A..{10**NUM_DIGITS - 1:0{NUM_DIGITS}d}Z)."
                )

            id_atual = f"{prefixo_atual}{idx_to_sufixo(proximo_idx)}"

            # Guarda em sessão para manter consistente com o ID definitivo no salvar
            st.session_state.id_preview = id_atual

            # Mostra somente o ID atual
            st.caption(f"ID atual: **{id_atual}**")

        except Exception as e:
            st.error(f"Erro ao calcular o ID: {e}")


    # Fluxo: Cadastrar
    if cadastrar and not st.session_state.ja_salvou:
        obrig = [
            caixa,
            conteudo,
            origem_depto,
            solicitante,
            responsavel,
            prateleira,
            local,
            estante,
            tipo_doc,
            origem_submissao,
        ]
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
                "Origem Documento Submissão": origem_submissao,
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
            update_sharepoint_file(
                file_name,
                df=df_final,             
                sheet_name="Arquivos",
                keep_existing=True
            )

            # registra histórico de SOLICITAÇÃO e ARQUIVAMENTO
            log_history(
                evento="SOLICITACAO_ARQUIVAMENTO",
                id_val=unique_id,
                solicitante_val=solicitante,
                responsavel_val=responsavel,
                data_val=datetime.now(),
                observacao_val="",
                conteudo_val=conteudo,
                local_val=local,
                prateleira_val=prateleira,
                estante_val=estante
            )
            log_history(
                evento="ARQUIVAMENTO",
                id_val=unique_id,
                solicitante_val=solicitante,
                responsavel_val=responsavel,
                data_val=datetime.now(),
                observacao_val="",
                conteudo_val=conteudo,
                local_val=local,
                prateleira_val=prateleira,
                estante_val=estante
            )
            # limpa prefixo aleatório para o próximo cadastro
            st.session_state["rand_prefix"] = None
            st.session_state["rand_tipo"] = None

            st.session_state.ja_salvou = True
            st.cache_data.clear()

            st.info(f"O ID gerado é: {unique_id}")
    else:
        st.session_state.ja_salvou = False





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
        # DF com ID_UP para facilitar comparações em maiúsculas
        df_ids_upper = df.assign(ID_UP=df["ID"].astype(str).str.upper())
        encontrados_df = df_ids_upper[df_ids_upper["ID_UP"].isin(ids_list)].copy()
        encontrados = encontrados_df["ID_UP"].tolist()
        faltando = [i for i in ids_list if i not in encontrados]

        # Feedback ao usuário
        if encontrados:
            st.success(f"{len(encontrados)} documento(s) localizado(s): {', '.join(encontrados)}")

            # --------- BLOQUEIO: Status = DESARQUIVADO ---------
            # Separa bloqueados e movíveis
            status_col = "Status" if "Status" in encontrados_df.columns else None
            if status_col:
                bloqueados_mask = encontrados_df[status_col].astype(str).str.upper().eq("DESARQUIVADO")
            else:
                bloqueados_mask = pd.Series([False] * len(encontrados_df), index=encontrados_df.index)

            bloqueados_df = encontrados_df[bloqueados_mask].copy()
            moveis_df    = encontrados_df[~bloqueados_mask].copy()

            # Tabela para bloqueados (não podem ser movimentados)
            if not bloqueados_df.empty:
                st.error(f"{len(bloqueados_df)} documento(s) com status DESARQUIVADO não podem ser movimentados. Listados abaixo:")
                # Formata Data Desarquivamento se existir
                if "Data Desarquivamento" in bloqueados_df.columns:
                    try:
                        bloqueados_df["Data Desarquivamento"] = pd.to_datetime(bloqueados_df["Data Desarquivamento"]).dt.strftime("%d/%m/%Y")
                    except Exception:
                        pass

                # Colunas pedidas (ignorando as que não existem)
                cols_desejadas = [
                    "ID",
                    "Tipo de Documento",
                    "Conteúdo da Caixa",
                    "Data Desarquivamento",
                    "Responsável Desarquivamento",
                ]
                cols_existentes = [c for c in cols_desejadas if c in bloqueados_df.columns]
                if cols_existentes:
                    st.dataframe(bloqueados_df[cols_existentes], use_container_width=True)
                else:
                    st.caption("Nenhuma das colunas esperadas para exibição foi encontrada nos dados.")

            # Mostra também a situação atual (local/estante/prateleira) dos elegíveis
            if not moveis_df.empty:
                try:
                    show_cols = [
                        "ID", "Local", "Estante", "Prateleira",
                        "Data Arquivamento", "Responsável Arquivamento"
                    ]
                    atual_df = df[df["ID"].astype(str).str.upper().isin(moveis_df["ID_UP"])].copy()
                    if "Data Arquivamento" in atual_df.columns:
                        try:
                            atual_df["Data Arquivamento"] = pd.to_datetime(atual_df["Data Arquivamento"]).dt.strftime("%d/%m/%Y")
                        except Exception:
                            pass
                    colunas_existentes = [c for c in show_cols if c in atual_df.columns]
                    if colunas_existentes:
                        st.subheader("📍 Localização atual")
                        st.dataframe(atual_df[colunas_existentes], use_container_width=True)
                except Exception:
                    pass

        if faltando:
            st.warning(f"Não encontrado(s): {', '.join(faltando)}")

        # ---------- UI para movimentar APENAS os elegíveis ----------
        moveis_ids = moveis_df["ID_UP"].tolist() if encontrados else []

        if moveis_ids:
            st.info(f"{len(moveis_ids)} documento(s) elegível(eis) para movimentação.")
            # Seleção da NOVA localização (aplicada a todos os IDs elegíveis)
            local = st.selectbox("Novo Local", list(estruturas.keys()))
            estantes_disp = [str(i + 1).zfill(3) for i in range(estruturas[local]["estantes"])]
            prateleiras_disp = [str(i + 1).zfill(3) for i in range(estruturas[local]["prateleiras"])]

            col1, col2 = st.columns(2)
            with col1:
                estante = st.selectbox("Nova Estante", estantes_disp)
            with col2:
                prateleira = st.selectbox("Nova Prateleira", prateleiras_disp)
            col3 = responsavel_operacao = st.selectbox(
                    "Responsável pela Operação", 
                    responsaveis,
                    key="sb_resp_operacao"
                )

            # Confirmar movimentação para TODOS os elegíveis
            if st.button("Confirmar Movimentação"):
                idxs = df[df["ID"].astype(str).str.upper().isin(moveis_ids)].index

                # --- (opcional) pegar origem antes de mudar, para registrar na observação ---
                cols_prev = [c for c in ["Local", "Estante", "Prateleira"] if c in df.columns]
                origem = ""
                if cols_prev and len(idxs) > 0:
                    orig_uniq = (
                        df.loc[idxs, cols_prev].astype(str).agg("/".join, axis=1).unique()
                    )
                    origem = orig_uniq[0] if len(orig_uniq) == 1 else ""

                # === 1) MONTAR UMA ÚNICA LINHA DE HISTÓRICO ===
                data_operacao = pd.Timestamp.now(tz="America/Sao_Paulo").strftime("%d/%m/%Y")
                ids_movidos = df.loc[idxs, "ID"].astype(str).tolist()
                ids_txt = ", ".join(ids_movidos)

                if origem:
                    observacao = f"{ids_txt} | {origem} → {local}/{estante}/{prateleira}"
                else:
                    observacao = f"{ids_txt} | Novos: {local}/{estante}/{prateleira}"

                registro_hist = {
                    "Data da Operação": data_operacao,
                    "Responsável": responsavel_operacao,
                    "Mudança": "Movimentação",
                    "ID": ids_txt,
                    "Conteúdo da Caixa": "",   # preencha se quiser agregar algo aqui
                    "Observação": observacao,
                }

                # === 2) OPÇÃO A: CARREGA/CRIA 'Historico' E APENDA UMA LINHA PRESERVANDO O QUE JÁ EXISTE ===
                try:
                    df_hist = carregar_excel(file_name, sheet_name="Historico")
                    if df_hist is None or not isinstance(df_hist, pd.DataFrame):
                        raise Exception("Historico inexistente")
                except Exception:
                    df_hist = pd.DataFrame(columns=[
                        "Data da Operação", "Responsável", "Mudança", "ID", "Conteúdo da Caixa", "Observação"
                    ])

                # adiciona 1 linha ao final (sem reescrever as anteriores)
                df_hist.loc[len(df_hist)] = registro_hist

                
                

                # === 3) APLICAR AS MUDANÇAS NA PLANILHA PRINCIPAL E SALVAR ===
                df.loc[idxs, "Local"] = local
                df.loc[idxs, "Estante"] = estante
                df.loc[idxs, "Prateleira"] = prateleira

                # salva a aba 'Historico' mantendo o resto do arquivo
                update_sharepoint_file(file_name, updates={
                    "Historico": df_hist,
                    "Arquivos": df
                    },
                    keep_existing=True
                )

                # Feedback pós-movimentação
                st.success(f"Movimentação concluída para: {', '.join(ids_movidos)}")
                if not bloqueados_df.empty:
                    st.info(
                        "Os seguintes IDs foram ignorados por estarem DESARQUIVADOS: "
                        + ", ".join(bloqueados_df["ID"].astype(str).tolist())
                    )


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
                            
                            # --- após aplicar a operação (antes de salvar o df principal) ---
                            novo_status = df.at[idx, "Status"]  # já atualizado acima
                            mudanca = f"{status_atual.title()} -> {str(novo_status).title()}"

                            # Capturar conteúdo da caixa do registro atual (se existir)
                            conteudo_caixa = ""
                            if "Conteúdo da Caixa" in df.columns:
                                conteudo_caixa = str(df.at[idx, "Conteúdo da Caixa"])

                            # Observação só existe em desarquivamento parcial
                            observacao = observacao_operacao.strip() if operacao_desarquivar else ""

                            # Monta registro do histórico
                            registro_hist = {
                                "Data da Operação": data_operacao.strftime("%d/%m/%Y"),
                                "Responsável": responsavel_operacao,
                                "Mudança": mudanca,
                                "ID": id_input,
                                "Conteúdo da Caixa": conteudo_caixa,
                                "Observação": observacao,
                            }

                            # === utilitário de leitura da planilha "Historico" ===
                            # Troque `load_sharepoint_file` pelo leitor que você já usa no app para ler sheets.
                            try:
                                df_hist = carregar_excel(file_name, sheet_name="Historico")
                                # Se vier vazio/None por alguma razão, inicializa
                                if df_hist is None or not isinstance(df_hist, pd.DataFrame):
                                    raise Exception("Historico inexistente")
                            except Exception:
                                df_hist = pd.DataFrame(columns=[
                                    "Data da Operação",
                                    "Responsável",
                                    "Mudança",
                                    "ID",
                                    "Conteúdo da Caixa",
                                    "Observação",
                                ])

                            # Anexa a nova linha e salva a aba Historico
                            df_hist = pd.concat([df_hist, pd.DataFrame([registro_hist])], ignore_index=True)

                            update_sharepoint_file(
                                file_name,
                                df=df, sheet_name="Arquivos",
                                df_hist=df_hist, history_sheet_name="Historico",
                                keep_existing=True
                            )



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
            desarquivados["Data Desarquivamento"],
            format="%d/%m/%Y",
            errors="coerce",
        )


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
            parciais = desarquivados[desarquivados["Observação Desarquivamento"].astype("string").fillna("").str.contains("parcial", case=False)]

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

        st.rerun()




#====================================#
#   CONSULTAR
# ===================================#
elif aba == "Consultar":
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.subheader("🔎 Consulta de Documentos")
        st.markdown("<br>", unsafe_allow_html=True)

    st.subheader("📄 Buscar por Codificação")
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
    st.markdown("<br>", unsafe_allow_html=True)

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
            st.dataframe(filtrado[["Status", "ID", "Codificação", "Conteúdo da Caixa", "Tipo de Documento", "Local", "Estante", "Prateleira", "Caixa", "Responsável Arquivamento", "Data Arquivamento"]])
        
    st.markdown("<br>", unsafe_allow_html=True)

    # ===== Consulta específica por ID =====
    st.subheader("🎯 Consulta específica")
    st.text("Veja toda informação referente ao documento")
    id_consulta = st.text_input("Informe o ID do documento", key="tx_consulta_id").strip().upper()
    if id_consulta:
        registro = df[df["ID"].astype(str).str.upper() == id_consulta].copy()
        if registro.empty:
            st.warning("ID não encontrado.")
        else:
            # Considera apenas a primeira linha correspondente
            linha = registro.iloc[0]
            colunas_preenchidas = []
            for col in linha.index.tolist():
                val = linha[col]
                if pd.isna(val):
                    continue
                # Trata strings vazias/espacos
                if isinstance(val, str) and val.strip() == "":
                    continue
                colunas_preenchidas.append(col)

            if not colunas_preenchidas:
                st.info("Nenhuma coluna preenchida para este registro.")
            else:
                df_mostrar = pd.DataFrame([linha[colunas_preenchidas].to_dict()])
                # Formata datas conhecidas, se existirem
                for c in [
                    "Data Arquivamento", "Data Desarquivamento", "Período Utilizado Início",
                    "Período Utilizado Fim", "Data Prevista de Descarte"
                ]:
                    if c in df_mostrar.columns:
                        try:
                            df_mostrar[c] = pd.to_datetime(df_mostrar[c]).dt.strftime("%d/%m/%Y")
                        except Exception:
                            pass
                st.dataframe(df_mostrar, use_container_width=True)


elif aba == "Editar":
    st.subheader("✏️ Editar Documentos")
    st.markdown("Pesquise por ID ou pela combinação de Local, Estante e Prateleira para atualizar registros existentes.")
    st.caption("As alterações feitas na tabela são salvas apenas após clicar em \"Salvar alterações\".")

    def _get_series(df_src: pd.DataFrame, coluna: str) -> pd.Series:
        if coluna in df_src.columns:
            return df_src[coluna]
        return pd.Series([""] * len(df_src), index=df_src.index)

    def _colunas_preenchidas(df_target: pd.DataFrame, obrigatorias=None) -> list:
        colunas_validas = []
        obrigatorias = set(obrigatorias or [])
        for coluna in df_target.columns:
            if coluna in obrigatorias:
                if coluna not in colunas_validas:
                    colunas_validas.append(coluna)
                continue
            serie = df_target[coluna]
            serie_sem_na = serie.dropna()
            if serie_sem_na.empty:
                continue
            possui_valor = False
            for valor in serie_sem_na:
                if isinstance(valor, str):
                    if valor.strip():
                        possui_valor = True
                        break
                else:
                    possui_valor = True
                    break
            if possui_valor:
                colunas_validas.append(coluna)
        for coluna in df_target.columns:
            if coluna in obrigatorias and coluna not in colunas_validas:
                colunas_validas.append(coluna)
        return colunas_validas

    def _formatar_valor(valor):
        if pd.isna(valor):
            return ""
        if isinstance(valor, str):
            return valor.strip()
        return str(valor)

    def _obter_primeiro_valor(serie: pd.Series, *chaves: str) -> str:
        for chave in chaves:
            if chave in serie:
                return serie.get(chave, "")
        return ""

    def _renderizar_editor(filtered_df: pd.DataFrame, key_prefix: str):
        if filtered_df.empty:
            st.info("Nenhum documento encontrado com os filtros selecionados.")
            return

        colunas_visiveis = _colunas_preenchidas(filtered_df, COLS_EDITAVEIS)
        if "ID" in filtered_df.columns and "ID" not in colunas_visiveis:
            colunas_visiveis.insert(0, "ID")

        editor_df = filtered_df[colunas_visiveis].copy()
        editor_df.insert(0, "__df_index", filtered_df.index)
        editor_df.reset_index(drop=True, inplace=True)

        # 🔒 Somente estas colunas poderão ser editadas
        COLS_EDITAVEIS = {"Status", "Conteúdo da Caixa"}

        # (opcional) lista de status para select — ajuste conforme seu domínio
        lista_status = ["Pendente", "Arquivado", "Em processamento", "Rearquivar", "Conferido"]

        # Configuração por coluna: tudo desabilitado, exceto Status e Conteúdo da Caixa
        column_config = {
            "__df_index": st.column_config.NumberColumn(
                "Linha",
                help="Identificador interno da linha. Não editar.",
                disabled=True,
            )
        }

        for col in editor_df.columns:
            if col == "__df_index":
                continue

            if col in COLS_EDITAVEIS:
                # Colunas permitidas para edição
                if col == "Conteúdo da Caixa":
                    column_config[col] = st.column_config.TextColumn(
                        "Conteúdo da Caixa",
                        help="Descreva/ajuste o conteúdo da caixa.",
                    )
            else:
                # Todas as demais colunas ficam somente leitura
                column_config[col] = st.column_config.Column(
                    col, disabled=True
                )

        with st.form(f"{key_prefix}_form"):
            edited_df = st.data_editor(
                editor_df,
                use_container_width=True,
                num_rows="fixed",
                key=f"{key_prefix}_editor",
                column_config=column_config,
                hide_index=True,
                # não use disabled=True aqui, senão trava tudo
            )
            col_a, col_b = st.columns([2, 1])
            with col_a:
                resp_alt = st.selectbox(
                    "Responsável pela alteração",
                    responsaveis,
                    key=f"{key_prefix}_responsavel",
                )
            with col_b:
                st.markdown(f"**Momento da alteração:** {datetime.now().strftime('%d/%m/%Y %H:%M')}")
            observacao_alt = st.text_area(
                "Observações adicionais (opcional)",
                key=f"{key_prefix}_observacao",
            )
            salvar_alt = st.form_submit_button("Salvar alterações", type="primary")


        if salvar_alt:
            if not resp_alt or not str(resp_alt).strip():
                st.warning("Selecione o responsável pela alteração antes de salvar.")
                return

            edited_df = pd.DataFrame(edited_df)
            if "__df_index" not in edited_df.columns:
                st.error("Não foi possível identificar as linhas editadas.")
                return

            edited_df.set_index("__df_index", inplace=True)
            edited_df.index = edited_df.index.astype(int)

            original_df = editor_df.copy()
            original_df.set_index("__df_index", inplace=True)
            original_df.index = original_df.index.astype(int)

            alteracoes = {}
            for idx in edited_df.index:
                if idx not in original_df.index:
                    continue
                mudancas = []
                for coluna in edited_df.columns:
                    valor_original = original_df.at[idx, coluna]
                    valor_novo = edited_df.at[idx, coluna]
                    if pd.isna(valor_original) and pd.isna(valor_novo):
                        continue
                    if _formatar_valor(valor_original) == _formatar_valor(valor_novo):
                        continue
                    mudancas.append((coluna, valor_original, valor_novo))
                if mudancas:
                    alteracoes[int(idx)] = mudancas

            if not alteracoes:
                st.info("Nenhuma alteração detectada.")
                return

            momento_alteracao = datetime.now()
            for idx, mudancas in alteracoes.items():
                for coluna, _, valor_novo in mudancas:
                    df.at[idx, coluna] = valor_novo

                linha_final = edited_df.loc[idx]
                descricao_alteracoes = "; ".join(
                    f"{coluna}: '{_formatar_valor(original_df.at[idx, coluna])}' → '{_formatar_valor(novo)}'"
                    for coluna, _, novo in mudancas
                )
                observacao_hist = f"Alterações: {descricao_alteracoes}"
                if observacao_alt and observacao_alt.strip():
                    observacao_hist += f". Observação do usuário: {observacao_alt.strip()}"

                log_history(
                    evento="EDIÇÃO",
                    id_val=str(linha_final.get("ID", "")),
                    solicitante_val=str(linha_final.get("Solicitante", "")),
                    responsavel_val=str(resp_alt),
                    data_val=momento_alteracao,
                    observacao_val=observacao_hist,
                    conteudo_val=str(linha_final.get("Conteúdo da Caixa", "")),
                    local_val=str(linha_final.get("Local", "")),
                    prateleira_val=str(linha_final.get("Prateleira", "")),
                    estante_val=str(linha_final.get("Estante", "")),
                )

            update_sharepoint_file(df, file_name, sheet_name="Arquivos", keep_existing=True)
            st.success(f"{len(alteracoes)} registro(s) atualizado(s).")

    id_busca = st.text_input("Pesquisar por ID", key="editar_busca_id").strip().upper()
    if id_busca:
        id_series = _get_series(df, "ID").astype(str).str.upper()
        filtro_id = id_series.str.contains(id_busca, na=False)
        resultados_id = df[filtro_id].copy() if hasattr(filtro_id, "__len__") else pd.DataFrame()
        if resultados_id.empty:
            st.info("Nenhum documento encontrado para o ID informado.")
        else:
            st.markdown(f"**Resultados para ID contendo {id_busca}:**")
            _renderizar_editor(resultados_id, "editar_por_id")

    local_sel = ""
    estante_sel = ""
    prateleira_sel = ""
    with st.expander("Pesquisar por Local, Estante e Prateleira"):
        col_local, col_estante, col_prateleira = st.columns(3)
        locais_disponiveis = [""] + sorted(_get_series(df, "Local").dropna().astype(str).str.strip().unique().tolist())
        with col_local:
            local_sel = st.selectbox("Local", locais_disponiveis, key="editar_local")

        base_estantes = df
        if local_sel:
            mask_local = _get_series(df, "Local").fillna("").astype(str).str.strip().str.upper() == local_sel.strip().upper()
            base_estantes = df[mask_local].copy()
        estantes_disponiveis = [""] + sorted(_get_series(base_estantes, "Estante").dropna().astype(str).str.strip().unique().tolist())
        with col_estante:
            estante_sel = st.selectbox("Estante", estantes_disponiveis, key="editar_estante")

        base_prateleiras = base_estantes
        if estante_sel:
            mask_estante = _get_series(base_estantes, "Estante").fillna("").astype(str).str.strip().str.upper() == estante_sel.strip().upper()
            base_prateleiras = base_estantes[mask_estante].copy()
        prateleiras_disponiveis = [""] + sorted(_get_series(base_prateleiras, "Prateleira").dropna().astype(str).str.strip().unique().tolist())
        with col_prateleira:
            prateleira_sel = st.selectbox("Prateleira", prateleiras_disponiveis, key="editar_prateleira")

    if local_sel and estante_sel and prateleira_sel:
        mask_local = _get_series(df, "Local").fillna("").astype(str).str.strip().str.upper() == local_sel.strip().upper()
        mask_estante = _get_series(df, "Estante").fillna("").astype(str).str.strip().str.upper() == estante_sel.strip().upper()
        mask_prateleira = _get_series(df, "Prateleira").fillna("").astype(str).str.strip().str.upper() == prateleira_sel.strip().upper()
        mask_combinado = mask_local & mask_estante & mask_prateleira
        resultados_combo = df[mask_combinado].copy()
        if resultados_combo.empty:
            st.info("Nenhum documento encontrado para a combinação selecionada.")
        else:
            st.markdown(f"**Resultados para {local_sel} / {estante_sel} / {prateleira_sel}:**")
            _renderizar_editor(resultados_combo, "editar_por_posicao")
    elif any([local_sel, estante_sel, prateleira_sel]):
        st.caption("Preencha Local, Estante e Prateleira para executar a pesquisa.")
    
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

elif aba == "Histórico":
    st.header("🕓 Histórico de Operações")

    # Carrega histórico
    hist, _ = get_history_df()
    if hist.empty:
        st.info("Nenhum histórico registrado ainda.")
    else:
        # Normaliza datas
        hist["Data da Operação"] = pd.to_datetime(hist["Data da Operação"], errors="coerce")

        # Filtros
        colf1, colf2, colf3 = st.columns(3)
        with colf1:
            f_id = st.text_input("ID")
        with colf2:
            f_evento = st.selectbox("Mudança", [""] + sorted(hist["Mudança"].dropna().unique().tolist()))
        with colf3:
            f_resp = st.selectbox("Responsável", [""] + sorted(hist["Responsável"].dropna().unique().tolist()))

        # Aplica filtros
        filtrado = hist.copy()
        if f_evento:
            filtrado = filtrado[filtrado["Mudança"] == f_evento]
        if f_resp:
            filtrado = filtrado[filtrado["Responsável"] == f_resp]
        if f_id:
            filtro_id = f_id.strip().upper()
            filtrado = filtrado[filtrado["ID"].astype(str).str.upper().str.contains(filtro_id, na=False)]


        filtrado = filtrado.sort_values("Data da Operação", ascending=False)
        # Formata data para exibição
        show = filtrado.copy()
        show["Data da Operação"] = pd.to_datetime(show["Data da Operação"]).dt.strftime("%d/%m/%Y %H:%M")
        # Define ordem de colunas priorizando as solicitadas, exibindo apenas as que existirem
        preferred_cols = [
            "Mudança", "Data da Operação", "ID", "Conteúdo da Caixa", "Local", "Prateleira", "Estante",
            "Solicitante", "Responsável", "Observação"
        ]
        cols_to_show = [c for c in preferred_cols if c in show.columns]
        if cols_to_show:
            st.dataframe(show[cols_to_show], use_container_width=True)
        else:
            st.dataframe(show, use_container_width=True)
