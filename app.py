import streamlit as st
import pandas as pd
import datetime
import io
import re
import time
import gspread
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Auditoria & Inventário - JBA (TESTE DRIVE)", layout="wide")

# --- CONEXÃO DIRETA COM GOOGLE SHEETS (GSPREAD) ---
def obter_conexao_sheets():
    try:
        url = st.secrets["gconfigs"]["spreadsheet_url"]
        gc = gspread.public_client()
        sh = gc.open_by_url(url)
        return sh
    except Exception as e:
        try:
            url = st.secrets["gconfigs"]["spreadsheet_url"]
            gc = gspread.api_client()
            return gc.open_by_url(url)
        except Exception:
            return None

def ler_aba(nome_aba):
    """Lê uma aba pública da planilha."""
    try:
        sh = obter_conexao_sheets()
        if sh:
            ws = sh.worksheet(nome_aba)
            dados = ws.get_all_records()
            return pd.DataFrame(dados).fillna("")
        return pd.DataFrame()
    except Exception:
        try:
            url = st.secrets["gconfigs"]["spreadsheet_url"]
            key = url.split("/d/")[1].split("/")[0]
            csv_url = f"https://docs.google.com/spreadsheets/d/{key}/gviz/tq?tqx=out:csv&sheet={nome_aba}"
            df = pd.read_csv(csv_url)
            return df.fillna("")
        except Exception:
            return pd.DataFrame()

def salvar_lote_aba(nome_aba, novos_dados_df):
    """Envia novos dados para a planilha."""
    try:
        df_ex = ler_aba(nome_aba)
        if not df_ex.empty:
            df_final = pd.concat([df_ex, novos_dados_df], ignore_index=True)
        else:
            df_final = novos_dados_df
        
        sh = obter_conexao_sheets()
        if sh:
            ws = sh.worksheet(nome_aba)
            ws.clear()
            ws.update([df_final.columns.values.tolist()] + df_final.values.tolist())
            st.cache_data.clear()
            return True
        return False
    except Exception as e:
        st.error(f"Erro ao salvar na aba {nome_aba}: {e}")
        return False

def atualizar_aba_completa(nome_aba, df_completo):
    """Substitui o conteúdo de uma aba."""
    try:
        sh = obter_conexao_sheets()
        if sh:
            ws = sh.worksheet(nome_aba)
            ws.clear()
            ws.update([df_completo.columns.values.tolist()] + df_completo.values.tolist())
            st.cache_data.clear()
            return True
        return False
    except Exception as e:
        st.error(f"Erro ao atualizar aba {nome_aba}: {e}")
        return False

# --- LISTA OFICIAL E UNIFICADA DE ESTOQUES JBA ---
LISTA_ESTOQUES_FIXA = [
    {"id": "1077", "desc": "JBA - CLASSE D"}, {"id": "1078", "desc": "JBA - COPA E COZINHA"},
    {"id": "1080", "desc": "JBA - DADOS - CLIENTE"}, {"id": "1082", "desc": "JBA - VIVO VITA - CLIENTE"},
    {"id": "1084", "desc": "JBA - EPI-EPC"}, {"id": "1086", "desc": "JBA - EQUIPAMENTOS"},
    {"id": "1088", "desc": "JBA - FERRAMENTAL"}, {"id": "1089", "desc": "JBA - KIT FERRAMENTAL CONTRATACOES"},
    {"id": "1090", "desc": "JBA - FERRAMENTAS DE CANTEIRO"}, {"id": "1104", "desc": "JBA - MATERIAL DE ESCRITORIO - SUPRIMENTOS DE INFORMATICA"},
    {"id": "1106", "desc": "JBA - MOBILIARIO"}, {"id": "1113", "desc": "1385 - MANUTENCAO JBA - CLIENTE"},
    {"id": "1118", "desc": "JBA - PROPRIO GERAL"}, {"id": "1122", "desc": "JBA - GRAND OBRAS IMPLANTACAO"},
    {"id": "1140", "desc": "JBA - SPEEDY/FTTX - CLIENTE"}, {"id": "1144", "desc": "1385 - MANUTENCAO JBA CLIENTE RESERVADO"},
    {"id": "1149", "desc": "JBA - UNIFORME"}, {"id": "2183", "desc": "1071 - BOL IMPLANTANCAO JBA - CLIENTE"},
    {"id": "2185", "desc": "JBA - PROPRIO FATURA B PLANTA EXTERNA - BDI"}, {"id": "2188", "desc": "1071 - IMPLANTACAO JBA CLIENTE RESERVADO"},
    {"id": "2190", "desc": "JBA - DEPARTAMENTO T.I"}, {"id": "2194", "desc": "JBA - KITS FERRAMENTAL - DEVOLUCAO"},
    {"id": "2197", "desc": "JBA - EQUIPAMENTOS TI"}, {"id": "2641", "desc": "1259 - IMPLANTACAO JBA - MATERIAL REUTILIZACAO"},
    {"id": "2643", "desc": "1724 - MANUTENCAO JBA - MATERIAL REUTILIZACAO"}, {"id": "2725", "desc": "JBA - RESERVA TIM"},
    {"id": "2983", "desc": "JBA - FORNECEDORES P/ MANUTENCAO - RECARGA"}, {"id": "3193", "desc": "JBA - PROPRIO MATERIAL REAPROVEITAVEL"},
    {"id": "3395", "desc": "LPA - FTTX - CLIENTE"}, {"id": "3484", "desc": "JBA - CELULARES DEFEITO"},
    {"id": "3546", "desc": "JBA - CELULARES"}
]
MAPA_ESTOQUES_DESC = {item['id']: item['desc'] for item in LISTA_ESTOQUES_FIXA}

def limpar_cache_aplicacao():
    st.cache_data.clear()

def limpar_documento(doc):
    return str(doc).strip().replace(".", "").replace("-", "").replace("/", "")

def extrair_id_estoque_do_nome(nome_inventario):
    numeros = re.findall(r'\b\d{4}\b', str(nome_inventario))
    for num in numeros:
        if num in MAPA_ESTOQUES_DESC:
            return num
    return ""

def converter_para_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Relatorio')
    return output.getvalue()

def gerar_relatorio_consolidado_excel(df_contagens, lista_estoques_ref):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        linhas_resumo = []
        for est in lista_estoques_ref:
            est_id = str(est["id"]).strip()
            df_est = df_contagens[df_contagens['id_estoque'].astype(str).str.strip() == est_id] if not df_contagens.empty else pd.DataFrame()
            
            total_itens = len(df_est)
            acertos = len(df_est[df_est['diferenca'] == 0]) if total_itens > 0 else 0
            erros = len(df_est[df_est['diferenca'] != 0]) if total_itens > 0 else 0
            qtd_segunda = len(df_est[df_est['fase_contagem'] == '2a Contagem']) if 'fase_contagem' in df_est.columns and total_itens > 0 else 0
            acuracidade_pct = (acertos / total_itens) if total_itens > 0 else 0.0
            u_data_est = str(df_est['data_hora'].max()) if not df_est.empty and 'data_hora' in df_est.columns else "—"
            
            obs_erros = []
            if not df_est.empty and 'observacao' in df_est.columns:
                obs_validas = df_est[df_est['observacao'].fillna('').str.strip() != '']['observacao'].unique().tolist()
                if obs_validas: obs_erros.append("; ".join(map(str, obs_validas)))
            if erros > 0 and not df_est.empty:
                erros_cods = df_est[df_est['diferenca'] != 0]['cod_produto'].unique().tolist()
                obs_erros.append(f"Divergências: {', '.join(map(str, erros_cods[:5]))}" + ("..." if len(erros_cods) > 5 else ""))
            
            linhas_resumo.append({
                "Id. Estoq. Físico": est_id, "Desc. Estoque Físico": est["desc"],
                "Total de itens": total_itens, "Acertos": acertos, "Erros": erros,
                "2º Contagem": qtd_segunda, "Acuracidade": acuracidade_pct,
                "Última Contagem": u_data_est, "Obs": " | ".join(obs_erros) if obs_erros else "Nenhuma observação"
            })
            
        df_resumo = pd.DataFrame(linhas_resumo)
        df_resumo.to_excel(writer, index=False, sheet_name='Acuracidade')
        
        worksheet = writer.sheets['Acuracidade']
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

        for row in range(2, len(df_resumo) + 2):
            worksheet[f'G{row}'].number_format = '0.00%'
            cell_erros = worksheet[f'E{row}']
            if cell_erros.value and int(cell_erros.value) > 0: cell_erros.fill = red_fill
            elif cell_erros.value == 0 and worksheet[f'C{row}'].value and int(worksheet[f'C{row}'].value) > 0: cell_erros.fill = green_fill

        if not df_contagens.empty:
            for est_id, group in df_contagens.groupby('id_estoque'):
                nome_aba = str(est_id).strip()[:31]
                if nome_aba:
                    group_limpo = group.drop(columns=['id'], errors='ignore')
                    group_limpo.to_excel(writer, index=False, sheet_name=nome_aba)
                    ws_est = writer.sheets[nome_aba]
                    rule_error = CellIsRule(operator='notEqual', formula=['0'], stopIfTrue=True, fill=red_fill)
                    rule_ok = CellIsRule(operator='equal', formula=['0'], stopIfTrue=True, fill=green_fill)
                    ws_est.conditional_formatting.add(f"I2:I{len(group_limpo)+1}", rule_error)
                    ws_est.conditional_formatting.add(f"I2:I{len(group_limpo)+1}", rule_ok)

    return output.getvalue()

# --- ESTILOS CSS ---
st.markdown("""
    <style>
    div.stButton > button:first-child[kind="primary"] { background-color: #d35400; border-color: #d35400; color: white; font-weight: bold; }
    div.stButton > button:first-child[kind="primary"]:hover { background-color: #e67e22; border-color: #e67e22; color: white; }
    .card-lateral { background-color: #1a233a; padding: 10px; border-radius: 8px; margin-bottom: 8px; border-left: 4px solid #2563eb; }
    .card-lateral-titulo { color: #93c5fd; font-size: 11px; font-weight: bold; }
    .card-lateral-valor { color: white; font-size: 22px; font-weight: bold; }
    .caixa-descricao-produto { background-color: #ebf8ff; border-left: 5px solid #3182ce; padding: 12px 16px; border-radius: 6px; font-size: 16px; font-weight: bold; color: #1a365d; margin: 10px 0; }
    </style>
""", unsafe_allow_html=True)

# INICIALIZAÇÃO DE ESTADOS
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'operador' not in st.session_state: st.session_state.operador = ""
if 'perfil_usuario' not in st.session_state: st.session_state.perfil_usuario = "Almoxarife"
if 'tela_acesso' not in st.session_state: st.session_state.tela_acesso = "login"
if 'contador_reset' not in st.session_state: st.session_state.contador_reset = 0
if 'contador_reset_sup' not in st.session_state: st.session_state.contador_reset_sup = 0
if 'bases_supervisor_por_inv' not in st.session_state: st.session_state.bases_supervisor_por_inv = {}
if 'pagina_historico' not in st.session_state: st.session_state.pagina_historico = 1

# MEMÓRIA RAM LOCAL PARA CONTAGEM EM LOTE
if 'buffer_ram_contagens' not in st.session_state:
    st.session_state.buffer_ram_contagens = []
if 'ultima_sincronizacao' not in st.session_state:
    st.session_state.ultima_sincronizacao = time.time()

# FUNÇÃO DE SINCRONIZAÇÃO EM LOTE RAM -> GOOGLE SHEETS
def sincronizar_ram_com_banco():
    if not st.session_state.buffer_ram_contagens:
        return 0
    
    df_novos = pd.DataFrame(st.session_state.buffer_ram_contagens)
    sucesso = salvar_lote_aba("contagens", df_novos)
    
    if sucesso:
        df_ultimas = ler_aba("ultima_contagem_estoques")
        novas_datas = []
        for item in st.session_state.buffer_ram_contagens:
            if item.get('id_estoque'):
                novas_datas.append({'id_estoque': item['id_estoque'], 'ultima_data': item['data_hora']})
        if novas_datas:
            df_novas_dt = pd.DataFrame(novas_datas)
            if not df_ultimas.empty:
                df_ultimas = pd.concat([df_ultimas, df_novas_dt], ignore_index=True).drop_duplicates(subset=['id_estoque'], keep='last')
            else:
                df_ultimas = df_novas_dt
            atualizar_aba_completa("ultima_contagem_estoques", df_ultimas)

        qtd_salva = len(st.session_state.buffer_ram_contagens)
        st.session_state.buffer_ram_contagens = []
        st.session_state.ultima_sincronizacao = time.time()
        limpar_cache_aplicacao()
        return qtd_salva
    return 0

# --- TELA DE LOGIN CENTRALIZADA ---
if not st.session_state.logged_in:
    col_vaz1, col_central, col_vaz2 = st.columns([1, 1.2, 1])
    with col_central:
        if st.session_state.tela_acesso == "login":
            st.title("🔒 Acesso ao Sistema JBA (Drive)")
            with st.form("login_form"):
                identificador = st.text_input("CPF ou E-mail")
                senha = st.text_input("Senha", type="password")
                if st.form_submit_button("Entrar no Sistema", type="primary", use_container_width=True):
                    id_limpo = identificador.strip()
                    doc_limpo = limpar_documento(id_limpo)
                    
                    df_usrs = ler_aba("usuarios")
                    if not df_usrs.empty:
                        df_usrs.columns = [str(c).strip().lower() for c in df_usrs.columns]
                        u_match = df_usrs[((df_usrs['email'].astype(str) == id_limpo) | (df_usrs['cpf'].astype(str) == doc_limpo)) & (df_usrs['senha'].astype(str) == senha.strip())]
                        if not u_match.empty:
                            st.session_state.logged_in = True
                            st.session_state.operador = u_match.iloc[0]['nome']
                            st.session_state.perfil_usuario = u_match.iloc[0]['perfil'] or "Almoxarife"
                            limpar_cache_aplicacao()
                            st.rerun()
                        else: st.error("❌ Credenciais incorretas.")
                    else:
                        if senha == "123":
                            st.session_state.logged_in = True
                            st.session_state.operador = "Administrador Tel"
                            st.session_state.perfil_usuario = "Administrador"
                            limpar_cache_aplicacao()
                            st.rerun()
                        else: st.error("❌ Base de usuários vazia no Google Sheets.")
            if st.button("📝 Criar nova conta de colaborador", use_container_width=True):
                st.session_state.tela_acesso = "cadastro"
                st.rerun()

        elif st.session_state.tela_acesso == "cadastro":
            st.title("📝 Cadastro de Colaborador")
            with st.form("cadastro_form"):
                novo_nome = st.text_input("Nome Completo")
                novo_cpf = st.text_input("CPF (Apenas números)")
                novo_email = st.text_input("E-mail")
                nova_senha = st.text_input("Senha", type="password")
                confirma_senha = st.text_input("Confirme a Senha", type="password")
                if st.form_submit_button("Finalizar Cadastro", type="primary", use_container_width=True):
                    cpf_l = limpar_documento(novo_cpf)
                    if not novo_nome or not cpf_l or not novo_email or not nova_senha: st.error("⚠️ Preencha todos os campos!")
                    elif nova_senha != confirma_senha: st.error("❌ Senhas divergentes!")
                    else:
                        df_novo_usr = pd.DataFrame([{
                            "nome": novo_nome.strip(), "cpf": cpf_l,
                            "email": novo_email.strip(), "senha": nova_senha, "perfil": "Almoxarife"
                        }])
                        salvar_lote_aba("usuarios", df_novo_usr)
                        st.success("✅ Cadastro realizado no Google Sheets!")
                        st.session_state.tela_acesso = "login"
                        st.rerun()
            if st.button("◀ Voltar para o Login"):
                st.session_state.tela_acesso = "login"
                st.rerun()

# --- APLICAÇÃO PRINCIPAL ---
else:
    df_inventarios = ler_aba("inventarios")
    if not df_inventarios.empty:
        df_inventarios.columns = [str(c).strip().lower() for c in df_inventarios.columns]

    df_inventarios_sup = ler_aba("inventarios_supervisor")
    if not df_inventarios_sup.empty:
        df_inventarios_sup.columns = [str(c).strip().lower() for c in df_inventarios_sup.columns]

    eh_supervisor = (st.session_state.perfil_usuario == "Administrador") or ("admin" in st.session_state.operador.lower())

    tempo_passado = time.time() - st.session_state.ultima_sincronizacao
    if tempo_passado >= 7200 and st.session_state.buffer_ram_contagens:
        qtd_auto = sincronizar_ram_com_banco()
        st.toast(f"🔄 Auto-sincronização realizada: {qtd_auto} itens salvos na planilha!")

    # --- BARRA LATERAL ---
    with st.sidebar:
        st.write(f"👤 **{st.session_state.operador}** ({st.session_state.perfil_usuario})")
        
        pendentes_ram = len(st.session_state.buffer_ram_contagens)
        if pendentes_ram > 0:
            st.warning(f"⚡ **{pendentes_ram} bips** guardados na RAM local!")
        else:
            st.success("🟢 Memória RAM sincronizada com Drive.")

        col_s1, col_s2 = st.columns(2)
        with col_s1:
            if st.button("🔄 Atualizar", use_container_width=True):
                limpar_cache_aplicacao()
                st.rerun()
        with col_s2:
            if st.button("🚪 Sair", use_container_width=True):
                if pendentes_ram > 0:
                    sincronizar_ram_com_banco()
                st.session_state.logged_in = False
                st.session_state.operador = ""
                limpar_cache_aplicacao()
                st.rerun()
            
        st.markdown("---")
        st.write("📁 **Seleção de Inventário**")
        if df_inventarios.empty or 'id' not in df_inventarios.columns:
            id_inventario_atual = None
            inventario_selected_obj = None
            id_pasta_limpo_base = ""
            st.info("Crie um inventário abaixo.")
        else:
            lista_inv = [f"{row['id']} – {row['nome']} ({row['status']})" for idx, row in df_inventarios.iterrows()]
            inventario_selected = st.selectbox("Selecione a Pasta", lista_inv, index=0, key="sb_pasta_ativa")
            id_inventario_atual = inventario_selected.split(" – ")[0]
            inventario_selected_obj = df_inventarios[df_inventarios['id'].astype(str) == id_inventario_atual].iloc[0]
            id_pasta_limpo_base = str(id_inventario_atual).replace("#", "").strip()

        # UPLOAD DA BASE
        st.write("📂 **Upload Base de Dados (.xlsx)**")
        if inventario_selected_obj is not None and str(inventario_selected_obj['status']) == "Aberto":
            uploader_key = f"func_excel_loader_{id_pasta_limpo_base if id_pasta_limpo_base else 'vazio'}_{st.session_state.contador_reset}"
            arquivo_excel = st.file_uploader("Suba o arquivo Excel (.xlsx)", type=["xlsx"], label_visibility="collapsed", key=uploader_key)
            
            if arquivo_excel is not None and id_pasta_limpo_base:
                with st.spinner("🚀 Processando base para o Google Drive..."):
                    df_upload_temp = pd.read_excel(arquivo_excel)
                    mapa_cols_upload = {str(col).replace('\xa0', ' ').strip().lower().replace(" ", "").replace(".", "").replace("ó", "o").replace("á", "a"): col for col in df_upload_temp.columns}

                    col_cod, col_desc = mapa_cols_upload.get("codproduto"), mapa_cols_upload.get("descproduto")
                    col_est, col_uni = mapa_cols_upload.get("descestoquefisico"), mapa_cols_upload.get("unidmedida")
                    col_qtd, col_idest = mapa_cols_upload.get("qtdestoque"), mapa_cols_upload.get("idestoquefisico")
                    col_lote, col_ativo = mapa_cols_upload.get("lote"), mapa_cols_upload.get("ativo")

                    est_id_fallback = extrair_id_estoque_do_nome(inventario_selected_obj['nome'])
                    est_desc_fallback = MAPA_ESTOQUES_DESC.get(est_id_fallback, 'ESTOQUE FÍSICO')

                    if not col_cod or not col_desc:
                        st.error("❌ Planilha fora do padrão JBA.")
                    else:
                        df_base_existente = ler_aba("itens_base_inventario")
                        if not df_base_existente.empty:
                            df_base_existente.columns = [str(c).strip().lower() for c in df_base_existente.columns]
                            df_base_existente = df_base_existente[df_base_existente['inventario_id'].astype(str) != id_pasta_limpo_base]

                        novas_linhas = []
                        for _, r in df_upload_temp.iterrows():
                            novas_linhas.append({
                                "inventario_id": id_pasta_limpo_base,
                                "cod_produto": str(r[col_cod]).strip() if col_cod and pd.notna(r[col_cod]) else '',
                                "desc_produto": str(r[col_desc]).strip() if col_desc and pd.notna(r[col_desc]) else '',
                                "desc_estoque_fisico": str(r[col_est]).strip() if col_est and pd.notna(r[col_est]) and str(r[col_est]).strip() != '' else est_desc_fallback,
                                "unid_medida": str(r[col_uni]).strip() if col_uni and pd.notna(r[col_uni]) else '',
                                "qtd_estoque": int(pd.to_numeric(r[col_qtd], errors='coerce') or 0) if col_qtd else 0,
                                "id_estoque_fisico": str(r[col_idest]).strip() if col_idest and pd.notna(r[col_idest]) and str(r[col_idest]).strip() != '' else est_id_fallback,
                                "lote": str(r[col_lote]).strip() if col_lote and pd.notna(r[col_lote]) and str(r[col_lote]).lower() != 'nan' else '',
                                "ativo": str(r[col_ativo]).strip() if col_ativo and pd.notna(r[col_ativo]) and str(r[col_ativo]).lower() != 'nan' else ''
                            })
                        
                        df_upload_final = pd.concat([df_base_existente, pd.DataFrame(novas_linhas)], ignore_index=True)
                        atualizar_aba_completa("itens_base_inventario", df_upload_final)
                        limpar_cache_aplicacao()
                        st.success("✅ Base Carregada no Google Sheets!")

        base_todas = ler_aba("itens_base_inventario")
        if not base_todas.empty:
            base_todas.columns = [str(c).strip().lower() for c in base_todas.columns]
            base_sistema_atual = base_todas[base_todas['inventario_id'].astype(str) == id_pasta_limpo_base] if id_pasta_limpo_base else None
        else:
            base_sistema_atual = None

        if inventario_selected_obj is not None and str(inventario_selected_obj['status']) == "Aberto" and base_sistema_atual is not None and not base_sistema_atual.empty:
            st.markdown("---")
            if st.button("🚀 Salvar Base e Iniciar 1ª Contagem", type="primary", use_container_width=True):
                df_inv_att = ler_aba("inventarios")
                df_inv_att.columns = [str(c).strip().lower() for c in df_inv_att.columns]
                df_inv_att.loc[df_inv_att['id'].astype(str) == str(id_inventario_atual), 'status'] = '1a Contagem'
                atualizar_aba_completa("inventarios", df_inv_att)
                limpar_cache_aplicacao()
                st.success("🔒 1ª Contagem liberada.")
                st.rerun()

        # CRIAÇÃO DE NOVO INVENTÁRIO
        with st.expander("➕ Criar Novo Inventário"):
            with st.form("form_novo", clear_on_submit=True):
                novo_nome = st.text_input("Nome do Inventário")
                if st.form_submit_button("Criar Pasta", type="primary"):
                    if not novo_nome.strip():
                        st.error("⚠️ Digite um nome para a pasta!")
                    else:
                        df_existentes = ler_aba("inventarios")
                        if not df_existentes.empty:
                            df_existentes.columns = [str(c).strip().lower() for c in df_existentes.columns]
                            ids_num = df_existentes['id'].astype(str).str.replace('#', '', regex=False)
                            maior_id = pd.to_numeric(ids_num, errors='coerce').max()
                            if pd.isna(maior_id): maior_id = 0
                            novo_id_str = f"#{int(maior_id) + 1}"
                        else:
                            novo_id_str = "#1"

                        df_novo_inv = pd.DataFrame([{
                            "id": novo_id_str,
                            "nome": novo_nome.strip(),
                            "data": datetime.date.today().strftime("%Y-%m-%d"),
                            "status": "Aberto",
                            "total_itens": 0,
                            "acuracidade_final": "0%"
                        }])

                        ok = salvar_lote_aba("inventarios", df_novo_inv)
                        if ok:
                            st.success(f"✅ Pasta {novo_id_str} criada no Google Sheets!")
                            limpar_cache_aplicacao()
                            time.sleep(1)
                            st.rerun()

        # FECHAMENTO DO INVENTÁRIO
        pode_fechar, itens_faltantes, itens_pendentes_2a = False, [], []
        if inventario_selected_obj is not None and str(inventario_selected_obj['status']) in ["1a Contagem", "2a Contagem"] and base_sistema_atual is not None:
            df_cnts_todas = ler_aba("contagens")
            if not df_cnts_todas.empty:
                df_cnts_todas.columns = [str(c).strip().lower() for c in df_cnts_todas.columns]
                df_cnts_pasta = df_cnts_todas[df_cnts_todas['inventario_id'].astype(str) == id_pasta_limpo_base]
            else:
                df_cnts_pasta = pd.DataFrame()
            
            set_contados_triade = set()
            if not df_cnts_pasta.empty:
                for _, r in df_cnts_pasta.iterrows():
                    c_str = str(r.get('cod_produto', '')).upper().strip()
                    l_str = str(r.get('lote', '')).upper().strip()
                    a_str = str(r.get('ativo', '')).upper().strip()
                    set_contados_triade.add(f"{c_str}_{l_str}_{a_str}")

            for r_ram in st.session_state.buffer_ram_contagens:
                if str(r_ram['inventario_id']) == id_pasta_limpo_base:
                    set_contados_triade.add(f"{str(r_ram['cod_produto']).upper().strip()}_{str(r_ram['lote']).upper().strip()}_{str(r_ram['ativo']).upper().strip()}")

            if str(inventario_selected_obj['status']) == "1a Contagem":
                for _, r_b in base_sistema_atual.iterrows():
                    c_b = str(r_b.get('cod_produto', '')).upper().strip()
                    l_b = str(r_b.get('lote', '')).upper().strip()
                    a_b = str(r_b.get('ativo', '')).upper().strip()
                    if f"{c_b}_{l_b}_{a_b}" not in set_contados_triade: itens_faltantes.append(c_b)
                if len(itens_faltantes) == 0: pode_fechar = True
            elif str(inventario_selected_obj['status']) == "2a Contagem":
                if not df_cnts_pasta.empty:
                    itens_pendentes_2a = [str(r['cod_produto']).upper().strip() for _, r in df_cnts_pasta.iterrows() if str(r.get('fase_contagem')) == '2a Contagem']
                if len(itens_pendentes_2a) == 0: pode_fechar = True

            st.markdown("---")
            def fechar_e_preservar_historico(id_inv, id_limpo):
                if pendentes_ram > 0:
                    sincronizar_ram_com_banco()
                df_cnts = ler_aba("contagens")
                if not df_cnts.empty:
                    df_cnts.columns = [str(c).strip().lower() for c in df_cnts.columns]
                    df_f = df_cnts[df_cnts['inventario_id'].astype(str) == id_limpo]
                else: df_f = pd.DataFrame()
                
                tot = len(df_f)
                acertos = len(df_f[pd.to_numeric(df_f['diferenca'], errors='coerce') == 0]) if tot > 0 else 0
                pct_acu = f"{(acertos / tot)*100:.1f}%" if tot > 0 else "0%"
                
                df_inv_up = ler_aba("inventarios")
                if not df_inv_up.empty:
                    df_inv_up.columns = [str(c).strip().lower() for c in df_inv_up.columns]
                    df_inv_up.loc[df_inv_up['id'].astype(str) == str(id_inv), 'status'] = 'Fechado'
                    df_inv_up.loc[df_inv_up['id'].astype(str) == str(id_inv), 'total_itens'] = tot
                    df_inv_up.loc[df_inv_up['id'].astype(str) == str(id_inv), 'acuracidade_final'] = pct_acu
                    atualizar_aba_completa("inventarios", df_inv_up)
                limpar_cache_aplicacao()

            if pode_fechar:
                if st.button("🔒 Fechar Inventário (100% Concluído)", use_container_width=True, type="primary"):
                    fechar_e_preservar_historico(id_inventario_atual, id_pasta_limpo_base)
                    st.success("✅ Inventário encerrado e arquivado!")
                    st.rerun()
            else:
                qtd_f = len(itens_faltantes) if str(inventario_selected_obj['status']) == "1a Contagem" else len(itens_pendentes_2a)
                st.warning(f"⏳ Faltam **{qtd_f}** itens para concluir.")
                if eh_supervisor:
                    if st.button("🚨 Forçar Fechamento Incompleto (ADMIN)", use_container_width=True):
                        fechar_e_preservar_historico(id_inventario_atual, id_pasta_limpo_base)
                        st.success("✅ Inventário encerrado!")
                        st.rerun()

        # KPIs SIDEBAR
        total_itens_base = len(base_sistema_atual) if base_sistema_atual is not None else 0
        df_cnt_cnt = ler_aba("contagens")
        if not df_cnt_cnt.empty:
            df_cnt_cnt.columns = [str(c).strip().lower() for c in df_cnt_cnt.columns]
            df_cnt_pasta = df_cnt_cnt[df_cnt_cnt['inventario_id'].astype(str) == id_pasta_limpo_base]
        else: df_cnt_pasta = pd.DataFrame()
        
        total_contados_cnt = len(df_cnt_pasta) + pendentes_ram
        
        st.markdown(f'<div class="card-lateral"><div class="card-lateral-titulo">📋 ITENS NA BASE</div><div class="card-lateral-valor">{total_itens_base}</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="card-lateral"><div class="card-lateral-titulo">✅ LANÇAMENTOS</div><div class="card-lateral-valor">{total_contados_cnt}</div></div>', unsafe_allow_html=True)

        st.markdown("---")
        if st.button("🚀 Sincronizar Agora com o Drive", type="primary", use_container_width=True):
            if pendentes_ram > 0:
                qtd = sincronizar_ram_com_banco()
                st.success(f"✅ {qtd} bips enviados para a planilha!")
            else:
                st.info("ℹ️ Nenhum bip pendente na memória RAM.")
            st.rerun()

    # --- DECLARAÇÃO DAS ABAS PRINCIPAIS ---
    lista_abas = ["🔍 Contar Item (RAM Modo Rápido)", "📊 Lançamentos & Base", "📈 Desempenho & Acuracidade", "📁 Histórico Geral"]
    if eh_supervisor: lista_abas.append("⚙️ Gestão ADM")
    
    abas_objs = st.tabs(lista_abas)
    aba_contar, aba_lancamentos, aba_desempenho, aba_historico = abas_objs[0], abas_objs[1], abas_objs[2], abas_objs[3]
    aba_adm = abas_objs[4] if eh_supervisor else None

    # --- ABA 1: CONTAR ITEM (RAM MODO RÁPIDO) ---
    with aba_contar:
        if pendentes_ram > 0:
            c_top1, c_top2 = st.columns([3, 1])
            c_top1.info(f"⚡ **Modo Ultra Rápido Ativo:** Você tem **{pendentes_ram} bips** guardados na RAM local.")
            with c_top2:
                df_ram_temp = pd.DataFrame(st.session_state.buffer_ram_contagens)
                st.download_button("📥 Backup RAM (.xlsx)", converter_para_excel(df_ram_temp), file_name=f"backup_ram_pasta_{id_pasta_limpo_base}.xlsx")

        if not id_inventario_atual or (base_sistema_atual is None and str(inventario_selected_obj['status']) == 'Aberto'):
            st.warning("⚠️ Selecione um inventário ativo e carregue a base na barra lateral.")
        elif str(inventario_selected_obj['status']) == "Aberto": st.warning("⚠️ Inventário em configuração. Libere a 1ª Contagem na barra lateral.")
        elif str(inventario_selected_obj['status']) == "Fechado": st.error("🔒 Inventário Fechado. Selecione um inventário ativo.")
        elif pode_fechar and str(inventario_selected_obj['status']) in ["1a Contagem", "2a Contagem"]:
            st.success("🎉 **100% dos itens desta fase já foram contados!** Você pode fechar o inventário na barra lateral.")
        else:
            codigo_input = st.text_input("💻 Bipar ou Digitar Código do Produto", value="", placeholder="Bipe a etiqueta aqui...", key=f"bip_{st.session_state.contador_reset}")
            if codigo_input:
                codigo_rastreio = str(codigo_input).upper().strip().split(" - ")[-1]
                matches_codigo = base_sistema_atual[base_sistema_atual['cod_produto'].astype(str).str.upper().str.strip() == codigo_rastreio]
                if matches_codigo.empty: st.error("❌ Código não cadastrado na planilha base!")
                else:
                    df_cnts_exist = ler_aba("contagens")
                    if not df_cnts_exist.empty:
                        df_cnts_exist.columns = [str(c).strip().lower() for c in df_cnts_exist.columns]
                        df_cnts_exist_pasta = df_cnts_exist[df_cnts_exist['inventario_id'].astype(str) == id_pasta_limpo_base]
                    else: df_cnts_exist_pasta = pd.DataFrame()
                    
                    set_ja_contados = set()
                    if not df_cnts_exist_pasta.empty:
                        for _, r in df_cnts_exist_pasta[df_cnts_exist_pasta['cod_produto'].astype(str).str.upper().str.strip() == codigo_rastreio].iterrows():
                            set_ja_contados.add(f"{str(r.get('lote', '')).strip().upper()}_{str(r.get('ativo', '')).strip().upper()}")
                    
                    for r_ram in st.session_state.buffer_ram_contagens:
                        if str(r_ram['cod_produto']).upper().strip() == codigo_rastreio:
                            set_ja_contados.add(f"{str(r_ram['lote']).strip().upper()}_{str(r_ram['ativo']).strip().upper()}")

                    status_pasta_atual = str(inventario_selected_obj['status'])

                    if status_pasta_atual == "1a Contagem":
                        matches_pendentes = [row_m for _, row_m in matches_codigo.iterrows() if f"{str(row_m.get('lote', '')).strip().upper()}_{str(row_m.get('ativo', '')).strip().upper()}" not in set_ja_contados]
                        if not matches_pendentes:
                            st.success(f"🎉 **Todos os ativos/lotes do produto {codigo_rastreio} já foram contabilizados!**")
                            pode_exibir_form = False
                        else:
                            matches_para_usar = pd.DataFrame(matches_pendentes)
                            pode_exibir_form = True
                    else:
                        matches_para_usar = matches_codigo
                        pode_exibir_form = True

                    if pode_exibir_form:
                        item = matches_para_usar.iloc[0]
                        lote_auto = str(item.get('lote', '')).strip()
                        ativo_auto = str(item.get('ativo', '')).strip()
                        id_est_limpo = str(item.get('id_estoque_fisico', '')).strip() or extrair_id_estoque_do_nome(inventario_selected_obj['nome'])
                        desc_est_limpo = str(item.get('desc_estoque_fisico', '')).strip() or MAPA_ESTOQUES_DESC.get(id_est_limpo, 'ESTOQUE FÍSICO')

                        c1, c2, c3 = st.columns(3)
                        c1.metric("Cód. Produto", str(item['cod_produto']))
                        c2.metric("Lote", lote_auto or "—")
                        c3.metric("Ativo", ativo_auto or "—")
                        
                        st.markdown(f'<div class="caixa-descricao-produto">📦 Descrição: {item["desc_produto"]}</div>', unsafe_allow_html=True)
                        st.info(f"📍 **Estoque Físico:** {desc_est_limpo} (ID: {id_est_limpo}) | **Unidade:** {item.get('unid_medida', 'UN')} | **Status:** {status_pasta_atual}")
                        
                        with st.form("form_lancar_qtd", clear_on_submit=True):
                            qtd_fisica = st.number_input("📦 Quantidade Contada Fisicamente:", min_value=0, step=1, value=0)
                            confirma_zero = st.checkbox("⚠️ Marque se este item REALMENTE NÃO EXISTE no estoque (Saldo Zero)")
                            obs = st.text_input("Observação (opcional)")
                            
                            if st.form_submit_button("⚡ Salvar na RAM (Instantâneo)", type="primary", use_container_width=True):
                                if qtd_fisica == 0 and not confirma_zero: st.error("⚠️ Para salvar quantidade 0, marque a confirmação amarela!")
                                else:
                                    qtd_sys = int(pd.to_numeric(item.get('qtd_estoque', 0), errors='coerce') or 0)
                                    dif = qtd_fisica - qtd_sys
                                    fase_gravar = "2a Contagem Concluida" if status_pasta_atual == "2a Contagem" else status_pasta_atual
                                    data_hora_agora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                                    st.session_state.buffer_ram_contagens.append({
                                        'inventario_id': id_pasta_limpo_base,
                                        'id_estoque': id_est_limpo,
                                        'desc_estoque': desc_est_limpo,
                                        'cod_produto': str(item['cod_produto']),
                                        'desc_produto': str(item['desc_produto']),
                                        'unid_medida': str(item.get('unid_medida', 'UN')),
                                        'qtd_sistema': qtd_sys,
                                        'qtd_contada': qtd_fisica,
                                        'diferenca': dif,
                                        'ativo': ativo_auto,
                                        'observacao': obs,
                                        'operador': st.session_state.operador,
                                        'data_hora': data_hora_agora,
                                        'lote': lote_auto,
                                        'fase_contagem': fase_gravar
                                    })

                                    if len(st.session_state.buffer_ram_contagens) >= 10:
                                        sincronizar_ram_com_banco()

                                    st.session_state.contador_reset += 1
                                    st.toast("⚡ Salvo na RAM instantaneamente!", icon="✅")
                                    st.rerun()

    # --- ABA 2: LANÇAMENTOS E ESPELHO BASE ---
    with aba_lancamentos:
        sub_aba1, sub_aba2 = st.tabs(["📋 Meus Lançamentos Nesta Pasta", "📄 Espelho Base do Saldo (Status Visual)"])
        with sub_aba1:
            df_cnts_todas = ler_aba("contagens")
            if not df_cnts_todas.empty:
                df_cnts_todas.columns = [str(c).strip().lower() for c in df_cnts_todas.columns]
                df_minhas = df_cnts_todas[df_cnts_todas['inventario_id'].astype(str) == id_pasta_limpo_base]
            else: df_minhas = pd.DataFrame()
            
            if pendentes_ram > 0:
                df_ram = pd.DataFrame(st.session_state.buffer_ram_contagens)
                df_ram_pasta = df_ram[df_ram['inventario_id'].astype(str) == id_pasta_limpo_base] if not df_ram.empty else pd.DataFrame()
                if not df_ram_pasta.empty:
                    df_minhas = pd.concat([df_ram_pasta, df_minhas], ignore_index=True)

            if not df_minhas.empty:
                m1, m2, m3 = st.columns(3)
                m1.metric("Lançamentos Totais", len(df_minhas))
                m2.metric("Com Divergência", len(df_minhas[pd.to_numeric(df_minhas['diferenca'], errors='coerce') != 0]))
                m3.metric("Total de Peças Contadas", int(pd.to_numeric(df_minhas['qtd_contada'], errors='coerce').sum()))
                st.download_button("📥 Exportar Lançamentos Filtrados para Excel", converter_para_excel(df_minhas), file_name=f"contagem_pasta_{id_pasta_limpo_base}.xlsx")
                st.dataframe(df_minhas, use_container_width=True, hide_index=True)
            else: st.info("Nenhum lançamento registrado nesta pasta.")
            
        with sub_aba2:
            if base_sistema_atual is not None:
                st.dataframe(base_sistema_atual, use_container_width=True, hide_index=True)
            else: st.info("Nenhuma base carregada.")

    # --- ABA 3: DESEMPENHO E ACURACIDADE ---
    with aba_desempenho:
        sub_d1, sub_d2 = st.tabs(["🔴 Desempenho & Prazos por Estoque", "📈 Acuracidade Auditada pelo Supervisor"])
        with sub_d1:
            st.subheader("🏆 Status de Atualização dos Estoques Físicos")
            df_ultimas_dt = ler_aba("ultima_contagem_estoques")
            mapa_datas = {}
            if not df_ultimas_dt.empty:
                df_ultimas_dt.columns = [str(c).strip().lower() for c in df_ultimas_dt.columns]
                for _, r in df_ultimas_dt.iterrows():
                    mapa_datas[str(r['id_estoque']).strip()] = str(r['ultima_data'])

            linhas_desempenho, hoje = [], datetime.datetime.now()
            b_count, a_count, c_count = 0, 0, 0
            
            for est in LISTA_ESTOQUES_FIXA:
                est_id = str(est["id"]).strip()
                u_data = mapa_datas.get(est_id, None)
                if u_data:
                    try:
                        str_dt = str(u_data).split(".")[0]
                        if len(str_dt) == 10: str_dt += " 12:00:00"
                        dt = datetime.datetime.strptime(str_dt, "%Y-%m-%d %H:%M:%S")
                        dias, dt_fmt = (hoje - dt).days, dt.strftime("%d/%m/%Y %H:%M")
                    except Exception: dias, dt_fmt = 999, "Sem histórico"
                else: dias, dt_fmt = 999, "Nunca Contado"
                
                if dias <= 7: status, b_count = "🟢 Em Dia", b_count + 1
                elif dias <= 14: status, a_count = "🟡 Necessário Auditar", a_count + 1
                else: status, c_count = "🔴 Crítico (+2 semanas)", c_count + 1
                    
                linhas_desempenho.append({"Id. Estoque": est_id, "Descrição do Estoque Físico": est["desc"], "Última Contagem": dt_fmt, "Dias Sem Contar": dias if dias != 999 else 999, "Criticidade": status})

            df_desempenho_full = pd.DataFrame(linhas_desempenho)
            k1, k2, k3 = st.columns(3)
            k1.metric("🟢 Em Dia (até 7 dias)", b_count)
            k2.metric("🟡 Necessário Auditar (8 a 14 dias)", a_count)
            k3.metric("🔴 Crítico (+14 dias)", c_count)
            st.dataframe(df_desempenho_full, use_container_width=True, hide_index=True)

        with sub_d2:
            st.subheader("📈 Tabela de Acuracidade Geral por Depósito")
            df_auds = ler_aba("auditorias_supervisor")
            if df_auds.empty: st.info("💡 Nenhuma amostragem coletada pelo supervisor.")
            else: st.dataframe(df_auds, use_container_width=True, hide_index=True)

    # --- ABA 4: HISTÓRICO GERAL (PAGINAÇÃO DE 10 EM 10 E ORDENAÇÃO DECRESCENTE) ---
    with aba_historico:
        st.title("📁 Arquivo Geral de Movimentações (Google Drive)")
        if df_inventarios.empty or 'id' not in df_inventarios.columns: 
            st.info("Nenhum inventário registrado.")
        else:
            df_inv_ordenados = df_inventarios.copy()
            df_inv_ordenados['id_num'] = pd.to_numeric(df_inv_ordenados['id'].astype(str).str.replace('#', '', regex=False), errors='coerce').fillna(0)
            df_inv_ordenados = df_inv_ordenados.sort_values(by=['data', 'id_num'], ascending=[False, False])

            itens_por_pagina = 10
            total_itens = len(df_inv_ordenados)
            total_paginas = max(1, (total_itens + itens_por_pagina - 1) // itens_por_pagina)

            if st.session_state.pagina_historico > total_paginas:
                st.session_state.pagina_historico = total_paginas

            inicio_idx = (st.session_state.pagina_historico - 1) * itens_por_pagina
            fatia_pastas = df_inv_ordenados.iloc[inicio_idx:inicio_idx + itens_por_pagina]

            for idx, inv in fatia_pastas.iterrows():
                id_proc = str(inv['id']).replace('#','').strip()
                df_cnts_todas = ler_aba("contagens")
                if not df_cnts_todas.empty:
                    df_cnts_todas.columns = [str(c).strip().lower() for c in df_cnts_todas.columns]
                    df_h = df_cnts_todas[df_cnts_todas['inventario_id'].astype(str) == id_proc]
                else: df_h = pd.DataFrame()
                
                tot_reg = len(df_h) if not df_h.empty else inv.get('total_itens', 0)
                acu_reg = inv.get('acuracidade_final', '—')

                c_exp, c_del = st.columns([8, 2])
                with c_exp:
                    with st.expander(f"📁 Pasta {inv['id']} - {inv['nome']} | Data: {inv['data']} | Status: {inv['status']} | Acuracidade: {acu_reg} ({tot_reg} itens contados)"):
                        if not df_h.empty:
                            st.download_button("📥 Baixar Planilha (.xlsx)", converter_para_excel(df_h), file_name=f"inventario_{id_proc}.xlsx", key=f"dl_hist_{id_proc}")
                            st.dataframe(df_h, use_container_width=True, hide_index=True)
                        else: st.info("Nenhum lançamento registrado nesta pasta.")
                with c_del:
                    if eh_supervisor and st.button("🗑️ Excluir Pasta", key=f"del_hist_inv_{inv['id']}", use_container_width=True):
                        df_inv_att = ler_aba("inventarios")
                        if not df_inv_att.empty:
                            df_inv_att.columns = [str(c).strip().lower() for c in df_inv_att.columns]
                            df_inv_att = df_inv_att[df_inv_att['id'].astype(str) != str(inv['id'])]
                            atualizar_aba_completa("inventarios", df_inv_att)
                        limpar_cache_aplicacao()
                        st.rerun()

            st.markdown("---")
            c_pag1, c_pag2, c_pag3 = st.columns([2, 3, 2])
            with c_pag1:
                if st.button("◀ Anterior", disabled=(st.session_state.pagina_historico <= 1), use_container_width=True):
                    st.session_state.pagina_historico -= 1
                    st.rerun()
            with c_pag2:
                st.markdown(f"<h5 style='text-align: center;'>Página {st.session_state.pagina_historico} de {total_paginas} ({total_itens} pastas)</h5>", unsafe_allow_html=True)
            with c_pag3:
                if st.button("Próxima ▶", disabled=(st.session_state.pagina_historico >= total_paginas), use_container_width=True):
                    st.session_state.pagina_historico += 1
                    st.rerun()

    # --- ABA 5: GESTÃO ADM ---
    if eh_supervisor and aba_adm is not None:
        with aba_adm:
            st.title("⚙️ Módulo de Gestão do Administrador (Google Drive)")
            opcao_adm = st.selectbox("Escolha o Módulo de Ação:", ["🚨 Liberar / Encerrar Divergências", "🔬 Auditoria Amostral (Supervisor)", "📊 Relatório Consolidado (Excel Gerencial)", "👥 Gestão de Usuários & Senhas"])
            st.markdown("---")

            if opcao_adm == "🚨 Liberar / Encerrar Divergências":
                st.subheader("Tratamento de Erros de Contagem da Equipe")
                df_cnts = ler_aba("contagens")
                if not df_cnts.empty:
                    df_cnts.columns = [str(c).strip().lower() for c in df_cnts.columns]
                    df_cnts['diferenca_num'] = pd.to_numeric(df_cnts['diferenca'], errors='coerce').fillna(0)
                    df_divs = df_cnts[(df_cnts['diferenca_num'] != 0) & (~df_cnts['fase_contagem'].astype(str).isin(['2a Contagem', 'Encerrado com Divergencia']))]
                    
                    if df_divs.empty:
                        st.success("🎉 Nenhuma divergência pendente no momento!")
                    else:
                        opcoes_div = [f"Item ID #{r.name} - {r['cod_produto']} - {r['desc_produto']} (Dif: {r['diferenca']})" for _, r in df_divs.iterrows()]
                        sel_div = st.selectbox("Selecione o Item com Divergência:", opcoes_div)
                        idx_sel = int(sel_div.split("Item ID #")[1].split(" - ")[0])
                        row_target = df_divs.loc[idx_sel]
                        
                        st.info(f"📊 Sistema: {row_target['qtd_sistema']} | Contado: {row_target['qtd_contada']} | Dif: {row_target['diferenca']}")
                        justificativa_adm = st.text_input("Justificativa:", value="Divergente")
                        
                        col_act1, col_act2 = st.columns(2)
                        with col_act1:
                            if st.button("🚨 Abrir 2ª Contagem", type="primary", use_container_width=True):
                                df_cnts.loc[idx_sel, 'fase_contagem'] = '2a Contagem'
                                df_cnts.loc[idx_sel, 'observacao'] = f"ADM: [2ª CONTAGEM] - {justificativa_adm}"
                                atualizar_aba_completa("contagens", df_cnts.drop(columns=['diferenca_num']))
                                st.success("✅ Enviado para 2ª Contagem!")
                                st.rerun()
                        with col_act2:
                            if st.button("🔒 Finalizar e Manter Divergência", use_container_width=True):
                                df_cnts.loc[idx_sel, 'fase_contagem'] = 'Encerrado com Divergencia'
                                df_cnts.loc[idx_sel, 'observacao'] = f"ADM: [ENCERRADO] - {justificativa_adm}"
                                atualizar_aba_completa("contagens", df_cnts.drop(columns=['diferenca_num']))
                                st.success("✅ Divergência encerrada!")
                                st.rerun()
                else:
                    st.info("Nenhuma contagem no sistema.")
