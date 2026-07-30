import streamlit as st
import pandas as pd
import datetime
import io
import re
import time
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="TESTE LOCAL - Auditoria & Inventário JBA", layout="wide")

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

# --- BANCO DE DADOS VIRTUAL APENAS NA MEMÓRIA DO STREAMLIT ---
if 'db_usuarios' not in st.session_state:
    st.session_state.db_usuarios = [
        {"id": 1, "nome": "Administrador Tel", "cpf": "00000000000", "email": "admin@tel.com.br", "senha": "123", "perfil": "Administrador"}
    ]
if 'db_inventarios' not in st.session_state:
    st.session_state.db_inventarios = []
if 'db_itens_base' not in st.session_state:
    st.session_state.db_itens_base = []
if 'db_contagens' not in st.session_state:
    st.session_state.db_contagens = []
if 'db_inventarios_sup' not in st.session_state:
    st.session_state.db_inventarios_sup = []
if 'db_auditorias_sup' not in st.session_state:
    st.session_state.db_auditorias_sup = []
if 'db_ultima_contagem_estoques' not in st.session_state:
    st.session_state.db_ultima_contagem_estoques = {}

# ESTADOS DA SESSÃO
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'operador' not in st.session_state: st.session_state.operador = ""
if 'perfil_usuario' not in st.session_state: st.session_state.perfil_usuario = "Almoxarife"
if 'tela_acesso' not in st.session_state: st.session_state.tela_acesso = "login"
if 'contador_reset' not in st.session_state: st.session_state.contador_reset = 0

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

# --- LOGIN ---
if not st.session_state.logged_in:
    col_vaz1, col_central, col_vaz2 = st.columns([1, 1.2, 1])
    with col_central:
        st.title("🧪 TESTE LOCAL - Acesso ao Sistema JBA")
        st.info("ℹ️ Este app está rodando 100% em memória RAM (sem nuvem).")
        with st.form("login_form"):
            identificador = st.text_input("CPF ou E-mail", value="admin@tel.com.br")
            senha = st.text_input("Senha", type="password", value="123")
            if st.form_submit_button("Entrar no Sistema", type="primary", use_container_width=True):
                id_limpo = identificador.strip()
                doc_limpo = limpar_documento(id_limpo)
                user = next((u for u in st.session_state.db_usuarios if (u['email'] == id_limpo or u['cpf'] == doc_limpo) and u['senha'] == senha), None)
                if user:
                    st.session_state.logged_in = True
                    st.session_state.operador = user['nome']
                    st.session_state.perfil_usuario = user['perfil']
                    st.rerun()
                else:
                    st.error("❌ Credenciais incorretas.")

# --- APLICAÇÃO PRINCIPAL ---
else:
    eh_supervisor = (st.session_state.perfil_usuario == "Administrador") or ("admin" in st.session_state.operador.lower())

    # --- BARRA LATERAL ---
    with st.sidebar:
        st.write(f"👤 **{st.session_state.operador}** ({st.session_state.perfil_usuario})")
        st.success("⚡ MODO LOCAL ATIVO (Zero Chamadas de Nuvem)")

        col_s1, col_s2 = st.columns(2)
        with col_s1:
            if st.button("🔄 Recarregar", use_container_width=True):
                st.rerun()
        with col_s2:
            if st.button("🚪 Sair", use_container_width=True):
                st.session_state.logged_in = False
                st.session_state.operador = ""
                st.rerun()
            
        st.markdown("---")
        st.write("📁 **Seleção de Inventário**")
        df_inventarios = pd.DataFrame(st.session_state.db_inventarios) if st.session_state.db_inventarios else pd.DataFrame()

        if df_inventarios.empty:
            id_inventario_atual = None
            inventario_selected_obj = None
            id_pasta_limpo_base = ""
            st.info("Crie um inventário abaixo.")
        else:
            lista_inv = [f"{row['id']} – {row['nome']} ({row['status']})" for idx, row in df_inventarios.iterrows()]
            inventario_selected = st.selectbox("Selecione a Pasta", lista_inv, index=0, key="sb_pasta_ativa")
            id_inventario_atual = inventario_selected.split(" – ")[0]
            inventario_selected_obj = df_inventarios[df_inventarios['id'] == id_inventario_atual].iloc[0]
            id_pasta_limpo_base = id_inventario_atual.replace("#", "").strip()

        # UPLOAD DA BASE
        st.write("📂 **Upload Base de Dados (.xlsx)**")
        if inventario_selected_obj is not None and inventario_selected_obj['status'] == "Aberto":
            uploader_key = f"func_excel_loader_{id_pasta_limpo_base}_{st.session_state.contador_reset}"
            arquivo_excel = st.file_uploader("Suba o arquivo Excel (.xlsx)", type=["xlsx"], label_visibility="collapsed", key=uploader_key)
            
            if arquivo_excel is not None and id_pasta_limpo_base:
                t_inicio = time.time()
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
                    # Limpa a base da pasta na RAM
                    st.session_state.db_itens_base = [item for item in st.session_state.db_itens_base if str(item['inventario_id']) != id_pasta_limpo_base]
                    
                    novos_itens = []
                    for idx_r, r in df_upload_temp.iterrows():
                        novos_itens.append({
                            'id': len(st.session_state.db_itens_base) + idx_r + 1,
                            'inventario_id': id_pasta_limpo_base,
                            'cod_produto': str(r[col_cod]).strip() if col_cod and pd.notna(r[col_cod]) else '',
                            'desc_produto': str(r[col_desc]).strip() if col_desc and pd.notna(r[col_desc]) else '',
                            'desc_estoque_fisico': str(r[col_est]).strip() if col_est and pd.notna(r[col_est]) and str(r[col_est]).strip() != '' else est_desc_fallback,
                            'unid_medida': str(r[col_uni]).strip() if col_uni and pd.notna(r[col_uni]) else '',
                            'qtd_estoque': int(pd.to_numeric(r[col_qtd], errors='coerce') or 0) if col_qtd else 0,
                            'id_estoque_fisico': str(r[col_idest]).strip() if col_idest and pd.notna(r[col_idest]) and str(r[col_idest]).strip() != '' else est_id_fallback,
                            'lote': str(r[col_lote]).strip() if col_lote and pd.notna(r[col_lote]) and str(r[col_lote]).lower() != 'nan' else '',
                            'ativo': str(r[col_ativo]).strip() if col_ativo and pd.notna(r[col_ativo]) and str(r[col_ativo]).lower() != 'nan' else ''
                        })
                    st.session_state.db_itens_base.extend(novos_itens)
                    t_fim = time.time()
                    st.success(f"✅ Base carregada em RAM em {t_fim - t_inicio:.2f} segundos! ({len(novos_itens)} itens)")

        # FILTRA BASE NA RAM
        df_base_ram = pd.DataFrame([i for i in st.session_state.db_itens_base if str(i['inventario_id']) == id_pasta_limpo_base]) if st.session_state.db_itens_base else pd.DataFrame()

        if inventario_selected_obj is not None and inventario_selected_obj['status'] == "Aberto" and not df_base_ram.empty:
            st.markdown("---")
            if st.button("🚀 Liberar 1ª Contagem", type="primary", use_container_width=True):
                for inv in st.session_state.db_inventarios:
                    if str(inv['id']).replace('#', '') == id_pasta_limpo_base:
                        inv['status'] = '1a Contagem'
                st.success("🔒 1ª Contagem liberada.")
                st.rerun()

        with st.expander("➕ Criar Novo Inventário"):
            with st.form("form_novo", clear_on_submit=True):
                novo_nome = st.text_input("Nome do Inventário")
                if st.form_submit_button("Criar Pasta", type="primary") and novo_nome:
                    maior_id = max([int(str(inv['id']).replace('#', '')) for inv in st.session_state.db_inventarios], default=0)
                    st.session_state.db_inventarios.append({
                        'id': f"#{maior_id + 1}",
                        'nome': novo_nome,
                        'data': datetime.date.today().strftime("%Y-%m-%d"),
                        'status': 'Aberto',
                        'total_itens': 0,
                        'acuracidade_final': '0%'
                    })
                    st.rerun()

        # KPIS LATERAIS
        total_itens_base = len(df_base_ram) if not df_base_ram.empty else 0
        contados_pasta = [c for c in st.session_state.db_contagens if str(c['inventario_id']) == id_pasta_limpo_base]
        
        st.markdown(f'<div class="card-lateral"><div class="card-lateral-titulo">📋 ITENS NA BASE</div><div class="card-lateral-valor">{total_itens_base}</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="card-lateral"><div class="card-lateral-titulo">✅ LANÇAMENTOS</div><div class="card-lateral-valor">{len(contados_pasta)}</div></div>', unsafe_allow_html=True)

    # --- ABA PRINCIPAL DE CONTAGEM ---
    tab_contar, tab_lancamentos, tab_historico = st.tabs(["🔍 Contar Item (100% RAM Local)", "📊 Lançamentos Efetivados", "📁 Histórico Geral"])

    with tab_contar:
        if not id_inventario_atual or df_base_ram.empty:
            st.warning("⚠️ Selecione um inventário ativo e suba a planilha base na barra lateral.")
        elif inventario_selected_obj['status'] == "Aberto":
            st.warning("⚠️ Inventário em configuração. Libere a 1ª Contagem na barra lateral.")
        else:
            codigo_input = st.text_input("💻 Bipar ou Digitar Código do Produto", value="", placeholder="Bipe a etiqueta aqui...", key=f"bip_{st.session_state.contador_reset}")
            
            if codigo_input:
                t_bip_inicio = time.time()
                codigo_rastreio = str(codigo_input).upper().strip().split(" - ")[-1]
                matches = df_base_ram[df_base_ram['cod_produto'].astype(str).str.upper().str.strip() == codigo_rastreio]
                
                if matches.empty:
                    st.error("❌ Código não cadastrado na planilha base!")
                else:
                    item = matches.iloc[0]
                    lote_auto = str(item['lote']).strip() if pd.notna(item['lote']) and str(item['lote']).lower() != 'nan' else ''
                    ativo_auto = str(item['ativo']).strip() if pd.notna(item['ativo']) and str(item['ativo']).lower() != 'nan' else ''
                    id_est_limpo = str(item['id_estoque_fisico']).strip() if pd.notna(item['id_estoque_fisico']) else extrair_id_estoque_do_nome(inventario_selected_obj['nome'])
                    desc_est_limpo = str(item['desc_estoque_fisico']).strip() if pd.notna(item['desc_estoque_fisico']) else MAPA_ESTOQUES_DESC.get(id_est_limpo, 'ESTOQUE FÍSICO')

                    c1, c2, c3 = st.columns(3)
                    c1.metric("Cód. Produto", str(item['cod_produto']))
                    c2.metric("Lote", lote_auto or "—")
                    c3.metric("Ativo", ativo_auto or "—")

                    st.markdown(f'<div class="caixa-descricao-produto">📦 Descrição: {item["desc_produto"]}</div>', unsafe_allow_html=True)
                    st.info(f"📍 **Estoque Físico:** {desc_est_limpo} | **Unidade:** {item['unid_medida']}")

                    with st.form("form_lancar_qtd_local", clear_on_submit=True):
                        qtd_fisica = st.number_input("📦 Quantidade Contada Fisicamente:", min_value=0, step=1, value=0)
                        confirma_zero = st.checkbox("⚠️ Confirma Saldo Zero para este item")
                        obs = st.text_input("Observação (opcional)")
                        
                        if st.form_submit_button("⚡ Salvar Instantâneo (RAM)", type="primary", use_container_width=True):
                            if qtd_fisica == 0 and not confirma_zero:
                                st.error("⚠️ Para salvar quantidade 0, marque a confirmação amarela!")
                            else:
                                t_salvar_ini = time.time()
                                qtd_sys = int(item['qtd_estoque'])
                                dif = qtd_fisica - qtd_sys
                                data_hora_agora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                                # Gravação instantânea na lista em memória
                                st.session_state.db_contagens.append({
                                    'id': len(st.session_state.db_contagens) + 1,
                                    'inventario_id': id_pasta_limpo_base,
                                    'id_estoque': id_est_limpo,
                                    'desc_estoque': desc_est_limpo,
                                    'cod_produto': str(item['cod_produto']),
                                    'desc_produto': str(item['desc_produto']),
                                    'unid_medida': str(item['unid_medida']),
                                    'qtd_sistema': qtd_sys,
                                    'qtd_contada': qtd_fisica,
                                    'diferenca': dif,
                                    'ativo': ativo_auto,
                                    'observacao': obs,
                                    'operador': st.session_state.operador,
                                    'data_hora': data_hora_agora,
                                    'lote': lote_auto,
                                    'fase_contagem': inventario_selected_obj['status']
                                })
                                
                                t_salvar_fim = time.time()
                                st.session_state.contador_reset += 1
                                st.toast(f"⚡ Salvo em RAM local em {(t_salvar_fim - t_salvar_ini)*1000:.1f} milissegundos!", icon="✅")
                                st.rerun()

    # --- ABA 2: LANÇAMENTOS ---
    with tab_lancamentos:
        df_lanc_local = pd.DataFrame([c for c in st.session_state.db_contagens if str(c['inventario_id']) == id_pasta_limpo_base])
        if not df_lanc_local.empty:
            st.download_button("📥 Baixar Lançamentos RAM (.xlsx)", converter_para_excel(df_lanc_local), file_name=f"lancamentos_ram_pasta_{id_pasta_limpo_base}.xlsx")
            st.dataframe(df_lanc_local, use_container_width=True, hide_index=True)
        else:
            st.info("Nenhum lançamento gravado em RAM nesta pasta.")

    # --- ABA 3: HISTÓRICO ---
    with tab_historico:
        st.dataframe(pd.DataFrame(st.session_state.db_inventarios), use_container_width=True, hide_index=True)
