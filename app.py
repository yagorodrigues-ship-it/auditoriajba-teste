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

# --- BANCO DE DADOS EM MEMÓRIA RAM DO STREAMLIT (SESSÃO LOCAL) ---
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

# ESTADOS DE CONTROLE DE INTERFACE
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'operador' not in st.session_state: st.session_state.operador = ""
if 'perfil_usuario' not in st.session_state: st.session_state.perfil_usuario = "Almoxarife"
if 'tela_acesso' not in st.session_state: st.session_state.tela_acesso = "login"
if 'contador_reset' not in st.session_state: st.session_state.contador_reset = 0
if 'contador_reset_sup' not in st.session_state: st.session_state.contador_reset_sup = 0
if 'bases_supervisor_por_inv' not in st.session_state: st.session_state.bases_supervisor_por_inv = {}

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

# ESTILOS CSS
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

# --- TELA DE LOGIN CENTRALIZADA ---
if not st.session_state.logged_in:
    col_vaz1, col_central, col_vaz2 = st.columns([1, 1.2, 1])
    with col_central:
        if st.session_state.tela_acesso == "login":
            st.title("🧪 TESTE LOCAL - Acesso ao Sistema JBA")
            st.info("ℹ️ Sistema rodando 100% em RAM local (Zero Nuvem).")
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
                    else: st.error("❌ Credenciais incorretas.")
            if st.button("📝 Criar nova conta de colaborador", use_container_width=True):
                st.session_state.tela_acesso = "cadastro"
                st.rerun()

        elif st.session_state.tela_acesso == "cadastro":
            st.title("📝 Cadastro de Colaborador (Local)")
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
                        st.session_state.db_usuarios.append({
                            "id": len(st.session_state.db_usuarios) + 1,
                            "nome": novo_nome.strip(),
                            "cpf": cpf_l,
                            "email": novo_email.strip(),
                            "senha": nova_senha,
                            "perfil": "Almoxarife"
                        })
                        st.success("✅ Cadastro realizado!")
                        st.session_state.tela_acesso = "login"
                        st.rerun()
            if st.button("◀ Voltar para o Login"):
                st.session_state.tela_acesso = "login"
                st.rerun()

# --- APLICAÇÃO PRINCIPAL ---
else:
    df_inventarios = pd.DataFrame(st.session_state.db_inventarios) if st.session_state.db_inventarios else pd.DataFrame()
    df_inventarios_sup = pd.DataFrame(st.session_state.db_inventarios_sup) if st.session_state.db_inventarios_sup else pd.DataFrame()
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
                t_ini = time.time()
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
                    # Remove base antiga da pasta na RAM
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
                    st.success(f"✅ Base carregada em RAM em {t_fim - t_ini:.2f}s ({len(novos_itens)} itens)")

        elif inventario_selected_obj is not None and inventario_selected_obj['status'] in ["1a Contagem", "2a Contagem"]:
            st.info(f"🔒 **Base Congelada ({inventario_selected_obj['status']})**.")
        else: st.info("🔒 Crie um inventário 'Aberto'.")

        # BUSCA BASE DA RAM
        df_base_ram = pd.DataFrame([i for i in st.session_state.db_itens_base if str(i['inventario_id']) == id_pasta_limpo_base]) if st.session_state.db_itens_base else pd.DataFrame()

        if inventario_selected_obj is not None and inventario_selected_obj['status'] == "Aberto" and not df_base_ram.empty:
            st.markdown("---")
            if st.button("🚀 Salvar Base e Iniciar 1ª Contagem", type="primary", use_container_width=True):
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

        # FECHAMENTO DO INVENTÁRIO
        pode_fechar, itens_faltantes, itens_pendentes_2a = False, [], []
        if inventario_selected_obj is not None and inventario_selected_obj['status'] in ["1a Contagem", "2a Contagem"] and not df_base_ram.empty:
            contagens_pasta = [c for c in st.session_state.db_contagens if str(c['inventario_id']) == id_pasta_limpo_base]
            set_contados_triade = {f"{str(c['cod_produto']).upper().strip()}_{str(c['lote']).upper().strip() if c['lote'] else ''}_{str(c['ativo']).upper().strip() if c['ativo'] else ''}" for c in contagens_pasta}

            if inventario_selected_obj['status'] == "1a Contagem":
                for _, r_b in df_base_ram.iterrows():
                    c_b, l_b, a_b = str(r_b['cod_produto']).upper().strip(), str(r_b['lote']).upper().strip() if pd.notna(r_b['lote']) and str(r_b['lote']).lower() != 'nan' else "", str(r_b['ativo']).upper().strip() if pd.notna(r_b['ativo']) and str(r_b['ativo']).lower() != 'nan' else ""
                    if f"{c_b}_{l_b}_{a_b}" not in set_contados_triade: itens_faltantes.append(c_b)
                if len(itens_faltantes) == 0: pode_fechar = True
            elif inventario_selected_obj['status'] == "2a Contagem":
                itens_pendentes_2a = [str(c['cod_produto']).upper().strip() for c in contagens_pasta if c['fase_contagem'] == '2a Contagem']
                if len(itens_pendentes_2a) == 0: pode_fechar = True

            st.markdown("---")
            def fechar_e_preservar_historico_local(id_inv, id_limpo):
                contagens_inv = [c for c in st.session_state.db_contagens if str(c['inventario_id']) == id_limpo]
                tot = len(contagens_inv)
                acertos = len([c for c in contagens_inv if c['diferenca'] == 0]) if tot > 0 else 0
                pct_acu = f"{(acertos / tot)*100:.1f}%" if tot > 0 else "0%"
                for inv in st.session_state.db_inventarios:
                    if str(inv['id']).replace('#', '') == id_limpo:
                        inv['status'] = 'Fechado'
                        inv['total_itens'] = tot
                        inv['acuracidade_final'] = pct_acu

            if pode_fechar:
                if st.button("🔒 Fechar Inventário (100% Concluído)", use_container_width=True, type="primary"):
                    fechar_e_preservar_historico_local(id_inventario_atual, id_pasta_limpo_base)
                    st.success("✅ Inventário encerrado e arquivado!")
                    st.rerun()
            else:
                qtd_f = len(itens_faltantes) if inventario_selected_obj['status'] == "1a Contagem" else len(itens_pendentes_2a)
                st.warning(f"⏳ Faltam **{qtd_f}** itens para concluir.")
                if eh_supervisor:
                    if st.button("🚨 Forçar Fechamento Incompleto (ADMIN)", use_container_width=True):
                        fechar_e_preservar_historico_local(id_inventario_atual, id_pasta_limpo_base)
                        st.success("✅ Inventário encerrado!")
                        st.rerun()

        # KPIs SIDEBAR
        total_itens_base = len(df_base_ram) if not df_base_ram.empty else 0
        total_contados_cnt = len([c for c in st.session_state.db_contagens if str(c['inventario_id']) == id_pasta_limpo_base])
        
        st.markdown(f'<div class="card-lateral"><div class="card-lateral-titulo">📋 ITENS NA BASE</div><div class="card-lateral-valor">{total_itens_base}</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="card-lateral"><div class="card-lateral-titulo">✅ LANÇAMENTOS</div><div class="card-lateral-valor">{total_contados_cnt}</div></div>', unsafe_allow_html=True)

    # --- DECLARAÇÃO DAS ABAS PRINCIPAIS ---
    lista_abas = ["🔍 Contar Item (RAM Modo Rápido)", "📊 Lançamentos & Base", "📈 Desempenho & Acuracidade", "📁 Histórico Geral"]
    if eh_supervisor: lista_abas.append("⚙️ Gestão ADM")
    
    abas_objs = st.tabs(lista_abas)
    aba_contar, aba_lancamentos, aba_desempenho, aba_historico = abas_objs[0], abas_objs[1], abas_objs[2], abas_objs[3]
    aba_adm = abas_objs[4] if eh_supervisor else None

    # --- ABA 1: CONTAR ITEM ---
    with aba_contar:
        if not id_inventario_atual or (df_base_ram.empty and inventario_selected_obj['status'] == 'Aberto'):
            st.warning("⚠️ Selecione um inventário ativo e carregue a base na barra lateral.")
        elif inventario_selected_obj['status'] == "Aberto": st.warning("⚠️ Inventário em configuração. Libere a 1ª Contagem na barra lateral.")
        elif inventario_selected_obj['status'] == "Fechado": st.error("🔒 Inventário Fechado. Selecione um inventário ativo.")
        elif pode_fechar and inventario_selected_obj['status'] in ["1a Contagem", "2a Contagem"]:
            st.success("🎉 **100% dos itens desta fase já foram contados!** Você pode fechar o inventário na barra lateral.")
        else:
            codigo_input = st.text_input("💻 Bipar ou Digitar Código do Produto", value="", placeholder="Bipe a etiqueta aqui...", key=f"bip_{st.session_state.contador_reset}")
            if codigo_input:
                codigo_rastreio = str(codigo_input).upper().strip().split(" - ")[-1]
                matches_codigo = df_base_ram[df_base_ram['cod_produto'].astype(str).str.upper().str.strip() == codigo_rastreio]
                
                if matches_codigo.empty: st.error("❌ Código não cadastrado na planilha base!")
                else:
                    contagens_pasta = [c for c in st.session_state.db_contagens if str(c['inventario_id']) == id_pasta_limpo_base and str(c['cod_produto']).strip().upper() == codigo_rastreio]
                    set_ja_contados = {f"{str(c['lote']).strip().upper() if c['lote'] else ''}_{str(c['ativo']).strip().upper() if c['ativo'] else ''}" for c in contagens_pasta}
                    status_pasta_atual = inventario_selected_obj['status']

                    if status_pasta_atual == "1a Contagem":
                        matches_pendentes = [row_m for _, row_m in matches_codigo.iterrows() if f"{str(row_m['lote']).strip().upper() if pd.notna(row_m['lote']) and str(row_m['lote']).lower() != 'nan' else ''}_{str(row_m['ativo']).strip().upper() if pd.notna(row_m['ativo']) and str(row_m['ativo']).lower() != 'nan' else ''}" not in set_ja_contados]
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
                        if len(matches_para_usar) > 1:
                            opcoes_item = [f"Ativo: {str(r['ativo']).strip() if pd.notna(r['ativo']) and str(r['ativo']).lower()!='nan' and str(r['ativo']).strip()!='' else 'Sem Ativo'} | Lote: {str(r['lote']).strip() if pd.notna(r['lote']) and str(r['lote']).lower()!='nan' and str(r['lote']).strip()!='' else 'Sem Lote'} (ID Linha #{r['id']})" for _, r in matches_para_usar.iterrows()]
                            item_sel_opcao = st.selectbox("Escolha o item específico PENDENTE:", opcoes_item)
                            item = matches_para_usar[matches_para_usar['id'] == int(item_sel_opcao.split("(ID Linha #")[1].replace(")", "").strip())].iloc[0]
                        else: item = matches_para_usar.iloc[0]

                        lote_auto = str(item['lote']).strip() if pd.notna(item['lote']) and str(item['lote']).lower() != 'nan' and str(item['lote']).strip() != '' else ''
                        ativo_auto = str(item['ativo']).strip() if pd.notna(item['ativo']) and str(item['ativo']).lower() != 'nan' and str(item['ativo']).strip() != '' else ''
                        id_est_limpo = str(item['id_estoque_fisico']).strip() if pd.notna(item['id_estoque_fisico']) and str(item['id_estoque_fisico']).strip() != '' else extrair_id_estoque_do_nome(inventario_selected_obj['nome'])
                        desc_est_limpo = str(item['desc_estoque_fisico']).strip() if pd.notna(item['desc_estoque_fisico']) and str(item['desc_estoque_fisico']).strip() != '' else MAPA_ESTOQUES_DESC.get(id_est_limpo, 'ESTOQUE FÍSICO')

                        c1, c2, c3 = st.columns(3)
                        c1.metric("Cód. Produto", str(item['cod_produto']))
                        c2.metric("Lote", lote_auto or "—")
                        c3.metric("Ativo", ativo_auto or "—")
                        
                        st.markdown(f'<div class="caixa-descricao-produto">📦 Descrição: {item["desc_produto"]}</div>', unsafe_allow_html=True)
                        st.info(f"📍 **Estoque Físico:** {desc_est_limpo} (ID: {id_est_limpo}) | **Unidade:** {item['unid_medida']} | **Status:** {status_pasta_atual}")
                        
                        with st.form("form_lancar_qtd", clear_on_submit=True):
                            qtd_fisica = st.number_input("📦 Quantidade Contada Fisicamente:", min_value=0, step=1, value=0)
                            confirma_zero = st.checkbox("⚠️ Marque se este item REALMENTE NÃO EXISTE no estoque (Saldo Zero)")
                            obs = st.text_input("Observação (opcional)")
                            
                            if st.form_submit_button("⚡ Salvar na RAM (Instantâneo)", type="primary", use_container_width=True):
                                if qtd_fisica == 0 and not confirma_zero: st.error("⚠️ Para salvar quantidade 0, marque a confirmação amarela!")
                                else:
                                    t_ini_save = time.time()
                                    qtd_sys = int(item['qtd_estoque'])
                                    dif = qtd_fisica - qtd_sys
                                    fase_gravar = "2a Contagem Concluida" if status_pasta_atual == "2a Contagem" else status_pasta_atual
                                    data_hora_agora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                                    # Atualiza ou insere na lista em memória RAM
                                    item_existente = next((c for c in st.session_state.db_contagens if str(c['inventario_id']) == id_pasta_limpo_base and str(c['cod_produto']) == str(item['cod_produto']) and c['lote'] == lote_auto and c['ativo'] == ativo_auto), None)
                                    
                                    if item_existente:
                                        item_existente['qtd_contada'] = qtd_fisica
                                        item_existente['diferenca'] = dif
                                        item_existente['observacao'] = obs
                                        item_existente['operador'] = st.session_state.operador
                                        item_existente['data_hora'] = data_hora_agora
                                        item_existente['fase_contagem'] = fase_gravar
                                    else:
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
                                            'fase_contagem': fase_gravar
                                        })

                                    if id_est_limpo:
                                        st.session_state.db_ultima_contagem_estoques[id_est_limpo] = data_hora_agora

                                    st.session_state.contador_reset += 1
                                    t_fim_save = time.time()
                                    st.toast(f"⚡ Salvo em RAM em {(t_fim_save - t_ini_save)*1000:.1f}ms!", icon="✅")
                                    st.rerun()

    # --- ABA 2: LANÇAMENTOS E ESPELHO BASE ---
    with aba_lancamentos:
        sub_aba1, sub_aba2 = st.tabs(["📋 Meus Lançamentos Nesta Pasta", "📄 Espelho Base do Saldo (Status Visual)"])
        with sub_aba1:
            df_minhas = pd.DataFrame([c for c in st.session_state.db_contagens if str(c['inventario_id']) == id_pasta_limpo_base])
            if not df_minhas.empty:
                m1, m2, m3 = st.columns(3)
                m1.metric("Lançamentos Totais", len(df_minhas))
                m2.metric("Com Divergência", len(df_minhas[df_minhas['diferenca'] != 0]))
                m3.metric("Total de Peças Contadas", int(df_minhas['qtd_contada'].sum()))
                st.download_button("📥 Exportar Lançamentos Filtrados para Excel", converter_para_excel(df_minhas), file_name=f"contagem_pasta_{id_pasta_limpo_base}.xlsx")
                st.dataframe(df_minhas[['cod_produto', 'desc_produto', 'desc_estoque', 'qtd_sistema', 'qtd_contada', 'diferenca', 'ativo', 'lote', 'observacao', 'operador', 'data_hora', 'fase_contagem']], use_container_width=True, hide_index=True)
            else: st.info("Nenhum lançamento registrado nesta pasta.")
            
        with sub_aba2:
            if not df_base_ram.empty:
                contagens_pasta = [c for c in st.session_state.db_contagens if str(c['inventario_id']) == id_pasta_limpo_base]
                mapa_contados = {f"{str(c['cod_produto']).upper().strip()}_{str(c['lote']).upper().strip() if c['lote'] else ''}_{str(c['ativo']).upper().strip() if c['ativo'] else ''}": c['operador'] for c in contagens_pasta}
                
                def obter_status(row):
                    c, l, a = str(row['cod_produto']).upper().strip(), str(row['lote']).upper().strip() if pd.notna(row['lote']) and str(row['lote']).lower()!='nan' else "", str(row['ativo']).upper().strip() if pd.notna(row['ativo']) and str(row['ativo']).lower()!='nan' else ""
                    key = f"{c}_{l}_{a}"
                    if key in mapa_contados: return f"🟩 Contabilizado por ({mapa_contados[key]})"
                    return "🟥 Não Contado"
                
                df_espelho = df_base_ram.copy()
                df_espelho['Status de Contagem'] = df_espelho.apply(obter_status, axis=1)
                st.dataframe(df_espelho[['Status de Contagem', 'cod_produto', 'desc_produto', 'desc_estoque_fisico', 'unid_medida', 'qtd_estoque', 'id_estoque_fisico', 'lote', 'ativo']], use_container_width=True, hide_index=True)
            else: st.info("Nenhuma base carregada.")

    # --- ABA 3: DESEMPENHO E ACURACIDADE ---
    with aba_desempenho:
        sub_d1, sub_d2 = st.tabs(["🔴 Desempenho & Prazos por Estoque", "📈 Acuracidade Auditada pelo Supervisor"])
        with sub_d1:
            st.subheader("🏆 Status de Atualização dos Estoques Físicos")
            mapa_datas = st.session_state.db_ultima_contagem_estoques.copy()
            
            linhas_desempenho, hoje = [], datetime.datetime.now()
            b_count, a_count, c_count = 0, 0, 0
            
            for est in LISTA_ESTOQUES_FIXA:
                est_id = str(est["id"]).strip()
                u_data = mapa_datas.get(est_id, None)
                if u_data:
                    try:
                        dt = datetime.datetime.strptime(u_data.split(".")[0], "%Y-%m-%d %H:%M:%S")
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
            
            col_f_stat, col_f_busca = st.columns([1, 1])
            filtro_criticidade = col_f_stat.selectbox("🎯 Filtrar por Criticidade:", ["Todos os Estoques", "🟢 Em Dia", "🟡 Necessário Auditar", "🔴 Crítico (+2 semanas)"])
            busca_estoque = col_f_busca.text_input("🔍 Pesquisar Estoque por Código ou Nome:")

            df_exibir = df_desempenho_full.copy()
            if filtro_criticidade != "Todos os Estoques": df_exibir = df_exibir[df_exibir['Criticidade'] == filtro_criticidade]
            if busca_estoque.strip():
                termo = busca_estoque.strip().lower()
                df_exibir = df_exibir[df_exibir['Id. Estoque'].str.lower().str.contains(termo) | df_exibir['Descrição do Estoque Físico'].str.lower().str.contains(termo)]

            df_exibir['Dias Sem Contar'] = df_exibir['Dias Sem Contar'].apply(lambda x: "—" if x == 999 else x)
            st.dataframe(df_exibir.sort_values(by=['Criticidade', 'Id. Estoque'], ascending=[False, True]), use_container_width=True, hide_index=True)

        with sub_d2:
            st.subheader("📈 Tabela de Acuracidade Geral por Depósito")
            df_auds = pd.DataFrame(st.session_state.db_auditorias_sup) if st.session_state.db_auditorias_sup else pd.DataFrame()
            if df_auds.empty: st.info("💡 Nenhuma amostragem coletada pelo supervisor.")
            else:
                linhas_acu = []
                for dep_id, grp in df_auds.groupby('id_estoque'):
                    tot = len(grp)
                    p_s = (len(grp[grp['diferenca'] == 0])/tot)*100
                    p_e = (len(grp[grp['etiqueta_correta'] == "Sim"])/tot)*100
                    p_l = (len(grp[grp['localizacao_correta'] == "Sim"])/tot)*100
                    linhas_acu.append({"CÓDIGO ESTOQUE": dep_id, "DESCRIÇÃO DO ESTOQUE": grp.iloc[0]['desc_estoque'] if 'desc_estoque' in grp.columns else "Não Informado", "ACURACIDADE SALDO": f"{'🟢' if p_s==100 else '🔴'} {p_s:.1f}%", "ACURACIDADE ETIQUETAS": f"{'🟢' if p_e==100 else '🔴'} {p_e:.1f}%", "ACURACIDADE LOCALIZAÇÃO": f"{'🟢' if p_l==100 else '🔴'} {p_l:.1f}%", "ITENS AUDITADOS": tot, "ÚLTIMA AUDITORIA": str(grp.iloc[0]['data_hora']).split(" ")[0] if 'data_hora' in grp.columns else ""})
                st.dataframe(pd.DataFrame(linhas_acu), use_container_width=True, hide_index=True)

            st.markdown("---")
            if not df_inventarios_sup.empty:
                for _, inv_s in df_inventarios_sup.iterrows():
                    df_hist_sup = pd.DataFrame([a for a in st.session_state.db_auditorias_sup if str(a['inventario_id']) == str(inv_s['id'])])
                    c_exp, c_del = st.columns([8, 2])
                    with c_exp:
                        with st.expander(f"📁 Pasta {inv_s['id']} – {inv_s['nome']} | Data: {inv_s['data']} | Status: {inv_s['status']} ({len(df_hist_sup)} itens auditados)"):
                            if not df_hist_sup.empty:
                                st.download_button("📥 Exportar Esta Auditoria para Excel", converter_para_excel(df_hist_sup), file_name=f"auditoria_{inv_s['id']}.xlsx", key=f"dl_sup_acu_{inv_s['id']}")
                                st.dataframe(df_hist_sup, use_container_width=True, hide_index=True)
                            else: st.info("Nenhum item auditado nesta pasta.")
                    with c_del:
                        if eh_supervisor and st.button("🗑️ Excluir Pasta", key=f"del_sup_f_{inv_s['id']}", use_container_width=True):
                            st.session_state.db_inventarios_sup = [i for i in st.session_state.db_inventarios_sup if str(i['id']) != str(inv_s['id'])]
                            st.session_state.db_auditorias_sup = [a for a in st.session_state.db_auditorias_sup if str(a['inventario_id']) != str(inv_s['id'])]
                            st.rerun()

    # --- ABA 4: HISTÓRICO GERAL ---
    with aba_historico:
        st.title("📁 Arquivo Geral de Movimentações")
        if df_inventarios.empty: st.info("Nenhum inventário registrado.")
        else:
            for idx, inv in df_inventarios.iterrows():
                id_proc = str(inv['id']).replace('#','').strip()
                df_h = pd.DataFrame([c for c in st.session_state.db_contagens if str(c['inventario_id']) == id_proc])
                tot_reg, acu_reg = len(df_h) if not df_h.empty else inv.get('total_itens', 0), inv.get('acuracidade_final', '—')

                c_exp, c_del = st.columns([8, 2])
                with c_exp:
                    with st.expander(f"📁 Pasta {inv['id']} - {inv['nome']} | Data: {inv['data']} | Status: {inv['status']} | Acuracidade: {acu_reg} ({tot_reg} itens contados)"):
                        if not df_h.empty:
                            st.download_button("📥 Baixar Planilha de Lançamentos (.xlsx)", converter_para_excel(df_h), file_name=f"inventario_{id_proc}.xlsx", key=f"dl_hist_{id_proc}")
                            st.dataframe(df_h, use_container_width=True, hide_index=True)
                        else: st.info("Nenhum lançamento registrado nesta pasta.")
                with c_del:
                    if eh_supervisor and st.button("🗑️ Excluir Pasta", key=f"del_hist_inv_{inv['id']}", use_container_width=True):
                        st.session_state.db_inventarios = [i for i in st.session_state.db_inventarios if str(i['id']).replace('#', '') != id_proc]
                        st.session_state.db_contagens = [c for c in st.session_state.db_contagens if str(c['inventario_id']) != id_proc]
                        st.session_state.db_itens_base = [b for b in st.session_state.db_itens_base if str(b['inventario_id']) != id_proc]
                        st.rerun()

    # --- ABA 5: GESTÃO ADM ---
    if eh_supervisor and aba_adm is not None:
        with aba_adm:
            st.title("⚙️ Módulo de Gestão do Administrador")
            opcao_adm = st.selectbox("Escolha o Módulo de Ação:", ["🚨 Liberar / Encerrar Divergências", "🔬 Auditoria Amostral (Supervisor)", "📊 Relatório Consolidado (Excel Gerencial)", "👥 Gestão de Usuários & Senhas"])
            st.markdown("---")

            if opcao_adm == "🚨 Liberar / Encerrar Divergências":
                st.subheader("Tratamento de Erros de Contagem da Equipe")
                divs_list = [c for c in st.session_state.db_contagens if c['diferenca'] != 0]
                
                if not divs_list:
                    st.success("🎉 Nenhuma divergência pendente em nenhuma pasta do sistema!")
                else:
                    pastas_divs = list(set([str(c['inventario_id']) for c in divs_list]))
                    inv_alvo_sel = st.selectbox("📁 Selecione a Pasta Com Divergência:", [f"Pasta #{p}" for p in pastas_divs])
                    id_pasta_div_target = inv_alvo_sel.replace("Pasta #", "").strip()
                    
                    items_div_pasta = [c for c in divs_list if str(c['inventario_id']) == id_pasta_div_target]
                    opcoes_items = [f"{c['cod_produto']} - {c['desc_produto']} (Dif: {c['diferenca']})" for c in items_div_pasta]
                    
                    item_alvo_sel = st.selectbox("📦 Selecione o Item Com Divergência:", opcoes_items)
                    idx_sel = opcoes_items.index(item_alvo_sel)
                    row_item_alvo = items_div_pasta[idx_sel]
                    
                    st.info(f"📊 **Resumo:** Sistema: **{row_item_alvo['qtd_sistema']}** | Contado: **{row_item_alvo['qtd_contada']}** | Diferença: **{row_item_alvo['diferenca']}** (Operador: {row_item_alvo['operador']})")
                    justificativa_adm = st.text_input("📝 Justificativa:")
                    
                    col_act1, col_act2 = st.columns(2)
                    with col_act1:
                        if st.button("🚨 Abrir 2ª Contagem para Almoxarife", type="primary", use_container_width=True):
                            if not justificativa_adm.strip(): st.error("⚠️ Digite uma justificativa!")
                            else:
                                row_item_alvo['fase_contagem'] = '2a Contagem'
                                row_item_alvo['qtd_contada'] = 0
                                row_item_alvo['diferenca'] = 0
                                row_item_alvo['observacao'] = f"ADM ({st.session_state.operador}): [LIBERADO 2ª CONTAGEM] - {justificativa_adm.strip()}"
                                for inv in st.session_state.db_inventarios:
                                    if str(inv['id']).replace('#', '') == id_pasta_div_target: inv['status'] = '2a Contagem'
                                st.success("✅ Liberado para 2ª Contagem!")
                                st.rerun()
                    with col_act2:
                        if st.button("🔒 Finalizar e Manter Divergência Atual", use_container_width=True):
                            if not justificativa_adm.strip(): st.error("⚠️ Digite uma justificativa!")
                            else:
                                row_item_alvo['fase_contagem'] = 'Encerrado com Divergencia'
                                row_item_alvo['observacao'] = f"ADM ({st.session_state.operador}): [ENCERRADO DIVERGÊNCIA] - {justificativa_adm.strip()}"
                                st.success("✅ Encerrado com divergência!")
                                st.rerun()

            elif opcao_adm == "🔬 Auditoria Amostral (Supervisor)":
                st.subheader("Módulo de Auditoria Amostral Própria")
                if df_inventarios_sup.empty: id_sup_act, inv_sup_obj = None, None
                else:
                    sel_s = st.selectbox("Selecione a Pasta de Auditoria Ativa:", [f"{r['id']} – {r['nome']} ({r['status']})" for _, r in df_inventarios_sup.iterrows()])
                    id_sup_act = sel_s.split(" – ")[0]
                    inv_sup_obj = df_inventarios_sup[df_inventarios_sup['id'] == id_sup_act].iloc[0]

                with st.expander("➕ Nova Pasta de Auditoria do Supervisor"):
                    with st.form("form_sup_new"):
                        nom_s = st.text_input("Nome da Pasta Amostral")
                        if st.form_submit_button("Criar Pasta Supervisor", type="primary") and nom_s:
                            m_id_s = max([int(str(i['id']).replace('SUP-#', '')) for i in st.session_state.db_inventarios_sup], default=0)
                            st.session_state.db_inventarios_sup.append({
                                'id': f"SUP-#{m_id_s + 1}", 'nome': nom_s, 'data': datetime.date.today().strftime("%Y-%m-%d"), 'status': 'Aberto'
                            })
                            st.rerun()

                arq_sup = st.file_uploader("Suba a planilha Excel de amostras (.xlsx)", type=["xlsx"], key="up_excel_sup")
                if arq_sup is not None and id_sup_act:
                    st.session_state.bases_supervisor_por_inv[id_sup_act] = pd.read_excel(arq_sup)
                    st.success("✅ Planilha anexa com sucesso!")

                base_sup_curr = st.session_state.bases_supervisor_por_inv.get(id_sup_act, None)

                if id_sup_act and base_sup_curr is not None:
                    cols_sup = list(base_sup_curr.columns)
                    def mapear_col_s(opcoes, idx_padrao):
                        for op in opcoes:
                            for c in cols_sup:
                                if op.lower().replace(" ", "").replace(".", "") in str(c).lower().replace(" ", "").replace(".", ""): return c
                        return cols_sup[idx_padrao] if idx_padrao < len(cols_sup) else cols_sup[0]

                    col_cod_s, col_desc_s = mapear_col_s(['códproduto', 'codproduto', 'codigo'], 0), mapear_col_s(['descproduto', 'descricao'], 1)
                    col_local_s, col_qtd_s, col_id_est_s = mapear_col_s(['descestoquefisico', 'localizacao'], 2), mapear_col_s(['qtdestoque', 'quantidade'], -1), mapear_col_s(['idestoquefísico', 'idestoque'], 0)

                    st.dataframe(base_sup_curr, use_container_width=True, hide_index=True)
                    item_combo_sup = st.selectbox("Selecione o material:", [f"{r[col_cod_s]} - {r[col_desc_s]}" for _, r in base_sup_curr.iterrows()])

                    if item_combo_sup:
                        cod_sup_clean = item_combo_sup.split(" - ")[0].strip()
                        row_s = base_sup_curr[base_sup_curr[col_cod_s].astype(str).str.strip() == cod_sup_clean].iloc[0]
                        
                        with st.form("form_auditar_item_completo", clear_on_submit=True):
                            c_f1, c_f2, c_f3 = st.columns(3)
                            q_aud = c_f1.number_input("Quantidade Real Encontrada Fisicamente:", min_value=0, step=1, value=0)
                            e_ok = c_f2.selectbox("A Etiqueta Física está Correta?", ["Sim", "Não"])
                            l_ok = c_f3.selectbox("O Endereçamento/Localização está Correto?", ["Sim", "Não"])
                            at_sup = st.text_input("Número do Ativo (Opcional)")
                            if st.form_submit_button("💾 Salvar Auditoria do Item", type="primary", use_container_width=True):
                                q_sys = int(pd.to_numeric(row_s[col_qtd_s], errors='coerce') or 0)
                                st.session_state.db_auditorias_sup.append({
                                    'inventario_id': id_sup_act,
                                    'id_estoque': str(row_s[col_id_est_s]).strip() if col_id_est_s in base_sup_curr.columns else "",
                                    'desc_estoque': str(row_s[col_local_s]) if col_local_s in base_sup_curr.columns else "Não Informado",
                                    'cod_produto': cod_sup_clean,
                                    'desc_produto': str(row_s[col_desc_s]),
                                    'qtd_sistema': q_sys,
                                    'qtd_auditada': q_aud,
                                    'diferenca': q_aud - q_sys,
                                    'etiqueta_correta': e_ok,
                                    'localizacao_correta': l_ok,
                                    'supervisor': st.session_state.operador,
                                    'data_hora': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    'ativo': at_sup.strip().upper()
                                })
                                st.success("✅ Auditoria registrada em RAM!")
                                st.rerun()

            elif opcao_adm == "📊 Relatório Consolidado (Excel Gerencial)":
                st.subheader("Gerar Planilha Gerencial por Período (Múltiplas Abas)")
                df_c_fil = pd.DataFrame(st.session_state.db_contagens)
                if df_c_fil.empty:
                    st.info("Nenhum lançamento gravado no sistema.")
                else:
                    bytes_ex = gerar_relatorio_consolidado_excel(df_c_fil, LISTA_ESTOQUES_FIXA)
                    st.download_button(
                        label="📥 Clique para Baixar o Relatório Consolidado (.xlsx)", 
                        data=bytes_ex, 
                        file_name="Relatorio_Gerencial_Local.xlsx", 
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                        use_container_width=True
                    )

            elif opcao_adm == "👥 Gestão de Usuários & Senhas":
                st.subheader("Gerenciamento de Colaboradores e Perfis")
                df_usrs = pd.DataFrame(st.session_state.db_usuarios)
                st.dataframe(df_usrs[['id', 'nome', 'cpf', 'email', 'perfil']], use_container_width=True, hide_index=True)
                
                c_u1, c_u2 = st.columns(2)
                with c_u1:
                    u_sel = st.selectbox("Escolha o Colaborador:", [f"{u['id']} - {u['nome']}" for u in st.session_state.db_usuarios])
                    n_senha, n_perfil = st.text_input("Nova Senha", type="password"), st.selectbox("Nível de Acesso:", ["Almoxarife", "Administrador"])
                    if st.button("🔄 Atualizar Dados do Usuário", type="primary", use_container_width=True):
                        uid = int(u_sel.split(" - ")[0])
                        usr_target = next((u for u in st.session_state.db_usuarios if u['id'] == uid), None)
                        if usr_target:
                            if n_senha.strip(): usr_target['senha'] = n_senha.strip()
                            usr_target['perfil'] = n_perfil
                            st.success("✅ Atualizado com sucesso!")
                            st.rerun()
                with c_u2:
                    u_del = st.selectbox("Remover Colaborador:", [f"{u['id']} - {u['nome']}" for u in st.session_state.db_usuarios], key="sb_del")
                    if st.button("❌ Confirmar Exclusão", type="primary", use_container_width=True):
                        uid = int(u_del.split(" - ")[0])
                        st.session_state.db_usuarios = [u for u in st.session_state.db_usuarios if u['id'] != uid]
                        st.success("✅ Usuário removido!")
                        st.rerun()
