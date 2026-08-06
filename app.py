import streamlit as st
import pandas as pd
from supabase import create_client, Client
import datetime

# ==========================================
# 1. CONFIGURAÇÃO DA PÁGINA E CONEXÃO SUPABASE
# ==========================================
st.set_page_config(page_title="Sistema JBA - Inventário", layout="wide", initial_sidebar_state="expanded")

SUPABASE_URL = st.secrets.get("SUPABASE_URL", "https://seu-projeto.supabase.co")
SUPABASE_KEY = st.secrets.get("SUPABASE_KEY", "sua-chave-anon-key")

@st.cache_resource
def init_supabase() -> Client:
    return create_client(SUPABASE_URL, SUPABASE_KEY)

try:
    supabase = init_supabase()
except Exception as e:
    st.error(f"Erro ao conectar ao Supabase: {e}")

# ==========================================
# 2. INICIALIZAÇÃO DE ESTADOS DA SESSÃO (RAM)
# ==========================================
if 'fila_bipagem' not in st.session_state:
    st.session_state['fila_bipagem'] = []

if 'usuario_logado' not in st.session_state:
    st.session_state['usuario_logado'] = None

if 'perfil_usuario' not in st.session_state:
    st.session_state['perfil_usuario'] = "Almoxarife"  # Perfil padrão

if 'pasta_ativa' not in st.session_state:
    st.session_state['pasta_ativa'] = None

# ==========================================
# 3. FUNÇÃO DE SINCRONIZAÇÃO EM LOTE (RAM -> NUVEM)
# ==========================================
def sincronizar_ram_com_supabase():
    """Descarrega os lançamentos salvos na memória em uma única requisição."""
    if st.session_state['fila_bipagem']:
        try:
            supabase.table("contagiosos").insert(st.session_state['fila_bipagem']).execute()
            total_sincronizado = len(st.session_state['fila_bipagem'])
            st.session_state['fila_bipagem'] = []  # Limpa a RAM local
            st.toast(f"✅ {total_sincronizado} itens sincronizados no Supabase!", icon="🚀")
            return True
        except Exception as e:
            st.error(f"Falha ao enviar lote para o banco: {e}")
            return False
    return False

# ==========================================
# 4. TELA DE LOGIN E AUTENTICAÇÃO DE USUÁRIO
# ==========================================
if not st.session_state['usuario_logado']:
    st.markdown("## 🔒 Acesso ao Sistema JBA")
    with st.form("form_login"):
        cpf_email = st.text_input("CPF ou E-mail")
        senha = st.text_input("Senha", type="password")
        btn_login = st.form_submit_button("Entrar no Sistema", type="primary")
        
        if btn_login:
            if cpf_email and senha:
                # Lógica de validação simples (pode ser vinculada à tabela 'usuarios')
                st.session_state['usuario_logado'] = cpf_email
                if "admin" in cpf_email.lower():
                    st.session_state['perfil_usuario'] = "Administrador"
                else:
                    st.session_state['perfil_usuario'] = "Almoxarife"
                st.rerun()
            else:
                st.warning("Preencha todos os campos para continuar.")
    st.stop()

# ==========================================
# 5. BARRA LATERAL (CONTROLE DE PASTAS E RAM)
# ==========================================
st.sidebar.markdown(f"👤 **{st.session_state['usuario_logado']}** ({st.session_state['perfil_usuario']})")

# Indicador de Status da Memória RAM
pendentes = len(st.session_state['fila_bipagem'])
if pendentes > 0:
    st.sidebar.warning(f"⚡ {pendentes} bips guardados na RAM local!")
    if st.sidebar.button("🚀 Sincronizar Agora com a Nuvem", type="primary"):
        sincronizar_ram_com_supabase()
        st.rerun()
else:
    st.sidebar.success("🟢 Memória RAM sincronizada.")

st.sidebar.divider()
st.sidebar.markdown("### 📁 Seleção de Inventário")

# Busca inventários no Supabase
try:
    res_inv = supabase.table("inventarios").select("id, nome, status").execute()
    inventarios_lista = res_inv.data if res_inv.data else []
except:
    inventarios_lista = []

opcoes_pasta = {f"#{i['id']} - {i['nome']} ({i['status']})": i for i in inventarios_lista}
pasta_selecionada_nome = st.sidebar.selectbox("Selecione a Pasta", options=[""] + list(opcoes_pasta.keys()))

if pasta_selecionada_nome:
    obj_pasta = opcoes_pasta[pasta_selecionada_nome]
    st.session_state['pasta_ativa'] = obj_pasta

# Abertura de Novo Inventário
with st.sidebar.expander("➕ Criar Novo Inventário"):
    nome_novo_inv = st.text_input("Nome/Código do Inventário:")
    if st.button("Salvar Inventário"):
        if nome_novo_inv:
            novo_inv = {"nome": nome_novo_inv, "status": "Aberto", "dados": str(datetime.date.today())}
            supabase.table("inventarios").insert(novo_inv).execute()
            st.sidebar.success("Inventário criado com sucesso!")
            st.rerun()

# ==========================================
# 6. MENU DE NAVEGAÇÃO PRINCIPAL (ABAS)
# ==========================================
aba1, aba2, aba3, aba4, aba5 = st.tabs([
    "💻 Contar Item (RAM Modo Rápido)",
    "📊 Lançamentos & Base",
    "📈 Desempenho & Acuracidade",
    "📁 Histórico Geral",
    "⚙️ Gestão ADM"
])

# ------------------------------------------
# ABA 1: BIPAGEM ULTRA RÁPIDA (MODO RAM)
# ------------------------------------------
with aba1:
    if not st.session_state['pasta_ativa']:
        st.info("⚠️ Selecione um inventário ativo na barra lateral para iniciar a contagem.")
    elif st.session_state['pasta_ativa']['status'] == "Fechado":
        st.error("🔒 Este inventário está Fechado. Selecione uma pasta em aberto para bipar.")
    else:
        st.subheader(f"Contagem Ativa: {st.session_state['pasta_ativa']['nome']}")
        
        with st.form("form_bipagem_rapida", clear_on_submit=True):
            cod_bip = st.text_input("💻 Bipar ou Digitar Código do Produto:")
            qtd_fisica = st.number_input("📦 Quantidade Contada Fisicamente:", min_value=0, value=1)
            saldo_zero = st.checkbox("⚠️ Marque se este item REALMENTE NÃO EXISTE no estoque (Saldo Zero)")
            
            btn_salvar_ram = st.form_submit_button("⚡ Salvar na RAM (Instantâneo)", type="primary")

        if btn_salvar_ram:
            if cod_bip:
                registro = {
                    "inventario_id": st.session_state['pasta_ativa']['id'],
                    "cod_produto": cod_bip.strip().upper(),
                    "qtd_contada": 0 if saldo_zero else qtd_fisica,
                    "operador": st.session_state['usuario_logado'],
                    "fase_contagem": "1a Contagem",
                    "data_hora": str(datetime.datetime.now())
                }
                # Gravação instantânea na RAM sem esperar resposta da rede
                st.session_state['fila_bipagem'].append(registro)
                st.success(f"Item **{cod_bip.strip().upper()}** registrado na RAM!")
                
                # Auto-sincronização a cada 10 bips
                if len(st.session_state['fila_bipagem']) >= 10:
                    sincronizar_ram_com_supabase()
            else:
                st.warning("Por favor, bipe ou informe um código válido.")

# ------------------------------------------
# ABA 2: LANÇAMENTOS E ESPELHO DE SALDO
# ------------------------------------------
with aba2:
    st.subheader("📊 Lançamentos Registrados na Pasta Ativa")
    if st.session_state['pasta_ativa']:
        inv_id = st.session_state['pasta_ativa']['id']
        res_bips = supabase.table("contagiosos").select("*").eq("inventario_id", inv_id).execute()
        
        if res_bips.data:
            df_bips = pd.DataFrame(res_bips.data)
            st.dataframe(df_bips, use_container_width=True)
            
            # Exportação direta para Excel
            st.download_button(
                label="📥 Exportar Lançamentos Filtrados para Excel",
                data=df_bips.to_csv(index=False).encode('utf-8'),
                file_name=f"lancamentos_pasta_{inv_id}.csv",
                mime="text/csv"
            )
        else:
            st.info("Nenhum lançamento gravado no banco de dados para esta pasta ainda.")

# ------------------------------------------
# ABA 3: DESEMPENHO E ACURACIDADE
# ------------------------------------------
with aba3:
    st.subheader("📈 Acuracidade e Prazos dos Estoques")
    st.markdown("Acompanhamento de qualidade das contagens e datas de última vistoria por depósito.")
    # Métricas gerais do painel
    col_kpi1, col_kpi2, col_kpi3 = st.columns(3)
    col_kpi1.metric("Estoques Em Dia", "10 Depósitos", delta="Conforme")
    col_kpi2.metric("Necessário Auditar", "0 Depósitos", delta_color="off")
    col_kpi3.metric("Status Crítico (+14 dias)", "21 Depósitos", delta="-Atenção")

# ------------------------------------------
# ABA 4: HISTÓRICO GERAL
# ------------------------------------------
with aba4:
    st.subheader("📁 Arquivo Geral de Movimentações")
    if inventarios_lista:
        df_hist = pd.DataFrame(inventarios_lista)
        st.dataframe(df_hist, use_container_width=True)

# ------------------------------------------
# ABA 5: GESTÃO ADM E RECONTAGEM (CORRIGIDA)
# ------------------------------------------
with aba5:
    if st.session_state['perfil_usuario'] != "Administrador":
        st.error("🚨 Acesso Restrito: Módulo exclusivo para Supervisores e Administradores.")
    else:
        st.subheader("⚙️ Módulo de Gestão do Administrador")
        
        acao_adm = st.selectbox("Escolha o Módulo de Ação:", [
            "🚨 Liberar / Encerrar Divergências (2ª Contagem)",
            "📊 Relatório Consolidado (Excel Gerencial)",
            "👥 Gestão de Usuários & Senhas"
        ])

        if "2ª Contagem" in acao_adm:
            st.markdown("### Tratamento de Erros de Contagem da Equipe")
            
            if st.session_state['pasta_ativa']:
                p_id = st.session_state['pasta_ativa']['id']
                
                # Busca itens da pasta com divergência
                res_div = supabase.table("contagiosos").select("*").eq("inventario_id", p_id).execute()
                
                if res_div.data:
                    df_div = pd.DataFrame(res_div.data)
                    lista_itens = df_div['cod_produto'].unique().tolist()
                    
                    item_selecionado = st.selectbox("Selecione o Item com Divergência nesta Pasta:", options=lista_itens)
                    justificativa = st.text_input("Informe a justificativa/observação (Obrigatório):", key="txt_justificativa")
                    
                    c1, c2 = st.columns(2)
                    
                    # BOTÃO 1: ABRIR 2ª CONTAGEM (Destrava e reabre para o almoxarife)
                    with c1:
                        if st.button("🚨 Abrir 2ª Contagem para Almoxarife", type="primary"):
                            if not justificativa.strip():
                                st.warning("Por favor, digite uma justificativa antes de abrir a 2ª contagem.")
                            else:
                                try:
                                    # 1. Atualiza o status do item na tabela
                                    supabase.table("contagiosos").update({
                                        "fase_contagem": "2a Contagem",
                                        "observacao": f"ADM: [LIBERADO 2ª CONTAGEM] - {justificativa}"
                                    }).eq("inventario_id", p_id).eq("cod_produto", item_selecionado).execute()
                                    
                                    # 2. Garante que a pasta permaneça em estado Aberto/2a Contagem
                                    supabase.table("inventarios").update({"status": "2a Contagem"}).eq("id", p_id).execute()
                                    
                                    st.success(f"2ª Contagem liberada com sucesso para o item {item_selecionado}!")
                                    st.rerun()
                                except Exception as err:
                                    st.error(f"Erro ao processar liberação: {err}")

                    # BOTÃO 2: FINALIZAR E MANTER DIVERGÊNCIA
                    with c2:
                        if st.button("🔒 Finalizar e Manter Divergência Atual"):
                            if not justificativa.strip():
                                st.warning("Por favor, digite uma justificativa para encerrar.")
                            else:
                                try:
                                    supabase.table("contagiosos").update({
                                        "observacao": f"ADM: [ENCERRADO COM DIVERGÊNCIA] - {justificativa}"
                                    }).eq("inventario_id", p_id).eq("cod_produto", item_selecionado).execute()
                                    
                                    st.info(f"Divergência encerrada e registrada para o item {item_selecionado}.")
                                    st.rerun()
                                except Exception as err:
                                    st.error(f"Erro ao encerrar divergência: {err}")
                else:
                    st.info("Nenhum lançamento encontrado nesta pasta para tratamento.")
            else:
                st.warning("Selecione uma pasta ativa no menu lateral antes de tratar as divergências.")
