import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import urllib.request
import json
import io
import os
from datetime import datetime

st.set_page_config(page_title="CAASI Imports - Gestão", page_icon="🚢", layout="wide")

# ==========================================
# SISTEMA DE LOGIN E SEGURANÇA
# ==========================================
def check_password():
    """Retorna `True` se o utilizador inserir a senha correta."""
    # Define a senha de acesso (Pode ser alterada futuramente)
    SENHA_SISTEMA = "caasi2026"

    def password_entered():
        """Verifica se a senha inserida está correta."""
        if st.session_state["password"] == SENHA_SISTEMA:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Limpa a senha por segurança
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct"):
        return True

    # Tela de Login (Se ainda não estiver logado)
    st.markdown("<h1 style='text-align: center; color: #1F4E78;'>🔐 CAASI Imports - Acesso Restrito</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>Por favor, insira a senha para aceder ao sistema.</p>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.text_input(
            "Senha", type="password", on_change=password_entered, key="password"
        )
        if "password_correct" in st.session_state and not st.session_state["password_correct"]:
            st.error("😕 Senha incorreta. Tente novamente.")
    
    return False

# Executa a verificação de login antes de carregar o resto da app
if not check_password():
    st.stop()  # Para a execução do código aqui se a senha não for inserida

# ==========================================
# CÓDIGO PRINCIPAL DA APP (Protegido)
# ==========================================

# Inicializar Variáveis na Sessão
if 'carrinho' not in st.session_state:
    st.session_state['carrinho'] = []

# --- FUNÇÕES DE BANCO DE DADOS (Excel Local) ---
DB_MASTERDATA = 'masterdata_caasi.xlsx'
DB_ESTOQUE = 'estoque_caasi.xlsx'


def carregar_dados(arquivo, colunas):
    if os.path.exists(arquivo):
        return pd.read_excel(arquivo)
    else:
        return pd.DataFrame(columns=colunas)

def salvar_dados(df, arquivo):
    df.to_excel(arquivo, index=False)

# Carregar os Bancos de Dados
df_masterdata = carregar_dados(DB_MASTERDATA, ['SKU', 'Nome_Produto', 'NCM', 'Preco_Alvo_USD'])
df_estoque = carregar_dados(DB_ESTOQUE, ['Data', 'SKU', 'Produto', 'Tipo_Movimento', 'Quantidade', 'Observacao'])

# --- BARRA LATERAL (MENU) ---
st.sidebar.title("📦 CAASI IMPORTS")
st.sidebar.markdown("Sistema Integrado de Importação e Estoque")
st.sidebar.markdown("---")
menu = st.sidebar.radio("Navegação Operacional", [
    "1. 📊 Cotação e Precificação", 
    "2. 🗃️ Masterdata (Produtos)", 
    "3. 🛠️ Portal de XML (Bling)", 
    "4. 📦 Controlo de Estoque"
])

# Busca Dólar Atualizado
try:
    url = 'https://economia.awesomeapi.com.br/last/USD-BRL'
    req = urllib.request.urlopen(url)
    data = json.loads(req.read())
    dolar_hoje = float(data['USDBRL']['bid'])
except:
    dolar_hoje = 5.35 # Fallback

# ==========================================
# MÓDULO 1: COTAÇÃO E PRECIFICAÇÃO
# ==========================================
if menu == "1. 📊 Cotação e Precificação":
    st.title("📊 Simulador de Viabilidade e Geração de Pedido (PO)")
    st.markdown("Analise o custo nacionalizado e adicione ao pedido para enviar à China.")

    with st.expander("📝 1. Dados do Produto", expanded=True):
        col_A, col_B, col_C = st.columns(3)
        with col_A:
            # Puxa sugestões do Masterdata
            produtos_cadastrados = df_masterdata['Nome_Produto'].tolist() if not df_masterdata.empty else ["Novo Produto..."]
            nome_produto = st.selectbox("Nome do Produto (Inglês)", ["Novo Produto..."] + produtos_cadastrados)
            
            # Auto-preenchimento de NCM
            ncm_sugerido = "39262000"
            if nome_produto != "Novo Produto..." and not df_masterdata.empty:
                ncm_sugerido = str(df_masterdata[df_masterdata['Nome_Produto'] == nome_produto]['NCM'].values[0])
            
            if nome_produto == "Novo Produto...":
                nome_produto = st.text_input("Digite o nome do produto")
                
            ncm = st.text_input("NCM", value=ncm_sugerido)
            formato_venda = st.text_input("Formato (Unit, Set)", value="Unit")
        with col_B:
            descricao = st.text_area("Descrição Detalhada", value="Descrição em inglês para a Invoice...", height=115)
        with col_C:
            link_fornecedor = st.text_input("Link Alibaba / 1688")
            detalhes_cores = st.text_area("Detalhes (Cores, Tamanhos)", value="200 - BLACK", height=45)

    with st.form("form_simulacao"):
        st.markdown("### 🧮 2. Estrutura de Custos")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("**Na Origem (USD)**")
            custo_usd = st.number_input("Custo Unitário (USD)", value=1.50, step=0.10)
            quantidade = st.number_input("Quantidade", value=1000, step=10)
            peso_unitario = st.number_input("Peso Unit. (KG)", value=0.050, step=0.010, format="%.3f")
            frete_usd = st.number_input("Frete Estimado (Total USD)", value=150.00, step=10.0)
            cambio = st.number_input(f"Dólar (Atual: {dolar_hoje:.2f})", value=dolar_hoje, step=0.05)
        with col2:
            st.markdown("**Impostos (%)**")
            aliq_ii = st.number_input("II (%)", value=18.0) / 100
            aliq_ipi = st.number_input("IPI (%)", value=9.75) / 100
            aliq_pis, aliq_cofins = 0.021, 0.0965
            aliq_icms = st.number_input("ICMS Estadual (%)", value=18.0) / 100
            siscomex_brl = st.number_input("Siscomex/Despesas (BRL)", value=250.00, step=10.0)
        with col3:
            st.markdown("**Venda (BRL)**")
            preco_venda = st.number_input("Preço de Venda (BRL)", value=35.00, step=1.0)
            taxa_ml = st.number_input("Taxa Marketplace (%)", value=16.0) / 100
            st.markdown("<br>", unsafe_allow_html=True)
            submit_button = st.form_submit_button("🔄 Calcular", type="primary", use_container_width=True)

    if submit_button:
        # Matemática
        peso_total = peso_unitario * quantidade
        vmld_brl = custo_usd * quantidade * cambio
        frete_brl = frete_usd * cambio
        valor_aduaneiro = vmld_brl + frete_brl
        
        vII = valor_aduaneiro * aliq_ii
        vIPI = (valor_aduaneiro + vII) * aliq_ipi
        vPIS, vCOFINS = valor_aduaneiro * aliq_pis, valor_aduaneiro * aliq_cofins
        
        base_icms = (valor_aduaneiro + vII + vIPI + vPIS + vCOFINS + siscomex_brl) / (1 - aliq_icms)
        vICMS = base_icms * aliq_icms
        
        custo_total_nacional = valor_aduaneiro + vII + vIPI + vPIS + vCOFINS + vICMS + siscomex_brl
        custo_unit = custo_total_nacional / quantidade if quantidade > 0 else 0
        
        lucro = (preco_venda * (1 - taxa_ml)) - custo_unit
        margem = (lucro / preco_venda) * 100 if preco_venda > 0 else 0

        st.session_state['ultimo_calculo'] = {
            'NAME': nome_produto, 'DETAILED DESCRIPTION': descricao, 'REFERENCE': link_fornecedor,
            'NCM': ncm, 'Sales Format': formato_venda, 'Quantity': quantidade, 'UNIT WEIGHT (KG)': peso_unitario,
            'UNIT PRICE': custo_usd, 'Total Product Cost': custo_usd*quantidade, 'DETAILS': detalhes_cores,
            'Custo BR': round(custo_unit, 2), 'Margem %': round(margem, 2)
        }

        m1, m2, m3 = st.columns(3)
        m1.metric("Custo Nacionalizado (Unid.)", f"R$ {custo_unit:.2f}")
        m2.metric("Lucro Líquido (Unid.)", f"R$ {lucro:.2f}")
        if margem >= 15: m3.metric("Margem", f"{margem:.1f}%", "Viável")
        else: m3.metric("Margem", f"{margem:.1f}%", "Apertada/Prejuízo", delta_color="inverse")

        if st.button("➕ Adicionar ao Pedido (PO)"):
            st.session_state['carrinho'].append(st.session_state['ultimo_calculo'])
            st.success("Adicionado com sucesso!")

    if len(st.session_state['carrinho']) > 0:
        st.markdown("---")
        st.subheader("🛒 Carrinho de Importação (Purchase Order)")
        df_po = pd.DataFrame(st.session_state['carrinho'])
        st.dataframe(df_po, use_container_width=True)
        
        if st.button("🗑️ Limpar"):
            st.session_state['carrinho'] = []
            st.rerun()
            
        buffer = io.BytesIO()
        df_po.to_excel(buffer, index=False)
        st.download_button("📥 Exportar Pedido (Excel) para China", buffer.getvalue(), "PO_CAASI.xlsx", type="primary")

# ==========================================
# MÓDULO 2: MASTERDATA (BANCO DE DADOS)
# ==========================================
elif menu == "2. 🗃️ Masterdata (Produtos)":
    st.title("🗃️ Gestor de Cadastros e NCMs")
    st.markdown("Base de conhecimento da CAASI. Registe os NCMs validados para não depender do chinês.")

    with st.form("form_masterdata"):
        col1, col2 = st.columns(2)
        sku = col1.text_input("SKU Interno (Opcional)", value=f"PRD-{len(df_masterdata)+1:04d}")
        nome = col2.text_input("Nome do Produto (Inglês ou Português)")
        ncm = col1.text_input("NCM Correto e Validado")
        preco_alvo = col2.number_input("Preço de Compra Alvo (USD)", step=0.10)
        
        if st.form_submit_button("Salvar na Base de Dados", type="primary"):
            novo_dado = pd.DataFrame([[sku, nome, ncm, preco_alvo]], columns=df_masterdata.columns)
            df_masterdata = pd.concat([df_masterdata, novo_dado], ignore_index=True)
            salvar_dados(df_masterdata, DB_MASTERDATA)
            st.success("Produto cadastrado na nuvem com sucesso!")
            st.rerun()

    st.markdown("---")
    st.subheader("Produtos Cadastrados")
    st.dataframe(df_masterdata, use_container_width=True)

# ==========================================
# MÓDULO 3: PORTAL DE ENTRADA XML
# ==========================================
elif menu == "3. 🛠️ Portal de XML (Bling)":
    st.title("🛠️ Portal de Integração Bling")
    
    aba1, aba2 = st.tabs(["1️⃣ Importação Formal (Corretor de Despachante)", "2️⃣ Importação Simplificada (Gerador via Planilha)"])
    
    # -- ABA 1: HIGIENIZADOR DE XML --
    with aba1:
        st.markdown("Arraste o XML do Despachante. O sistema arranja os arredondamentos e a ST (Erro 531 e 932).")
        uploaded_xml = st.file_uploader("Arquivo XML (Despachante)", type=['xml'])

        if uploaded_xml:
            ET.register_namespace('', 'http://www.portalfiscal.inf.br/nfe')
            tree = ET.parse(uploaded_xml)
            root = tree.getroot()
            ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

            soma_vBC, soma_vICMS = 0.0, 0.0
            
            # Auditoria NCM e Correções
            for det in root.findall('.//nfe:det', ns):
                imposto = det.find('nfe:imposto', ns)
                if imposto is not None:
                    icms = imposto.find('.//nfe:ICMS/*', ns)
                    if icms is not None:
                        modBCST = icms.find('nfe:modBCST', ns)
                        if modBCST is not None: modBCST.text = '3'
                        
                        vBC = round(float(icms.find('nfe:vBC', ns).text), 2)
                        vICMS = round(float(icms.find('nfe:vICMS', ns).text), 2)
                        icms.find('nfe:vBC', ns).text = f"{vBC:.2f}"
                        icms.find('nfe:vICMS', ns).text = f"{vICMS:.2f}"
                        soma_vBC += vBC; soma_vICMS += vICMS

            total = root.find('.//nfe:total/nfe:ICMSTot', ns)
            if total is not None:
                total.find('nfe:vBC', ns).text = f"{soma_vBC:.2f}"
                total.find('nfe:vICMS', ns).text = f"{soma_vICMS:.2f}"

            xml_str = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            st.success("✅ XML corrigido com sucesso!")
            st.download_button("📥 Baixar XML Higienizado", xml_str, "XML_Pronto_Bling.xml", "text/xml", type="primary")

    # -- ABA 2: GERADOR DE SIMPLIFICADA (DIR/UPS) --
    with aba2:
        st.markdown("Arraste a **Commercial Invoice (Excel)** da China para gerar o XML com 60% de II + ICMS.")
        uploaded_csv = st.file_uploader("Arquivo Excel (Invoice)", type=['xlsx', 'csv'])
        
        c1, c2, c3 = st.columns(3)
        dir_dolar = c1.number_input("Dólar Base (DIR)", value=5.3338, format="%.4f")
        dir_frete = c2.number_input("Frete/Seguro Total (BRL)", value=0.0)
        dir_icms = c3.number_input("Alíquota ICMS Estado (%)", value=17.0) / 100
        
        if uploaded_csv and st.button("🚀 Processar e Gerar XML Automático", type="primary"):
            try:
                df_inv = pd.read_excel(uploaded_csv)
                
                # Tenta localizar a linha onde começa a tabela de produtos
                start_row = -1
                for i, row in df_inv.iterrows():
                    if 'DESCRIPTION OF GOODS' in str(row.values).upper() or 'NAME' in str(row.values).upper():
                        start_row = i
                        break
                
                if start_row != -1:
                    df_inv = pd.read_excel(uploaded_csv, skiprows=start_row+1)
                
                # Procura colunas chave
                cols = [str(c).upper().strip() for c in df_inv.columns]
                df_inv.columns = cols
                
                # Identifica as colunas certas dinamicamente
                col_nome = [c for c in cols if 'NAME' in c or 'DESCRIPTION' in c][0]
                col_ncm = [c for c in cols if 'HS CODE' in c or 'NCM' in c][0]
                col_qty = [c for c in cols if 'QTY' in c or 'QUANTITY' in c][0]
                col_total = [c for c in cols if 'TOTAL' in c and 'COST' in c or 'AMOUNT' in c][0]

                df_inv = df_inv.dropna(subset=[col_qty, col_total])
                
                # Geração Estrutura XML Básica NFe 4.0
                nfe = ET.Element("NFe", xmlns="http://www.portalfiscal.inf.br/nfe")
                infNFe = ET.SubElement(nfe, "infNFe", Id="NFeGeradaCaasi", versao="4.00")
                
                soma_prod_brl = 0
                soma_bc_icms = 0
                soma_icms = 0
                soma_ii = 0
                
                for idx, row in df_inv.iterrows():
                    det = ET.SubElement(infNFe, "det", nItem=str(idx+1))
                    prod = ET.SubElement(det, "prod")
                    ET.SubElement(prod, "cProd").text = f"IMP-{idx+1:03d}"
                    ET.SubElement(prod, "xProd").text = str(row[col_nome])[:120]
                    ET.SubElement(prod, "NCM").text = str(row[col_ncm]).replace('.', '').strip()[:8]
                    ET.SubElement(prod, "CFOP").text = "3102"
                    ET.SubElement(prod, "uCom").text = "UN"
                    ET.SubElement(prod, "qCom").text = str(row[col_qty])
                    
                    vProd_usd = float(row[col_total])
                    vProd_brl = vProd_usd * dir_dolar
                    ET.SubElement(prod, "vProd").text = f"{vProd_brl:.2f}"
                    
                    # Cálculo DIR: 60% II + ICMS por dentro
                    vII = vProd_brl * 0.60
                    # Rateio básico do frete para BC do ICMS
                    rateio_frete = dir_frete / len(df_inv) 
                    base_icms = (vProd_brl + vII + rateio_frete) / (1 - dir_icms)
                    vICMS = base_icms * dir_icms
                    
                    soma_prod_brl += vProd_brl
                    soma_bc_icms += base_icms
                    soma_icms += vICMS
                    soma_ii += vII
                    
                    imposto = ET.SubElement(det, "imposto")
                    icms = ET.SubElement(imposto, "ICMS")
                    icmssn = ET.SubElement(icms, "ICMSSN900")
                    ET.SubElement(icmssn, "orig").text = "1"
                    ET.SubElement(icmssn, "CSOSN").text = "900"
                    ET.SubElement(icmssn, "modBC").text = "3"
                    ET.SubElement(icmssn, "vBC").text = f"{base_icms:.2f}"
                    ET.SubElement(icmssn, "pICMS").text = f"{dir_icms*100:.2f}"
                    ET.SubElement(icmssn, "vICMS").text = f"{vICMS:.2f}"
                    
                    ii_tag = ET.SubElement(imposto, "II")
                    ET.SubElement(ii_tag, "vBC").text = f"{vProd_brl:.2f}"
                    ET.SubElement(ii_tag, "vDespAdu").text = "0.00"
                    ET.SubElement(ii_tag, "vII").text = f"{vII:.2f}"
                    ET.SubElement(ii_tag, "vIOF").text = "0.00"

                # Fechar Totais
                total = ET.SubElement(infNFe, "total")
                icmstot = ET.SubElement(total, "ICMSTot")
                ET.SubElement(icmstot, "vBC").text = f"{soma_bc_icms:.2f}"
                ET.SubElement(icmstot, "vICMS").text = f"{soma_icms:.2f}"
                ET.SubElement(icmstot, "vProd").text = f"{soma_prod_brl:.2f}"
                ET.SubElement(icmstot, "vII").text = f"{soma_ii:.2f}"
                ET.SubElement(icmstot, "vNF").text = f"{(soma_prod_brl + soma_ii + soma_icms):.2f}" # Aproximado para rascunho

                xml_saida = ET.tostring(nfe, encoding='utf-8', xml_declaration=True)
                st.success("✅ Matriz XML Simplificada Gerada com Sucesso! Faça a importação no Bling e finalize o Destinatário.")
                st.download_button("📥 Baixar XML Simplificado (DIR)", xml_saida, "XML_Simplificada_UPS.xml", "text/xml", type="primary")

            except Exception as e:
                st.error(f"Erro na leitura da planilha. Garanta que é o arquivo Excel padrão da China. Detalhe: {e}")

# ==========================================
# MÓDULO 4: CONTROLE DE ESTOQUE
# ==========================================
elif menu == "4. 📦 Controlo de Estoque":
    st.title("📦 Inventário e Entradas/Saídas")
    st.markdown("Controle as quantidades disponíveis. Alimente via XML ou baixe com vendas do Mercado Pago.")

    c_ent, c_sai = st.columns(2)
    with c_ent:
        st.subheader("Registrar Movimentação")
        with st.form("form_movimento"):
            data_mov = datetime.now().strftime("%d/%m/%Y")
            
            prods = df_masterdata['Nome_Produto'].tolist() if not df_masterdata.empty else ["Cadastre produtos no Masterdata"]
            prod_sel = st.selectbox("Produto", prods)
            
            tipo_mov = st.radio("Tipo", ["Entrada (Importação)", "Saída (Venda)"], horizontal=True)
            qtd_mov = st.number_input("Quantidade", min_value=1, step=1)
            obs = st.text_input("Observação (Ex: Venda ML, Chegada UPS)")
            
            if st.form_submit_button("Lançar no Estoque", type="primary"):
                mult = 1 if tipo_mov.startswith("Ent") else -1
                novo_mov = pd.DataFrame([[data_mov, "SKU", prod_sel, tipo_mov, qtd_mov * mult, obs]], columns=df_estoque.columns)
                df_estoque = pd.concat([df_estoque, novo_mov], ignore_index=True)
                salvar_dados(df_estoque, DB_ESTOQUE)
                st.success("Estoque atualizado!")
                st.rerun()

    with c_sai:
        st.subheader("Baixa em Lote (Mercado Livre/Bling)")
        arquivo_vendas = st.file_uploader("Subir Relatório de Vendas (CSV/Excel)", type=['csv', 'xlsx'])
        if arquivo_vendas:
            st.info("Funcionalidade Pronta para Mapeamento: O sistema lerá os SKUs vendidos e fará a baixa automática de todos de uma vez.")

    st.markdown("---")
    st.subheader("📊 Posição Atual do Estoque")
    
    if not df_estoque.empty:
        # Agrupa pelo nome do produto somando as quantidades
        saldo_atual = df_estoque.groupby('Produto')['Quantidade'].sum().reset_index()
        saldo_atual.columns = ['Produto', 'Saldo Disponível']
        
        c1, c2 = st.columns([2, 3])
        with c1:
            st.dataframe(saldo_atual, use_container_width=True)
        with c2:
            st.markdown("**Histórico Recente de Movimentações**")
            st.dataframe(df_estoque.tail(10).iloc[::-1], use_container_width=True) # Mostra os ultimos 10 invertidos
    else:
        st.info("Nenhuma movimentação registada no momento.")
