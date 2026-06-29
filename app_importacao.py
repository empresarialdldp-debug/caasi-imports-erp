import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import urllib.request
import json
import io
import os
import re
import random
from datetime import datetime
import google.generativeai as genai

# Leitor de PDF necessário para a DIR e Recibos (Nota de Débito)
try:
    import PyPDF2
    pypdf_installed = True
except ImportError:
    pypdf_installed = False

st.set_page_config(page_title="CAASI Imports - Gestão", page_icon="🚢", layout="wide")

erro_ia_msg = ""
try:
    if "GEMINI_API_KEY" in st.secrets:
        genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
        ia_configurada = True
    else:
        ia_configurada = False
        erro_ia_msg = "A chave 'GEMINI_API_KEY' não foi encontrada nos Secrets do Streamlit."
except Exception as e:
    ia_configurada = False
    erro_ia_msg = f"Erro ao configurar a IA: {str(e)}"

# ==========================================
# SISTEMA DE LOGIN E SEGURANÇA
# ==========================================
def check_password():
    """Retorna `True` se o utilizador inserir a senha correta."""
    SENHA_SISTEMA = "caasi2026"

    def password_entered():
        if st.session_state["password"] == SENHA_SISTEMA:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Limpa a senha por segurança
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct"):
        return True

    # Tela de Login
    st.markdown("<h1 style='text-align: center; color: #1F4E78;'>🔐 CAASI Imports - Acesso Restrito</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>Por favor, insira a senha para aceder ao sistema.</p>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.text_input("Senha", type="password", on_change=password_entered, key="password")
        if "password_correct" in st.session_state and not st.session_state["password_correct"]:
            st.error("😕 Senha incorreta. Tente novamente.")
    
    return False

# Executa a verificação de login antes de carregar o resto da app
if not check_password():
    st.stop()

if 'carrinho' not in st.session_state:
    st.session_state['carrinho'] = []

# --- FUNÇÕES DE BANCO DE DADOS (Excel Local) ---
DB_MASTERDATA = 'masterdata_caasi.xlsx'
DB_ESTOQUE = 'estoque_caasi.xlsx'

def carregar_dados(arquivo, colunas):
    if os.path.exists(arquivo):
        return pd.read_excel(arquivo)
    else:
        df_vazio = pd.DataFrame(columns=colunas)
        df_vazio.to_excel(arquivo, index=False)
        return df_vazio

def salvar_dados(df, arquivo):
    df.to_excel(arquivo, index=False)

df_masterdata = carregar_dados(DB_MASTERDATA, ['SKU', 'Nome_Produto', 'NCM', 'Preco_Alvo_USD'])
df_estoque = carregar_dados(DB_ESTOQUE, ['Data', 'SKU', 'Produto', 'Tipo_Movimento', 'Quantidade', 'Observacao'])

st.sidebar.title("📦 CAASI IMPORTS")
st.sidebar.markdown("Sistema Integrado de Importação e Estoque")
st.sidebar.markdown("---")
menu = st.sidebar.radio("Navegação Operacional", [
    "1. 📊 Cotação e Precificação", 
    "2. 🗃️ Masterdata (Produtos)", 
    "3. 🛠️ Portal de XML (Bling)", 
    "4. 📦 Controlo de Estoque",
    "5. 🟡 Inteligência Mercado Livre"
])

try:
    url = 'https://economia.awesomeapi.com.br/last/USD-BRL'
    req = urllib.request.urlopen(url)
    data = json.loads(req.read())
    dolar_hoje = float(data['USDBRL']['bid'])
except:
    dolar_hoje = 5.35 # Fallback

if not ia_configurada:
    st.error(f"⚠️ A Inteligência Artificial está desativada! Motivo: {erro_ia_msg}")

# ==========================================
# MÓDULO 1: COTAÇÃO E PRECIFICAÇÃO
# ==========================================
if menu == "1. 📊 Cotação e Precificação":
    st.title("📊 Cotação, Precificação e Mercado")
    st.markdown("Importe cotações antigas, analise o mercado com IA e calcule a viabilidade exata.")

    # --- 1. IMPORTAR PLANILHA DE COTAÇÕES ---
    with st.expander("📂 1. Importar Planilha de Cotação (Validação)", expanded=False):
        st.markdown("Suba uma planilha de cotação antiga ou recebida do fornecedor para visualizar os dados.")
        arquivo_cotacao = st.file_uploader("Arquivo Excel ou CSV", type=['xlsx', 'csv'], key="import_cotacao")
        if arquivo_cotacao:
            try:
                df_cot = pd.read_excel(arquivo_cotacao) if arquivo_cotacao.name.endswith('.xlsx') else pd.read_csv(arquivo_cotacao)
                st.dataframe(df_cot, use_container_width=True)
                st.success("Planilha carregada com sucesso! Use os dados acima para preencher a simulação abaixo.")
            except Exception as e:
                st.error(f"Erro ao ler arquivo: {e}")

    # --- 2. INTELIGÊNCIA DE MERCADO E PRODUTO ---
    with st.expander("🧠 2. Inteligência de Mercado (Mercado Livre)", expanded=True):
        st.markdown("Digite o nome do produto livremente e peça à IA para analisar os preços e concorrência no Brasil.")
        
        # Digitação livre e direta do Marcelo
        nome_produto = st.text_input("Nome do Produto a ser cotado:", placeholder="Ex: Capa de Silicone iPhone 15, Lanterna Tática X900...")
        
        col_pesq1, col_pesq2 = st.columns(2)
        with col_pesq1:
            if st.button("🤖 Analisar Mercado com IA", type="primary", use_container_width=True):
                if not nome_produto:
                    st.warning("Por favor, digite o nome de um produto primeiro.")
                elif ia_configurada:
                    with st.spinner(f"A IA está pesquisando a base do Mercado Livre para '{nome_produto}'..."):
                        try:
                            model = genai.GenerativeModel('gemini-2.5-flash')
                            prompt = f"""
                            Aja como o Diretor de Pricing da CAASI Imports, especialista no Mercado Livre Brasil.
                            Faça uma análise rápida e direta sobre o produto: '{nome_produto}'.
                            Traga os seguintes dados em tópicos (estimativas realistas baseadas no mercado atual):
                            1. Preço médio de venda atual (R$).
                            2. Nível de concorrência (Baixo, Médio, Alto).
                            3. Preço sugerido (R$) para entrar competitivo e ganhar a Buy Box dos top 5.
                            4. Palavras-chave mais buscadas para esse produto no ML.
                            """
                            resposta = model.generate_content(prompt)
                            st.info(resposta.text)
                        except Exception as e:
                            st.error(f"Erro de Comunicação com a IA: {e}")
                else:
                    st.error("A IA não conseguiu conectar. Verifique os Secrets.")
                    
        with col_pesq2:
            if nome_produto:
                termo_ml = nome_produto.replace(" ", "-")
                link_ml = f"https://lista.mercadolivre.com.br/{termo_ml}"
                st.markdown(f'<a href="{link_ml}" target="_blank"><button style="background-color:#ffe600; color:#333; border:none; padding:10px 20px; text-align:center; text-decoration:none; display:inline-block; font-size:16px; margin:4px 2px; cursor:pointer; border-radius:8px; width:100%; font-weight:bold;">📦 Ver este produto direto no ML</button></a>', unsafe_allow_html=True)

        col_B, col_C = st.columns(2)
        with col_B:
            ncm = st.text_input("NCM (Opcional)", value="39262000")
            formato_venda = st.text_input("Formato (Unit, Set, Pairs)", value="Unit")
        with col_C:
            link_fornecedor = st.text_input("Link Alibaba / Fornecedor")
            descricao = st.text_area("Detalhes Rápidos (Para o Carrinho)", value="Cores sortidas, embalagem plástica", height=68)

    # --- 3. CÁLCULO E PRECIFICAÇÃO ---
    with st.form("form_simulacao"):
        st.markdown("### 🧮 3. Estrutura de Custos")
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
            taxa_ml = st.number_input("Taxa ML Aproximada (%)", value=16.0) / 100
            st.markdown("<br>", unsafe_allow_html=True)
            submit_button = st.form_submit_button("🔄 Calcular", type="primary", use_container_width=True)

    if submit_button:
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
            'UNIT PRICE': custo_usd, 'TARGET': '', 'Total Product Cost': custo_usd*quantidade,
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
        st.subheader("🛒 Carrinho de Importação (PO)")
        df_po = pd.DataFrame(st.session_state['carrinho'])
        st.dataframe(df_po[['NAME', 'Quantity', 'UNIT PRICE', 'Total Product Cost', 'Margem %']], use_container_width=True)
        
        c_limpar, c_export = st.columns(2)
        with c_limpar:
            if st.button("🗑️ Limpar"):
                st.session_state['carrinho'] = []
                st.rerun()
        with c_export:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_po.drop(columns=['Custo BR', 'Margem %']).to_excel(writer, sheet_name='INVOICE', index=False)
            st.download_button("📥 Exportar Pedido (Excel) para China", buffer.getvalue(), "PO_CAASI.xlsx", type="primary", use_container_width=True)

# ==========================================
# MÓDULO 2: MASTERDATA (BANCO DE DADOS)
# ==========================================
elif menu == "2. 🗃️ Masterdata (Produtos)":
    st.title("🗃️ Gestor de Cadastros e NCMs")
    st.markdown("Base de conhecimento da CAASI. Registe os NCMs validados.")

    st.markdown("### 🤖 Consultor Fiscal IA (Buscador de NCM)")
    if ia_configurada:
        with st.expander("Não sabe o NCM? Descreva o produto e pergunte à IA:", expanded=False):
            desc_produto = st.text_input("Descreva o produto (Ex: Isca artificial de silicone)")
            if st.button("🔍 Consultar NCM e Regras", type="secondary"):
                if desc_produto:
                    with st.spinner("A IA está a analisar..."):
                        try:
                            model = genai.GenerativeModel('gemini-2.5-flash')
                            prompt = f"Aja como Despachante Aduaneiro no Brasil. Produto: '{desc_produto}'. Qual a classificação fiscal (NCM) adequada? Dê o NCM 8 dígitos, breve justificativa e informe se exige LI."
                            st.info(model.generate_content(prompt).text)
                        except Exception as e:
                            st.error(f"Erro IA: {e}")
    else:
        st.warning("IA não conectada no Secrets.")

    with st.form("form_masterdata"):
        col1, col2 = st.columns(2)
        sku = col1.text_input("SKU Interno (Opcional)", value=f"PRD-{len(df_masterdata)+1:04d}")
        nome = col2.text_input("Nome do Produto")
        ncm = col1.text_input("NCM Correto")
        preco_alvo = col2.number_input("Preço Alvo (USD)", step=0.10)
        
        if st.form_submit_button("Salvar na Base", type="primary"):
            novo_dado = pd.DataFrame([[sku, nome, ncm, preco_alvo]], columns=df_masterdata.columns)
            df_masterdata = pd.concat([df_masterdata, novo_dado], ignore_index=True)
            salvar_dados(df_masterdata, DB_MASTERDATA)
            st.success("Produto cadastrado!")
            st.rerun()

    st.dataframe(df_masterdata, use_container_width=True)

# ==========================================
# MÓDULO 3: PORTAL DE ENTRADA XML
# ==========================================
elif menu == "3. 🛠️ Portal de XML (Bling)":
    st.title("🛠️ Portal de Integração Bling")
    aba1, aba2 = st.tabs(["1️⃣ Corretor de XML (Despachante)", "2️⃣ Gerador XML Inteligente (Simplificada DIR)"])
    
    with aba1:
        st.markdown("Arraste o XML do Despachante para corrigir arredondamentos (Erro 531 e 932).")
        uploaded_xml = st.file_uploader("Arquivo XML", type=['xml'])
        if uploaded_xml:
            ET.register_namespace('', 'http://www.portalfiscal.inf.br/nfe')
            tree = ET.parse(uploaded_xml)
            root = tree.getroot()
            ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
            soma_vBC, soma_vICMS = 0.0, 0.0
            
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
            st.success("✅ XML corrigido!")
            st.download_button("📥 Baixar XML Higienizado", xml_str, "XML_Bling.xml", "text/xml")

    with aba2:
        st.markdown("### Geração Profissional de XML para Importação Simplificada (Conforme DIR)")
        st.markdown("A Inteligência Artificial lerá a DIR e o Recibo de Impostos (PDFs) e os cruzará com a Invoice da China, garantindo que o XML gerado tenha os valores aduaneiros e a Tag `<DI>` perfeitos, e formatados como 'Fornecedor Estrangeiro' conforme manual do Bling.")
        
        if not pypdf_installed:
            st.error("⚠️ ERRO CRÍTICO: O sistema precisa do pacote 'PyPDF2' para ler PDFs. Adicione 'PyPDF2' ao arquivo requirements.txt no GitHub.")
        
        col_up1, col_up2, col_up3 = st.columns(3)
        uploaded_csv = col_up1.file_uploader("1. Excel (Invoice)", type=['xlsx', 'csv'])
        uploaded_dir = col_up2.file_uploader("2. PDF da DIR", type=['pdf'])
        uploaded_recibo = col_up3.file_uploader("3. PDF do Recibo/Imposto", type=['pdf'])
        
        col_imp1, col_imp2, col_imp3 = st.columns(3)
        nome_fornecedor = col_imp1.text_input("Fornecedor (Remetente)", value="SHANDONG SHY CLOUD TECH CO")
        dir_icms = col_imp2.number_input("ICMS do Estado (%)", value=18.0) / 100
        dir_outras = col_imp3.number_input("Outras Despesas DHL/UPS (BRL)", value=91.08)
        
        st.markdown("### Configurações de Transporte")
        col_transp1, col_transp2 = st.columns(2)
        transportadora = col_transp1.selectbox("Transportadora (Courier)", ["UPS DO BRASIL", "DHL EXPRESS BRASIL", "FEDEX BRASIL", "Nenhum/Outro"])
        qtd_volumes = col_transp2.number_input("Quantidade de Volumes", value=1, min_value=1, step=1)
        
        st.warning("""
        📝 **CHECKLIST OBRIGATÓRIA PÓS-IMPORTAÇÃO NO BLING:**
        
        O formato XML da SEFAZ aceita apenas dados fiscais (valores). Opções internas do Bling **não podem** ser enviadas via XML. 
        Após importar o XML gerado por este sistema para dentro do Bling, você **PRECISA** abrir a nota importada e realizar estes 4 passos manualmente em todos os itens:
        
        1. **Aba Importação:** Selecione a opção `Somar ICMS: Sim`.
        2. **Aba Outros:** Em 'Presumido no cálculo do PIS/COFINS', marque `Sim`.
        3. **Aba Outros:** Em 'Tipo do item', selecione `Mercadoria para Revenda`.
        4. **Aba Estoque:** Marque `Não cadastrar` (para evitar duplicar o saldo de estoque).
        """)
        
        if 'numero_nfe_atual' not in st.session_state:
            st.session_state['numero_nfe_atual'] = 100009
            
        numero_nfe = st.number_input("Número da NFe (Inicia em 100009)", value=st.session_state['numero_nfe_atual'], step=1)
        
        if uploaded_csv and uploaded_dir and uploaded_recibo:
            if st.button("🚀 Extrair Dados Aduaneiros e Gerar XML (Padrão Bling)", type="primary", use_container_width=True):
                if not ia_configurada:
                    st.error("A Inteligência Artificial precisa estar configurada para extrair os dados dos PDFs. Verifique os Secrets.")
                else:
                    try:
                        with st.spinner("La IA está extrayendo el Número de la DIR, Dólar PTAX, Flete e Impuestos pagados..."):
                            # 1. Extrair Texto dos PDFs com PyPDF2
                            texto_dir = ""
                            leitor_dir = PyPDF2.PdfReader(uploaded_dir)
                            for page in leitor_dir.pages: texto_dir += page.extract_text() + "\n"
                            
                            texto_recibo = ""
                            leitor_recibo = PyPDF2.PdfReader(uploaded_recibo)
                            for page in leitor_recibo.pages: texto_recibo += page.extract_text() + "\n"
                            
                            # 2. IA Audita os Documentos
                            model = genai.GenerativeModel('gemini-2.5-flash')
                            prompt = f"""
                            Aja como um auditor fiscal especialista em comércio exterior. Extraia os dados numéricos cruciais dos seguintes textos extraídos da DIR e Recibo de Importação (FedEx/UPS/DHL).
                            TEXTO DA DIR E RECIBO:
                            {texto_dir[:3000]}
                            {texto_recibo[:3000]}
                            
                            Retorne APENAS um JSON válido e limpo com as seguintes chaves (sem formatação extra):
                            "numero_dir": (string, pegue os dígitos numéricos da DIR/Declaração. Ex: "260024703330"),
                            "data_desembaraco": (string, data do registro formato YYYY-MM-DD. Ex: "2026-02-06"),
                            "local_desembaraco": (string, pegue o local, Ex: "VIRACOPOS"),
                            "uf_desembaraco": (string, UF do local de desembaraço, Ex: "SP"),
                            "valor_frete_brl": (float, Valor Frete em R$, sem símbolo. Ex: 1174.26),
                            "valor_ii_brl": (float, Imposto Importação I.I. pago em R$. Ex: 2247.25),
                            "taxa_dolar": (float, calcule com alta precisão dividindo o Valor Total Remessa BRL pelo Valor Total Remessa USD. Ex: 2571.16 / 489.0)
                            """
                            resposta = model.generate_content(prompt)
                            
                            json_str = resposta.text
                            match = re.search(r'```json(.*?)```', json_str, re.DOTALL)
                            if match: json_str = match.group(1)
                            else: json_str = json_str.replace("```", "")
                            
                            dados_dir = json.loads(json_str.strip())
                            
                            st.success(f"**Dados da DIR Extraídos pela IA:** Número: {dados_dir['numero_dir']} | I.I. Pago: R$ {dados_dir['valor_ii_brl']} | Frete: R$ {dados_dir['valor_frete_brl']} | Dólar: {dados_dir['taxa_dolar']:.4f}")
                            
                            # 3. Cruzar com a Planilha (Excel Invoice) de Forma Ultrarrobusta
                            df_inv = pd.read_excel(uploaded_csv) if uploaded_csv.name.endswith('.xlsx') else pd.read_csv(uploaded_csv)
                            start_row = -1
                            
                            # Filtra as linhas de cabeçalho ignorando nomes de pessoas/empresas
                            for i, row in df_inv.iterrows():
                                row_str = " ".join([str(val).upper() for val in row.values if pd.notna(val)])
                                
                                tem_desc = any(kw in row_str for kw in ['DESC', 'GOOD', 'ITEM', 'PROD', 'NAME', 'ARTICLE'])
                                tem_qtd = any(kw in row_str for kw in ['QTY', 'QUANT', 'PCS', 'PIECE'])
                                tem_valor = any(kw in row_str for kw in ['TOTAL', 'AMOUNT', 'PRICE', 'COST', 'VALUE'])
                                se_nome_ignorar = any(kw in row_str for kw in ['STELLA', 'MARIA', 'SHIPPER', 'CONSIGNEE', 'ATTN', 'COMPANY'])
                                
                                if tem_desc and (tem_qtd or tem_valor) and not se_nome_ignorar:
                                    start_row = i
                                    break
                            
                            # Fallback caso a busca acima seja rígida demais
                            if start_row == -1:
                                for i, row in df_inv.iterrows():
                                    row_str = " ".join([str(val).upper() for val in row.values if pd.notna(val)])
                                    if 'QTY' in row_str or 'QUANTITY' in row_str:
                                        start_row = i
                                        break
                                        
                            if start_row != -1:
                                df_inv = pd.read_excel(uploaded_csv, skiprows=start_row+1) if uploaded_csv.name.endswith('.xlsx') else pd.read_csv(uploaded_csv, skiprows=start_row+1)
                            
                            cols = [str(c).upper().strip() for c in df_inv.columns]
                            df_inv.columns = cols
                            
                            col_nome = next((c for c in cols if any(kw in c for kw in ['DESC', 'GOOD', 'ITEM', 'PROD', 'NAME', 'ARTICLE'])), None)
                            col_ncm = next((c for c in cols if any(kw in c for kw in ['HS', 'NCM', 'CODE', 'CODIGO'])), None)
                            col_qty = next((c for c in cols if any(kw in c for kw in ['QTY', 'QUANT', 'PCS', 'PIECE'])), None)
                            
                            col_total = next((c for c in cols if any(kw in c for kw in ['TOTAL', 'AMOUNT', 'VALOR'])), None)
                            if not col_total:
                                col_total = next((c for c in cols if any(kw in c for kw in ['PRICE', 'COST', 'VALUE'])), None)
                                
                            if not col_nome or not col_qty or not col_total:
                                raise ValueError(f"Las columnas no fueron reconocidas. Verifique si la fila de encabezado contiene palabras como Description, Qty, Total. Columnas leídas: {cols}")
                            
                            # Força a conversão para números e remove linhas de totais textuais do rodapé
                            df_inv[col_qty] = pd.to_numeric(df_inv[col_qty].astype(str).str.replace(r'[^\d.]', '', regex=True), errors='coerce')
                            df_inv[col_total] = pd.to_numeric(df_inv[col_total].astype(str).str.replace(r'[^\d.]', '', regex=True), errors='coerce')
                            df_inv = df_inv.dropna(subset=[col_qty, col_total])
                            
                            total_produtos_usd = df_inv[col_total].astype(float).sum()
                            
                            nfe = ET.Element("NFe", xmlns="[http://www.portalfiscal.inf.br/nfe](http://www.portalfiscal.inf.br/nfe)")
                            
                            # Gerar ID da NFe com 44 dígitos simulados
                            chave_nfe = f"3125124410256200011155001{numero_nfe:09d}12345678"
                            infNFe = ET.SubElement(nfe, "infNFe", Id=f"NFe{chave_nfe}", versao="4.00")
                            
                            # IDE - Identificação da Nota
                            ide = ET.SubElement(infNFe, "ide")
                            ET.SubElement(ide, "cUF").text = "31" # MG
                            ET.SubElement(ide, "cNF").text = str(random.randint(10000000, 99999999))
                            ET.SubElement(ide, "natOp").text = "Compra de mercadoria"
                            ET.SubElement(ide, "mod").text = "55"
                            ET.SubElement(ide, "serie").text = "1"
                            ET.SubElement(ide, "nNF").text = str(numero_nfe)
                            data_atual = datetime.now().strftime("%Y-%m-%dT%H:%M:%S-03:00")
                            ET.SubElement(ide, "dhEmi").text = data_atual
                            ET.SubElement(ide, "dhSaiEnt").text = data_atual
                            ET.SubElement(ide, "tpNF").text = "0" # 0 = Entrada (Importação)
                            ET.SubElement(ide, "idDest").text = "3" # 3 = Operação com exterior
                            ET.SubElement(ide, "cMunFG").text = "3106200" # BH
                            ET.SubElement(ide, "tpImp").text = "1"
                            ET.SubElement(ide, "tpEmis").text = "1"
                            ET.SubElement(ide, "cDV").text = "3"
                            ET.SubElement(ide, "tpAmb").text = "1"
                            ET.SubElement(ide, "finNFe").text = "1"
                            ET.SubElement(ide, "indFinal").text = "0"
                            ET.SubElement(ide, "indPres").text = "9" # 9 = Não presencial
                            ET.SubElement(ide, "indIntermed").text = "0"
                            ET.SubElement(ide, "procEmi").text = "0"
                            ET.SubElement(ide, "verProc").text = "Bling 1.1"
                            
                            # EMIT - CAASI (Sua empresa é a Emitente da Nota de Entrada)
                            emit = ET.SubElement(infNFe, "emit")
                            ET.SubElement(emit, "CNPJ").text = "44102562000111"
                            ET.SubElement(emit, "xNome").text = "CAASI IMPORTACAO E COMERCIO LTDA"
                            ET.SubElement(emit, "xFant").text = "CAASI IMPORTS"
                            enderEmit = ET.SubElement(emit, "enderEmit")
                            ET.SubElement(enderEmit, "xLgr").text = "RUA JOSE FERREIRA LOPES"
                            ET.SubElement(enderEmit, "nro").text = "197"
                            ET.SubElement(enderEmit, "xCpl").text = "CASA"
                            ET.SubElement(enderEmit, "xBairro").text = "JARDIM FLORENCA"
                            ET.SubElement(enderEmit, "cMun").text = "3106200"
                            ET.SubElement(enderEmit, "xMun").text = "Belo Horizonte"
                            ET.SubElement(enderEmit, "UF").text = "MG"
                            ET.SubElement(enderEmit, "CEP").text = "31680110"
                            ET.SubElement(enderEmit, "cPais").text = "1058"
                            ET.SubElement(enderEmit, "xPais").text = "Brasil"
                            ET.SubElement(enderEmit, "fone").text = "31996326633"
                            ET.SubElement(emit, "IE").text = "0041882320093"
                            ET.SubElement(emit, "CRT").text = "1" # 1 = Simples Nacional
                            
                            # DEST - FORNECEDOR ESTRANGEIRO (Remetente real da Mercadoria)
                            # ATENÇÃO: No layout da SEFAZ NÃO EXISTE tag <remetente>. 
                            # Em notas de ENTRADA (tpNF=0), a tag <dest> é lida como o FORNECEDOR (Remetente)
                            dest = ET.SubElement(infNFe, "dest")
                            ET.SubElement(dest, "idEstrangeiro").text = "00000"
                            ET.SubElement(dest, "xNome").text = nome_fornecedor.strip()[:60]
                            enderDest = ET.SubElement(dest, "enderDest")
                            ET.SubElement(enderDest, "xLgr").text = "EXTERIOR"
                            ET.SubElement(enderDest, "nro").text = "SN"
                            ET.SubElement(enderDest, "xBairro").text = "EXTERIOR"
                            ET.SubElement(enderDest, "cMun").text = "9999999"
                            ET.SubElement(enderDest, "xMun").text = "EXTERIOR"
                            ET.SubElement(enderDest, "UF").text = "EX"
                            ET.SubElement(enderDest, "cPais").text = "1600"
                            ET.SubElement(enderDest, "xPais").text = "CHINA, REPUBLICA POPULAR"
                            ET.SubElement(dest, "indIEDest").text = "9" # 9 = Não Contribuinte

                            soma_prod_brl = soma_bc_icms = soma_icms = soma_ii = soma_frete = soma_outras = 0
                            aliq_ii = 0.60
                            
                            for idx, row in df_inv.iterrows():
                                vProd_usd = float(row[col_total])
                                qtd_item = float(row[col_qty])
                                proporcao = vProd_usd / total_produtos_usd if total_produtos_usd > 0 else 0
                                
                                # CÁLCULO EXATO DA IMPORTAÇÃO (Frete embutido no custo do produto)
                                frete_total_usd = float(dados_dir['valor_frete_brl']) / float(dados_dir['taxa_dolar'])
                                rateio_frete_usd = frete_total_usd * proporcao
                                
                                vProd_total_usd = vProd_usd + rateio_frete_usd
                                vProd_brl = vProd_total_usd * float(dados_dir['taxa_dolar'])
                                
                                rateio_outras_despesas = float(dir_outras) * proporcao
                                
                                # Imposto de Importação (60% da Base que é o Produto já com o frete)
                                vII = vProd_brl * aliq_ii
                                
                                # ICMS (Base = Produto + II) / (1 - aliq_icms). As Outras Despesas ficam FORA da Base!
                                base_icms = (vProd_brl + vII) / (1 - dir_icms)
                                vICMS = base_icms * dir_icms
                                
                                det = ET.SubElement(infNFe, "det", nItem=str(idx+1))
                                prod = ET.SubElement(det, "prod")
                                ET.SubElement(prod, "cProd").text = f"IMP-{(idx+1):03d}"
                                ET.SubElement(prod, "cEAN").text = "SEM GTIN"
                                ET.SubElement(prod, "xProd").text = str(row[col_nome])[:120]
                                
                                ncm_val = str(row[col_ncm]).replace('.', '').strip()[:8] if col_ncm and pd.notna(row[col_ncm]) else "00000000"
                                ET.SubElement(prod, "NCM").text = ncm_val
                                
                                ET.SubElement(prod, "CFOP").text = "3102"
                                ET.SubElement(prod, "uCom").text = "UN"
                                ET.SubElement(prod, "qCom").text = f"{qtd_item:.4f}"
                                
                                # Valor Unitário calculado
                                vUnCom = vProd_brl / qtd_item if qtd_item > 0 else vProd_brl
                                ET.SubElement(prod, "vUnCom").text = f"{vUnCom:.10f}"
                                
                                # Valor Total do Produto na Nota
                                ET.SubElement(prod, "vProd").text = f"{vProd_brl:.2f}"
                                ET.SubElement(prod, "cEANTrib").text = "SEM GTIN"
                                
                                ET.SubElement(prod, "uTrib").text = "UN"
                                ET.SubElement(prod, "qTrib").text = f"{qtd_item:.4f}"
                                ET.SubElement(prod, "vUnTrib").text = f"{vUnCom:.10f}"
                                
                                # Outras Despesas (lançadas separadas e fora da base do ICMS)
                                ET.SubElement(prod, "vOutro").text = f"{rateio_outras_despesas:.2f}"
                                ET.SubElement(prod, "indTot").text = "1"
                                
                                # Tag da DI
                                di = ET.SubElement(prod, "DI")
                                ET.SubElement(di, "nDI").text = str(dados_dir['numero_dir'])
                                ET.SubElement(di, "dDI").text = str(dados_dir['data_desembaraco'])
                                ET.SubElement(di, "xLocDesemb").text = str(dados_dir['local_desembaraco']).strip()
                                ET.SubElement(di, "UFDesemb").text = str(dados_dir['uf_desembaraco']).strip()
                                ET.SubElement(di, "dDesemb").text = str(dados_dir['data_desembaraco'])
                                ET.SubElement(di, "tpViaTransp").text = "4" # Aéreo
                                ET.SubElement(di, "tpIntermedio").text = "1"
                                ET.SubElement(di, "cExportador").text = "40601426132"
                                
                                adi = ET.SubElement(di, "adi")
                                ET.SubElement(adi, "nAdicao").text = str(idx+1)
                                ET.SubElement(adi, "nSeqAdic").text = "1"
                                ET.SubElement(adi, "cFabricante").text = "N/A"
                                ET.SubElement(adi, "nDraw").text = "0"
                                
                                soma_prod_brl += vProd_brl
                                soma_bc_icms += base_icms
                                soma_icms += vICMS
                                soma_ii += vII
                                soma_outras += rateio_outras_despesas
                                
                                imposto = ET.SubElement(det, "imposto")
                                vTotTrib = vII + vICMS
                                ET.SubElement(imposto, "vTotTrib").text = f"{vTotTrib:.2f}"
                                
                                # ICMS CSOSN 900
                                icms = ET.SubElement(imposto, "ICMS")
                                icmssn = ET.SubElement(icms, "ICMSSN900")
                                ET.SubElement(icmssn, "orig").text = "1" 
                                ET.SubElement(icmssn, "CSOSN").text = "900"
                                ET.SubElement(icmssn, "modBC").text = "0" # Conforme o seu XML validado!
                                ET.SubElement(icmssn, "vBC").text = f"{base_icms:.2f}"
                                ET.SubElement(icmssn, "pICMS").text = f"{dir_icms*100:.4f}"
                                ET.SubElement(icmssn, "vICMS").text = f"{vICMS:.2f}"
                                ET.SubElement(icmssn, "pCredSN").text = "0.00"
                                ET.SubElement(icmssn, "vCredICMSSN").text = "0.00"
                                
                                # IPI (CST 49)
                                ipi = ET.SubElement(imposto, "IPI")
                                ET.SubElement(ipi, "cEnq").text = "999"
                                ipitrib = ET.SubElement(ipi, "IPITrib")
                                ET.SubElement(ipitrib, "CST").text = "49"
                                ET.SubElement(ipitrib, "vBC").text = "0.00"
                                ET.SubElement(ipitrib, "pIPI").text = "0.00"
                                ET.SubElement(ipitrib, "vIPI").text = "0.00"
                                
                                # II
                                ii_tag = ET.SubElement(imposto, "II")
                                ET.SubElement(ii_tag, "vBC").text = f"{vProd_brl:.2f}"
                                ET.SubElement(ii_tag, "vDespAdu").text = "0.00"
                                ET.SubElement(ii_tag, "vII").text = f"{vII:.2f}"
                                ET.SubElement(ii_tag, "vIOF").text = "0.00"
                                
                                # PIS (CST 99)
                                pis = ET.SubElement(imposto, "PIS")
                                pisoutr = ET.SubElement(pis, "PISOutr")
                                ET.SubElement(pisoutr, "CST").text = "99"
                                ET.SubElement(pisoutr, "vBC").text = "0.00"
                                ET.SubElement(pisoutr, "pPIS").text = "0.00"
                                ET.SubElement(pisoutr, "vPIS").text = "0.00"
                                
                                # COFINS (CST 99)
                                cofins = ET.SubElement(imposto, "COFINS")
                                cofinsoutr = ET.SubElement(cofins, "COFINSOutr")
                                ET.SubElement(cofinsoutr, "CST").text = "99"
                                ET.SubElement(cofinsoutr, "vBC").text = "0.00"
                                ET.SubElement(cofinsoutr, "pCOFINS").text = "0.00"
                                ET.SubElement(cofinsoutr, "vCOFINS").text = "0.00"
                                
                            total = ET.SubElement(infNFe, "total")
                            icmstot = ET.SubElement(total, "ICMSTot")
                            ET.SubElement(icmstot, "vBC").text = f"{soma_bc_icms:.2f}"
                            ET.SubElement(icmstot, "vICMS").text = f"{soma_icms:.2f}"
                            ET.SubElement(icmstot, "vICMSDeson").text = "0.00"
                            ET.SubElement(icmstot, "vFCP").text = "0.00"
                            ET.SubElement(icmstot, "vBCST").text = "0.00"
                            ET.SubElement(icmstot, "vST").text = "0.00"
                            ET.SubElement(icmstot, "vFCPST").text = "0.00"
                            ET.SubElement(icmstot, "vFCPSTRet").text = "0.00"
                            ET.SubElement(icmstot, "vProd").text = f"{soma_prod_brl:.2f}"
                            ET.SubElement(icmstot, "vFrete").text = "0.00"
                            ET.SubElement(icmstot, "vSeg").text = "0.00"
                            ET.SubElement(icmstot, "vDesc").text = "0.00"
                            ET.SubElement(icmstot, "vII").text = f"{soma_ii:.2f}"
                            ET.SubElement(icmstot, "vIPI").text = "0.00"
                            ET.SubElement(icmstot, "vIPIDevol").text = "0.00"
                            ET.SubElement(icmstot, "vPIS").text = "0.00"
                            ET.SubElement(icmstot, "vCOFINS").text = "0.00"
                            ET.SubElement(icmstot, "vOutro").text = f"{soma_outras:.2f}"
                            
                            # Valor Total da NF
                            v_nf_total = soma_prod_brl + soma_ii + soma_outras + soma_icms
                            ET.SubElement(icmstot, "vNF").text = f"{v_nf_total:.2f}"
                            ET.SubElement(icmstot, "vTotTrib").text = f"{soma_icms + soma_ii:.2f}"

                            # Bloco de Transporte e Volumes
                            transp = ET.SubElement(infNFe, "transp")
                            ET.SubElement(transp, "modFrete").text = "0" # 0 = Contratação do Frete por conta do Remetente (CIF)
                            
                            transporta = ET.SubElement(transp, "transporta")
                            if transportadora == "UPS DO BRASIL":
                                ET.SubElement(transporta, "CNPJ").text = "74155052000173"
                                ET.SubElement(transporta, "xNome").text = "UPS DO BRASIL REMESSAS EXPRESSAS LTDA"
                                ET.SubElement(transporta, "IE").text = "114953497113"
                                ET.SubElement(transporta, "xEnder").text = "R. Dom Aguirre, 554"
                                ET.SubElement(transporta, "xMun").text = "SAO PAULO"
                                ET.SubElement(transporta, "UF").text = "SP"
                            elif transportadora == "DHL EXPRESS BRASIL":
                                ET.SubElement(transporta, "CNPJ").text = "58118019000108"
                                ET.SubElement(transporta, "xNome").text = "DHL EXPRESS (BRAZIL) LTDA"
                                ET.SubElement(transporta, "IE").text = "112613589110"
                                ET.SubElement(transporta, "xEnder").text = "AV. OTAVIANO ALVES DE LIMA, 4000"
                                ET.SubElement(transporta, "xMun").text = "SAO PAULO"
                                ET.SubElement(transporta, "UF").text = "SP"
                            elif transportadora == "FEDEX BRASIL":
                                ET.SubElement(transporta, "CNPJ").text = "10970887000102"
                                ET.SubElement(transporta, "xNome").text = "FEDERAL EXPRESS CORPORATION"
                                ET.SubElement(transporta, "IE").text = "111425110118"
                                ET.SubElement(transporta, "xEnder").text = "R. JOAO PRESTES MAIA, 200"
                                ET.SubElement(transporta, "xMun").text = "SAO PAULO"
                                ET.SubElement(transporta, "UF").text = "SP"

                            vol = ET.SubElement(transp, "vol")
                            ET.SubElement(vol, "qVol").text = str(qtd_volumes)
                            ET.SubElement(vol, "esp").text = "Caixa(s)"
                            ET.SubElement(vol, "pesoL").text = "0.000"
                            ET.SubElement(vol, "pesoB").text = "0.000"
                            
                            # Bloco de Pagamento
                            pag = ET.SubElement(infNFe, "pag")
                            detPag = ET.SubElement(pag, "detPag")
                            ET.SubElement(detPag, "tPag").text = "01" # 01 = Dinheiro
                            ET.SubElement(detPag, "vPag").text = f"{v_nf_total:.2f}"
                            
                            xml_saida = ET.tostring(nfe, encoding='utf-8', xml_declaration=True)
                            
                            st.session_state['numero_nfe_atual'] = numero_nfe + 1
                            
                            st.success(f"✅ Matriz XML Integrada (DIR {dados_dir['numero_dir']}) Gerada com Sucesso!")
                            st.download_button("📥 Baixar XML para Bling", xml_saida, f"XML_CAASI_{dados_dir['numero_dir']}.xml", "text/xml", type="primary")

                    except Exception as e:
                        st.error(f"Erro Crítico ao gerar o XML Integrado: Verifique os PDFs ou a Planilha. Detalhe técnico: {e}")

# ==========================================
# MÓDULO 4: CONTROLE DE ESTOQUE
# ==========================================
elif menu == "4. 📦 Controlo de Estoque":
    st.title("📦 Inventário e Entradas/Saídas")

    with st.expander("🚀 Importar Estoque Inicial (Marco Zero - T0)", expanded=False):
        arquivo_t0 = st.file_uploader("Planilha de Estoque Inicial", type=['xlsx', 'csv'], key="t0")
        if arquivo_t0 and st.button("Carregar Saldo T0", type="primary"):
            try:
                df_t0 = pd.read_excel(arquivo_t0) if arquivo_t0.name.endswith('.xlsx') else pd.read_csv(arquivo_t0)
                col_prod = next((c for c in df_t0.columns if 'PRODUTO' in str(c).upper() or 'DESCRI' in str(c).upper()), None)
                col_qtd = next((c for c in df_t0.columns if 'QUANT' in str(c).upper() or 'QTD' in str(c).upper()), None)
                if col_prod and col_qtd:
                    for _, row in df_t0.iterrows():
                        novo_mov = pd.DataFrame([[datetime.now().strftime("%d/%m/%Y"), "SKU-T0", row[col_prod], "Entrada (T0)", row[col_qtd], "Carga Inicial"]], columns=df_estoque.columns)
                        df_estoque = pd.concat([df_estoque, novo_mov], ignore_index=True)
                    salvar_dados(df_estoque, DB_ESTOQUE)
                    st.success("Estoque T0 carregado!")
                    st.rerun()
                else:
                    st.error("Faltam colunas de Produto/Quantidade.")
            except Exception as e:
                st.error(f"Erro: {e}")

    c_ent, c_sai = st.columns(2)
    with c_ent:
        with st.form("form_movimento"):
            data_mov = datetime.now().strftime("%d/%m/%Y")
            prods = df_masterdata['Nome_Produto'].tolist() if not df_masterdata.empty else ["Sem produtos"]
            prod_sel = st.selectbox("Produto", prods)
            tipo_mov = st.radio("Tipo", ["Entrada (Importação)", "Saída (Venda)"], horizontal=True)
            qtd_mov = st.number_input("Quantidade", min_value=1, step=1)
            obs = st.text_input("Observação")
            if st.form_submit_button("Lançar Movimento", type="primary"):
                mult = 1 if "Entrada" in tipo_mov else -1
                novo_mov = pd.DataFrame([[data_mov, "SKU-REG", prod_sel, tipo_mov, qtd_mov * mult, obs]], columns=df_estoque.columns)
                df_estoque = pd.concat([df_estoque, novo_mov], ignore_index=True)
                salvar_dados(df_estoque, DB_ESTOQUE)
                st.success("Atualizado!")
                st.rerun()

    with c_sai:
        st.info("Baixa em Lote via Mercado Pago (Mapeamento em Breve)")

    st.markdown("---")
    if not df_estoque.empty:
        saldo = df_estoque.groupby('Produto')['Quantidade'].sum().reset_index()
        saldo.columns = ['Produto', 'Saldo']
        c1, c2 = st.columns([2, 3])
        c1.dataframe(saldo, use_container_width=True)
        c2.dataframe(df_estoque.tail(10).iloc[::-1], use_container_width=True)

# ==========================================
# MÓDULO 5: INTELIGÊNCIA MERCADO LIVRE
# ==========================================
elif menu == "5. 🟡 Inteligência Mercado Livre":
    st.title("🟡 Painel de Inteligência - Mercado Livre")
    st.markdown("Conecte a conta do Mercado Livre para análise de Vendas, Estoque FULL e Campanhas via IA.")

    # Área de Configuração da API do Mercado Livre
    with st.expander("⚙️ Configuração de Integração (API Mercado Livre)", expanded=False):
        st.warning("Para ativar a integração real, crie um Aplicativo no painel **developers.mercadolivre.com.br** e insira as chaves nos Secrets do Streamlit.")
        st.code("""
# Adicione isto no painel Secrets do Streamlit:
ML_APP_ID = "seu_app_id_aqui"
ML_SECRET_KEY = "sua_secret_key_aqui"
        """, language="python")
        st.info("Status da Integração: Aguardando Autenticação OAuth2.0")

    st.markdown("---")
    st.subheader("📊 Dashboard Gerencial (Modo Simulação para Demonstração)")
    
    # Simulando os dados que viriam da API do Mercado Livre
    col_dash1, col_dash2, col_dash3 = st.columns(3)
    col_dash1.metric("Vendas (Últimos 30 dias)", "R$ 45.230,00", "+12.5% vs Mês Anterior")
    col_dash2.metric("Custo Mercado Ads", "R$ 3.150,00", "-5% vs Mês Anterior", delta_color="inverse")
    col_dash3.metric("Itens no FULL (Total)", "1.250 un", "Estoque Saudável")

    st.markdown("### 🧠 Análise Estratégica da IA (Gemini)")
    st.markdown("A Inteligência Artificial cruza os dados do seu FULL com as vendas para sugerir ações de precificação e redução de custos.")
    
    if st.button("Executar Auditoria de Conta via IA", type="primary"):
        if ia_configurada:
            with st.spinner("A IA está cruzando dados de anúncios, custos de Ads e tempo de estoque no FULL..."):
                try:
                    model = genai.GenerativeModel('gemini-2.5-flash')
                    
                    # Prompt simulando o envio dos dados extraídos da API do ML
                    prompt = """
                    Aja como um Assessor Estratégico de E-commerce. Recebemos os seguintes dados da API do Mercado Livre do cliente (CAASI):
                    - Anúncio A (Capa de Silicone): 800 unidades no FULL (paradas há 45 dias). Vendas caíram 20%. ACOS do Mercado Ads está em 35% (muito alto). Preço atual R$ 35,00.
                    - Anúncio B (Lanterna Tática): 50 unidades no FULL. Vendendo 10 por dia. Preço R$ 80,00.
                    
                    Faça uma análise rápida, direta e focada em lucro. Sugira:
                    1. Ação para o Anúncio A (baixar preço? pausar ads?).
                    2. Ação para o Anúncio B (risco de ruptura de estoque).
                    3. Alerta geral sobre taxas de armazenagem do FULL.
                    """
                    
                    resposta = model.generate_content(prompt)
                    st.success("Auditoria Concluída!")
                    st.write(resposta.text)
                except Exception as e:
                    st.error(f"Erro ao gerar análise da IA: {e}")
        else:
            st.warning("IA não configurada no Secrets. O cérebro do sistema está desativado.")
