import streamlit as st
import re
from utils import extrair_dados_pedido, gerar_declaracao_pdf

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Jumbo CDP - Declaração Automática",
    page_icon="📦",
    layout="centered"
)

# --- ESTILO CSS (Para deixar o botão gigante e centralizado) ---
st.markdown("""
    <style>
    div.stButton > button {
        width: 100%;
        height: 3em;
        font-size: 20px;
        font-weight: bold;
    }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    </style>
""", unsafe_allow_html=True)

# --- CABEÇALHO ---
st.title("📦 Gerador de Declaração")
st.caption("Sistema Automatizado Jumbo CDP")

# --- UPLOAD ---
uploaded_file = st.file_uploader("Arraste o Pedido PDF aqui", type="pdf", label_visibility="collapsed")
TEMPLATE_FILENAME = "Formulario Declaracao de Conteudo - A4.pdf"

if uploaded_file:
    try:
        # --- PROCESSAMENTO AUTOMÁTICO (ZERO CLIQUES) ---
        with st.spinner("⚙️ Processando dados e gerando documento..."):

            # 1. Extração
            dados = extrair_dados_pedido(uploaded_file)

            # 2. Geração do PDF
            pdf_final = gerar_declaracao_pdf(dados, TEMPLATE_FILENAME)

            # 3. Definição do Nome do Arquivo (Baseado no upload)
            nome_original = uploaded_file.name
            match_numero = re.search(r"(\d+)", nome_original)

            if match_numero:
                num_pedido = match_numero.group(1)
            else:
                num_pedido = dados.get('numero_pedido', 'S_NUMERO')

            nome_download = f"Pedido_{num_pedido}.pdf"

        # --- ÁREA DE DOWNLOAD (EM DESTAQUE) ---
        if pdf_final:
            st.success("✅ Documento pronto!")

            # Botão Principal - Ocupa largura total
            st.download_button(
                label=f"⬇️ BAIXAR DECLARAÇÃO ({num_pedido})",
                data=pdf_final,
                file_name=nome_download,
                mime="application/pdf",
                type="primary"
            )

            st.markdown("---")

            # --- RESUMO VISUAL (Para conferência rápida) ---
            st.subheader("📋 Resumo do Processamento")

            col1, col2 = st.columns(2)

            with col1:
                st.markdown("### 📤 Remetente")
                st.write(f"**Nome:** {dados.get('remetente_nome')}")
                st.write(f"**Endereço:** {dados.get('remetente_end')}")
                st.write(f"**Local:** {dados.get('remetente_cidade')} - {dados.get('remetente_uf')}")
                st.write(f"**CEP:** {dados.get('remetente_cep')}")

            with col2:
                st.markdown("### 📥 Destinatário")
                st.write(f"**Nome:** {dados.get('destinatario_nome')}")
                st.write(f"**Endereço:** {dados.get('destinatario_end')}")
                st.write(f"**Local:** {dados.get('destinatario_cidade')} - {dados.get('destinatario_uf')}")
                st.write(f"**CEP:** {dados.get('destinatario_cep')}")

            # Dados Extras (Discretos)
            with st.expander(f"Ver {len(dados['itens'])} Itens e Detalhes"):
                c_a, c_b = st.columns(2)
                c_a.write(f"**Peso Total:** {dados.get('peso_pedido')}")
                c_b.write(f"**Data:** {dados.get('data_dia')}/{dados.get('data_mes')}/{dados.get('data_ano')}")
                st.table(dados["itens"])

        else:
            st.error("⚠️ Erro crítico: O arquivo modelo (template) não foi encontrado na pasta.")

    except Exception as e:
        st.error(f"❌ Ocorreu um erro ao ler o arquivo: {e}")
else:
    # Espaço vazio clean quando não tem arquivo
    st.info("Aguardando arquivo...")
