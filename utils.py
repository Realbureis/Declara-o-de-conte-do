import io
import re
import pdfplumber
from datetime import datetime, timedelta  # <--- Adicionado timedelta
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

# Lista de UFs do Brasil
LISTA_UFS = [
    "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS",
    "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC",
    "SP", "SE", "TO"
]

MESES = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
    7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}


def limpar_lixo(texto):
    """Remove caracteres indesejados e palavras proibidas"""
    if not texto: return ""

    # Remove frases que costumam sujar os campos
    lixo = [
        "Somente Cx Correio", "Somente Cx", "Somente", "Correio",
        "620 Dia para Sedex", "Dia para Sedex", "Dia para", "620 Dia",
        "Peso Máximo", "Peso do pedido", "Caixa Máxima", "Cx=5",
        "Obs.", "Não", "620", "Visitante"
    ]
    for item in lixo:
        texto = re.sub(re.escape(item), '', texto, flags=re.IGNORECASE)

    # Remove pontuação solta
    texto = texto.replace(":", "")
    texto = texto.strip(" :-,.")

    return re.sub(r'\s+', ' ', texto).strip()


def extrair_valor_sequencial(texto_completo, chave, indice_desejado, paradas):
    """
    Procura a N-ésima ocorrência de uma chave e pega o texto até a próxima parada.
    """
    txt = texto_completo.replace('\n', ' ')

    ocorrencias = [m.end() for m in re.finditer(re.escape(chave), txt, re.IGNORECASE)]

    if len(ocorrencias) <= indice_desejado:
        return ""

    inicio = ocorrencias[indice_desejado]
    texto_corte = txt[inicio:]

    menor_indice = len(texto_corte)

    for parada in paradas:
        idx = texto_corte.lower().find(parada.lower())
        if idx != -1 and idx < menor_indice:
            menor_indice = idx

    valor_bruto = texto_corte[:menor_indice]
    return limpar_lixo(valor_bruto)


def separar_cidade_uf(texto_cidade_uf):
    """Separa 'Hortolândia - SP' em Cidade e UF"""
    cid = texto_cidade_uf
    uf = ""
    for estado in LISTA_UFS:
        if re.search(r'[- ]' + estado + r'\b', texto_cidade_uf, re.IGNORECASE):
            uf = estado
            break
    if uf:
        cid = re.sub(r'[- ]*\b' + uf + r'\b', '', texto_cidade_uf, flags=re.IGNORECASE)
    return cid.strip().title(), uf


def extrair_dados_pedido(pdf_file):
    dados = {
        "remetente_nome": "", "remetente_end": "", "remetente_cidade": "", "remetente_uf": "", "remetente_cep": "",
        "destinatario_nome": "", "destinatario_end": "", "destinatario_cidade": "", "destinatario_uf": "",
        "destinatario_cep": "",
        "itens": [], "peso_pedido": "", "numero_pedido": "", "data_dia": "", "data_mes": "", "data_ano": ""
    }

    # --- DATA AUTOMÁTICA (AGORA - 3h) ---
    agora = datetime.now() - timedelta(hours=3)
    dados["data_dia"] = str(agora.day)
    dados["data_mes"] = MESES[agora.month]
    dados["data_ano"] = str(agora.year)

    with pdfplumber.open(pdf_file) as pdf:
        page = pdf.pages[0]
        text = page.extract_text() or ""
        text_flat = text.replace('\n', '  ')

        # === LISTA DE PARADAS ===
        paradas_gerais = [
            "Endereço:", "Bairro:", "Cidade-UF:", "CEP:", "Fone:",
            "E-mail:", "Peso", "Caixa", "Dia para", "Visitante", "Preso"
        ]

        # 1. NOMES
        dados["remetente_nome"] = extrair_valor_sequencial(text, "Nome:", 0, paradas_gerais)
        dados["destinatario_nome"] = extrair_valor_sequencial(text, "Nome:", 1, paradas_gerais)

        # 2. ENDEREÇOS
        dados["destinatario_end"] = extrair_valor_sequencial(text, "Endereço:", 0, paradas_gerais)
        dados["remetente_end"] = extrair_valor_sequencial(text, "Endereço:", 1, paradas_gerais)

        # 3. CIDADE-UF
        raw_cid_dest = extrair_valor_sequencial(text, "Cidade-UF:", 0, paradas_gerais)
        dados["destinatario_cidade"], dados["destinatario_uf"] = separar_cidade_uf(raw_cid_dest)

        raw_cid_rem = extrair_valor_sequencial(text, "Cidade-UF:", 1, paradas_gerais)
        dados["remetente_cidade"], dados["remetente_uf"] = separar_cidade_uf(raw_cid_rem)

        # 4. CEP
        todos_ceps = re.findall(r"(\d{5}[-]?\d{3})", text)
        if len(todos_ceps) >= 1: dados["destinatario_cep"] = todos_ceps[0]
        if len(todos_ceps) >= 2: dados["remetente_cep"] = todos_ceps[1]

        # === DADOS GERAIS ===
        m_num = re.search(r"Pedido N°:\s*(\d+)", text)
        if m_num: dados["numero_pedido"] = m_num.group(1)
        m_peso = re.search(r"Peso do pedido:\s*(.*?)\n", text)
        if m_peso: dados["peso_pedido"] = m_peso.group(1).strip()

        # A data extraída do PDF foi removida daqui, usamos a data automática acima.

        # === ITENS ===
        for table in page.extract_tables():
            if not table: continue
            for row in table:
                if not row or not row[0]: continue
                col0 = str(row[0]).strip().replace(',', '')
                if col0.isdigit():
                    try:
                        qtd = str(row[0]).strip()
                        partes_nome = []
                        for i in range(1, 4):
                            if i < len(row) and row[i]:
                                txt = str(row[i]).strip()
                                if txt.upper() in ["ALIMENTOS", "HIGIENE", "LIMPEZA", "VESTUÁRIOS", "DIVERSOS"]: break
                                if re.match(r'^\d+,\d+$', txt): break
                                partes_nome.append(txt)
                        nome = " ".join(partes_nome).replace('\n', ' ').strip()
                        if nome: dados["itens"].append({"qtd": qtd, "nome": nome, "valor": ""})
                    except:
                        pass
    return dados


def gerar_declaracao_pdf(dados, template_path, ajustes=None):
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)
    can.setFont("Helvetica", 10)

    # === COORDENADAS ===
    coord = {
        "rem_nome_x": 45, "rem_nome_y": 726, "rem_end_x": 63, "rem_end_y": 709,
        "rem_cid_x": 50, "rem_cid_y": 673, "rem_uf_x": 245, "rem_uf_y": 673, "rem_cep_x": 35, "rem_cep_y": 655,
        "dest_nome_x": 325, "dest_nome_y": 726, "dest_end_x": 343, "dest_end_y": 709,
        "dest_cid_x": 330, "dest_cid_y": 673, "dest_uf_x": 540, "dest_uf_y": 673, "dest_cep_x": 325, "dest_cep_y": 655,
        "itens_y_inicio": 600, "itens_espacamento": 17, "item_desc_x": 65, "item_qtd_x": 405,
        "item_desc_x_2": 230, "item_qtd_x_2": 450, "sep_desc_x": 220, "sep_qtd_x": 430,

        "peso_x": 395, "peso_y": 327,
        "data_y": 190,
        "data_cidade_x": 50,  # <--- AJUSTADO PARA 50
        "data_dia_x": 145, "data_mes_x": 215, "data_ano_x": 325
    }

    if ajustes: coord.update(ajustes)

    # Remetente
    can.drawString(coord["rem_nome_x"], coord["rem_nome_y"], str(dados["remetente_nome"])[:50])
    can.drawString(coord["rem_end_x"], coord["rem_end_y"], str(dados["remetente_end"])[:60])
    can.drawString(coord["rem_cid_x"], coord["rem_cid_y"], str(dados["remetente_cidade"]))
    can.drawString(coord["rem_uf_x"], coord["rem_uf_y"], str(dados["remetente_uf"]))
    can.drawString(coord["rem_cep_x"], coord["rem_cep_y"], str(dados["remetente_cep"]))

    # Destinatário
    if dados["destinatario_nome"]:
        can.drawString(coord["dest_nome_x"], coord["dest_nome_y"], str(dados["destinatario_nome"])[:50])
        can.drawString(coord["dest_end_x"], coord["dest_end_y"], str(dados["destinatario_end"])[:60])
        can.drawString(coord["dest_cid_x"], coord["dest_cid_y"], str(dados["destinatario_cidade"]))
        can.drawString(coord["dest_uf_x"], coord["dest_uf_y"], str(dados["destinatario_uf"]))
        can.drawString(coord["dest_cep_x"], coord["dest_cep_y"], str(dados["destinatario_cep"]))

    # Itens
    if len(dados["itens"]) > 15:
        can.setLineWidth(0.5)
        h = 15 * coord["itens_espacamento"]
        can.line(coord["sep_desc_x"], coord["itens_y_inicio"] + 10, coord["sep_desc_x"],
                 coord["itens_y_inicio"] - h + 5)
        can.line(coord["sep_qtd_x"], coord["itens_y_inicio"] + 10, coord["sep_qtd_x"], coord["itens_y_inicio"] - h + 5)

    can.setFont("Helvetica", 9)
    for i, item in enumerate(dados["itens"]):
        if i >= 30: break
        if i < 15:
            x_n = coord["item_desc_x"];
            x_q = coord["item_qtd_x"];
            y = coord["itens_y_inicio"] - (i * coord["itens_espacamento"]);
            lim = 45
        else:
            x_n = coord["item_desc_x_2"];
            x_q = coord["item_qtd_x_2"];
            y = coord["itens_y_inicio"] - ((i - 15) * coord["itens_espacamento"]);
            lim = 30
        can.drawString(x_n, y, item["nome"][:lim])
        can.drawString(x_q, y, item["qtd"])

    # Rodapé
    can.setFont("Helvetica-Bold", 10)
    can.drawString(coord["peso_x"], coord["peso_y"], str(dados["peso_pedido"]))
    can.setFont("Helvetica", 10)
    can.drawString(coord["data_cidade_x"], coord["data_y"], "São Paulo")
    can.drawString(coord["data_dia_x"], coord["data_y"], str(dados["data_dia"]))
    can.drawString(coord["data_mes_x"], coord["data_y"], str(dados["data_mes"]))
    can.drawString(coord["data_ano_x"], coord["data_y"], str(dados["data_ano"]))

    can.save()
    packet.seek(0)
    try:
        new_pdf = PdfReader(packet)
        existing_pdf = PdfReader(template_path)
        output = PdfWriter()
        page = existing_pdf.pages[0]
        page.merge_page(new_pdf.pages[0])
        output.add_page(page)
        output_stream = io.BytesIO()
        output.write(output_stream)
        output_stream.seek(0)
        return output_stream
    except Exception:
        return None
