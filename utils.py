import io
import re
import pdfplumber
from datetime import datetime, timedelta
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.pdfmetrics import stringWidth

# Lista de UFs
LISTA_UFS = [
    "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS",
    "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC",
    "SP", "SE", "TO"
]

MESES = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
    7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}


def remover_fones(texto):
    """
    Remove qualquer trecho que pareça um telefone com DDD (XX)
    Isso impede que o telefone seja confundido com CEP.
    """
    if not texto: return ""
    # Remove padrões como (11) 99999-9999
    texto_limpo = re.sub(r'\(\d{2}\)[\s\d-]*', '', texto)
    # Remove qualquer linha que tenha parenteses soltos
    texto_limpo = texto_limpo.replace('(', '').replace(')', '')
    return texto_limpo


def limpar_lixo_nome(texto):
    if not texto: return ""
    # Remove Fone explicitamente
    texto = texto.split("Fone")[0]
    texto = texto.split("Endereço")[0]
    texto = texto.split("Endereço:")[0]
    texto = texto.split("E-mail")[0]
    return texto.strip()


def limpar_lixo_endereco(texto):
    if not texto: return ""
    lixo = [
        "Somente Cx Correio", "Somente Cx", "Somente", "Correio",
        "620 Dia para Sedex", "Dia para Sedex", "Dia para", "620 Dia",
        "Peso Máximo", "Peso do pedido", "Caixa Máxima", "Cx=5",
        "Obs.", "Não", "620", "Visitante"
    ]
    for item in lixo:
        texto = re.sub(re.escape(item), '', texto, flags=re.IGNORECASE)

    texto = re.sub(r'\s\d+$', '', texto)
    texto = texto.replace(":", "")
    return texto.strip()


def limpar_nome_produto(texto):
    """Remove unidades e sujeira do nome do produto"""
    if not texto: return ""
    termos_unidade = [r"\bUnid\.\b", r"\bUnid\b", r"\bPct\b", r"\bL\b", r"\bUnidade\b", r"\buni\b", r"\bG\b"]
    novo_texto = texto
    for termo in termos_unidade:
        novo_texto = re.sub(termo, "", novo_texto, flags=re.IGNORECASE)
    return re.sub(r'\s+', ' ', novo_texto).strip(" .")


def extrair_valor_sequencial(texto_completo, chave, indice_desejado, paradas):
    txt = texto_completo.replace('\n', ' ')

    # Acha todas as ocorrências
    ocorrencias = [m.end() for m in re.finditer(re.escape(chave), txt, re.IGNORECASE)]
    if len(ocorrencias) <= indice_desejado: return ""

    inicio = ocorrencias[indice_desejado]
    texto_corte = txt[inicio:]

    menor_indice = len(texto_corte)
    for parada in paradas:
        idx = texto_corte.lower().find(parada.lower())
        if idx != -1 and idx < menor_indice: menor_indice = idx

    valor_bruto = texto_corte[:menor_indice]
    return limpar_lixo_endereco(valor_bruto) if "Endereço" in chave else limpar_lixo_nome(valor_bruto)


def separar_cidade_uf(texto_cidade_uf):
    cid = texto_cidade_uf;
    uf = ""
    for estado in LISTA_UFS:
        if re.search(r'[- ]' + estado + r'\b', texto_cidade_uf, re.IGNORECASE):
            uf = estado;
            break
    if uf: cid = re.sub(r'[- ]*\b' + uf + r'\b', '', texto_cidade_uf, flags=re.IGNORECASE)
    return cid.strip().title(), uf


def extrair_dados_pedido(pdf_file):
    dados = {
        "remetente_nome": "", "remetente_end": "", "remetente_cidade": "", "remetente_uf": "", "remetente_cep": "",
        "destinatario_nome": "", "destinatario_end": "", "destinatario_cidade": "", "destinatario_uf": "",
        "destinatario_cep": "",
        "itens": [], "peso_pedido": "", "numero_pedido": "", "data_dia": "", "data_mes": "", "data_ano": ""
    }

    agora = datetime.now() - timedelta(hours=3)
    dados["data_dia"] = str(agora.day);
    dados["data_mes"] = MESES[agora.month];
    dados["data_ano"] = str(agora.year)

    with pdfplumber.open(pdf_file) as pdf:
        page = pdf.pages[0]
        text = page.extract_text() or ""

        # --- 1. LIMPEZA PRÉVIA DE TELEFONES ---
        # Removemos qualquer coisa entre parenteses para não confundir com CEP
        text_sem_fone = remover_fones(text)

        paradas = ["Endereço:", "Bairro:", "Cidade-UF:", "CEP:", "Fone:", "E-mail:", "Peso", "Caixa", "Dia para",
                   "Visitante", "Preso"]

        # 2. NOMES (Sequencial)
        dados["remetente_nome"] = extrair_valor_sequencial(text, "Nome:", 0, paradas)
        dados["destinatario_nome"] = extrair_valor_sequencial(text, "Nome:", 1, paradas)

        # 3. ENDEREÇOS (Sequencial)
        dados["destinatario_end"] = extrair_valor_sequencial(text, "Endereço:", 0, paradas)
        dados["remetente_end"] = extrair_valor_sequencial(text, "Endereço:", 1, paradas)

        # 4. CIDADE-UF
        raw_cid_dest = extrair_valor_sequencial(text, "Cidade-UF:", 0, paradas)
        dados["destinatario_cidade"], dados["destinatario_uf"] = separar_cidade_uf(raw_cid_dest)

        raw_cid_rem = extrair_valor_sequencial(text, "Cidade-UF:", 1, paradas)
        dados["remetente_cidade"], dados["remetente_uf"] = separar_cidade_uf(raw_cid_rem)

        # 5. CEP (USANDO O TEXTO LIMPO SEM FONE)
        # Procura apenas no texto que já teve os telefones removidos
        ceps = re.findall(r"(\d{5}[-]?\d{3})", text_sem_fone)

        if len(ceps) >= 1: dados["destinatario_cep"] = ceps[0]
        if len(ceps) >= 2: dados["remetente_cep"] = ceps[1]

        # 6. DADOS GERAIS
        m_num = re.search(r"Pedido N°:\s*(\d+)", text)
        if m_num: dados["numero_pedido"] = m_num.group(1)
        m_peso = re.search(r"Peso do pedido:\s*(.*?)\n", text)
        if m_peso: dados["peso_pedido"] = m_peso.group(1).strip()

        # 7. ITENS
        for table in page.extract_tables():
            if not table: continue
            for row in table:
                if not row or not row[0]: continue
                if str(row[0]).strip().replace(',', '').isdigit():
                    try:
                        qtd = str(row[0]).strip()
                        partes_nome = []
                        for i in range(1, 4):
                            if i < len(row) and row[i]:
                                txt = str(row[i]).strip()
                                if txt.upper() in ["ALIMENTOS", "HIGIENE", "LIMPEZA", "VESTUÁRIOS", "DIVERSOS",
                                                   "CIGARROS", "PAPELARIA"]: break
                                if re.match(r'^\d+,\d+$', txt): break
                                partes_nome.append(txt)

                        nome_bruto = " ".join(partes_nome).replace('\n', ' ').strip()
                        nome_limpo = limpar_nome_produto(nome_bruto)
                        if nome_limpo:
                            dados["itens"].append({"qtd": qtd, "nome": nome_limpo, "valor": ""})
                    except:
                        pass
    return dados


def desenhar_texto_quebrado(can, texto, x_ini, y_ini, x_limite, x_retorno, y_retorno):
    fonte = "Helvetica"
    tamanho = 10
    largura_total = stringWidth(texto, fonte, tamanho)
    largura_disponivel = x_limite - x_ini

    if largura_total <= largura_disponivel:
        can.drawString(x_ini, y_ini, texto)
    else:
        palavras = texto.split(" ")
        linha1 = ""
        linha2 = ""
        for palavra in palavras:
            teste_linha = linha1 + " " + palavra if linha1 else palavra
            if stringWidth(teste_linha, fonte, tamanho) < largura_disponivel:
                linha1 = teste_linha
            else:
                idx = len(linha1.split(" "))
                linha2 = " ".join(palavras[idx:])
                if not linha2: linha2 = palavra
                break
        can.drawString(x_ini, y_ini, linha1)
        can.drawString(x_retorno, y_retorno, linha2)


def gerar_declaracao_pdf(dados, template_path, ajustes=None):
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)
    can.setFont("Helvetica", 10)

    c = {
        "rem_nome_x": 45, "rem_nome_y": 726,
        "rem_end_x": 63, "rem_end_y": 709,
        "rem_limit_x": 280, "rem_ret_x": 17, "rem_ret_y": 692,
        "rem_cid_x": 50, "rem_cid_y": 673, "rem_uf_x": 245, "rem_uf_y": 673, "rem_cep_x": 35, "rem_cep_y": 655,

        "dest_nome_x": 325, "dest_nome_y": 726,
        "dest_end_x": 343, "dest_end_y": 709,
        "dest_limit_x": 570, "dest_ret_x": 297, "dest_ret_y": 692,
        "dest_cid_x": 330, "dest_cid_y": 673, "dest_uf_x": 540, "dest_uf_y": 673, "dest_cep_x": 325, "dest_cep_y": 655,

        "itens_y_inicio": 600, "itens_espacamento": 17, "item_desc_x": 60, "item_qtd_x": 405,
        "item_desc_x_2": 230, "item_qtd_x_2": 450, "sep_desc_x": 220, "sep_qtd_x": 430,
        "peso_x": 395, "peso_y": 327, "data_y": 190, "data_cidade_x": 50, "data_dia_x": 145, "data_mes_x": 215,
        "data_ano_x": 325
    }
    if ajustes: c.update(ajustes)

    can.drawString(c["rem_nome_x"], c["rem_nome_y"], str(dados["remetente_nome"])[:50])
    desenhar_texto_quebrado(can, str(dados["remetente_end"]), c["rem_end_x"], c["rem_end_y"], c["rem_limit_x"],
                            c["rem_ret_x"], c["rem_ret_y"])
    can.drawString(c["rem_cid_x"], c["rem_cid_y"], str(dados["remetente_cidade"]))
    can.drawString(c["rem_uf_x"], c["rem_uf_y"], str(dados["remetente_uf"]))
    can.drawString(c["rem_cep_x"], c["rem_cep_y"], str(dados["remetente_cep"]))

    if dados["destinatario_nome"]:
        can.drawString(c["dest_nome_x"], c["dest_nome_y"], str(dados["destinatario_nome"])[:50])
        desenhar_texto_quebrado(can, str(dados["destinatario_end"]), c["dest_end_x"], c["dest_end_y"],
                                c["dest_limit_x"], c["dest_ret_x"], c["dest_ret_y"])
        can.drawString(c["dest_cid_x"], c["dest_cid_y"], str(dados["destinatario_cidade"]))
        can.drawString(c["dest_uf_x"], c["dest_uf_y"], str(dados["destinatario_uf"]))
        can.drawString(c["dest_cep_x"], c["dest_cep_y"], str(dados["destinatario_cep"]))

    if len(dados["itens"]) > 15:
        can.setLineWidth(0.5)
        h = 15 * c["itens_espacamento"]
        can.line(c["sep_desc_x"], c["itens_y_inicio"] + 10, c["sep_desc_x"], c["itens_y_inicio"] - h + 5)
        can.line(c["sep_qtd_x"], c["itens_y_inicio"] + 10, c["sep_qtd_x"], c["itens_y_inicio"] - h + 5)

    can.setFont("Helvetica", 9)
    for i, item in enumerate(dados["itens"]):
        if i >= 30: break
        if i < 15:
            x_n = c["item_desc_x"];
            x_q = c["item_qtd_x"];
            y = c["itens_y_inicio"] - (i * c["itens_espacamento"]);
            lim = 45
        else:
            x_n = c["item_desc_x_2"];
            x_q = c["item_qtd_x_2"];
            y = c["itens_y_inicio"] - ((i - 15) * c["itens_espacamento"]);
            lim = 30
        can.drawString(x_n, y, item["nome"][:lim])
        can.drawString(x_q, y, item["qtd"])

    can.setFont("Helvetica-Bold", 10)
    can.drawString(c["peso_x"], c["peso_y"], str(dados["peso_pedido"]))
    can.setFont("Helvetica", 10)
    can.drawString(c["data_cidade_x"], c["data_y"], "São Paulo")
    can.drawString(c["data_dia_x"], c["data_y"], str(dados["data_dia"]))
    can.drawString(c["data_mes_x"], c["data_y"], str(dados["data_mes"]))
    can.drawString(c["data_ano_x"], c["data_y"], str(dados["data_ano"]))

    can.save()
    packet.seek(0)
    try:
        new_pdf = PdfReader(packet);
        existing_pdf = PdfReader(template_path);
        output = PdfWriter()
        page = existing_pdf.pages[0];
        page.merge_page(new_pdf.pages[0]);
        output.add_page(page)
        output_stream = io.BytesIO();
        output.write(output_stream);
        output_stream.seek(0)
        return output_stream
    except Exception:
        return None
