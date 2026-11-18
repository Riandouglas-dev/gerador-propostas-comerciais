import os
import sys
import time
import threading
import urllib.parse
import webbrowser
import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd
import qrcode
from reportlab.lib.pagesizes import letter
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, white

def obter_diretorio_base():
    """Retorna o diretório onde o executável (.exe) ou script (.py) está localizado."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = obter_diretorio_base()

# Configurações de Arquivos (sempre na mesma pasta do .exe)
PLANILHA_PADRAO_PATH = os.path.join(BASE_DIR, "modelo_planilha_servicos.xlsx")
PDF_PROPOSTA_PATH = os.path.join(BASE_DIR, "Proposta_Servicos_Empresa.pdf")
CONTADOR_PATH = os.path.join(BASE_DIR, "contador.txt")
QR_CODE_PATH = os.path.join(BASE_DIR, "qrcode_site.png")

def garantir_planilha_modelo(caminho_planilha):
    """Cria a planilha de exemplo caso o arquivo padrão não exista."""
    if not os.path.exists(caminho_planilha):
        dados_iniciais = {
            "Serviço": [
                "Troca do filtro diesel de alta eficiência", 
                "Troca do óleo lubrificante sintético (15W40)", 
                "Limpeza química e descontaminação de tanque de combustível", 
                "Troca do filtro lubrificante primário"
            ],
            "Custo_Base": [150.00, 300.00, 500.00, 120.00]
        }
        pd.DataFrame(dados_iniciais).to_excel(caminho_planilha, index=False)

garantir_planilha_modelo(PLANILHA_PADRAO_PATH)

def formatar_moeda_br(valor):
    """Formata um float no padrão monetário brasileiro (ex: 1.250,50)."""
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def numero_por_extenso_reais(valor):
    """Converte um valor numérico (float) para moeda em português por extenso sem bibliotecas externas."""
    unidades = ["", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove"]
    especiais = ["dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove"]
    dezenas = ["", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa"]
    centenas = ["", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos"]

    def grupo_ate_999(n):
        if n == 0:
            return ""
        if n == 100:
            return "cem"
        partes = []
        c = n // 100
        d = (n % 100) // 10
        u = n % 10
        if c > 0:
            partes.append(centenas[c])
        if d == 1:
            partes.append(especiais[u])
        else:
            if d > 1:
                partes.append(dezenas[d])
            if u > 0:
                partes.append(unidades[u])
        return " e ".join(partes)

    valor_arredondado = round(float(valor), 2)
    inteiro = int(valor_arredondado)
    centavos = int(round((valor_arredondado - inteiro) * 100))

    resultado = []

    if inteiro > 0:
        milhoes = inteiro // 1_000_000
        milhares = (inteiro % 1_000_000) // 1000
        resto = inteiro % 1000

        partes_inteiro = []
        if milhoes > 0:
            txt_m = grupo_ate_999(milhoes)
            partes_inteiro.append(f"{txt_m} {'milhão' if milhoes == 1 else 'milhões'}")
        if milhares > 0:
            txt_k = "mil" if milhares == 1 else f"{grupo_ate_999(milhares)} mil"
            partes_inteiro.append(txt_k)
        if resto > 0:
            txt_r = grupo_ate_999(resto)
            if partes_inteiro and (resto < 100 or resto % 100 == 0):
                partes_inteiro.append(f"e {txt_r}")
            else:
                partes_inteiro.append(txt_r)

        texto_reais = " ".join(partes_inteiro)
        if milhoes > 0 and milhares == 0 and resto == 0:
            texto_reais += " de"
        moeda = "real" if inteiro == 1 else "reais"
        resultado.append(f"{texto_reais} {moeda}")

    if centavos > 0:
        txt_c = grupo_ate_999(centavos)
        moeda_c = "centavo" if centavos == 1 else "centavos"
        resultado.append(f"{txt_c} {moeda_c}")

    if not resultado:
        return "zero reais"

    return " e ".join(resultado)

def obter_proximo_numero_proposta():
    if not os.path.exists(CONTADOR_PATH):
        with open(CONTADOR_PATH, "w", encoding="utf-8") as f:
            f.write("164")
    with open(CONTADOR_PATH, "r", encoding="utf-8") as f:
        conteudo = f.read().strip()
    ultimo_numero = int(conteudo) if conteudo.isdigit() else 0
    proximo_numero = ultimo_numero + 1
    with open(CONTADOR_PATH, "w", encoding="utf-8") as f:
        f.write(str(proximo_numero))
    return proximo_numero

def validar_e_formatar_whatsapp(whatsapp_input):
    """
    Valida e padroniza o número de telefone brasileiro para o padrão internacional (55 + DDD + Número).
    Retorna uma tupla: (numero_internacional, numero_exibicao)
    """
    digitos = "".join([c for c in str(whatsapp_input) if c.isdigit()])

    if digitos.startswith("55") and len(digitos) in (12, 13):
        digitos_locais = digitos[2:]
    else:
        digitos_locais = digitos

    if len(digitos_locais) not in (10, 11):
        raise ValueError(
            f"Número de WhatsApp inválido: '{whatsapp_input}'.\n"
            "Informe o DDD + número (ex: (83) 99999-9999)."
        )

    ddd = digitos_locais[:2]
    numero = digitos_locais[2:]

    if ddd.startswith("0"):
        raise ValueError(f"DDD inválido: '{ddd}'. Não inclua o zero inicial no DDD.")

    numero_internacional = f"55{ddd}{numero}"
    if len(numero) == 9:
        numero_exibicao = f"({ddd}) {numero[:5]}-{numero[5:]}"
    else:
        numero_exibicao = f"({ddd}) {numero[:4]}-{numero[4:]}"

    return numero_internacional, numero_exibicao

def desenhar_capa_executiva(c, width, height, nome_cliente, whatsapp_formatado, num_proposta, ano_atual, data_formatada):
    """Desenha uma capa corporativa de alto impacto na Página 1 do PDF."""
    cor_navy = HexColor("#0F172A")
    cor_navy_mid = HexColor("#1E293B")
    cor_accent = HexColor("#2563EB")
    cor_accent_light = HexColor("#60A5FA")
    cor_texto_suave = HexColor("#94A3B8")
    cor_borda_escura = HexColor("#334155")

    c.setFillColor(cor_navy)
    c.rect(0, 0, width, height, fill=True, stroke=False)

    c.saveState()
    c.setFillColor(cor_navy_mid)
    c.circle(width - 40, height - 60, 210, fill=True, stroke=False)
    c.setStrokeColor(cor_borda_escura)
    c.setLineWidth(1)
    c.circle(width - 40, height - 60, 260, fill=False, stroke=True)
    c.restoreState()

    c.setFillColor(cor_accent)
    c.rect(0, 0, 14, height, fill=True, stroke=False)

    margem_x = 58
    largura_util = width - margem_x - 48

    topo_y = height - 95
    c.setFillColor(cor_accent)
    c.roundRect(margem_x, topo_y - 6, 42, 42, 8, fill=True, stroke=False)
    c.setFillColor(white)
    c.circle(margem_x + 21, topo_y + 15, 9, fill=True, stroke=False)
    c.setFillColor(cor_navy)
    c.circle(margem_x + 21, topo_y + 15, 4, fill=True, stroke=False)

    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 22)
    c.drawString(margem_x + 54, topo_y + 16, "EMPRESA")
    c.setFillColor(cor_texto_suave)
    c.setFont("Helvetica", 9.5)
    c.drawString(margem_x + 54, topo_y + 2, "ENGENHARIA & MANUTENÇÃO INDUSTRIAL")

    badge_w = 165
    badge_h = 46
    badge_x = width - 48 - badge_w
    badge_y = topo_y - 8
    c.setFillColor(cor_navy_mid)
    c.setStrokeColor(cor_borda_escura)
    c.setLineWidth(1)
    c.roundRect(badge_x, badge_y, badge_w, badge_h, 6, fill=True, stroke=True)

    c.setFillColor(cor_accent_light)
    c.setFont("Helvetica-Bold", 7.5)
    c.drawCentredString(badge_x + badge_w / 2, badge_y + 30, "DOCUMENTO OFICIAL")
    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(badge_x + badge_w / 2, badge_y + 12, f"Nº {num_proposta}/{ano_atual}")

    centro_y = height * 0.56
    c.setFillColor(cor_accent)
    c.rect(margem_x, centro_y + 68, 56, 4, fill=True, stroke=False)

    c.setFillColor(cor_accent_light)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(margem_x, centro_y + 46, "APRESENTAÇÃO TÉCNICA E ORÇAMENTÁRIA")

    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 38)
    c.drawString(margem_x, centro_y + 4, "PROPOSTA")
    c.drawString(margem_x, centro_y - 38, "COMERCIAL")

    c.setFillColor(cor_texto_suave)
    c.setFont("Helvetica", 12)
    c.drawString(margem_x, centro_y - 72, "Soluções Especializadas em Manutenção Preventiva, Corretiva")
    c.drawString(margem_x, centro_y - 89, "e Gestão de Eficiência Operacional")

    card_y = 145
    card_h = 115
    c.setFillColor(cor_navy_mid)
    c.setStrokeColor(cor_borda_escura)
    c.setLineWidth(1)
    c.roundRect(margem_x, card_y, largura_util, card_h, 8, fill=True, stroke=True)

    c.setFillColor(cor_accent)
    c.roundRect(margem_x, card_y, 5, card_h, 2, fill=True, stroke=False)

    c.setFillColor(cor_accent_light)
    c.setFont("Helvetica-Bold", 8.5)
    c.drawString(margem_x + 22, card_y + 88, "PREPARADO EXCLUSIVAMENTE PARA")

    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 17)
    c.drawString(margem_x + 22, card_y + 64, nome_cliente)

    c.setStrokeColor(cor_borda_escura)
    c.setLineWidth(0.8)
    c.line(margem_x + 22, card_y + 48, margem_x + largura_util - 22, card_y + 48)

    c.setFillColor(cor_texto_suave)
    c.setFont("Helvetica-Bold", 7.5)
    c.drawString(margem_x + 22, card_y + 30, "CONTATO PRINCIPAL")
    c.drawString(margem_x + 210, card_y + 30, "DATA DE EMISSÃO")
    c.drawString(margem_x + 365, card_y + 30, "VALIDADE")

    c.setFillColor(white)
    c.setFont("Helvetica", 9.5)
    c.drawString(margem_x + 22, card_y + 15, whatsapp_formatado)
    c.drawString(margem_x + 210, card_y + 15, data_formatada)
    c.drawString(margem_x + 365, card_y + 15, "10 dias corridos")

    c.setStrokeColor(cor_borda_escura)
    c.setLineWidth(0.8)
    c.line(margem_x, 82, margem_x + largura_util, 82)

    c.setFillColor(cor_texto_suave)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(margem_x, 62, "EMPRESA SOLUÇÕES EM ENGENHARIA LTDA")
    c.setFont("Helvetica", 8)
    c.drawString(margem_x, 49, "CNPJ: 00.000.000/0001-00  •  www.exemplo-empresa.com.br")
    c.drawRightString(margem_x + largura_util, 55, "CONFIDENCIAL • USO COMERCIAL")

    c.showPage()

def desenhar_cabecalho_moderno(c, width, height, num_proposta, ano_atual, data_formatada):
    """Desenha um cabeçalho editorial escuro com logotipo vetorial e badge da proposta."""
    cor_navy = HexColor("#0F172A")
    cor_accent = HexColor("#2563EB")
    cor_badge_bg = HexColor("#1E293B")
    cor_texto_suave = HexColor("#94A3B8")

    altura_header = 105
    c.setFillColor(cor_navy)
    c.rect(0, height - altura_header, width, altura_header, fill=True, stroke=False)

    c.setFillColor(cor_accent)
    c.rect(0, height - altura_header - 3, width, 3, fill=True, stroke=False)

    margem_x = 45
    logo_y = height - 62

    c.setFillColor(cor_accent)
    c.roundRect(margem_x, logo_y - 4, 32, 32, 6, fill=True, stroke=False)
    c.setFillColor(white)
    c.circle(margem_x + 16, logo_y + 12, 7, fill=True, stroke=False)
    c.setFillColor(cor_navy)
    c.circle(margem_x + 16, logo_y + 12, 3, fill=True, stroke=False)

    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(margem_x + 42, logo_y + 13, "EMPRESA")
    c.setFillColor(cor_texto_suave)
    c.setFont("Helvetica", 8)
    c.drawString(margem_x + 42, logo_y + 1, "ENGENHARIA & MANUTENÇÃO INDUSTRIAL")

    badge_w = 175
    badge_h = 54
    badge_x = width - margem_x - badge_w
    badge_y = height - 78

    c.setFillColor(cor_badge_bg)
    c.setStrokeColor(HexColor("#334155"))
    c.setLineWidth(0.8)
    c.roundRect(badge_x, badge_y, badge_w, badge_h, 6, fill=True, stroke=True)

    c.setFillColor(HexColor("#60A5FA"))
    c.setFont("Helvetica-Bold", 7.5)
    c.drawRightString(badge_x + badge_w - 12, badge_y + 38, "PROPOSTA COMERCIAL")

    c.setFillColor(white)
    c.setFont("Helvetica-Bold", 14)
    c.drawRightString(badge_x + badge_w - 12, badge_y + 21, f"Nº {num_proposta}/{ano_atual}")

    c.setFillColor(cor_texto_suave)
    c.setFont("Helvetica", 8)
    c.drawRightString(badge_x + badge_w - 12, badge_y + 8, f"Emissão: {data_formatada}")

def desenhar_assinatura_ficticia(c, x, y):
    """Desenha uma rubrica vetorial moderna para fins de portfólio."""
    c.saveState()
    c.setStrokeColor(HexColor("#1D4ED8"))
    c.setLineWidth(1.5)
    p = c.beginPath()
    p.moveTo(x + 8, y + 10)
    p.curveTo(x + 22, y + 32, x + 32, y + 4, x + 48, y + 24)
    p.curveTo(x + 62, y + 34, x + 74, y + 8, x + 94, y + 18)
    p.moveTo(x + 12, y + 12)
    p.lineTo(x + 106, y + 15)
    c.drawPath(p, stroke=1, fill=0)
    c.restoreState()

def disparar_whatsapp_web(telefone, nome_cliente, num_proposta, ano_atual, valor_formatado, valor_extenso):
    """Monta a mensagem comercial formatada e abre o WhatsApp Web automaticamente."""
    mensagem = (
        f"OLA, TUDO BEM? *{nome_cliente}*!\n\n"
        f"Segue a nossa *Proposta Comercial Nº {num_proposta}/{ano_atual}* referente aos serviços solicitados.\n\n"
        f"💰 *Valor Total:* R$ {valor_formatado} _({valor_extenso})_\n\n"
        f"📄 *Todos os detalhes técnicos, condições comerciais e prazos estão no documento PDF em anexo.*\n\n"
        f"Ficamos à disposição para qualquer dúvida!"
    )
    mensagem_codificada = urllib.parse.quote(mensagem)
    url_whatsapp = f"https://web.whatsapp.com/send?phone={telefone}&text={mensagem_codificada}"
    webbrowser.open(url_whatsapp)

def gerar_proposta_pdf_e_enviar(nome_cliente, whatsapp_input, caminho_planilha, enviar_zap, callback_status):
    """Executa o fluxo completo de leitura do Excel, geração do PDF com capa e disparo opcional no WhatsApp."""
    telefone, whatsapp_formatado = validar_e_formatar_whatsapp(whatsapp_input)

    callback_status("Lendo planilha de serviços...")
    try:
        df = pd.read_excel(caminho_planilha)
    except PermissionError:
        raise RuntimeError("Feche a planilha Excel antes de gerar a proposta!")
    except Exception as err:
        raise RuntimeError(f"Erro ao ler a planilha Excel: {err}")

    if "Serviço" not in df.columns or "Custo_Base" not in df.columns:
        raise ValueError("A planilha deve conter as colunas 'Serviço' e 'Custo_Base'.")

    df["Valor_Final"] = df["Custo_Base"] * 2.20 * 1.13
    total_geral = df["Valor_Final"].sum()
    total_formatado_br = formatar_moeda_br(total_geral)
    total_extenso = numero_por_extenso_reais(total_geral)
    num_proposta = obter_proximo_numero_proposta()
    ano_atual = time.strftime("%Y")
    data_atual_formatada = time.strftime("%d/%m/%Y")

    callback_status("Gerando QR Code institucional...")
    try:
        qr = qrcode.QRCode(version=1, box_size=10, border=1)
        qr.add_data("https://exemplo-empresa.com.br")
        qr.make(fit=True)
        img_qr = qr.make_image(fill_color="#0F172A", back_color="white")
        img_qr.save(QR_CODE_PATH)
    except Exception:
        pass

    callback_status("Renderizando PDF com Capa Executiva...")
    c = canvas.Canvas(PDF_PROPOSTA_PATH, pagesize=letter)
    width, height = letter
    margem_x = 45
    largura_util = width - (margem_x * 2)

    cor_navy = HexColor("#0F172A")
    cor_accent = HexColor("#2563EB")
    cor_texto_principal = HexColor("#1E293B")
    cor_texto_secundario = HexColor("#64748B")
    cor_card_bg = HexColor("#F8FAFC")
    cor_borda = HexColor("#E2E8F0")
    cor_alerta_bg = HexColor("#FFFBEB")
    cor_alerta_borda = HexColor("#F59E0B")
    cor_alerta_texto = HexColor("#92400E")

    # PÁGINA 1: Capa Executiva
    desenhar_capa_executiva(
        c=c,
        width=width,
        height=height,
        nome_cliente=nome_cliente,
        whatsapp_formatado=whatsapp_formatado,
        num_proposta=num_proposta,
        ano_atual=ano_atual,
        data_formatada=data_atual_formatada
    )

    # PÁGINA 2: Escopo Técnico e Investimentos
    desenhar_cabecalho_moderno(c, width, height, num_proposta, ano_atual, data_atual_formatada)

    y_card = height - 185
    altura_card = 58
    c.setFillColor(cor_card_bg)
    c.setStrokeColor(cor_borda)
    c.setLineWidth(0.8)
    c.roundRect(margem_x, y_card, largura_util, altura_card, 6, fill=True, stroke=True)

    c.setFillColor(cor_accent)
    c.roundRect(margem_x, y_card, 4, altura_card, 2, fill=True, stroke=False)

    c.setFillColor(cor_accent)
    c.setFont("Helvetica-Bold", 7.5)
    c.drawString(margem_x + 16, y_card + 41, "PREPARADO PARA")

    c.setFillColor(cor_navy)
    c.setFont("Helvetica-Bold", 11.5)
    c.drawString(margem_x + 16, y_card + 25, nome_cliente)

    c.setFillColor(cor_texto_secundario)
    c.setFont("Helvetica", 8.5)
    c.drawString(margem_x + 16, y_card + 11, "A/C: Gerência Administrativa e Operacional")

    c.setFillColor(cor_texto_secundario)
    c.setFont("Helvetica-Bold", 7.5)
    c.drawRightString(width - margem_x - 16, y_card + 41, "CONTATO & LOCALIDADE")

    c.setFillColor(cor_texto_principal)
    c.setFont("Helvetica", 8.5)
    c.drawRightString(width - margem_x - 16, y_card + 26, f"WhatsApp: {whatsapp_formatado}")
    c.drawRightString(width - margem_x - 16, y_card + 12, "Cidade Exemplo - UF")

    y_pos = y_card - 24
    c.setFillColor(cor_texto_principal)
    c.setFont("Helvetica", 9.5)
    c.drawString(margem_x, y_pos, "Apresentamos abaixo o escopo técnico e os investimentos previstos para manutenção preventiva e corretiva:")

    styles = getSampleStyleSheet()
    style_header_left = ParagraphStyle('HLeft', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8.5, textColor=white)
    style_header_center = ParagraphStyle('HCenter', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8.5, alignment=1, textColor=white)
    style_header_right = ParagraphStyle('HRight', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8.5, alignment=2, textColor=white)

    style_item_num = ParagraphStyle('ItemNum', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, alignment=1, textColor=cor_texto_secundario)
    style_item_desc = ParagraphStyle('ItemDesc', parent=styles['Normal'], fontName='Helvetica', fontSize=9.5, leading=13, textColor=cor_texto_principal)
    style_item_val = ParagraphStyle('ItemVal', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9.5, alignment=2, textColor=cor_navy)
    style_alerta = ParagraphStyle('TxtAlerta', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, leading=12, textColor=cor_alerta_texto)

    dados_tabela = [[
        Paragraph("ITEM", style_header_center), 
        Paragraph("DESCRIÇÃO DO ESCOPO TÉCNICO", style_header_left), 
        Paragraph("VALOR (R$)", style_header_right)
    ]]

    estilos_dinamicos = [
        ('BACKGROUND', (0, 0), (-1, 0), cor_navy),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
        ('LINEBELOW', (0, 0), (-1, 0), 2, cor_accent)
    ]

    linha_idx = 1
    item_contador = 1
    for ind, row in df.iterrows():
        texto_servico = str(row["Serviço"])
        valor_formatado = formatar_moeda_br(row["Valor_Final"])

        bg_cor = cor_card_bg if item_contador % 2 == 0 else white
        estilos_dinamicos.append(('BACKGROUND', (0, linha_idx), (-1, linha_idx), bg_cor))
        estilos_dinamicos.append(('LINEBELOW', (0, linha_idx), (-1, linha_idx), 0.5, cor_borda))
        estilos_dinamicos.append(('TOPPADDING', (0, linha_idx), (-1, linha_idx), 9))
        estilos_dinamicos.append(('BOTTOMPADDING', (0, linha_idx), (-1, linha_idx), 9))

        dados_tabela.append([
            Paragraph(f"{item_contador:02d}", style_item_num),
            Paragraph(texto_servico, style_item_desc),
            Paragraph(valor_formatado, style_item_val)
        ])
        linha_idx += 1
        item_contador += 1

        if "limpeza" in texto_servico.lower() and "tanque" in texto_servico.lower():
            texto_obs = "<b>NOTA TÉCNICA:</b> É obrigatória a substituição do óleo diesel contaminado por combustível novo. O fornecimento do insumo não está contemplado neste orçamento."
            estilos_dinamicos.append(('BACKGROUND', (0, linha_idx), (-1, linha_idx), cor_alerta_bg))
            estilos_dinamicos.append(('SPAN', (0, linha_idx), (-1, linha_idx)))
            estilos_dinamicos.append(('LINEBEFORE', (0, linha_idx), (0, linha_idx), 3, cor_alerta_borda))
            estilos_dinamicos.append(('LINEBELOW', (0, linha_idx), (-1, linha_idx), 0.5, cor_borda))
            estilos_dinamicos.append(('TOPPADDING', (0, linha_idx), (-1, linha_idx), 7))
            estilos_dinamicos.append(('BOTTOMPADDING', (0, linha_idx), (-1, linha_idx), 7))
            dados_tabela.append([Paragraph(texto_obs, style_alerta), "", ""])
            linha_idx += 1

    largura_colunas = [45, 377, 100]
    tabela_servicos = Table(dados_tabela, colWidths=largura_colunas)
    tabela_servicos.setStyle(TableStyle(estilos_dinamicos))

    _, tabela_height = tabela_servicos.wrap(largura_util, height)
    y_pos -= (tabela_height + 16)
    tabela_servicos.drawOn(c, margem_x, y_pos)

    altura_total_box = 52
    y_pos -= (altura_total_box + 16)

    c.setFillColor(HexColor("#EFF6FF"))
    c.setStrokeColor(HexColor("#BFDBFE"))
    c.setLineWidth(1)
    c.roundRect(margem_x, y_pos, largura_util, altura_total_box, 6, fill=True, stroke=True)

    c.setFillColor(cor_navy)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(margem_x + 16, y_pos + 31, "INVESTIMENTO TOTAL DA PROPOSTA")

    c.setFillColor(cor_texto_secundario)
    c.setFont("Helvetica", 8)
    c.drawString(margem_x + 16, y_pos + 14, "Valores expressos em Reais (BRL) com impostos inclusos.")

    c.setFillColor(cor_accent)
    c.setFont("Helvetica-Bold", 15)
    c.drawRightString(width - margem_x - 16, y_pos + 28, f"R$ {total_formatado_br}")

    c.setFillColor(cor_texto_principal)
    c.setFont("Helvetica-Oblique", 8.5)
    c.drawRightString(width - margem_x - 16, y_pos + 12, f"({total_extenso.capitalize()})")

    y_pos -= 24
    c.setFillColor(cor_navy)
    c.setFont("Helvetica-Bold", 9.5)
    c.drawString(margem_x, y_pos, "CONDIÇÕES COMERCIAIS E GARANTIAS")

    altura_cond = 46
    y_pos -= (altura_cond + 8)
    gap = 10
    largura_card_cond = (largura_util - (gap * 2)) / 3

    condicoes = [
        ("GARANTIA TÉCNICA", "90 dias regulamentares sobre os serviços executados."),
        ("VALIDADE DA PROPOSTA", "10 dias corridos a partir da data de emissão."),
        ("FATURAMENTO", "À vista via PIX/TED ou em até 2x fixas.")
    ]

    for i, (titulo_c, desc_c) in enumerate(condicoes):
        cx = margem_x + i * (largura_card_cond + gap)
        c.setFillColor(cor_card_bg)
        c.setStrokeColor(cor_borda)
        c.setLineWidth(0.7)
        c.roundRect(cx, y_pos, largura_card_cond, altura_cond, 5, fill=True, stroke=True)

        c.setFillColor(cor_accent)
        c.setFont("Helvetica-Bold", 7.5)
        c.drawString(cx + 10, y_pos + 31, titulo_c)

        c.setFillColor(cor_texto_principal)
        c.setFont("Helvetica", 8)
        palavras = desc_c.split()
        meio = len(palavras) // 2 + 1
        linha1 = " ".join(palavras[:meio])
        linha2 = " ".join(palavras[meio:])
        c.drawString(cx + 10, y_pos + 18, linha1)
        c.drawString(cx + 10, y_pos + 8, linha2)

    y_assinatura = 88
    c.setFillColor(cor_texto_secundario)
    c.setFont("Helvetica", 8.5)
    c.drawString(margem_x, y_assinatura + 42, "Atenciosamente,")

    desenhar_assinatura_ficticia(c, margem_x, y_assinatura)

    c.setStrokeColor(HexColor("#CBD5E1"))
    c.setLineWidth(0.8)
    c.line(margem_x, y_assinatura - 2, margem_x + 170, y_assinatura - 2)

    c.setFillColor(cor_navy)
    c.setFont("Helvetica-Bold", 9.5)
    c.drawString(margem_x, y_assinatura - 14, "Carlos Mendes")
    c.setFillColor(cor_texto_secundario)
    c.setFont("Helvetica", 8)
    c.drawString(margem_x, y_assinatura - 24, "Diretor Técnico • Empresa Engenharia")

    if os.path.exists(QR_CODE_PATH):
        qr_size = 56
        qr_x = width - margem_x - qr_size
        qr_y = y_assinatura - 16
        c.setStrokeColor(cor_borda)
        c.setLineWidth(0.8)
        c.roundRect(qr_x - 4, qr_y - 4, qr_size + 8, qr_size + 8, 4, fill=False, stroke=True)
        c.drawImage(QR_CODE_PATH, qr_x, qr_y, width=qr_size, height=qr_size)

        c.setFillColor(cor_texto_secundario)
        c.setFont("Helvetica", 7.5)
        c.drawRightString(qr_x - 12, qr_y + 32, "Valide este documento ou")
        c.drawRightString(qr_x - 12, qr_y + 22, "acesse nosso portfólio digital:")
        c.setFillColor(cor_accent)
        c.setFont("Helvetica-Bold", 8)
        c.drawRightString(qr_x - 12, qr_y + 10, "www.exemplo-empresa.com.br")

    altura_footer = 38
    c.setFillColor(cor_card_bg)
    c.rect(0, 0, width, altura_footer, fill=True, stroke=False)
    c.setStrokeColor(cor_borda)
    c.setLineWidth(0.8)
    c.line(0, altura_footer, width, altura_footer)

    c.setFillColor(cor_texto_principal)
    c.setFont("Helvetica-Bold", 7.5)
    c.drawCentredString(width / 2, 22, "EMPRESA SOLUÇÕES EM ENGENHARIA LTDA  •  CNPJ: 00.000.000/0001-00")
    c.setFillColor(cor_texto_secundario)
    c.setFont("Helvetica", 7.5)
    c.drawCentredString(width / 2, 10, "Av. das Indústrias, 1234 - Distrito Industrial, Cidade Exemplo - UF  |  Tel: (00) 0000-0000  |  contato@exemplo-empresa.com.br")

    c.save()
    callback_status("Abrindo PDF gerado...")
    webbrowser.open(PDF_PROPOSTA_PATH)

    if enviar_zap:
        time.sleep(1.2)
        callback_status("Abrindo WhatsApp Web...")
        disparar_whatsapp_web(
            telefone=telefone,
            nome_cliente=nome_cliente,
            num_proposta=num_proposta,
            ano_atual=ano_atual,
            valor_formatado=total_formatado_br,
            valor_extenso=total_extenso
        )

    return num_proposta, ano_atual, total_formatado_br


# ==============================================================================
# INTERFACE GRÁFICA MODERNA COM TKINTER (FONTES COM TAMANHOS INTEIROS)
# ==============================================================================
class AplicativoPropostasGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Empresa Engenharia • Gerador de Propostas Comerciais")
        self.root.geometry("560x530")
        self.root.resizable(False, False)
        self.root.configure(bg="#F1F5F9")

        # Paleta de Cores da Interface
        self.cor_navy = "#0F172A"
        self.cor_accent = "#2563EB"
        self.cor_accent_hover = "#1D4ED8"
        self.cor_bg = "#F1F5F9"
        self.cor_card = "#FFFFFF"
        self.cor_texto = "#1E293B"
        self.cor_subtexto = "#64748B"
        self.cor_borda = "#CBD5E1"
        self.cor_sucesso = "#16A34A"

        self.centralizar_janela()
        self.construir_interface()

    def centralizar_janela(self):
        self.root.update_idletasks()
        largura = 560
        altura = 530
        x = (self.root.winfo_screenwidth() // 2) - (largura // 2)
        y = (self.root.winfo_screenheight() // 2) - (altura // 2)
        self.root.geometry(f"{largura}x{altura}+{x}+{y}")

    def construir_interface(self):
        # 1. Cabeçalho Superior Escuro
        header = tk.Frame(self.root, bg=self.cor_navy, height=88)
        header.pack(fill="x")
        header.pack_propagate(False)

        faixa_azul = tk.Frame(self.root, bg=self.cor_accent, height=4)
        faixa_azul.pack(fill="x")

        tk.Label(
            header,
            text="GERADOR DE PROPOSTAS COMERCIAIS",
            font=("Segoe UI", 14, "bold"),
            fg="#FFFFFF",
            bg=self.cor_navy
        ).pack(anchor="w", padx=28, pady=(18, 2))

        tk.Label(
            header,
            text="Empresa Engenharia & Manutenção Industrial  •  PDF Executivo + WhatsApp",
            font=("Segoe UI", 9),
            fg="#94A3B8",
            bg=self.cor_navy
        ).pack(anchor="w", padx=28)

        # 2. Container Principal (Card Branco)
        card = tk.Frame(
            self.root,
            bg=self.cor_card,
            highlightbackground="#E2E8F0",
            highlightthickness=1
        )
        card.pack(fill="both", expand=True, padx=24, pady=20)

        # Campo 1: Nome do Cliente
        tk.Label(
            card,
            text="Nome do Cliente / Razão Social *",
            font=("Segoe UI", 10, "bold"),
            fg=self.cor_texto,
            bg=self.cor_card
        ).pack(anchor="w", padx=22, pady=(18, 4))

        self.entry_cliente = tk.Entry(
            card,
            font=("Segoe UI", 11),
            fg=self.cor_texto,
            bg="#F8FAFC",
            relief="solid",
            bd=1
        )
        self.entry_cliente.pack(fill="x", padx=22, ipady=6)
        self.entry_cliente.insert(0, "Cliente Exemplo Ltda")

        # Campo 2: WhatsApp do Contato
        tk.Label(
            card,
            text="WhatsApp do Contato (com DDD) *",
            font=("Segoe UI", 10, "bold"),
            fg=self.cor_texto,
            bg=self.cor_card
        ).pack(anchor="w", padx=22, pady=(14, 4))

        self.entry_whatsapp = tk.Entry(
            card,
            font=("Segoe UI", 11),
            fg=self.cor_texto,
            bg="#F8FAFC",
            relief="solid",
            bd=1
        )
        self.entry_whatsapp.pack(fill="x", padx=22, ipady=6)
        self.entry_whatsapp.insert(0, "(83) 99999-9999")

        # Campo 3: Seleção da Planilha Excel
        tk.Label(
            card,
            text="Planilha de Serviços (.xlsx)",
            font=("Segoe UI", 10, "bold"),
            fg=self.cor_texto,
            bg=self.cor_card
        ).pack(anchor="w", padx=22, pady=(14, 4))

        frame_planilha = tk.Frame(card, bg=self.cor_card)
        frame_planilha.pack(fill="x", padx=22)

        self.var_planilha = tk.StringVar(value=PLANILHA_PADRAO_PATH)
        self.entry_planilha = tk.Entry(
            frame_planilha,
            textvariable=self.var_planilha,
            font=("Segoe UI", 9),
            fg=self.cor_subtexto,
            bg="#F8FAFC",
            relief="solid",
            bd=1
        )
        self.entry_planilha.pack(side="left", fill="x", expand=True, ipady=5)

        btn_procurar = tk.Button(
            frame_planilha,
            text="Selecionar...",
            font=("Segoe UI", 9, "bold"),
            fg=self.cor_texto,
            bg="#E2E8F0",
            activebackground="#CBD5E1",
            relief="flat",
            cursor="hand2",
            command=self.selecionar_planilha
        )
        btn_procurar.pack(side="left", padx=(6, 0), ipady=4, ipadx=8)

        btn_abrir_excel = tk.Button(
            frame_planilha,
            text="Editar Itens",
            font=("Segoe UI", 9, "bold"),
            fg=self.cor_accent,
            bg="#EFF6FF",
            activebackground="#DBEAFE",
            relief="flat",
            cursor="hand2",
            command=self.abrir_planilha_excel
        )
        btn_abrir_excel.pack(side="left", padx=(6, 0), ipady=4, ipadx=8)

        # Opção Checkbox: Abrir WhatsApp Web automaticamente
        self.var_enviar_zap = tk.BooleanVar(value=True)
        chk_zap = tk.Checkbutton(
            card,
            text="Abrir WhatsApp Web automaticamente com a mensagem pronta",
            variable=self.var_enviar_zap,
            font=("Segoe UI", 10),
            fg=self.cor_texto,
            bg=self.cor_card,
            activebackground=self.cor_card,
            selectcolor=self.cor_card,
            cursor="hand2"
        )
        chk_zap.pack(anchor="w", padx=20, pady=(14, 10))

        # Botão Principal de Ação
        self.btn_gerar = tk.Button(
            card,
            text="GERAR PROPOSTA PDF E ENVIAR",
            font=("Segoe UI", 11, "bold"),
            fg="#FFFFFF",
            bg=self.cor_accent,
            activebackground=self.cor_accent_hover,
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            command=self.iniciar_geracao
        )
        self.btn_gerar.pack(fill="x", padx=22, pady=(6, 12), ipady=10)

        # Barra de Status
        self.label_status = tk.Label(
            card,
            text="Pronto para gerar a proposta.",
            font=("Segoe UI", 9),
            fg=self.cor_subtexto,
            bg=self.cor_card
        )
        self.label_status.pack(pady=(2, 10))

    def selecionar_planilha(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione a Planilha de Serviços",
            filetypes=[("Planilhas Excel", "*.xlsx *.xls")]
        )
        if arquivo:
            self.var_planilha.set(arquivo)

    def abrir_planilha_excel(self):
        caminho = self.var_planilha.get().strip()
        if os.path.exists(caminho):
            webbrowser.open(caminho)
        else:
            messagebox.showwarning("Aviso", "Arquivo de planilha não encontrado.")

    def atualizar_status(self, texto, cor=None):
        self.root.after(0, lambda: self.label_status.config(
            text=texto,
            fg=cor if cor else self.cor_subtexto
        ))

    def iniciar_geracao(self):
        nome_cliente = self.entry_cliente.get().strip()
        whatsapp = self.entry_whatsapp.get().strip()
        caminho_planilha = self.var_planilha.get().strip()
        enviar_zap = self.var_enviar_zap.get()

        if not nome_cliente:
            messagebox.showwarning("Campo Obrigatório", "Por favor, informe o nome do cliente.")
            self.entry_cliente.focus()
            return

        if not whatsapp:
            messagebox.showwarning("Campo Obrigatório", "Por favor, informe o número de WhatsApp com DDD.")
            self.entry_whatsapp.focus()
            return

        if not os.path.exists(caminho_planilha):
            messagebox.showerror("Erro", "A planilha selecionada não foi encontrada.")
            return

        self.btn_gerar.config(state="disabled", text="GERANDO PROPOSTA...")
        self.atualizar_status("Iniciando processamento...", self.cor_accent)

        thread = threading.Thread(
            target=self.executar_em_segundo_plano,
            args=(nome_cliente, whatsapp, caminho_planilha, enviar_zap),
            daemon=True
        )
        thread.start()

    def executar_em_segundo_plano(self, nome_cliente, whatsapp, caminho_planilha, enviar_zap):
        try:
            num_prop, ano, total_br = gerar_proposta_pdf_e_enviar(
                nome_cliente=nome_cliente,
                whatsapp_input=whatsapp,
                caminho_planilha=caminho_planilha,
                enviar_zap=enviar_zap,
                callback_status=self.atualizar_status
            )
            self.atualizar_status(
                f"Proposta Nº {num_prop}/{ano} gerada com sucesso! (Total: R$ {total_br})",
                self.cor_sucesso
            )
        except Exception as err:
            self.atualizar_status("Falha ao gerar a proposta.", "#DC2626")
            self.root.after(0, lambda: messagebox.showerror("Erro", str(err)))
        finally:
            self.root.after(0, lambda: self.btn_gerar.config(
                state="normal",
                text="GERAR PROPOSTA PDF E ENVIAR"
            ))

if __name__ == "__main__":
    root = tk.Tk()
    app = AplicativoPropostasGUI(root)
    root.mainloop()
