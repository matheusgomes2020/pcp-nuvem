import streamlit as st
import pandas as pd
import math
import re
import os
import io
import fitz  # PyMuPDF
import pdfplumber
import ezdxf
from pypdf import PdfWriter
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from regras_engenharia import calcular_mesa_parametrica, calcular_estante_parametrica, calcular_prateleira_parede

# Configuração da Página Web
st.set_page_config(page_title="ERP AçoNobre", page_icon="🏭", layout="wide")

# ==========================================
# 1. MOTOR DE ENGENHARIA, LASER E DXF
# ==========================================
PARAMETROS_LASER_INOX = {
    0.6: (7800, 1.0), 0.8: (7800, 1.0), 1.0: (7800, 1.0), 1.2: (6800, 1.0),
    1.5: (6300, 1.1), 2.0: (5300, 1.1), 2.5: (4500, 1.2), 3.0: (3800, 1.2),
    3.18: (3528, 1.3), 4.0: (2450, 1.4), 4.75: (2000, 1.5), 5.0: (1600, 1.8),
    6.35: (1200, 2.0), 8.0: (500, 2.5), 10.0: (350, 3.0), 12.7: (225, 4.0)
}

def to_comma_str(val):
    if not val or val == "-": return ""
    try: 
        fval = float(str(val).replace(',', '.'))
        if fval.is_integer(): return str(int(fval))
        return f"{fval:.2f}".replace('.', ',')
    except: 
        return str(val).replace('.', ',')

def extrair_dados_do_dxf(caminho_dxf):
    try:
        doc = ezdxf.readfile(caminho_dxf)
        msp = doc.modelspace()
        perimetro = 0.0
        entradas_peck = 1
        for entity in msp:
            tipo = entity.dxftype()
            if tipo == 'LINE':
                p1, p2 = entity.dxf.start, entity.dxf.end
                perimetro += math.sqrt((p2.x - p1.x)**2 + (p2.y - p1.y)**2)
            elif tipo == 'ARC':
                raio = entity.dxf.radius
                ang_inicio = entity.dxf.start_angle
                ang_fim = entity.dxf.end_angle
                delta = (ang_fim - ang_inicio) % 360
                perimetro += 2 * math.pi * raio * (delta / 360)
            elif tipo == 'CIRCLE':
                perimetro += 2 * math.pi * entity.dxf.radius
                entradas_peck += 1
            elif tipo in ['LWPOLYLINE', 'POLYLINE']:
                if hasattr(entity, 'closed') and entity.closed: entradas_peck += 1
                try:
                    from ezdxf import path
                    perimetro += path.make_path(entity).length()
                except: pass
        return perimetro, entradas_peck
    except Exception as e:
        return 0.0, 0

def calcular_tempo_e_custo_laser(espessura_str, perimetro_mm, entradas):
    try:
        esp = float(str(espessura_str).replace(',', '.'))
        avanco, peck = PARAMETROS_LASER_INOX.get(esp, (5000, 1.5))
        tempo_corte = perimetro_mm / avanco
        tempo_furo = (entradas * peck) / 60.0
        return round(tempo_corte + tempo_furo, 2)
    except: return 0.0

def calcular_distancia(valA, valB, dict_coords):
    min_dist = float('inf')
    for pA in dict_coords[valA]:
        for pB in dict_coords[valB]:
            dist = math.sqrt((pA['x'] - pB['x'])**2 + (pA['y'] - pB['y'])**2)
            if dist < min_dist: min_dist = dist
    return min_dist

def extrair_dados_tecnicos(caminho_pdf):
    dados = {
        "massa": "", "esp": "", "tipo": "CHAPA LISA", 
        "c_plan": "", "l_plan": "", "c_fin": "", "l_fin": "", "a_fin": "",
        "planificado": "-", "finalizado": "-"
    }
    numeros_dict = {} 
    coord_secao = None
    
    doc = fitz.open(caminho_pdf)
    pagina = doc[0]
    nome_arquivo = caminho_pdf.replace('\\', '/').split('/')[-1]
    codigo_peca = nome_arquivo.lower().replace('.pdf', '')
    
    texto_completo = pagina.get_text("text")
    if re.search(r'(DOBRA|PARA BAIXO|PARA CIMA|90°)', texto_completo, re.IGNORECASE):
        dados["tipo"] = "CHAPA DOBRADA"
        
    m = re.search(r'([\d.,]+)\s*Kg', texto_completo, re.IGNORECASE)
    if m: dados["massa"] = m.group(1).replace(',', '.')
    
    e = re.search(r'Espessura[:\s\n]*([\d.,]+)\s*(?:mm)?', texto_completo, re.IGNORECASE)
    if e: dados["esp"] = e.group(1).replace(',', '.')

    blocos = pagina.get_text("dict")["blocks"]
    for b in blocos:
        if "lines" in b:
            for linha in b["lines"]:
                for span in linha["spans"]:
                    texto = span["text"].strip()
                    x_centro = (span["bbox"][0] + span["bbox"][2]) / 2
                    y_centro = (span["bbox"][1] + span["bbox"][3]) / 2
                    if re.search(r'(SEÇÃO|SECAO|DETALHE)', texto, re.IGNORECASE):
                        coord_secao = (x_centro, y_centro)
                    if texto == codigo_peca or texto.replace(',', '.') == codigo_peca.replace(',', '.'):
                        continue
                    if re.match(r'^\d+([.,]\d+)?$', texto):
                        val = float(texto.replace(',', '.'))
                        if 5.0 < val < 3000.0:
                            if str(val) != dados["massa"] and str(val) != dados["esp"]:
                                if val not in numeros_dict:
                                    numeros_dict[val] = []
                                numeros_dict[val].append({'x': x_centro, 'y': y_centro})

    nums = sorted(list(numeros_dict.keys()), reverse=True)

    if dados["tipo"] != "CHAPA LISA" and len(nums) >= 3:
        if nums[0] > nums[1] * 1.3:
            comp_plan = comp_fin = nums[0]
            grupos = []
            for n in nums[1:]:
                if not grupos or n < grupos[-1][0] * 0.85:
                    grupos.append([n])
                else: grupos[-1].append(n)
            larg_plan = grupos[0][0] if len(grupos) > 0 else 0
            larg_fin = grupos[1][0] if len(grupos) > 1 else larg_plan
            h_cands = [g[0] for g in grupos[2:]]
        elif len(nums) >= 4:
            d1, d2, d3, d4 = nums[:4]
            distA = calcular_distancia(d1, d2, numeros_dict) + calcular_distancia(d3, d4, numeros_dict)
            distB = calcular_distancia(d1, d3, numeros_dict) + calcular_distancia(d2, d4, numeros_dict)
            distC = calcular_distancia(d1, d4, numeros_dict) + calcular_distancia(d2, d3, numeros_dict)
            min_dist = min(distA, distB, distC)
            if min_dist == distA: par1, par2 = (d1, d2), (d3, d4)
            elif min_dist == distB: par1, par2 = (d1, d3), (d2, d4)
            else: par1, par2 = (d1, d4), (d2, d3)
            plan_pair = par1 if d1 in par1 else par2
            fin_pair  = par2 if d1 in par1 else par1
            comp_plan, larg_plan = max(plan_pair), min(plan_pair)
            comp_fin, larg_fin   = max(fin_pair), min(fin_pair)
            h_cands = [n for n in nums[4:] if n < larg_fin * 0.85]
        else:
            comp_plan = comp_fin = nums[0]
            larg_plan = larg_fin = nums[1]
            h_cands = [n for n in nums[2:] if n < larg_fin * 0.85]
            
        alt_fin = 0
        if h_cands:
            if coord_secao:
                menor_dist = float('inf')
                for val in h_cands:
                    dist = calcular_distancia(val, val, {val: numeros_dict[val], val: [{'x': coord_secao[0], 'y': coord_secao[1]}]})
                    if dist < menor_dist:
                        menor_dist = dist
                        alt_fin = val
            else:
                alt_fin = h_cands[0] 
                
        dados["c_plan"], dados["l_plan"] = comp_plan, larg_plan
        dados["c_fin"], dados["l_fin"], dados["a_fin"] = comp_fin, larg_fin, alt_fin
    else:
        if len(nums) >= 2:
            dados["c_plan"] = dados["c_fin"] = nums[0]
            dados["l_plan"] = dados["l_fin"] = nums[1]
            dados["a_fin"] = 0

    return dados

PADRAO_CODIGO = re.compile(r"([A-Za-z]{2,4}[-\.]\w+[-\.]\d{2,4}(?:\.\d+)?(?:-R\d{2})?|\d{3}[.\-_]\d{6}|00[-\.]\d{2,4})", re.IGNORECASE)

def normalizar_codigo(codigo):
    codigo = str(codigo).strip().upper()
    if re.match(r'^\d{3}[.\-_]\d{6}$', codigo): return re.sub(r"[\-_]", ".", codigo)
    return codigo 

def extrair_codigo_do_nome_arquivo(nome_arquivo):
    nome_sem_extensao = os.path.splitext(nome_arquivo)[0]
    match = PADRAO_CODIGO.search(nome_sem_extensao)
    if match: return normalizar_codigo(match.group(1))
    return None

def escanear_pasta(pasta):
    mapa = {}
    if not pasta or not os.path.exists(pasta): return mapa
    for nome_arquivo in sorted(os.listdir(pasta)):
        if nome_arquivo.lower().endswith((".pdf", ".dxf")):
            codigo = extrair_codigo_do_nome_arquivo(nome_arquivo)
            if codigo: 
                if codigo not in mapa: mapa[codigo] = {}
                ext = ".pdf" if nome_arquivo.lower().endswith(".pdf") else ".dxf"
                mapa[codigo][ext] = nome_arquivo
    return mapa

def ler_planilha_entrada(caminho):
    df = pd.read_excel(caminho, dtype=str)
    pedidos = {}
    dados_comerciais = {}
    for index, row in df.iterrows():
        pedido = str(row.iloc[0]).strip()
        item_num = str(row.iloc[4]).strip()
        codigo_norm = normalizar_codigo(str(row.iloc[7]).replace("'", "").strip())
        chave_unica = f"{pedido}_{item_num}_{codigo_norm}"
        if pedido not in pedidos: pedidos[pedido] = []
        pedidos[pedido].append((item_num, codigo_norm))
        of_bruta = str(row.iloc[8]).strip()
        of_limpa = of_bruta.replace(" 00:00:00", "").replace(" 00:00", "")
        if " 00:" in of_limpa: of_limpa = of_limpa.split(" 00:")[0].strip()
        dados_comerciais[chave_unica] = {
            "cliente": str(row.iloc[1]).strip(), "inicio": str(row.iloc[2]).strip(),
            "entrega": str(row.iloc[3]).strip(), "desc": str(row.iloc[5]).strip(),
            "qtd": str(row.iloc[6]).strip(), "of": of_limpa
        }
    return pedidos, dados_comerciais

def extrair_tabela_de_materiais(caminho_pdf):
    pecas_dict = {} 
    cod_pai = extrair_codigo_do_nome_arquivo(os.path.basename(caminho_pdf))
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            tabelas = pagina.extract_tables()
            for tabela in tabelas:
                if not tabela: continue
                for linha in tabela:
                    cols = [str(c).replace('\n', ' ').strip() if c else "" for c in linha]
                    cols_validas = [c for c in cols if c and c != "-"]
                    if not cols_validas: continue
                    linha_str = "  ".join(cols_validas).upper()
                    if "CONJUNTO" in linha_str or "CONJ." in linha_str or "CONJ " in linha_str: continue
                    matches = PADRAO_CODIGO.findall(linha_str)
                    cod_peca = None
                    if matches:
                        for m in matches:
                            c_norm = normalizar_codigo(m)
                            if c_norm != cod_pai: 
                                cod_peca = c_norm
                                break
                    else:
                        for c, p in pecas_dict.items():
                            desc_atual = p["desc"].upper()
                            if desc_atual != "-" and len(desc_atual) > 3:
                                if desc_atual in linha_str or any(desc_atual in col for col in cols_validas):
                                    cod_peca = c
                                    break
                    if not cod_peca: continue
                    if cod_peca not in pecas_dict:
                        pecas_dict[cod_peca] = {"n": "-", "cod": cod_peca, "desc": "-", "mat": "-", "esp": "-", "qtd": "-", "_raw_esp": ""}
                    p = pecas_dict[cod_peca]
                    if cols_validas[0].isdigit() and len(cols_validas[0]) <= 3:
                        if p["n"] == "-": p["n"] = cols_validas[0]
                    if cols_validas[-1].isdigit() and len(cols_validas[-1]) <= 3:
                        if p["qtd"] == "-": p["qtd"] = cols_validas[-1]
                    elif len(cols_validas) > 1 and cols_validas[-2].isdigit() and len(cols_validas[-2]) <= 3:
                        if p["qtd"] == "-": p["qtd"] = cols_validas[-2]
                    for col in cols_validas:
                        col_up = col.upper()
                        if cod_peca in col_up:
                            desc_limpa = col_up.replace(cod_peca, "")
                            desc_limpa = re.sub(r'-?R\d{2}-?', '', desc_limpa).strip(" -/.") 
                            if desc_limpa and len(desc_limpa) > 2:
                                if p["desc"] == "-" or len(desc_limpa) > len(p["desc"]): p["desc"] = desc_limpa
                        elif not col_up.isdigit() and not re.match(r'^[\d.,]+\s*(MM|mm)?$', col_up):
                            if not any(x in col_up for x in ["INOX", "AISI", "CARBONO", "GALV", "POLIDO", "ESC", "CHAPA"]):
                                if p["desc"] == "-" and len(col_up) > 3: p["desc"] = col_up
                        if any(x in col_up for x in ["INOX", "AISI", "CARBONO", "GALV", "POLIDO", "ESC"]):
                            if p["mat"] == "-" or len(col_up) > len(p["mat"]): p["mat"] = col_up
                        match_esp = re.match(r'^([\d.,]+)\s*(MM|mm)?$', col_up)
                        if match_esp:
                            val_str = match_esp.group(1)
                            tem_mm = bool(match_esp.group(2))
                            is_decimal = (',' in val_str or '.' in val_str)
                            is_num_item_ou_qtd = (val_str == p["n"] or val_str == p["qtd"])
                            if tem_mm or is_decimal or not is_num_item_ou_qtd:
                                try:
                                    f_val = float(val_str.replace(',', '.'))
                                    if 0.3 <= f_val <= 20.0: 
                                        if p["esp"] == "-" or (tem_mm and "MM" not in str(p.get("_raw_esp", ""))):
                                            p["esp"] = val_str
                                            p["_raw_esp"] = col_up
                                except: pass
    for p in pecas_dict.values():
        if p["mat"] == "-":
            m_busca = re.search(r'(INOX|AISI|CARBONO|GALV)\s*\d*\s*(ESC|POL|ESCOVADO|POLIDO)?', p["desc"])
            if m_busca: p["mat"] = m_busca.group(0)
        if p["esp"] == "-":
            e_busca = re.search(r'(\d[.,]\d+)\s*MM', p["desc"] + " " + p["mat"])
            if e_busca: p["esp"] = e_busca.group(1)
        if not re.search(r'(INOX|AISI|CARBONO|GALV)', p["mat"]):
            num_match = re.search(r'(\d[.,]\d+)', p["mat"])
            if num_match and p["esp"] == "-": p["esp"] = num_match.group(1)
            if p["mat"] != "-": p["mat"] = "-" 
        if p["mat"] != "-":
            mat_str = p["mat"].upper()
            mat_str = re.sub(r'\b(CHAPA|ESP|DE|MM|\d[.,]\d+)\b', '', mat_str).strip() 
            mat_str = mat_str.replace("AISI", "INOX")
            mat_str = re.sub(r'INOX(\d)', r'INOX \1', mat_str) 
            mat_str = re.sub(r'(\d)(ESC|POL|ESCOVADO|POLIDO)', r'\1 \2', mat_str) 
            mat_str = mat_str.replace("ESCOVADO", "ESC").replace("POLIDO", "POL")
            mat_str = re.sub(r'\s+', ' ', mat_str).strip()
            p["mat"] = mat_str if mat_str else "-"
        if p["esp"] != "-":
            esp_str = str(p["esp"]).upper().replace("MM", "").strip()
            try:
                val_e = float(esp_str.replace(',', '.'))
                if 0.3 <= val_e <= 20.0: p["esp"] = f"{int(val_e)},0" if val_e.is_integer() else str(val_e).replace('.', ',')
                else: p["esp"] = "-"
            except: p["esp"] = "-"
        p["desc"] = re.sub(r'\s+', ' ', p["desc"]).strip()

    lista_pecas = list(pecas_dict.values())
    def ordenar_por_numero(peca):
        try: return int(peca['n'])
        except ValueError: return 9999
    lista_pecas.sort(key=ordenar_por_numero)
    return lista_pecas

def gerar_conferencia_item(n_item, cod_produto, pecas, pasta_item, num_pedido, dados_comerciais):
    wb = Workbook()
    ws = wb.active
    fonte_p = Font(name='Calibri', size=11)
    fonte_neg = Font(name='Calibri', size=11, bold=True)
    fonte_titulo = Font(name='Calibri', size=14, bold=True)
    borda = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    alinhamento_centro = Alignment(horizontal="center", vertical="center")
    alinhamento_esq = Alignment(horizontal="left", vertical="center")
    
    linha = 2
    ws.merge_cells(f"B{linha}:J{linha}")
    cliente = dados_comerciais.get('cliente', 'AVULSO')
    ws.cell(linha, 2, f"PEDIDO {num_pedido} - {cliente}").font = fonte_titulo
    ws.cell(linha, 2).alignment = alinhamento_centro
    ws.cell(linha, 2).fill = PatternFill(start_color="DDDDDD", fill_type="solid")
    for col in range(2, 11): ws.cell(linha, col).border = borda
    ws.row_dimensions[linha].height = 18.75
    linha += 1
    
    ws.cell(linha, 2, "OF").font = fonte_neg; ws.cell(linha, 2).alignment = alinhamento_centro
    ws.cell(linha, 3, "CÓD").font = fonte_neg; ws.cell(linha, 3).alignment = alinhamento_centro
    ws.merge_cells(f"D{linha}:I{linha}")
    ws.cell(linha, 4, "DESCRIÇÃO").font = fonte_neg; ws.cell(linha, 4).alignment = alinhamento_centro
    ws.cell(linha, 10, "QTD").font = fonte_neg; ws.cell(linha, 10).alignment = alinhamento_centro
    for col in range(2, 11): ws.cell(linha, col).border = borda
    ws.row_dimensions[linha].height = 18.75
    linha += 1
    
    ws.cell(linha, 2, dados_comerciais.get('of', '-')).font = fonte_neg; ws.cell(linha, 2).alignment = alinhamento_centro
    ws.cell(linha, 3, cod_produto).font = fonte_neg; ws.cell(linha, 3).alignment = alinhamento_centro
    ws.merge_cells(f"D{linha}:I{linha}")
    ws.cell(linha, 4, dados_comerciais.get('desc', '-')).font = fonte_neg; ws.cell(linha, 4).alignment = alinhamento_esq
    try: qtd_pai = int(float(str(dados_comerciais.get('qtd', '1')).replace(',', '.')))
    except: qtd_pai = 1
    ws.cell(linha, 10, qtd_pai).font = fonte_neg; ws.cell(linha, 10).alignment = alinhamento_centro
    for col in range(2, 11): ws.cell(linha, col).border = borda
    ws.row_dimensions[linha].height = 18.75
    linha += 1
    
    headers = ["TIPO", "Nº", "CÓDIGO", "ITENS", "QTD", "QTD OF", "ENVIADA", "SALDO", "✔️"]
    for i, h in enumerate(headers, 2):
        cel = ws.cell(linha, i, h); cel.font = fonte_neg; cel.fill = PatternFill(start_color="DDDDDD", fill_type="solid")
        cel.border = borda; cel.alignment = alinhamento_centro
    ws.row_dimensions[linha].height = 18.75
    linha += 1
    
    for p in pecas:
        tipo = "INSUMO" if p['cod'].startswith("1") else "MATERIAL"
        ws.cell(linha, 2, tipo).font = fonte_p; ws.cell(linha, 2).alignment = alinhamento_centro
        ws.cell(linha, 3, p['n']).font = fonte_p; ws.cell(linha, 3).alignment = alinhamento_centro
        ws.cell(linha, 4, p['cod']).font = fonte_p; ws.cell(linha, 4).alignment = alinhamento_centro 
        ws.cell(linha, 5, p['desc']).font = fonte_p; ws.cell(linha, 5).alignment = alinhamento_esq
        qtd_peca = to_comma_str(p['qtd'])
        ws.cell(linha, 6, qtd_peca).font = fonte_p; ws.cell(linha, 6).alignment = alinhamento_centro
        ws.cell(linha, 7, f"=F{linha}*$J$4").font = fonte_p; ws.cell(linha, 7).alignment = alinhamento_centro
        for col in range(2, 11): ws.cell(linha, col).border = borda
        ws.row_dimensions[linha].height = 18.75
        linha += 1

    linha += 2
    ws.cell(linha, 2, "Conferido por: _________________________________________   Data: ____/____/____").font = fonte_neg
    linha += 2
    ws.cell(linha, 2, "Recebido por:  _________________________________________   Data: ____/____/____").font = fonte_neg

    for col_let, w in zip(['A','B','C','D','E','F','G','H','I','J'], [2, 13, 14, 14, 40, 8, 12, 12, 10, 6]): ws.column_dimensions[col_let].width = w
    ws.page_setup.fitToPage = True      
    ws.page_setup.orientation = "landscape" 
    
    nome_arquivo = f"LISTA-CONFERÊNCIA-{num_pedido}-{n_item}-{cod_produto}.xlsx"
    wb.save(os.path.join(pasta_item, nome_arquivo))

def atualizar_excel_inteligente(dados_novos, caminho_global, num_item):
    df_novo = pd.DataFrame(dados_novos)
    if os.path.exists(caminho_global):
        try:
            df_existente = pd.read_excel(caminho_global, dtype=str)
            df_existente = df_existente[df_existente["ITEM"] != str(num_item)]
            df_final = pd.concat([df_existente, df_novo], ignore_index=True)
        except: df_final = df_novo
    else: df_final = df_novo

    formulas_peso_bruto = []
    formulas_desc_nome = []
    formulas_desc_desc = []
    formulas_desc_ficha = []

    for i in range(len(df_final)):
        linha_excel = i + 2 
        formulas_peso_bruto.append(f"=S{linha_excel}*T{linha_excel}*N{linha_excel}*0.000008")
        formulas_desc_nome.append(f'=L{linha_excel} & " - " & M{linha_excel} & " " & N{linha_excel} & "MM"')
        formulas_desc_desc.append(f'=L{linha_excel} & " - " & M{linha_excel} & " - " & N{linha_excel} & "MM - " & P{linha_excel} & "X" & Q{linha_excel} & "X" & R{linha_excel} & "MM"')
        formulas_desc_ficha.append(f'=L{linha_excel} & " - " & M{linha_excel} & " - " & N{linha_excel} & "MM - " & P{linha_excel} & "X" & Q{linha_excel} & "X" & R{linha_excel} & "MM - PLAN(" & S{linha_excel} & "X" & T{linha_excel} & "MM)"')
    
    df_final["PESO BRUTO"] = formulas_peso_bruto
    df_final["DESC PRODUTO NOME"] = formulas_desc_nome
    df_final["DESC PRODUTO DESCRIÇÃO"] = formulas_desc_desc
    df_final["DESC PRODUTO FICHA"] = formulas_desc_ficha

    with pd.ExcelWriter(caminho_global, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False)
        planilha_g = writer.sheets['Sheet1']
        for i in range(1, len(df_final.columns) + 1): planilha_g.column_dimensions[get_column_letter(i)].width = 22

def processar_item_unico(n_item, cod_pai, num_pedido, dados_comerciais, pasta_pedido, origem_pdf, origem_dxf, mapa_pdf, mapa_dxf, opcoes, log_widget):
    log_widget.write(f"\n▶ Iniciando Item {n_item} -> Produto: {cod_pai}\n")
    pasta_item = os.path.join(pasta_pedido, f"{n_item} - {cod_pai}")
    os.makedirs(pasta_item, exist_ok=True)
    
    if cod_pai not in mapa_pdf:
        log_widget.write(f"❌ ERRO: PDF pai ({cod_pai}) não encontrado na pasta origem.\n")
        return []

    caminho_pai = os.path.join(origem_pdf, mapa_pdf[cod_pai][".pdf"])
    caminho_pai_copiado = os.path.join(pasta_item, mapa_pdf[cod_pai][".pdf"])
    shutil.copy2(caminho_pai, caminho_pai_copiado)
    
    log_widget.write(f"  📄 Lendo desenho pai: {mapa_pdf[cod_pai]['.pdf']}\n")
    pecas = extrair_tabela_de_materiais(caminho_pai)
    log_widget.write(f"  🔍 Encontradas {len(pecas)} peças na tabela. Extraindo dados (PDF/DXF)...\n")
    
    dados_pcp = []
    merger = PdfWriter() if opcoes["pdf"] else None
    if merger: merger.append(caminho_pai_copiado)
    
    if opcoes["pcp"] or opcoes["pdf"]:
        for p in pecas:
            eng_dados = {}
            metros_corte = "-"
            qtd_entradas = "-"
            tempo_laser_min = "-"
            tempo_laser_relogio = "-"
            
            log_widget.write(f"    🔸 Lendo medidas de: {p['cod']} - {p['desc'][:15]}...\n")
            
            if p['cod'] in mapa_pdf:
                caminho_filho = os.path.join(origem_pdf, mapa_pdf[p['cod']][".pdf"])
                caminho_dxf_real = ""
                if p['cod'] in mapa_dxf and ".dxf" in mapa_dxf[p['cod']]:
                    caminho_dxf_real = os.path.join(origem_dxf, mapa_dxf[p['cod']][".dxf"])

                if opcoes["pcp"]:
                    try: 
                        eng_dados = extrair_dados_tecnicos(caminho_filho)
                    except Exception as e: 
                        log_widget.write(f"      [AVISO] Falha ao ler medidas do PDF {p['cod']}: {e}\n")
                    
                    if caminho_dxf_real:
                        try:
                            peri_mm, entradas = extrair_dados_do_dxf(caminho_dxf_real)
                            if peri_mm > 0:
                                metros_corte = to_comma_str(round(peri_mm / 1000.0, 3))
                                qtd_entradas = str(entradas)
                                tempo_calc = calcular_tempo_e_custo_laser(p['esp'], peri_mm, entradas)
                                tempo_laser_min = to_comma_str(tempo_calc)
                                
                                minutos_relogio = int(math.floor(tempo_calc))
                                segundos_relogio = int(round((tempo_calc - minutos_relogio) * 60))
                                tempo_laser_relogio = f"{minutos_relogio:02d}:{segundos_relogio:02d}s"
                                
                                log_widget.write(f"      ✂️ DXF Lido: {metros_corte} m | {entradas} Furos | Tempo: {tempo_laser_relogio} ({tempo_laser_min} min).\n")
                        except Exception as e:
                            log_widget.write(f"      [AVISO] Erro no DXF: {e}\n")

                if opcoes["pdf"]:
                    caminho_filho_copiado = os.path.join(pasta_item, mapa_pdf[p['cod']][".pdf"])
                    shutil.copy2(caminho_filho, caminho_filho_copiado)
                    merger.append(caminho_filho_copiado)
            
            if opcoes["pcp"]:
                dados_pcp.append({
                    "PEDIDO": num_pedido, "CLIENTE": dados_comerciais.get('cliente', ''), 
                    "DATA INICIO": dados_comerciais.get('inicio', ''), "DATA ENTREGA": dados_comerciais.get('entrega', ''),
                    "OF": dados_comerciais.get('of', ''), "ITEM": str(n_item), "CÓDIGO": cod_pai,
                    "DESCRIÇÃO": dados_comerciais.get('desc', ''), "QTD": to_comma_str(dados_comerciais.get('qtd', '1')),
                    "Nº": p['n'], "CÓDIGO PEÇA": p['cod'], 
                    "DESCRIÇÃO PEÇA": p['desc'], "MATERIAL": p['mat'], 
                    "ESPESSURA": to_comma_str(p['esp']), "QTD PEÇA": to_comma_str(p['qtd']),
                    "COMP": to_comma_str(eng_dados.get("c_fin", "")), "LARG": to_comma_str(eng_dados.get("l_fin", "")),
                    "ALT": to_comma_str(eng_dados.get("a_fin", "")), "COMP PL": to_comma_str(eng_dados.get("c_plan", "")),
                    "LARG PL": to_comma_str(eng_dados.get("l_plan", "")), "MASSA": to_comma_str(eng_dados.get("massa", "0,00")),
                    "PESO BRUTO": "",
                    "CORTE (M)": metros_corte,
                    "ENTRADAS (Furos)": qtd_entradas,
                    "TEMPO LASER (Min)": tempo_laser_min,
                    "TEMPO LASER (Relógio)": tempo_laser_relogio,
                    "DESC PRODUTO NOME": "",
                    "DESC PRODUTO DESCRIÇÃO": "",
                    "DESC PRODUTO FICHA": ""
                })
    
    if opcoes["pdf"] and merger:
        caminho_unificado = os.path.join(pasta_item, f"{cod_pai}-COMPLETO.pdf")
        merger.write(caminho_unificado)
        merger.close()
        log_widget.write("  ✅ PDF Unificado gerado.\n")

    if opcoes["conf"]:
        gerar_conferencia_item(n_item, cod_pai, pecas, pasta_item, num_pedido, dados_comerciais)
        log_widget.write("  ✅ Lista de Conferência gerada.\n")
        
    return dados_pcp


# ==========================================
# 2. VARIÁVEIS DE SESSÃO E MAPAS
# ==========================================
if 'carrinho' not in st.session_state: st.session_state.carrinho = []
if 'custo_chapa' not in st.session_state:
    st.session_state.custo_chapa = {
        "INOX 304": {0.6: 29.01, 0.8: 33.16, 1.0: 30.74, 1.2: 28.66, 1.5: 29.00, 2.0: 30.96},
        "INOX 430": {0.6: 23.32, 0.8: 19.41, 1.0: 18.91, 1.2: 19.42, 1.5: 19.38, 2.0: 19.38},
        "INOX 201": {0.6: 15.00, 0.8: 16.00, 1.0: 16.50, 1.2: 17.00}
    }
if 'custo_tubo' not in st.session_state:
    st.session_state.custo_tubo = {"TUBO-RED-38": 18.47, "TUBO-RED-25": 12.08}

mapa_tampo = {"Mesa Lisa": "LISA", "Mesa com Encosto": "ENCOSTO", "Pia com Cuba": "PIA", "Prateleira Parede": "PRAT_PAREDE", "Estante Lisa": "ESTANTE_LISA", "Estante Gradeada": "ESTANTE_GRADEADA"}
mapa_base = {"Contraventamento": "CONTRAVENTAMENTO", "Prat. Lisa": "PRAT_LISA", "Prat. Lisa Dupla": "PRAT_LISA_DUPLA", "Prat. Gradeada": "PRAT_GRADEADA", "Prat. Gradeada Dupla": "PRAT_GRADEADA_DUPLA"}

# ==========================================
# 3. INTERFACE PRINCIPAL
# ==========================================
st.title("🏭 AçoNobre ERP - Painel Web")

tab_projetista, tab_pcp = st.tabs(["🛒 Projetista Virtual", "⚙️ Gerador de PCP"])

# ----------------- ABA PROJETISTA -----------------
with tab_projetista:
    col_esq, col_dir = st.columns([1, 2])

    with col_esq:
        st.subheader("📦 Montar Pedido")
        pedido_num = st.text_input("Nº do Pedido (Ex: 6885)")

        aba_param, aba_livre = st.tabs(["📐 Catálogo Paramétrico", "✏️ Peça Avulsa"])
        
        with aba_param:
            c1, c2, c3 = st.columns(3)
            qtd = c1.number_input("QTD", min_value=1, value=1)
            comp = c2.number_input("Comp. (mm)", value=0.0, step=10.0)
            larg = c3.number_input("Larg. (mm)", value=0.0, step=10.0)

            tampo_nome = st.selectbox("Selecione o Produto", list(mapa_tampo.keys()))
            cod_tampo = mapa_tampo[tampo_nome]

            alt = 0.0; planos_val = 4; base_nome = "Contraventamento"; cod_base = "CONTRAVENTAMENTO"
            
            if "ESTANTE" in cod_tampo:
                c4, c5 = st.columns(2)
                alt = c4.number_input("Altura (mm)", value=0.0, step=10.0)
                planos_val = c5.number_input("Qtd Planos", min_value=1, value=4)
                lbl_tampo = "Mat. dos Planos"
                lbl_base = "Mat. das Colunas"
            elif cod_tampo == "PRAT_PAREDE":
                lbl_tampo = "Material Principal"
            else:
                c4, c5 = st.columns(2)
                alt = c4.number_input("Altura (mm)", value=0.0, step=10.0)
                base_nome = c5.selectbox("Tipo de Base", list(mapa_base.keys()))
                cod_base = mapa_base[base_nome]
                lbl_tampo = "Material do Tampo"
                lbl_base = "Material da Base"

            st.markdown("---")
            st.markdown("**Materiais e Espessuras**")
            
            c_mat1, c_mat2 = st.columns(2)
            mat_tampo = c_mat1.selectbox(lbl_tampo, ["INOX 304", "INOX 430", "INOX 201"])
            esp_tampo = c_mat2.selectbox(f"Esp. {lbl_tampo}", ["Esp. Padrão", "0.8", "1.0", "1.2", "1.5", "2.0"])

            mat_base = "INOX 430"; esp_base = "Esp. Padrão"
            if cod_tampo != "PRAT_PAREDE" and cod_base != "CONTRAVENTAMENTO":
                idx_mat = ["INOX 304", "INOX 430", "INOX 201"].index(mat_tampo)
                c_base1, c_base2 = st.columns(2)
                mat_base = c_base1.selectbox(lbl_base, ["INOX 304", "INOX 430", "INOX 201"], index=idx_mat)
                esp_base = c_base2.selectbox(f"Esp. {lbl_base}", ["Esp. Padrão", "0.8", "1.0", "1.2", "1.5", "2.0"])

            st.markdown("---")
            st.markdown("**Engenharia de Reforços**")
            
            lbl_ref1 = "Ref. Plano" if "ESTANTE" in cod_tampo else "Reforço" if cod_tampo == "PRAT_PAREDE" else "Ref. Tampo"
            
            c_r1, c_r2, c_r3 = st.columns(3)
            qtd_ref = c_r1.selectbox(f"Qtd {lbl_ref1}", ["Padrão", "0", "1", "2", "3", "4", "5", "6"])
            idx_ref1 = ["INOX 430", "INOX 304", "INOX 201"].index(mat_tampo) if mat_tampo in ["INOX 430", "INOX 304", "INOX 201"] else 0
            mat_ref = c_r2.selectbox(f"Mat. {lbl_ref1}", ["INOX 430", "INOX 304", "INOX 201"], index=idx_ref1)
            esp_ref = c_r3.selectbox(f"Esp. {lbl_ref1}", ["0.6", "0.8", "1.0", "1.2", "1.5"], index=1)

            mat_ref_prat = "INOX 430"; esp_ref_prat = "0.8"; qtd_ref_prat = "Padrão"
            if "PRAT" in cod_base and cod_tampo != "PRAT_PAREDE" and "ESTANTE" not in cod_tampo:
                lbl_ref_p = "Ref/Prat" if "DUPLA" in cod_base else "Ref. Prat"
                c_p1, c_p2, c_p3 = st.columns(3)
                qtd_ref_prat = c_p1.selectbox(f"Qtd {lbl_ref_p}", ["Padrão", "0", "1", "2", "3", "4", "5", "6"])
                idx_ref2 = ["INOX 430", "INOX 304", "INOX 201"].index(mat_base) if mat_base in ["INOX 430", "INOX 304", "INOX 201"] else 0
                mat_ref_prat = c_p2.selectbox(f"Mat. {lbl_ref_p}", ["INOX 430", "INOX 304", "INOX 201"], index=idx_ref2)
                esp_ref_prat = c_p3.selectbox(f"Esp. {lbl_ref_p}", ["0.6", "0.8", "1.0", "1.2", "1.5"], index=1)

            if st.button("➕ Adicionar ao Pedido", type="primary", use_container_width=True):
                item = {
                    "num": len(st.session_state.carrinho)+1, "tipo": "parametrico", "qtd": int(qtd), 
                    "comp": comp, "larg": larg, "alt": alt, "planos": int(planos_val), 
                    "tampo_nome": tampo_nome, "tampo_cod": cod_tampo, 
                    "base_nome": base_nome, "base_cod": cod_base, 
                    "mat_tampo": mat_tampo, "esp_tampo": esp_tampo, 
                    "mat_base": mat_base, "esp_base": esp_base,
                    "mat_ref": mat_ref, "esp_ref": esp_ref, "qtd_ref": qtd_ref,
                    "mat_ref_prat": mat_ref_prat, "esp_ref_prat": esp_ref_prat, "qtd_ref_prat": qtd_ref_prat
                }
                if "ESTANTE" in item["tampo_cod"]: item["desc_carrinho"] = f"{item['tampo_nome']} {item['planos']} Planos ({int(item['comp'])}x{int(item['larg'])}x{int(item['alt'])}) - {item['mat_tampo']}"
                elif item["tampo_cod"] == "PRAT_PAREDE": item["desc_carrinho"] = f"{item['tampo_nome']} ({int(item['comp'])}x{int(item['larg'])}) - {item['mat_tampo']}"
                else: 
                    if item["base_cod"] == "CONTRAVENTAMENTO": item["desc_carrinho"] = f"{item['tampo_nome']} c/ {item['base_nome']} ({int(item['comp'])}x{int(item['larg'])}x{int(item['alt'])}) - {item['mat_tampo']}"
                    else: item["desc_carrinho"] = f"{item['tampo_nome']} c/ {item['base_nome']} ({int(item['comp'])}x{int(item['larg'])}x{int(item['alt'])}) - Tampo {item['mat_tampo']} / Prat. {item['mat_base']}"
                st.session_state.carrinho.append(item)
                st.rerun()

        st.markdown("### 🛒 Carrinho")
        if st.session_state.carrinho:
            for idx, it in enumerate(st.session_state.carrinho):
                st.info(f"**{it['qtd']}x** {it['desc_carrinho']}")
            if st.button("❌ Limpar Carrinho", use_container_width=True):
                st.session_state.carrinho = []
                st.rerun()
        else:
            st.warning("Carrinho vazio.")

    # ----------------- TELA DE RESULTADOS DO PROJETISTA -----------------
    with col_dir:
        st.subheader("📊 Dashboard de Fabricação")
        if st.session_state.carrinho:
            dados_para_df = []
            resumo_geral_chapas = {}
            resumo_geral_tubos = {}
            peso_total_pedido = 0.0

            for item in st.session_state.carrinho:
                pecas_base = []
                if item["tipo"] == "parametrico":
                    if "ESTANTE" in item["tampo_cod"]:
                        tipo_est = "LISA" if "LISA" in item["tampo_cod"] else "GRADEADA"
                        pb_cruas = calcular_estante_parametrica(item["comp"], item["larg"], item["alt"], planos=item.get("planos", 4), tipo=tipo_est, material=item["mat_tampo"])
                    elif item["tampo_cod"] == "PRAT_PAREDE":
                        pb_cruas = calcular_prateleira_parede(item["comp"], item["larg"], item["mat_tampo"])
                    else:
                        pb_cruas = calcular_mesa_parametrica(item["comp"], item["larg"], item["alt"], item["tampo_cod"], item["base_cod"])

                    molde_reforco = None
                    for p in pb_cruas:
                        if p["CÓDIGO"] in ["SAPATA", "PARAFUSO", "CUBA_PADRAO"]:
                            p["MAT_CUSTOM"] = "-"
                        else:
                            is_reforco = any(x in p["DESC"].upper() for x in ["REFORÇO", "REFORCO", "OMEGA"])
                            if is_reforco and "MAIOR" not in p["DESC"].upper(): molde_reforco = p.copy()
                            if item["tampo_cod"] == "LISA" and is_reforco and "MAIOR" in p["DESC"].upper(): continue
                                
                            if is_reforco:
                                is_traseiro = False; is_prat = False; is_plano = False; is_tampo = False
                                if "MAIOR" in p["DESC"].upper():
                                    p["DESC"] = "REFORÇO TRASEIRO DO TAMPO"; is_traseiro = True
                                elif "PRAT" in p["DESC"].upper() or "BASE" in p["DESC"].upper():
                                    p["DESC"] = "REFORÇO PRATELEIRA"; is_prat = True
                                elif "PLANO" in p["DESC"].upper() or "ESTANTE" in item["tampo_cod"]:
                                    p["DESC"] = "REFORÇO PLANO"; is_plano = True
                                else:
                                    p["DESC"] = "REFORÇO TAMPO"; is_tampo = True
                                    
                                if is_traseiro:
                                    p["MAT_CUSTOM"] = item["mat_tampo"]
                                    p["ESP"] = 0.8
                                elif is_prat:
                                    p["MAT_CUSTOM"] = item.get("mat_ref_prat", "INOX 430")
                                    try: p["ESP"] = float(item.get("esp_ref_prat", "0.8"))
                                    except: p["ESP"] = 0.8
                                    qtd = item.get("qtd_ref_prat", "Padrão")
                                    if qtd != "Padrão":
                                        if qtd == "0": continue
                                        mult = 2 if "DUPLA" in item.get("base_cod", "") else 1
                                        p["QTD"] = int(qtd) * mult
                                elif is_plano:
                                    p["MAT_CUSTOM"] = item.get("mat_ref", "INOX 430")
                                    try: p["ESP"] = float(item.get("esp_ref", "0.8"))
                                    except: p["ESP"] = 0.8
                                    qtd = item.get("qtd_ref", "Padrão")
                                    if qtd != "Padrão":
                                        if qtd == "0": continue
                                        p["QTD"] = int(qtd) * item.get("planos", 4)
                                elif is_tampo:
                                    p["MAT_CUSTOM"] = item.get("mat_ref", "INOX 430")
                                    try: p["ESP"] = float(item.get("esp_ref", "0.8"))
                                    except: p["ESP"] = 0.8
                                    qtd = item.get("qtd_ref", "Padrão")
                                    if qtd != "Padrão":
                                        if qtd == "0": continue
                                        p["QTD"] = int(qtd)
                            else:
                                e_base = False
                                if "TUBO" in p["CÓDIGO"] or any(x in p["DESC"].upper() for x in ["PRAT", "GRADE", "PERNA", "CONTRA", "TRAVESSA", "COLUNA"]):
                                    e_base = True
                                if item["tampo_cod"] == "PRAT_PAREDE": e_base = False
                                    
                                p["MAT_CUSTOM"] = item["mat_base"] if e_base else item["mat_tampo"]
                                esp_ap = item["esp_base"] if e_base else item["esp_tampo"]
                                if esp_ap != "Esp. Padrão":
                                    try: p["ESP"] = float(esp_ap)
                                    except: pass
                            
                            if p.get("MAT_CUSTOM") != "-" and "TUBO" not in p["CÓDIGO"]:
                                if p["COMP PL"] != "-" and p["LARG PL"] != "-":
                                    try: p["PESO UNIT"] = float(p["COMP PL"]) * float(p["LARG PL"]) * float(p["ESP"]) * 0.000008
                                    except: pass
                        pecas_base.append(p)
                    
                    if "PRAT" in item.get("base_cod", "") and not any(p["DESC"] == "REFORÇO PRATELEIRA" for p in pecas_base):
                        if molde_reforco:
                            ref_prat = molde_reforco.copy()
                            ref_prat["DESC"] = "REFORÇO PRATELEIRA"
                            ref_prat["MAT_CUSTOM"] = item.get("mat_ref_prat", "INOX 430")
                            try: ref_prat["ESP"] = float(item.get("esp_ref_prat", "0.8"))
                            except: ref_prat["ESP"] = 0.8
                            qtd = item.get("qtd_ref_prat", "Padrão")
                            mult = 2 if "DUPLA" in item.get("base_cod", "") else 1
                            if qtd != "Padrão":
                                if qtd == "0": ref_prat = None
                                else: ref_prat["QTD"] = int(qtd) * mult
                            else: ref_prat["QTD"] = ref_prat.get("QTD", 1) * mult
                                
                            if ref_prat:
                                if ref_prat["COMP PL"] != "-" and ref_prat["LARG PL"] != "-":
                                    try: ref_prat["PESO UNIT"] = float(ref_prat["COMP PL"]) * float(ref_prat["LARG PL"]) * float(ref_prat["ESP"]) * 0.000008
                                    except: pass
                                pecas_base.append(ref_prat)

                for p in pecas_base:
                    qtd_final = p["QTD"] * item["qtd"]
                    peso_total_final = p["PESO UNIT"] * qtd_final
                    mat_peca = p.get("MAT_CUSTOM", "-")
                    med = f"{p['COMP PL']}x{p['LARG PL']}" if p['LARG PL'] != "-" else f"{p['COMP PL']}mm" if p['COMP PL'] != "-" else "-"
                    esp_str = str(p['ESP']).replace('.', ',') if p['ESP'] > 0 else "-"

                    if "CHAPA" in p["CÓDIGO"] or (p["COMP PL"] != "-" and p["LARG PL"] != "-"):
                        if p["ESP"] > 0:
                            chave = (mat_peca, p["ESP"])
                            resumo_geral_chapas[chave] = resumo_geral_chapas.get(chave, 0) + peso_total_final
                    elif "TUBO" in p["CÓDIGO"]:
                        resumo_geral_tubos[p["CÓDIGO"]] = resumo_geral_tubos.get(p["CÓDIGO"], 0) + (float(p["COMP PL"]) * qtd_final)

                    dados_para_df.append({
                        "ITEM": f"Item {item['num']}", "QTD": qtd_final, "CÓDIGO PEÇA": p['CÓDIGO'], 
                        "DESCRIÇÃO": p['DESC'], "MATERIAL": mat_peca, "ESPESSURA": esp_str, 
                        "MEDIDA CORTE": med, "PESO UNIT (KG)": round(p['PESO UNIT'],2), "PESO TOTAL (KG)": round(peso_total_final,2)
                    })

            custo_tot_chapas = sum([peso * st.session_state.custo_chapa.get(mat, {}).get(esp, 0.0) for (mat, esp), peso in resumo_geral_chapas.items()])
            custo_tot_tubos = sum([(c_total / 1000.0) * st.session_state.custo_tubo.get(cod, 0.0) for cod, c_total in resumo_geral_tubos.items()])
            custo_total_geral = custo_tot_chapas + custo_tot_tubos

            st.success(f"💰 **Custo Total Fabril:** R$ {custo_total_geral:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            
            g1, g2 = st.columns(2)
            with g1:
                st.markdown("**📦 Consumo de Chapas**")
                for (mat, esp), peso in sorted(resumo_geral_chapas.items()):
                    st.write(f"• {mat} {esp}mm: **{peso:.2f} KG**")
            with g2:
                st.markdown("**📏 Consumo de Tubos**")
                for cod, c_total in resumo_geral_tubos.items():
                    mts = c_total / 1000.0
                    barras = mts / 6.0
                    st.write(f"• {cod}: **{mts:.2f} m** ({barras:.1f} un)")

            df_lista = pd.DataFrame(dados_para_df)
            st.dataframe(df_lista, use_container_width=True, hide_index=True)

            if not df_lista.empty:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_lista.to_excel(writer, index=False, sheet_name='Lista_Corte')
                excel_data = output.getvalue()
                
                st.download_button(
                    label="📥 Baixar Excel do PCP",
                    data=excel_data,
                    file_name=f"PCP_Pedido_{pedido_num}.xlsx" if pedido_num else "PCP_Orcamento.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

# ----------------- ABA PCP (ARQUIVOS LOCAIS / REDE) -----------------
with tab_pcp:
    st.subheader("⚙️ Gerador de PCP (Pastas de Rede / Físicas)")
    st.info("⚠️ *Aviso: O Streamlit lerá as pastas abaixo se você estiver executando o comando 'streamlit run' em uma máquina que tenha acesso à rede (Ex: \\Servidor).*")

    col1, col2 = st.columns(2)
    pasta_pdf = col1.text_input("Pasta Origem (PDFs):", value=r"\\Servidor\aco_nobre\5 - ENG DESENV\00-PROJETOS\NOVA CODIFICAÇÃO\PDF")
    pasta_dxf = col2.text_input("Pasta Origem (DXFs):", value=r"\\Servidor\aco_nobre\5 - ENG DESENV\00-PROJETOS\NOVA CODIFICAÇÃO\DXF")
    pasta_destino = st.text_input("Pasta Destino (Resultados):", placeholder=r"C:\Users\Acn.Gomes\Downloads")

    st.markdown("---")
    modo = st.radio("Modo de Processamento:", ["Processamento em Lote", "Atualização Avulsa"], horizontal=True)

    planilha_path = ""
    ped_avulso = item_avulso = cod_avulso = qtd_avulso = ""

    if modo == "Processamento em Lote":
        planilha_path = st.text_input("Caminho do arquivo Excel Comercial (.xlsx):", placeholder=r"C:\Users\Acn.Gomes\Downloads\comercial.xlsx")
    else:
        c1, c2, c3, c4 = st.columns(4)
        ped_avulso = c1.text_input("Nº Pedido")
        item_avulso = c2.text_input("Item")
        cod_avulso = c3.text_input("Código Pai")
        qtd_avulso = c4.text_input("Qtd")

    c_opt1, c_opt2, c_opt3 = st.columns(3)
    chk_pdf = c_opt1.checkbox("Unir PDFs e DXFs", value=True)
    chk_conf = c_opt2.checkbox("Gerar Conferência", value=True)
    chk_pcp = c_opt3.checkbox("Atualizar Excel", value=True)

    if st.button("🚀 INICIAR PROCESSAMENTO PCP", type="primary", use_container_width=True):
        if not pasta_pdf or not pasta_destino:
            st.error("Preencha as pastas de Origem (PDFs) e Destino.")
        else:
            # Emulador do Log Visual para o Streamlit
            log_placeholder = st.empty()
            class StLogger:
                def __init__(self, ph):
                    self.ph = ph
                    self.log = ""
                def write(self, txt):
                    self.log += txt
                    self.ph.text(self.log)
            logger = StLogger(log_placeholder)
            
            opcoes = {"pdf": chk_pdf, "conf": chk_conf, "pcp": chk_pcp}
            
            if modo == "Processamento em Lote":
                if not planilha_path:
                    st.error("Preencha o caminho da planilha para o lote.")
                else:
                    try:
                        pedidos, comerciais = ler_planilha_entrada(planilha_path)
                        mapa_pdf = escanear_pasta(pasta_pdf)
                        mapa_dxf = escanear_pasta(pasta_dxf) if pasta_dxf else {}
                        
                        logger.write("▶ Modo Lote Iniciado...\n")
                        for num_pedido, itens in pedidos.items():
                            pasta_pedido = os.path.join(pasta_destino, num_pedido)
                            caminho_global = os.path.join(pasta_pedido, f"LISTA-ITENS-{num_pedido}.xlsx")
                            for n_item, cod_pai in itens:
                                dados = comerciais.get(f"{num_pedido}_{n_item}_{cod_pai}", {})
                                dados_pcp = processar_item_unico(n_item, cod_pai, num_pedido, dados, pasta_pedido, pasta_pdf, pasta_dxf, mapa_pdf, mapa_dxf, opcoes, logger)
                                if opcoes["pcp"] and dados_pcp:
                                    atualizar_excel_inteligente(dados_pcp, caminho_global, n_item)
                                    logger.write(f"  ✅ Tabela do Item {n_item} injetada na base global.\n")
                        logger.write("\n✅ LOTE CONCLUÍDO!\n")
                        st.success("Lote concluído com sucesso!")
                    except Exception as e:
                        logger.write(f"\n❌ ERRO CRÍTICO: {str(e)}\n")
                        st.error(f"Erro no processamento: {str(e)}")
            else:
                if not ped_avulso or not item_avulso or not cod_avulso:
                    st.error("Preencha Pedido, Item e Código para a atualização avulsa.")
                else:
                    try:
                        logger.write(f"▶ Modo Avulso Iniciado para Item {item_avulso} (Cod: {cod_avulso})...\n")
                        mapa_pdf = escanear_pasta(pasta_pdf)
                        mapa_dxf = escanear_pasta(pasta_dxf) if pasta_dxf else {}
                        dados_mock = {"cliente": "AVULSO", "qtd": qtd_avulso if qtd_avulso else "1", "desc": "Processamento Isolado"}
                        pasta_pedido = os.path.join(pasta_destino, ped_avulso)
                        caminho_global = os.path.join(pasta_pedido, f"LISTA-ITENS-{ped_avulso}.xlsx")
                        
                        dados_pcp = processar_item_unico(item_avulso, normalizar_codigo(cod_avulso), ped_avulso, dados_mock, pasta_pedido, pasta_pdf, pasta_dxf, mapa_pdf, mapa_dxf, opcoes, logger)
                        if opcoes["pcp"] and dados_pcp:
                            atualizar_excel_inteligente(dados_pcp, caminho_global, item_avulso)
                            logger.write(f"  ✅ Planilha Global do Pedido {ped_avulso} ATUALIZADA.\n")
                        logger.write("\n✅ ITEM AVULSO CONCLUÍDO!\n")
                        st.success("Item avulso processado com sucesso!")
                    except Exception as e:
                        logger.write(f"\n❌ ERRO CRÍTICO: {str(e)}\n")
                        st.error(f"Erro no processamento: {str(e)}")
