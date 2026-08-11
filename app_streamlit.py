import streamlit as st
import os
import re
import shutil
import math
import io
import pandas as pd
import pdfplumber
import fitz  # PyMuPDF

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from pypdf import PdfWriter

# Importa a sua engenharia blindada do arquivo local!
from regras_engenharia import calcular_mesa_parametrica, calcular_estante_parametrica, calcular_prateleira_parede

# Importando bibliotecas de Nesting e Gráficos para o Projetista
try:
    from rectpack import newPacker, PackingMode, PackingBin
    import matplotlib.pyplot as plt
    import matplotlib.patches as patches
    HAS_NESTING = True
except ImportError:
    HAS_NESTING = False

# ==========================================
# 1. CONFIGURAÇÃO DA PÁGINA E CSS CUSTOMIZADO
# ==========================================
# O "initial_sidebar_state='collapsed'" garante que não sobre nenhum resquício da barra lateral
st.set_page_config(page_title="AçoNobre ERP", layout="wide", page_icon="🏭", initial_sidebar_state="collapsed")

# Injeção de CSS para deixar o visual idêntico às imagens (Cards escuros, Banner Azul, Textos Amarelos)
st.markdown("""
<style>
    .box-azul {
        background-color: #2980b9;
        color: white;
        padding: 25px;
        border-radius: 8px;
        text-align: center;
        margin-bottom: 20px;
    }
    .box-azul h4 { color: #ecf0f1; margin: 0; font-size: 16px; font-weight: normal; letter-spacing: 1px; }
    .box-azul h1 { color: white; margin: 5px 0 0 0; font-size: 45px; font-weight: bold; }
    
    .card-dark {
        background-color: #2b2b2b;
        padding: 20px;
        border-radius: 8px;
        margin-bottom: 15px;
        border: 1px solid #3d3d3d;
    }
    .card-dark h4 { color: #ffffff; font-size: 15px; font-weight: bold; display: flex; align-items: center; gap: 10px; margin-bottom: 15px;}
    .card-dark p { color: #dddddd; font-size: 14px; margin-bottom: 10px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;}
    .card-dark hr { border-top: 1px solid #444; margin: 15px 0; }
    
    .item-title { color: #f1c40f; font-weight: bold; font-size: 16px; margin-bottom: 15px; }
    .item-linha { font-family: 'Courier New', Courier, monospace; font-size: 14px; color: #cccccc; margin-bottom: 10px; margin-left: 10px; line-height: 1.5; }
    
    /* Centraliza os radio buttons do menu superior */
    div.row-widget.stRadio > div { flex-direction: row; justify-content: center; background-color: #1e1e1e; padding: 10px; border-radius: 10px; }
</style>
""", unsafe_allow_html=True)

# Inicializa variáveis de memória
if 'carrinho' not in st.session_state: st.session_state.carrinho = []
if 'pecas_temp_composto' not in st.session_state: st.session_state.pecas_temp_composto = []

mapa_tampo = {"Mesa Lisa": "LISA", "Mesa com Encosto": "ENCOSTO", "Pia com Cuba": "PIA", "Prateleira Parede": "PRAT_PAREDE", "Estante Lisa": "ESTANTE_LISA", "Estante Gradeada": "ESTANTE_GRADEADA"}
mapa_base = {"Contraventamento": "CONTRAVENTAMENTO", "Prat. Lisa": "PRAT_LISA", "Prat. Lisa Dupla": "PRAT_LISA_DUPLA", "Prat. Gradeada": "PRAT_GRADEADA", "Prat. Gradeada Dupla": "PRAT_GRADEADA_DUPLA"}

@st.cache_data
def carregar_tabela_precos():
    custo_chapa = {"INOX 304": {0.6: 29.01, 0.8: 33.16, 1.0: 30.74, 1.2: 28.66, 1.5: 29.00, 2.0: 30.96}, "INOX 430": {0.6: 23.32, 0.8: 19.41, 1.0: 18.91, 1.2: 19.42, 1.5: 19.38, 2.0: 19.38}}
    custo_tubo = {"TUBO-RED-38": 18.47, "TUBO-RED-25": 12.08}
    try:
        df = pd.read_excel("CHAPAS.xlsx", sheet_name="CHAPAS")
        custo_chapa["INOX 304"] = dict(zip(df[df['Aisi']==304]['Espessura'], df[df['Aisi']==304]['Custo última compra']))
        custo_chapa["INOX 430"] = dict(zip(df[df['Aisi']==430]['Espessura'], df[df['Aisi']==430]['Custo última compra']))
        df_t = pd.read_excel("CHAPAS.xlsx", sheet_name="TUBOS")
        for _, r in df_t.iterrows():
            if "016.201.020" in str(r['Código']): custo_tubo["TUBO-RED-38"] = float(r['Custo última compra'])
            elif "016.201.014" in str(r['Código']): custo_tubo["TUBO-RED-25"] = float(r['Custo última compra'])
    except: pass
    return custo_chapa, custo_tubo

custo_chapa, custo_tubo = carregar_tabela_precos()

# ==========================================
# 2. FUNÇÕES DO MOTOR PCP E PDF
# ==========================================
def calcular_distancia(valA, valB, dict_coords):
    min_dist = float('inf')
    for pA in dict_coords[valA]:
        for pB in dict_coords[valB]:
            dist = math.sqrt((pA['x'] - pB['x'])**2 + (pA['y'] - pB['y'])**2)
            if dist < min_dist: min_dist = dist
    return min_dist

def extrair_dados_tecnicos(caminho_pdf):
    dados = {"massa": "", "esp": "", "tipo": "CHAPA LISA", "c_plan": "", "l_plan": "", "c_fin": "", "l_fin": "", "a_fin": "", "planificado": "-", "finalizado": "-"}
    numeros_dict = {} 
    coord_secao = None
    doc = fitz.open(caminho_pdf)
    pagina = doc[0]
    nome_arquivo = caminho_pdf.replace('\\', '/').split('/')[-1]
    codigo_peca = nome_arquivo.lower().replace('.pdf', '')
    texto_completo = pagina.get_text("text")
    if re.search(r'(DOBRA|PARA BAIXO|PARA CIMA|90°)', texto_completo, re.IGNORECASE): dados["tipo"] = "CHAPA DOBRADA"
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
                    if re.search(r'(SEÇÃO|SECAO|DETALHE)', texto, re.IGNORECASE): coord_secao = (x_centro, y_centro)
                    if texto == codigo_peca or texto.replace(',', '.') == codigo_peca.replace(',', '.'): continue
                    if re.match(r'^\d+([.,]\d+)?$', texto):
                        val = float(texto.replace(',', '.'))
                        if 5.0 < val < 3000.0:
                            if str(val) != dados["massa"] and str(val) != dados["esp"]:
                                if val not in numeros_dict: numeros_dict[val] = []
                                numeros_dict[val].append({'x': x_centro, 'y': y_centro})

    nums = sorted(list(numeros_dict.keys()), reverse=True)
    if dados["tipo"] != "CHAPA LISA" and len(nums) >= 3:
        if nums[0] > nums[1] * 1.3:
            comp_plan = comp_fin = nums[0]
            grupos = []
            for n in nums[1:]:
                if not grupos or n < grupos[-1][0] * 0.85: grupos.append([n])
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
                        menor_dist = dist; alt_fin = val
            else: alt_fin = h_cands[0] 
                
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
    for nome_arquivo in sorted(os.listdir(pasta)):
        if nome_arquivo.lower().endswith(".pdf"):
            codigo = extrair_codigo_do_nome_arquivo(nome_arquivo)
            if codigo: mapa[codigo] = nome_arquivo
    return mapa

def ler_planilha_entrada(df):
    df = df.astype(str)
    pedidos, dados_comerciais = {}, {}
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
                                cod_peca = c_norm; break
                    else:
                        for c, p in pecas_dict.items():
                            desc_atual = p["desc"].upper()
                            if desc_atual != "-" and len(desc_atual) > 3:
                                if desc_atual in linha_str or any(desc_atual in col for col in cols_validas):
                                    cod_peca = c; break
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
                            val_str = match_esp.group(1); tem_mm = bool(match_esp.group(2)); is_decimal = (',' in val_str or '.' in val_str); is_num_item_ou_qtd = (val_str == p["n"] or val_str == p["qtd"])
                            if tem_mm or is_decimal or not is_num_item_ou_qtd:
                                try:
                                    f_val = float(val_str.replace(',', '.'))
                                    if 0.3 <= f_val <= 20.0: 
                                        if p["esp"] == "-" or (tem_mm and "MM" not in str(p.get("_raw_esp", ""))):
                                            p["esp"] = val_str; p["_raw_esp"] = col_up
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
    fonte_p = Font(name='Calibri', size=11); fonte_neg = Font(name='Calibri', size=11, bold=True); fonte_titulo = Font(name='Calibri', size=14, bold=True)
    borda = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    alinhamento_centro = Alignment(horizontal="center", vertical="center"); alinhamento_esq = Alignment(horizontal="left", vertical="center")
    
    linha = 2
    ws.merge_cells(f"B{linha}:J{linha}")
    cliente = dados_comerciais.get('cliente', 'AVULSO')
    ws.cell(linha, 2, f"PEDIDO {num_pedido} - {cliente}").font = fonte_titulo
    ws.cell(linha, 2).alignment = alinhamento_centro; ws.cell(linha, 2).fill = PatternFill(start_color="DDDDDD", fill_type="solid")
    for col in range(2, 11): ws.cell(linha, col).border = borda
    ws.row_dimensions[linha].height = 18.75
    linha += 1
    
    ws.cell(linha, 2, "OF").font = fonte_neg; ws.cell(linha, 2).alignment = alinhamento_centro
    ws.cell(linha, 3, "CÓD").font = fonte_neg; ws.cell(linha, 3).alignment = alinhamento_centro
    ws.merge_cells(f"D{linha}:I{linha}"); ws.cell(linha, 4, "DESCRIÇÃO").font = fonte_neg; ws.cell(linha, 4).alignment = alinhamento_centro
    ws.cell(linha, 10, "QTD").font = fonte_neg; ws.cell(linha, 10).alignment = alinhamento_centro
    for col in range(2, 11): ws.cell(linha, col).border = borda
    ws.row_dimensions[linha].height = 18.75; linha += 1
    
    ws.cell(linha, 2, dados_comerciais.get('of', '-')).font = fonte_neg; ws.cell(linha, 2).alignment = alinhamento_centro
    ws.cell(linha, 3, cod_produto).font = fonte_neg; ws.cell(linha, 3).alignment = alinhamento_centro
    ws.merge_cells(f"D{linha}:I{linha}"); ws.cell(linha, 4, dados_comerciais.get('desc', '-')).font = fonte_neg; ws.cell(linha, 4).alignment = alinhamento_esq
    try: qtd_pai = int(float(str(dados_comerciais.get('qtd', '1')).replace(',', '.')))
    except: qtd_pai = 1
    ws.cell(linha, 10, qtd_pai).font = fonte_neg; ws.cell(linha, 10).alignment = alinhamento_centro
    for col in range(2, 11): ws.cell(linha, col).border = borda
    ws.row_dimensions[linha].height = 18.75; linha += 1
    
    headers = ["TIPO", "Nº", "CÓDIGO", "ITENS", "QTD", "QTD OF", "ENVIADA", "SALDO", "✔️"]
    for i, h in enumerate(headers, 2):
        cel = ws.cell(linha, i, h); cel.font = fonte_neg; cel.fill = PatternFill(start_color="DDDDDD", fill_type="solid")
        cel.border = borda; cel.alignment = alinhamento_centro
    ws.row_dimensions[linha].height = 18.75; linha += 1
    
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
        ws.row_dimensions[linha].height = 18.75; linha += 1

    linha += 2; ws.cell(linha, 2, "Conferido por: _________________________________________   Data: ____/____/____").font = fonte_neg
    linha += 2; ws.cell(linha, 2, "Recebido por:  _________________________________________   Data: ____/____/____").font = fonte_neg

    for col_let, w in zip(['A','B','C','D','E','F','G','H','I','J'], [2, 13, 14, 14, 40, 8, 12, 12, 10, 6]): ws.column_dimensions[col_let].width = w
    ws.page_setup.fitToPage = True; ws.page_setup.orientation = "landscape" 
    wb.save(os.path.join(pasta_item, f"LISTA-CONFERÊNCIA-{num_pedido}-{n_item}-{cod_produto}.xlsx"))

def atualizar_excel_inteligente(dados_novos, caminho_global, num_item):
    df_novo = pd.DataFrame(dados_novos)
    if os.path.exists(caminho_global):
        try:
            df_existente = pd.read_excel(caminho_global, dtype=str)
            df_existente = df_existente[df_existente["ITEM"] != str(num_item)]
            df_final = pd.concat([df_existente, df_novo], ignore_index=True)
        except Exception: df_final = df_novo
    else: df_final = df_novo

    formulas_peso_bruto = []
    for i in range(len(df_final)):
        linha_excel = i + 2 
        formulas_peso_bruto.append(f"=S{linha_excel}*T{linha_excel}*N{linha_excel}*0.000008")
    df_final["PESO BRUTO"] = formulas_peso_bruto

    with pd.ExcelWriter(caminho_global, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False)
        planilha_g = writer.sheets['Sheet1']
        for i in range(1, len(df_final.columns) + 1): planilha_g.column_dimensions[get_column_letter(i)].width = 22

def to_comma_str(val):
    if not val or val == "-": return ""
    try: 
        fval = float(str(val).replace(',', '.'))
        if fval.is_integer(): return str(int(fval))
        return str(fval).replace('.', ',')
    except: return str(val).replace('.', ',')

def processar_item_unico(n_item, cod_pai, num_pedido, dados_comerciais, pasta_pedido, origem, mapa, opcoes, log_widget):
    log_widget.text(f"▶ Inciando Item {n_item} -> Produto: {cod_pai}")
    pasta_item = os.path.join(pasta_pedido, f"{n_item} - {cod_pai}")
    os.makedirs(pasta_item, exist_ok=True)
    
    if cod_pai not in mapa:
        log_widget.text(f"❌ ERRO: PDF pai ({cod_pai}) não encontrado na pasta origem.")
        return []

    caminho_pai = os.path.join(origem, mapa[cod_pai])
    caminho_pai_copiado = os.path.join(pasta_item, mapa[cod_pai])
    shutil.copy2(caminho_pai, caminho_pai_copiado)
    
    log_widget.text(f"  📄 Lendo ficheiro pai: {mapa[cod_pai]}")
    pecas = extrair_tabela_de_materiais(caminho_pai)
    log_widget.text(f"  🔍 Encontradas {len(pecas)} peças. Extraindo dados (PCP/Geometria)...")
    
    dados_pcp = []
    merger = PdfWriter() if opcoes["pdf"] else None
    if merger: merger.append(caminho_pai_copiado)
    
    if opcoes["pcp"] or opcoes["pdf"]:
        for p in pecas:
            eng_dados = {}
            if p['cod'] in mapa:
                caminho_filho = os.path.join(origem, mapa[p['cod']])
                if opcoes["pdf"]:
                    caminho_filho_copiado = os.path.join(pasta_item, mapa[p['cod']])
                    shutil.copy2(caminho_filho, caminho_filho_copiado)
                    merger.append(caminho_filho_copiado)
                if opcoes["pcp"]:
                    try: eng_dados = extrair_dados_tecnicos(caminho_filho)
                    except Exception: pass
            
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
                    "PESO BRUTO": "" 
                })
    
    if opcoes["pdf"] and merger:
        merger.write(os.path.join(pasta_item, f"{cod_pai}-COMPLETO.pdf"))
        merger.close()
        log_widget.text("  ✅ PDF Unificado gerado.")

    if opcoes["conf"]:
        gerar_conferencia_item(n_item, cod_pai, pecas, pasta_item, num_pedido, dados_comerciais)
        log_widget.text("  ✅ Lista de Conferência gerada.")
    return dados_pcp

# ==========================================
# 3. CABEÇALHO E MENU SUPERIOR (SEM SIDEBAR)
# ==========================================
st.markdown("<h2 style='text-align: center; color: #f1c40f; margin-bottom: 0;'>🏭 AçoNobre ERP</h2>", unsafe_allow_html=True)

# Tira a barra lateral inteira e transforma o menu de navegação em abas horizontais no topo
menu = st.radio(
    "Navegação Principal", 
    ["🛒 Projetista Virtual", "⚙️ Gerador de PCP"], 
    horizontal=True, 
    label_visibility="collapsed"
)
st.markdown("---")

# ==========================================
# 4. TELA: PROJETISTA VIRTUAL
# ==========================================
if menu == "🛒 Projetista Virtual":
    col_esq, col_dir = st.columns([1.3, 2.7])
    
    # ------------------ COLUNA ESQUERDA (INPUTS) ------------------
    with col_esq:
        st.markdown("#### 📦 DADOS DO PEDIDO")
        inp_pedido_nome = st.text_input("Nº do Pedido (Ex: 6885)", key="inp_pedido_global")
        
        aba_param, aba_comp, aba_livre = st.tabs(["📐 Catálogo Paramétrico", "🧩 Item Composto", "☕ Peça Avulsa"])
        
        # --- ABA: CATÁLOGO PARAMÉTRICO ---
        with aba_param:
            c_qtd, c_planos = st.columns(2)
            qtd = c_qtd.number_input("QTD", min_value=1, value=1)
            planos = c_planos.number_input("Planos", min_value=1, value=4)
            
            c1, c2, c3 = st.columns(3)
            comp = c1.number_input("C(mm)", min_value=100.0, value=1200.0, step=50.0)
            larg = c2.number_input("L(mm)", min_value=100.0, value=600.0, step=50.0)
            alt = c3.number_input("A(mm)", min_value=10.0, value=900.0, step=50.0)
            
            tampo_nome = st.selectbox("Tampo", list(mapa_tampo.keys()), label_visibility="collapsed")
            base_nome = st.selectbox("Base", list(mapa_base.keys()), label_visibility="collapsed")
            
            c_mat1, c_esp1 = st.columns([1, 1.5])
            mat_tampo = c_mat1.selectbox("Tampo Mat", ["INOX 304", "INOX 430"])
            esp_tampo = c_esp1.selectbox("Tampo Esp", ["Esp. Padrão", "0.8", "1.0", "1.2", "1.5", "2.0"])
            
            c_mat2, c_esp2 = st.columns([1, 1.5])
            mat_base = c_mat2.selectbox("Base Mat", ["INOX 304", "INOX 430"])
            esp_base = c_esp2.selectbox("Base Esp", ["Esp. Padrão", "0.8", "1.0", "1.2", "1.5", "2.0"])
            
            if st.button("➕ Adicionar ao Pedido", use_container_width=True, type="primary"):
                item = {
                    "num": len(st.session_state.carrinho) + 1, "tipo": "parametrico", "qtd": int(qtd),
                    "comp": float(comp), "larg": float(larg), "alt": float(alt), "planos": int(planos),
                    "tampo_nome": tampo_nome, "tampo_cod": mapa_tampo[tampo_nome],
                    "base_nome": base_nome, "base_cod": mapa_base[base_nome],
                    "mat_tampo": mat_tampo, "esp_tampo": esp_tampo, "mat_base": mat_base, "esp_base": esp_base
                }
                if "ESTANTE" in item["tampo_cod"]: item["desc_carrinho"] = f"{tampo_nome} {planos} Planos ({int(comp)}x{int(larg)}x{int(alt)})"
                elif item["tampo_cod"] == "PRAT_PAREDE": item["desc_carrinho"] = f"{tampo_nome} ({int(comp)}x{int(larg)})"
                else: item["desc_carrinho"] = f"{tampo_nome} c/ {base_nome} ({int(comp)}x{int(larg)}x{int(alt)})"
                st.session_state.carrinho.append(item)
                st.rerun()

        # --- ABA: ITEM COMPOSTO ---
        with aba_comp:
            c_qtd_item, c_nome_item = st.columns([1, 3])
            qtd_item_comp = c_qtd_item.number_input("QTD", min_value=1, value=1, key="qtd_comp")
            nome_item_comp = c_nome_item.text_input("Produto Modular", placeholder="Produto Modular", label_visibility="collapsed")
            
            c_nome_pc, c_qtd_pc = st.columns([3, 1])
            nome_pc = c_nome_pc.text_input("Nome Peça", placeholder="Nome Peça", label_visibility="collapsed")
            qtd_pc = c_qtd_pc.number_input("Qtd", min_value=1, value=1, key="qtd_pc_comp", label_visibility="collapsed")
            
            c_c, c_l, c_mat, c_esp = st.columns([1.5, 1.5, 2, 2])
            comp_pc = c_c.number_input("C(mm)", min_value=1.0, value=100.0, step=10.0, key="c_comp", label_visibility="collapsed")
            larg_pc = c_l.number_input("L(mm)", min_value=1.0, value=100.0, step=10.0, key="l_comp", label_visibility="collapsed")
            mat_pc = c_mat.selectbox("Mat", ["304", "430"], key="mat_comp", label_visibility="collapsed")
            esp_pc = c_esp.selectbox("Esp", ["0.6", "0.8", "1.0", "1.2", "1.5", "2.0"], index=2, key="esp_comp", label_visibility="collapsed")
            
            if st.button("⏬ Inserir Peça", use_container_width=True):
                if nome_pc:
                    st.session_state.pecas_temp_composto.append({
                        "nome": nome_pc, "qtd": int(qtd_pc), "comp": float(comp_pc), "larg": float(larg_pc),
                        "mat": f"INOX {mat_pc}", "esp": float(esp_pc)
                    })
                    st.rerun()
                
            if st.session_state.pecas_temp_composto:
                df_temp = pd.DataFrame(st.session_state.pecas_temp_composto)
                st.dataframe(df_temp[["qtd", "nome", "mat"]].rename(columns={"qtd":"QTD", "nome":"PEÇA", "mat":"MAT"}), use_container_width=True, hide_index=True)
                
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("✅ EMPACOTAR ITEM E ADD", type="primary", use_container_width=True):
                if st.session_state.pecas_temp_composto and nome_item_comp:
                    item = {
                        "num": len(st.session_state.carrinho) + 1, "tipo": "composto",
                        "nome_item": nome_item_comp, "qtd_item": int(qtd_item_comp), "qtd": int(qtd_item_comp),
                        "pecas": list(st.session_state.pecas_temp_composto),
                        "desc_carrinho": f"🧩 {nome_item_comp}"
                    }
                    st.session_state.carrinho.append(item)
                    st.session_state.pecas_temp_composto = [] 
                    st.rerun()
                else:
                    st.warning("Insira um nome para o produto e adicione peças na lista!")

        # --- ABA: PEÇA AVULSA ---
        with aba_livre:
            c_qtd_livre, c_nome_livre = st.columns([1, 3])
            qtd_livre = c_qtd_livre.number_input("QTD", min_value=1, value=1, key="qtd_livre")
            nome_livre = c_nome_livre.text_input("Nome da Peça", placeholder="Ex: Chapa de Proteção")
            
            c_c_livre, c_l_livre = st.columns(2)
            comp_livre = c_c_livre.number_input("C Planif (mm)", min_value=1.0, value=500.0, step=10.0, key="c_livre")
            larg_livre = c_l_livre.number_input("L Planif (mm)", min_value=1.0, value=500.0, step=10.0, key="l_livre")
            
            c_mat_livre, c_esp_livre = st.columns(2)
            mat_livre = c_mat_livre.selectbox("Material", ["INOX 304", "INOX 430"], key="mat_livre")
            esp_livre = c_esp_livre.selectbox("Espessura", ["0.6", "0.8", "1.0", "1.2", "1.5", "2.0"], index=2, key="esp_livre")
            
            if st.button("➕ Adicionar Peça Solta", type="primary", use_container_width=True):
                if nome_livre:
                    item = {
                        "num": len(st.session_state.carrinho) + 1, "tipo": "livre", "qtd": int(qtd_livre),
                        "nome_peca": nome_livre, "comp_pl": float(comp_livre), "larg_pl": float(larg_livre),
                        "esp": float(esp_livre), "material": mat_livre, "desc_carrinho": f"✏️ {nome_livre}"
                    }
                    st.session_state.carrinho.append(item)
                    st.rerun()

        # --- VISOR DO CARRINHO ---
        st.markdown("<hr>", unsafe_allow_html=True)
        if st.session_state.carrinho:
            df_carrinho = pd.DataFrame([{"Nº": i['num'], "QTD": i.get('qtd', i.get('qtd_item', 1)), "DESCRIÇÃO": i['desc_carrinho']} for i in st.session_state.carrinho])
            st.dataframe(df_carrinho, hide_index=True, use_container_width=True)
            
        c_btn_limpar, c_btn_gerar = st.columns([1, 2])
        if c_btn_limpar.button("❌ Limpar", use_container_width=True):
            st.session_state.carrinho = []
            st.rerun()
            
        # O botão principal que dispara o cálculo
        gerar_calc = c_btn_gerar.button("🚀 CALCULAR PROJETO", type="primary", use_container_width=True)
            
    # ------------------ COLUNA DIREITA (DASHBOARD) ------------------
    with col_dir:
        
        if not st.session_state.carrinho:
            st.info("👈 Adicione itens ao pedido no menu à esquerda e clique em **CALCULAR PROJETO**.")
        elif gerar_calc:
            with st.spinner('Acessando o Motor de Engenharia e Gerando Custos...'):
                dados_para_df = []
                pecas_para_nesting_global = [] 
                resumo_geral_chapas = {}
                resumo_geral_tubos = {}
                linhas_resumo_excel = []
                peso_total_pedido = 0.0
                custo_total_geral = 0.0
                
                # Prepara os blocos de HTML para a aba Dashboard (Estilo idêntico à imagem)
                html_detalhamento = "<h4>📄 DETALHAMENTO DE FABRICAÇÃO POR ITEM</h4>"
                
                for item in st.session_state.carrinho:
                    if item["tipo"] == "parametrico":
                        if "ESTANTE" in item["tampo_cod"]:
                            tipo_est = "LISA" if "LISA" in item["tampo_cod"] else "GRADEADA"
                            pb_cruas = calcular_estante_parametrica(item["comp"], item["larg"], item["alt"], planos=item["planos"], tipo=tipo_est, material=item["mat_tampo"])
                        elif item["tampo_cod"] == "PRAT_PAREDE":
                            pb_cruas = calcular_prateleira_parede(item["comp"], item["larg"], item["mat_tampo"])
                        else:
                            pb_cruas = calcular_mesa_parametrica(item["comp"], item["larg"], item["alt"], item["tampo_cod"], item["base_cod"])
                    
                    elif item["tipo"] == "composto":
                        pb_cruas = []
                        for pt in item["pecas"]:
                            peso_un = pt["comp"] * pt["larg"] * pt["esp"] * 0.000008
                            pb_cruas.append({"CÓDIGO": "CHAPA_LIVRE", "DESC": pt["nome"], "QTD": pt["qtd"], "COMP PL": pt["comp"], "LARG PL": pt["larg"], "ESP": pt["esp"], "PESO UNIT": peso_un, "MAT_CUSTOM": pt["mat"]})
                    
                    elif item["tipo"] == "livre":
                        pb_cruas = []
                        peso_un = item["comp_pl"] * item["larg_pl"] * item["esp"] * 0.000008
                        pb_cruas.append({"CÓDIGO": "CHAPA_LIVRE", "DESC": item["nome_peca"], "QTD": 1, "COMP PL": item["comp_pl"], "LARG PL": item["larg_pl"], "ESP": item["esp"], "PESO UNIT": peso_un, "MAT_CUSTOM": item["material"]})

                    peso_item_total = 0.0
                    custo_unit_item_chapa = 0.0
                    custo_unit_item_tubo = 0.0

                    # Inicia a "Carta" de Detalhamento do Item
                    html_detalhamento += f"<div class='card-dark'><div class='item-title'>ITEM {item['num']}: {item['qtd']}x {item['desc_carrinho']}</div>"
                    
                    for p in pb_cruas:
                        qtd_final = p["QTD"] * item["qtd"]
                        
                        if item["tipo"] == "parametrico":
                            if "TUBO" in p["CÓDIGO"]: p["MAT_CUSTOM"] = "INOX 201 ESC"; p["ESP"] = 1.2
                            elif p["CÓDIGO"] in ["SAPATA", "PARAFUSO", "CUBA_PADRAO"]: p["MAT_CUSTOM"] = "-"
                            else:
                                if "MAT_CUSTOM" not in p: p["MAT_CUSTOM"] = item["mat_tampo"]
                        
                        custo_peca_un = 0.0
                        mat_peca = p["MAT_CUSTOM"]
                        
                        # Custo Chapa
                        if "CHAPA" in p["CÓDIGO"]:
                            peso_peca = p.get("PESO UNIT", 0) * qtd_final
                            if p.get("COMP PL", "-") != "-" and p.get("LARG PL", "-") != "-":
                                for _ in range(qtd_final): pecas_para_nesting_global.append({'nome': p['DESC'], 'material': f"{mat_peca} {p['ESP']}mm", 'comp': float(p["COMP PL"]), 'larg': float(p["LARG PL"])})
                            if p.get("ESP", 0) > 0:
                                custo_peca_un = p["PESO UNIT"] * custo_chapa.get(mat_peca, {}).get(p["ESP"], 0.0)
                                custo_unit_item_chapa += custo_peca_un * p["QTD"]
                                chave = (mat_peca, p["ESP"])
                                resumo_geral_chapas[chave] = resumo_geral_chapas.get(chave, 0) + peso_peca
                            linha_med = f"{p['PESO UNIT']:.2f} KG un, | {peso_peca:.2f} KG tot,".replace('.',',')

                        # Custo Tubo
                        elif "TUBO" in p["CÓDIGO"]:
                            metros_unit = float(p.get("COMP PL", 0)) / 1000.0
                            metros_tot = metros_unit * qtd_final
                            peso_peca = 0
                            custo_peca_un = metros_unit * custo_tubo.get(p["CÓDIGO"], 0.0)
                            custo_unit_item_tubo += custo_peca_un * p["QTD"]
                            resumo_geral_tubos[p["CÓDIGO"]] = resumo_geral_tubos.get(p["CÓDIGO"], 0) + (float(p["COMP PL"]) * qtd_final)
                            linha_med = f"{metros_unit:.2f} M un, | {metros_tot:.2f} M tot,".replace('.',',')
                        else:
                            peso_peca = 0
                            linha_med = "- | -"

                        peso_item_total += peso_peca
                        mat_str = f"{mat_peca} - " if mat_peca != "-" else ""
                        custo_str = f"Custo un: R$ {custo_peca_un:.2f}".replace('.',',')
                        
                        # Linha do item no HTML
                        html_detalhamento += f"<div class='item-linha'>&#x25AB; {qtd_final}x {mat_str}{p['DESC']} | {linha_med} | {custo_str}</div>"
                        
                        dados_para_df.append({
                            "ITEM": f"Item {item['num']}", "QTD": qtd_final, "CÓDIGO": p["CÓDIGO"],
                            "DESCRIÇÃO": p["DESC"], "MATERIAL": p["MAT_CUSTOM"], "ESP": p.get("ESP", "-"),
                            "MEDIDA (C x L)": f"{p.get('COMP PL','-')} x {p.get('LARG PL','-')}", "PESO TOT (KG)": round(peso_peca, 2)
                        })
                        
                    c_unit = custo_unit_item_chapa + custo_unit_item_tubo
                    c_tot = c_unit * item["qtd"]
                    peso_total_pedido += peso_item_total
                    html_detalhamento += f"<hr><div class='item-linha' style='color:#fff;'><b>CUSTO TOTAL DO ITEM: R$ {c_tot:.2f} | Peso: {peso_item_total:.2f} KG</b></div></div>".replace('.',',')

                # Cálculo Final do Dashboard
                custo_tot_chapas = sum([peso * custo_chapa.get(mat, {}).get(esp, 0.0) for (mat, esp), peso in resumo_geral_chapas.items()])
                custo_tot_tubos = sum([(c_total / 1000.0) * custo_tubo.get(cod, 0.0) for cod, c_total in resumo_geral_tubos.items()])
                custo_total_geral = custo_tot_chapas + custo_tot_tubos
                
                # Renderiza ABAS Principais dos Resultados
                tab_lista, tab_dash = st.tabs(["✂️ Lista PCP e Mapas", "📊 Dashboard Financeiro"])
                
                with tab_dash:
                    # Banner Azul Exato da Imagem 2
                    titulo_pedido = f"CUSTO DO PEDIDO {inp_pedido_nome}" if inp_pedido_nome else "CUSTO DO PEDIDO AVULSO"
                    st.markdown(f"""
                    <div class="box-azul">
                        <h4>⏱️ {titulo_pedido}</h4>
                        <h1>R$ {custo_total_geral:,.2f}</h1>
                    </div>
                    """.replace(',', 'X').replace('.', ',').replace('X', '.'), unsafe_allow_html=True)
                    
                    # Box de Chapas Inox
                    if resumo_geral_chapas:
                        html_chapas = "<div class='card-dark'><h4>📦 CONSUMO DE CHAPAS INOX</h4>"
                        for (mat, esp), peso in sorted(resumo_geral_chapas.items()):
                            prc = custo_chapa.get(mat, {}).get(esp, 0.0)
                            ct = peso * prc
                            html_chapas += f"<p>• {mat} {esp}mm ➔ {peso:.2f} KG | (R$ {prc:.2f}/kg) | Subtotal: R$ {ct:.2f}</p>".replace('.',',')
                            linhas_resumo_excel.append({"TIPO": "CHAPA", "MATERIAL": f"{mat} {esp}mm", "QTD": f"{peso:.2f} KG".replace('.', ','), "CUSTO (R$)": round(ct, 2)})
                        html_chapas += f"<p style='font-weight:bold; margin-top:10px;'>Peso Total em Chapas: {peso_total_pedido:.2f} KG</p></div>".replace('.',',')
                        st.markdown(html_chapas, unsafe_allow_html=True)
                            
                    # Box de Tubos
                    if resumo_geral_tubos:
                        html_tubos = "<div class='card-dark'><h4>📎 CONSUMO DE TUBOS E PERFIS</h4>"
                        for cod, c_total in resumo_geral_tubos.items():
                            mts = c_total / 1000.0
                            barras = mts / 6.0
                            prc = custo_tubo.get(cod, 0.0)
                            ct = mts * prc
                            html_tubos += f"<p>• {cod} ➔ {mts:.2f} Mts ({barras:.1f} barras) | (R$ {prc:.2f}/m) | Subtotal: R$ {ct:.2f}</p>".replace('.',',')
                            linhas_resumo_excel.append({"TIPO": "TUBO", "MATERIAL": f"{cod}", "QTD": f"{mts:.2f} M ({barras:.1f} un)".replace('.', ','), "CUSTO (R$)": round(ct, 2)})
                        html_tubos += "</div>"
                        st.markdown(html_tubos, unsafe_allow_html=True)
                        
                    # Box de Nesting (Aproveitamento)
                    if HAS_NESTING and pecas_para_nesting_global:
                        html_nest = f"<div class='card-dark'><h4>🧩 ESTIMATIVA DE ENCAIXE (CHAPAS 3000x1250)</h4>"
                        agrup_nesting = {}
                        for pc in pecas_para_nesting_global:
                            chave = pc['material']
                            if chave not in agrup_nesting: agrup_nesting[chave] = []
                            agrup_nesting[chave].append((pc['comp'], pc['larg']))

                        dados_para_graficos = {}
                        for mat_nome, dimensoes in agrup_nesting.items():
                            packer = newPacker(mode=PackingMode.Offline, bin_algo=PackingBin.Global, rotation=True)
                            for idx, (c, l) in enumerate(dimensoes): packer.add_rect(width=c + 5, height=l + 5, rid=idx) 
                            for _ in range(50): packer.add_bin(width=3000, height=1250)
                            packer.pack()
                            
                            chp_usadas = len(packer)
                            area_pc = sum([c * l for c, l in dimensoes])
                            area_ch = chp_usadas * (3000 * 1250)
                            aprov = (area_pc / area_ch) * 100 if area_ch > 0 else 0
                            
                            html_nest += f"<p>• {mat_nome}: Será preciso puxar <b>{chp_usadas} chapa(s)</b> do estoque. (Aproveitamento: {aprov:.1f}%)</p>".replace('.',',')
                            dados_para_graficos[mat_nome] = packer
                            
                        html_nest += "</div>"
                        st.markdown(html_nest, unsafe_allow_html=True)
                        
                    # Printa o detalhamento completo dos itens injetando o HTML que geramos lá em cima
                    st.markdown(html_detalhamento, unsafe_allow_html=True)

                with tab_lista:
                    st.markdown("#### ✂️ Lista de Corte Detalhada")
                    df_lista = pd.DataFrame(dados_para_df)
                    st.dataframe(df_lista, hide_index=True, use_container_width=True)
                    
                    # GERADOR DE EXCEL REAL
                    df_resumo = pd.DataFrame(linhas_resumo_excel)
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df_lista.to_excel(writer, sheet_name='Lista_Corte_PCP', index=False)
                        if not df_resumo.empty: df_resumo.to_excel(writer, sheet_name='Resumo_Financeiro', index=False)
                    
                    st.download_button(label="📥 Baixar Excel Completo (PCP e Custos)", data=buffer.getvalue(), file_name=f'PCP_AcoNobre_{inp_pedido_nome}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)
                    
                    # Renderiza Mapas de Corte abaixo da lista
                    if HAS_NESTING and pecas_para_nesting_global:
                        st.markdown("---")
                        st.markdown("#### 👁️ Mapas de Corte")
                        for mat_nome, packer in dados_para_graficos.items():
                            with st.expander(f"Ver Mapas - {mat_nome}"):
                                for b_idx, bin in enumerate(packer):
                                    fig, ax = plt.subplots(figsize=(10, 5))
                                    ax.set_xlim(0, 3000); ax.set_ylim(0, 1250)
                                    ax.set_title(f"Plano de Corte - {mat_nome} | Chapa {b_idx + 1}")
                                    ax.add_patch(patches.Rectangle((0, 0), 3000, 1250, facecolor='#ecf0f1', edgecolor='black'))
                                    for rect in bin: ax.add_patch(patches.Rectangle((rect.x, rect.y), rect.width, rect.height, facecolor='#2980b9', edgecolor='white'))
                                    plt.gca().set_aspect('equal', adjustable='box'); plt.tight_layout()
                                    st.pyplot(fig); plt.close(fig) 

# ==========================================
# 5. TELA: GERADOR DE PCP (LOTE E AVULSO)
# ==========================================
elif menu == "⚙️ Gerador de PCP":
    st.title("⚙️ Gerador de PCP - Tratamento de Engenharia")
    st.markdown("Diferente do Projetista, este módulo processa os **PDFs reais** das suas pastas locais.")

    with st.expander("📂 Configuração de Pastas (Obrigatório)", expanded=True):
        col1, col2 = st.columns(2)
        pasta_origem = col1.text_input("Pasta Origem (Onde estão os PDFs)", value=r"C:\AcoNobre\Origem")
        pasta_destino = col2.text_input("Pasta Destino (Onde os resultados serão salvos)", value=r"C:\AcoNobre\Destino")

    st.markdown("### Ações Desejadas")
    col_chk1, col_chk2, col_chk3 = st.columns(3)
    chk_pdf = col_chk1.checkbox("Unir PDFs", value=True)
    chk_conf = col_chk2.checkbox("Gerar Conferência (Excel)", value=True)
    chk_pcp = col_chk3.checkbox("Atualizar Excel Global do Pedido", value=True)
    
    opcoes_selecionadas = {"pdf": chk_pdf, "conf": chk_conf, "pcp": chk_pcp}

    aba_lote, aba_avulso = st.tabs(["📑 Processamento em Lote (Planilha)", "📌 Atualização Avulsa (Item Único)"])

    with aba_lote:
        arquivo_planilha = st.file_uploader("Selecione a Planilha Comercial do Pedido (.xlsx)", type=["xlsx"])
        
        if st.button("🚀 PROCESSAR LOTE COMPLETO", type="primary", use_container_width=True):
            if not os.path.isdir(pasta_origem) or not os.path.isdir(pasta_destino):
                st.error("❌ Os caminhos das pastas Origem ou Destino são inválidos. Verifique se as pastas existem no seu computador.")
            elif arquivo_planilha is None:
                st.warning("⚠️ Faça o upload da planilha do comercial para iniciar o lote.")
            else:
                log_container = st.empty()
                log_container.info("⏳ Iniciando processamento em lote... Lendo pasta e arquivos.")
                try:
                    df_comercial = pd.read_excel(arquivo_planilha, dtype=str)
                    pedidos, comerciais = ler_planilha_entrada(df_comercial)
                    mapa = escanear_pasta(pasta_origem)
                    
                    for num_pedido, itens in pedidos.items():
                        pasta_pedido = os.path.join(pasta_destino, num_pedido)
                        caminho_global = os.path.join(pasta_pedido, f"LISTA-ITENS-{num_pedido}.xlsx")
                        
                        for n_item, cod_pai in itens:
                            dados = comerciais.get(f"{num_pedido}_{n_item}_{cod_pai}", {})
                            dados_pcp = processar_item_unico(n_item, cod_pai, num_pedido, dados, pasta_pedido, pasta_origem, mapa, opcoes_selecionadas, log_container)
                            
                            if chk_pcp and dados_pcp:
                                atualizar_excel_inteligente(dados_pcp, caminho_global, n_item)
                                log_container.success(f"  ✅ Tabela do Item {n_item} injetada na base global.")
                                
                    st.balloons()
                    st.success("🎉 Processamento de Lote concluído com sucesso!")
                except Exception as e:
                    st.error(f"❌ Ocorreu um erro crítico durante o processamento: {e}")

    with aba_avulso:
        c1, c2, c3, c4 = st.columns(4)
        ped_avulso = c1.text_input("Nº Pedido")
        item_avulso = c2.text_input("Item")
        cod_avulso = c3.text_input("Código Pai (Ex: 00-1234)")
        qtd_avulso = c4.text_input("Qtd")
        
        if st.button("🚀 PROCESSAR ITEM AVULSO", type="primary", use_container_width=True):
            if not os.path.isdir(pasta_origem) or not os.path.isdir(pasta_destino):
                st.error("❌ Os caminhos das pastas Origem ou Destino são inválidos.")
            elif not ped_avulso or not cod_avulso:
                st.warning("⚠️ Preencha pelo menos o Nº do Pedido e o Código Pai.")
            else:
                log_container = st.empty()
                log_container.info("⏳ Iniciando processamento avulso...")
                try:
                    mapa = escanear_pasta(pasta_origem)
                    dados_mock = {"cliente": "AVULSO", "qtd": qtd_avulso if qtd_avulso else "1", "desc": "Processamento Isolado"}
                    pasta_pedido = os.path.join(pasta_destino, ped_avulso)
                    caminho_global = os.path.join(pasta_pedido, f"LISTA-ITENS-{ped_avulso}.xlsx")
                    
                    dados_pcp = processar_item_unico(item_avulso, normalizar_codigo(cod_avulso), ped_avulso, dados_mock, pasta_pedido, pasta_origem, mapa, opcoes_selecionadas, log_container)
                    
                    if chk_pcp and dados_pcp:
                        atualizar_excel_inteligente(dados_pcp, caminho_global, item_avulso)
                        log_container.success(f"  ✅ Planilha Global do Pedido {ped_avulso} atualizada.")
                        
                    st.success("🎉 Processamento do Item Avulso concluído!")
                except Exception as e:
                    st.error(f"❌ Erro ao processar o item avulso: {e}")