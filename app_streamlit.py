import streamlit as st
import pandas as pd
import math
import re
import io

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

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
st.set_page_config(page_title="Projetista AçoNobre", layout="wide", page_icon="🏭", initial_sidebar_state="collapsed")

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
    
    div.row-widget.stRadio > div { flex-direction: row; justify-content: center; background-color: #1e1e1e; padding: 10px; border-radius: 10px; }
</style>
""", unsafe_allow_html=True)

# Memória do navegador
if 'carrinho' not in st.session_state: st.session_state.carrinho = []
if 'pecas_temp_composto' not in st.session_state: st.session_state.pecas_temp_composto = []

mapa_tampo = {"Mesa Lisa": "LISA", "Mesa com Encosto": "ENCOSTO", "Pia com Cuba": "PIA", "Prateleira Parede": "PRAT_PAREDE", "Estante Lisa": "ESTANTE_LISA", "Estante Gradeada": "ESTANTE_GRADEADA"}
mapa_base = {"Contraventamento": "CONTRAVENTAMENTO", "Prat. Lisa": "PRAT_LISA", "Prat. Lisa Dupla": "PRAT_LISA_DUPLA", "Prat. Gradeada": "PRAT_GRADEADA", "Prat. Gradeada Dupla": "PRAT_GRADEADA_DUPLA"}
dim_chapa_x = 3000
dim_chapa_y = 1250

@st.cache_data
def carregar_tabela_precos():
    custo_chapa = {"INOX 304": {0.6: 29.01, 0.8: 33.16, 1.0: 30.74, 1.2: 28.66, 1.5: 29.00, 2.0: 30.96}, 
                   "INOX 430": {0.6: 23.32, 0.8: 19.41, 1.0: 18.91, 1.2: 19.42, 1.5: 19.38, 2.0: 19.38},
                   "INOX 201": {0.6: 15.00, 0.8: 16.00, 1.0: 16.50, 1.2: 17.00}}
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
# 2. CABEÇALHO DO SISTEMA
# ==========================================
st.markdown("<h2 style='text-align: center; color: #f1c40f; margin-bottom: 15px;'>🛒 Projetista Virtual - AçoNobre</h2>", unsafe_allow_html=True)

# ==========================================
# 3. TELA: PROJETISTA VIRTUAL
# ==========================================
col_esq, col_dir = st.columns([1.3, 2.7])

with col_esq:
    st.markdown("#### 📦 DADOS DO PEDIDO")
    inp_pedido_nome = st.text_input("Nº do Pedido (Ex: 6885)", key="inp_pedido_global")
    
    aba_param, aba_comp, aba_livre = st.tabs(["📐 Catálogo Paramétrico", "🧩 Item Composto", "☕ Peça Avulsa"])
    
    with aba_param:
        c_qtd, c_planos = st.columns(2)
        qtd = c_qtd.number_input("QTD", min_value=1, value=1)
        planos = c_planos.number_input("Planos (P/ Estantes)", min_value=1, value=4)
        
        c1, c2, c3 = st.columns(3)
        comp = c1.number_input("C(mm)", min_value=100.0, value=1200.0, step=50.0)
        larg = c2.number_input("L(mm)", min_value=100.0, value=600.0, step=50.0)
        alt = c3.number_input("A(mm)", min_value=0.0, value=900.0, step=50.0)
        
        tampo_nome = st.selectbox("Produto Principal", list(mapa_tampo.keys()))
        cod_tampo = mapa_tampo[tampo_nome]
        
        # Inteligência da Base: Oculta a escolha se for Estante ou Prateleira Parede
        if "ESTANTE" in cod_tampo or cod_tampo == "PRAT_PAREDE":
            base_nome = "Contraventamento"
            cod_base = "CONTRAVENTAMENTO"
        else:
            base_nome = st.selectbox("Tipo de Base", list(mapa_base.keys()))
            cod_base = mapa_base[base_nome]
        
        st.markdown("---")
        # Nomenclaturas Visuais Inteligentes
        lbl_tampo = "Mat. dos Planos" if "ESTANTE" in cod_tampo else "Mat. Principal" if cod_tampo == "PRAT_PAREDE" else "Tampo Mat"
        lbl_base = "Mat. das Colunas" if "ESTANTE" in cod_tampo else "Base Mat"
        
        c_mat1, c_esp1 = st.columns([1, 1.5])
        mat_tampo = c_mat1.selectbox(lbl_tampo, ["INOX 304", "INOX 430", "INOX 201"])
        esp_tampo = c_esp1.selectbox("Espessura", ["Esp. Padrão", "0.8", "1.0", "1.2", "1.5", "2.0"], key="esp_t1")
        
        # Index sincronizado para a Base acompanhar o Tampo
        try: idx_base = ["INOX 304", "INOX 430", "INOX 201"].index(mat_tampo)
        except: idx_base = 1
        
        mat_base = "INOX 430"; esp_base = "Esp. Padrão"
        if cod_tampo != "PRAT_PAREDE" and cod_base != "CONTRAVENTAMENTO":
            c_mat2, c_esp2 = st.columns([1, 1.5])
            mat_base = c_mat2.selectbox(lbl_base, ["INOX 304", "INOX 430", "INOX 201"], index=idx_base)
            esp_base = c_esp2.selectbox("Espessura", ["Esp. Padrão", "0.8", "1.0", "1.2", "1.5", "2.0"], key="esp_b1")
            
        # --- ENGENHARIA DE REFORÇOS ---
        st.markdown("---")
        st.markdown("**Engenharia de Reforços**")
        
        # Reforço Principal (Tampo / Planos)
        lbl_ref1 = "Ref. Plano" if "ESTANTE" in cod_tampo else "Reforço" if cod_tampo == "PRAT_PAREDE" else "Ref. Tampo"
        cr1, cr2, cr3 = st.columns([1.5, 2, 1.5])
        qtd_ref = cr1.selectbox(f"Qtd {lbl_ref1}", ["Padrão", "0", "1", "2", "3", "4", "5", "6"])
        mat_ref = cr2.selectbox(f"Mat {lbl_ref1}", ["INOX 430", "INOX 304", "INOX 201"], index=["INOX 430", "INOX 304", "INOX 201"].index(mat_tampo) if mat_tampo in ["INOX 430", "INOX 304", "INOX 201"] else 0)
        esp_ref = cr3.selectbox(f"Esp {lbl_ref1}", ["0.6", "0.8", "1.0", "1.2", "1.5"], index=1)

        # Reforço Secundário (Prateleiras) - Só aparece se houver Prateleira na Base!
        mat_ref_prat = "INOX 430"; esp_ref_prat = "0.8"; qtd_ref_prat = "Padrão"
        if "PRAT" in cod_base and cod_tampo != "PRAT_PAREDE" and "ESTANTE" not in cod_tampo:
            lbl_ref_p = "Ref/Prat" if "DUPLA" in cod_base else "Ref. Prat"
            cp1, cp2, cp3 = st.columns([1.5, 2, 1.5])
            qtd_ref_prat = cp1.selectbox(f"Qtd {lbl_ref_p}", ["Padrão", "0", "1", "2", "3", "4", "5", "6"])
            mat_ref_prat = cp2.selectbox(f"Mat {lbl_ref_p}", ["INOX 430", "INOX 304", "INOX 201"], index=["INOX 430", "INOX 304", "INOX 201"].index(mat_base) if mat_base in ["INOX 430", "INOX 304", "INOX 201"] else 0)
            esp_ref_prat = cp3.selectbox(f"Esp {lbl_ref_p}", ["0.6", "0.8", "1.0", "1.2", "1.5"], index=1)
        
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("➕ Adicionar ao Pedido", use_container_width=True, type="primary"):
            item = {
                "num": len(st.session_state.carrinho) + 1, "tipo": "parametrico", "qtd": int(qtd),
                "comp": float(comp), "larg": float(larg), "alt": float(alt), "planos": int(planos),
                "tampo_nome": tampo_nome, "tampo_cod": cod_tampo,
                "base_nome": base_nome, "base_cod": cod_base,
                "mat_tampo": mat_tampo, "esp_tampo": esp_tampo, "mat_base": mat_base, "esp_base": esp_base,
                "mat_ref": mat_ref, "esp_ref": esp_ref, "qtd_ref": qtd_ref,
                "mat_ref_prat": mat_ref_prat, "esp_ref_prat": esp_ref_prat, "qtd_ref_prat": qtd_ref_prat
            }
            if "ESTANTE" in item["tampo_cod"]: item["desc_carrinho"] = f"{tampo_nome} {planos} Planos ({int(comp)}x{int(larg)}x{int(alt)}) - {mat_tampo}"
            elif item["tampo_cod"] == "PRAT_PAREDE": item["desc_carrinho"] = f"{tampo_nome} ({int(comp)}x{int(larg)}) - {mat_tampo}"
            else: 
                if cod_base == "CONTRAVENTAMENTO": item["desc_carrinho"] = f"{tampo_nome} c/ {base_nome} ({int(comp)}x{int(larg)}x{int(alt)}) - {mat_tampo}"
                else: item["desc_carrinho"] = f"{tampo_nome} c/ {base_nome} ({int(comp)}x{int(larg)}x{int(alt)}) - Tmp {mat_tampo} | Bs {mat_base}"
            
            st.session_state.carrinho.append(item)
            st.rerun()

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
        mat_pc = c_mat.selectbox("Mat", ["304", "430", "201"], key="mat_comp", label_visibility="collapsed")
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

    with aba_livre:
        c_qtd_livre, c_nome_livre = st.columns([1, 3])
        qtd_livre = c_qtd_livre.number_input("QTD", min_value=1, value=1, key="qtd_livre")
        nome_livre = c_nome_livre.text_input("Nome da Peça", placeholder="Ex: Chapa de Proteção")
        
        c_c_livre, c_l_livre = st.columns(2)
        comp_livre = c_c_livre.number_input("C Planif (mm)", min_value=1.0, value=500.0, step=10.0, key="c_livre")
        larg_livre = c_l_livre.number_input("L Planif (mm)", min_value=1.0, value=500.0, step=10.0, key="l_livre")
        
        c_mat_livre, c_esp_livre = st.columns(2)
        mat_livre = c_mat_livre.selectbox("Material", ["INOX 304", "INOX 430", "INOX 201"], key="mat_livre")
        esp_livre = c_esp_livre.selectbox("Espessura", ["0.6", "0.8", "1.0", "1.2", "1.5", "2.0"], index=2, key="esp_livre")
        
        if st.button("➕ Adicionar Peça Solta", type="primary", use_container_width=True):
            if nome_livre:
                item = {
                    "num": len(st.session_state.carrinho) + 1, "tipo": "livre", "qtd": int(qtd_livre),
                    "nome_peca": nome_livre, "comp_pl": float(comp_livre), "larg_pl": float(larg_livre),
                    "esp": float(esp_livre), "material": mat_livre, "desc_carrinho": f"✏️ {nome_livre} - {mat_livre}"
                }
                st.session_state.carrinho.append(item)
                st.rerun()

    st.markdown("<hr>", unsafe_allow_html=True)
    if st.session_state.carrinho:
        df_carrinho = pd.DataFrame([{"Nº": i['num'], "QTD": i.get('qtd', i.get('qtd_item', 1)), "DESCRIÇÃO": i['desc_carrinho']} for i in st.session_state.carrinho])
        st.dataframe(df_carrinho, hide_index=True, use_container_width=True)
        
    c_btn_limpar, c_btn_gerar = st.columns([1, 2])
    if c_btn_limpar.button("❌ Limpar", use_container_width=True):
        st.session_state.carrinho = []
        st.rerun()
        
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
            
            html_detalhamento = "<h4>📄 DETALHAMENTO DE FABRICAÇÃO POR ITEM</h4>"
            
            for item in st.session_state.carrinho:
                pecas_base = []
                
                # --- GERAÇÃO DA GEOMETRIA BRUTA ---
                if item["tipo"] == "parametrico":
                    if "ESTANTE" in item["tampo_cod"]:
                        tipo_est = "LISA" if "LISA" in item["tampo_cod"] else "GRADEADA"
                        pb_cruas = calcular_estante_parametrica(item["comp"], item["larg"], item["alt"], planos=item["planos"], tipo=tipo_est, material=item["mat_tampo"])
                    elif item["tampo_cod"] == "PRAT_PAREDE":
                        pb_cruas = calcular_prateleira_parede(item["comp"], item["larg"], item["mat_tampo"])
                    else:
                        pb_cruas = calcular_mesa_parametrica(item["comp"], item["larg"], item["alt"], item["tampo_cod"], item["base_cod"])
                
                    molde_reforco = None

                    # --- LAÇO INTELIGENTE DE REFORÇOS E MATERIAIS ---
                    for p in pb_cruas:
                        if p["CÓDIGO"] in ["SAPATA", "PARAFUSO", "CUBA_PADRAO"]:
                            p["MAT_CUSTOM"] = "-"
                        else:
                            is_reforco = any(x in p["DESC"].upper() for x in ["REFORÇO", "REFORCO", "OMEGA"])
                            if is_reforco and "MAIOR" not in p["DESC"].upper(): molde_reforco = p.copy()
                            if item["tampo_cod"] == "LISA" and is_reforco and "MAIOR" in p["DESC"].upper(): continue
                                
                            if is_reforco:
                                is_traseiro = False; is_prat = False; is_plano = False; is_tampo = False
                                
                                # Nomeia os reforços
                                if "MAIOR" in p["DESC"].upper(): p["DESC"] = "REFORÇO TRASEIRO DO TAMPO"; is_traseiro = True
                                elif "PRAT" in p["DESC"].upper() or "BASE" in p["DESC"].upper(): p["DESC"] = "REFORÇO PRATELEIRA"; is_prat = True
                                elif "PLANO" in p["DESC"].upper() or "ESTANTE" in item["tampo_cod"]: p["DESC"] = "REFORÇO PLANO"; is_plano = True
                                else: p["DESC"] = "REFORÇO TAMPO"; is_tampo = True
                                    
                                # Aplica regras de material e espessura
                                if is_traseiro:
                                    p["MAT_CUSTOM"] = item["mat_tampo"] # Intocável da Caldeiraria
                                    p["ESP"] = 0.8
                                elif is_prat:
                                    p["MAT_CUSTOM"] = item.get("mat_ref_prat", "INOX 430")
                                    try: p["ESP"] = float(item.get("esp_ref_prat", "0.8"))
                                    except: p["ESP"] = 0.8
                                    qtd_r = item.get("qtd_ref_prat", "Padrão")
                                    if qtd_r != "Padrão":
                                        if qtd_r == "0": continue
                                        mult = 2 if "DUPLA" in item.get("base_cod", "") else 1
                                        p["QTD"] = int(qtd_r) * mult
                                elif is_plano:
                                    p["MAT_CUSTOM"] = item.get("mat_ref", "INOX 430")
                                    try: p["ESP"] = float(item.get("esp_ref", "0.8"))
                                    except: p["ESP"] = 0.8
                                    qtd_r = item.get("qtd_ref", "Padrão")
                                    if qtd_r != "Padrão":
                                        if qtd_r == "0": continue
                                        p["QTD"] = int(qtd_r) * item.get("planos", 4)
                                elif is_tampo:
                                    p["MAT_CUSTOM"] = item.get("mat_ref", "INOX 430")
                                    try: p["ESP"] = float(item.get("esp_ref", "0.8"))
                                    except: p["ESP"] = 0.8
                                    qtd_r = item.get("qtd_ref", "Padrão")
                                    if qtd_r != "Padrão":
                                        if qtd_r == "0": continue
                                        p["QTD"] = int(qtd_r)
                            else:
                                e_base = False
                                if "TUBO" in p["CÓDIGO"] or any(x in p["DESC"].upper() for x in ["PRAT", "GRADE", "PERNA", "CONTRA", "TRAVESSA", "COLUNA"]): e_base = True
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
                    
                    # Cria o reforço da prateleira inferior caso a engenharia crua não tenha gerado
                    if "PRAT" in item.get("base_cod", "") and not any(p["DESC"] == "REFORÇO PRATELEIRA" for p in pecas_base):
                        if molde_reforco:
                            ref_prat = molde_reforco.copy()
                            ref_prat["DESC"] = "REFORÇO PRATELEIRA"
                            ref_prat["MAT_CUSTOM"] = item.get("mat_ref_prat", "INOX 430")
                            try: ref_prat["ESP"] = float(item.get("esp_ref_prat", "0.8"))
                            except: ref_prat["ESP"] = 0.8
                            qtd_r = item.get("qtd_ref_prat", "Padrão")
                            mult = 2 if "DUPLA" in item.get("base_cod", "") else 1
                            if qtd_r != "Padrão":
                                if qtd_r == "0": ref_prat = None
                                else: ref_prat["QTD"] = int(qtd_r) * mult
                            else: ref_prat["QTD"] = ref_prat.get("QTD", 1) * mult
                                
                            if ref_prat:
                                if ref_prat["COMP PL"] != "-" and ref_prat["LARG PL"] != "-":
                                    try: ref_prat["PESO UNIT"] = float(ref_prat["COMP PL"]) * float(ref_prat["LARG PL"]) * float(ref_prat["ESP"]) * 0.000008
                                    except: pass
                                pecas_base.append(ref_prat)

                elif item["tipo"] == "composto":
                    for pt in item["pecas"]:
                        peso_un = pt["comp"] * pt["larg"] * pt["esp"] * 0.000008
                        pecas_base.append({"CÓDIGO": "CHAPA_LIVRE", "DESC": pt["nome"], "QTD": pt["qtd"], "COMP PL": pt["comp"], "LARG PL": pt["larg"], "ESP": pt["esp"], "PESO UNIT": peso_un, "MAT_CUSTOM": pt["mat"]})
                
                elif item["tipo"] == "livre":
                    peso_un = item["comp_pl"] * item["larg_pl"] * item["esp"] * 0.000008
                    pecas_base.append({"CÓDIGO": "CHAPA_LIVRE", "DESC": item["nome_peca"], "QTD": 1, "COMP PL": item["comp_pl"], "LARG PL": item["larg_pl"], "ESP": item["esp"], "PESO UNIT": peso_un, "MAT_CUSTOM": item["material"]})

                # --- RENDERIZAÇÃO DO CARD HTML DO ITEM ---
                qtd_multiplicador = item.get("qtd", item.get("qtd_item", 1))
                peso_item_total = 0.0
                custo_unit_item_chapa = 0.0
                custo_unit_item_tubo = 0.0

                html_detalhamento += f"<div class='card-dark'><div class='item-title'>ITEM {item['num']}: {qtd_multiplicador}x {item['desc_carrinho']}</div>"
                
                for p in pecas_base:
                    qtd_final = p["QTD"] * qtd_multiplicador
                    custo_peca_un = 0.0
                    mat_peca = p.get("MAT_CUSTOM", "-")
                    
                    if "CHAPA" in p["CÓDIGO"] or (p.get("COMP PL", "-") != "-" and p.get("LARG PL", "-") != "-"):
                        peso_peca = p.get("PESO UNIT", 0) * qtd_final
                        if p.get("LARG PL", "-") != "-":
                            for _ in range(qtd_final): pecas_para_nesting_global.append({'nome': p['DESC'], 'material': f"{mat_peca} {p.get('ESP','-')}mm", 'comp': float(p["COMP PL"]), 'larg': float(p["LARG PL"])})
                        if p.get("ESP", 0) > 0:
                            custo_peca_un = p["PESO UNIT"] * custo_chapa.get(mat_peca, {}).get(p["ESP"], 0.0)
                            custo_unit_item_chapa += custo_peca_un * p["QTD"]
                            chave = (mat_peca, p["ESP"])
                            resumo_geral_chapas[chave] = resumo_geral_chapas.get(chave, 0) + peso_peca
                        linha_med = f"{p.get('PESO UNIT',0):.2f} KG un, | {peso_peca:.2f} KG tot,".replace('.',',')

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
                    
                    html_detalhamento += f"<div class='item-linha'>&#x25AB; {qtd_final}x {mat_str}{p['DESC']} | {linha_med} | {custo_str}</div>"
                    
                    dados_para_df.append({
                        "ITEM": f"Item {item['num']}", "QTD": qtd_final, "CÓDIGO": p["CÓDIGO"],
                        "DESCRIÇÃO": p["DESC"], "MATERIAL": p.get("MAT_CUSTOM", "-"), "ESP": p.get("ESP", "-"),
                        "MEDIDA (C x L)": f"{p.get('COMP PL','-')} x {p.get('LARG PL','-')}", "PESO TOT (KG)": round(peso_peca, 2)
                    })
                    
                c_unit = custo_unit_item_chapa + custo_unit_item_tubo
                c_tot = c_unit * qtd_multiplicador
                peso_total_pedido += peso_item_total
                html_detalhamento += f"<hr><div class='item-linha' style='color:#fff;'><b>CUSTO TOTAL DO ITEM: R$ {c_tot:.2f} | Peso: {peso_item_total:.2f} KG</b></div></div>".replace('.',',')

            # --- TOTAIS E DASHBOARD ---
            custo_tot_chapas = sum([peso * custo_chapa.get(mat, {}).get(esp, 0.0) for (mat, esp), peso in resumo_geral_chapas.items()])
            custo_tot_tubos = sum([(c_total / 1000.0) * custo_tubo.get(cod, 0.0) for cod, c_total in resumo_geral_tubos.items()])
            custo_total_geral = custo_tot_chapas + custo_tot_tubos
            
            tab_lista, tab_dash = st.tabs(["✂️ Lista PCP e Mapas", "📊 Dashboard"])
            
            with tab_dash:
                titulo_pedido = f"CUSTO DO PEDIDO {inp_pedido_nome}" if inp_pedido_nome else "CUSTO DO PEDIDO AVULSO"
                st.markdown(f"""
                <div class="box-azul">
                    <h4>⏱️ {titulo_pedido}</h4>
                    <h1>R$ {custo_total_geral:,.2f}</h1>
                </div>
                """.replace(',', 'X').replace('.', ',').replace('X', '.'), unsafe_allow_html=True)
                
                if resumo_geral_chapas:
                    html_chapas = "<div class='card-dark'><h4>📦 CONSUMO DE CHAPAS INOX</h4>"
                    for (mat, esp), peso in sorted(resumo_geral_chapas.items()):
                        prc = custo_chapa.get(mat, {}).get(esp, 0.0)
                        ct = peso * prc
                        html_chapas += f"<p>• {mat} {esp}mm ➔ {peso:.2f} KG | (R$ {prc:.2f}/kg) | Subtotal: R$ {ct:.2f}</p>".replace('.',',')
                        linhas_resumo_excel.append({"TIPO": "CHAPA", "MATERIAL": f"{mat} {esp}mm", "QTD": f"{peso:.2f} KG".replace('.', ','), "CUSTO (R$)": round(ct, 2)})
                    html_chapas += f"<p style='font-weight:bold; margin-top:10px;'>Peso Total em Chapas: {peso_total_pedido:.2f} KG</p></div>".replace('.',',')
                    st.markdown(html_chapas, unsafe_allow_html=True)
                        
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
                    
                if HAS_NESTING and pecas_para_nesting_global:
                    html_nest = f"<div class='card-dark'><h4>🧩 ESTIMATIVA DE ENCAIXE (CHAPAS {dim_chapa_x}x{dim_chapa_y})</h4>"
                    agrup_nesting = {}
                    for pc in pecas_para_nesting_global:
                        chave = pc['material']
                        if chave not in agrup_nesting: agrup_nesting[chave] = []
                        agrup_nesting[chave].append((pc['comp'], pc['larg']))

                    dados_para_graficos = {}
                    for mat_nome, dimensoes in agrup_nesting.items():
                        packer = newPacker(mode=PackingMode.Offline, bin_algo=PackingBin.Global, rotation=True)
                        for idx, (c, l) in enumerate(dimensoes): packer.add_rect(width=c + 5, height=l + 5, rid=idx) 
                        for _ in range(50): packer.add_bin(width=dim_chapa_x, height=dim_chapa_y)
                        packer.pack()
                        
                        chp_usadas = len(packer)
                        area_pc = sum([c * l for c, l in dimensoes])
                        area_ch = chp_usadas * (dim_chapa_x * dim_chapa_y)
                        aprov = (area_pc / area_ch) * 100 if area_ch > 0 else 0
                        
                        html_nest += f"<p>• {mat_nome}: Será preciso puxar <b>{chp_usadas} chapa(s)</b> do estoque. (Aproveitamento: {aprov:.1f}%)</p>".replace('.',',')
                        dados_para_graficos[mat_nome] = packer
                        
                    html_nest += "</div>"
                    st.markdown(html_nest, unsafe_allow_html=True)
                    
                st.markdown(html_detalhamento, unsafe_allow_html=True)

            with tab_lista:
                st.markdown("#### ✂️ Lista de Corte Detalhada")
                df_lista = pd.DataFrame(dados_para_df)
                st.dataframe(df_lista, hide_index=True, use_container_width=True)
                
                df_resumo = pd.DataFrame(linhas_resumo_excel)
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_lista.to_excel(writer, sheet_name='Lista_Corte_PCP', index=False)
                    if not df_resumo.empty: df_resumo.to_excel(writer, sheet_name='Resumo_Financeiro', index=False)
                
                st.download_button(label="📥 Baixar Excel Completo (PCP e Custos)", data=buffer.getvalue(), file_name=f'PCP_AcoNobre_{inp_pedido_nome}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)
                
                if HAS_NESTING and pecas_para_nesting_global:
                    st.markdown("---")
                    st.markdown("#### 👁️ Mapas de Corte")
                    for mat_nome, packer in dados_para_graficos.items():
                        with st.expander(f"Ver Mapas - {mat_nome}"):
                            for b_idx, bin in enumerate(packer):
                                fig, ax = plt.subplots(figsize=(10, 5))
                                ax.set_xlim(0, dim_chapa_x); ax.set_ylim(0, dim_chapa_y)
                                ax.set_title(f"Plano de Corte - {mat_nome} | Chapa {b_idx + 1}")
                                ax.add_patch(patches.Rectangle((0, 0), dim_chapa_x, dim_chapa_y, facecolor='#ecf0f1', edgecolor='black'))
                                for rect in bin: ax.add_patch(patches.Rectangle((rect.x, rect.y), rect.width, rect.height, facecolor='#2980b9', edgecolor='white'))
                                plt.gca().set_aspect('equal', adjustable='box'); plt.tight_layout()
                                st.pyplot(fig); plt.close(fig)
