import streamlit as st
import pandas as pd
import math
import re
import io

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from regras_engenharia import calcular_mesa_parametrica, calcular_estante_parametrica, calcular_prateleira_parede

try:
    from rectpack import newPacker, PackingMode, PackingBin
    import matplotlib.pyplot as plt
    import matplotlib.patches as patches
    HAS_NESTING = True
except ImportError:
    HAS_NESTING = False

# ==========================================
# CONFIGURAÇÃO DA PÁGINA E CSS
# ==========================================
st.set_page_config(page_title="Projetista AçoNobre", layout="wide", page_icon="🏭", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    .box-azul {
        background-color: #2980b9;
        color: white;
        padding: 20px;
        border-radius: 8px;
        text-align: center;
        margin-bottom: 15px;
    }
    .box-azul h4 { color: #ecf0f1; margin: 0; font-size: 15px; font-weight: normal; letter-spacing: 1px; }
    .box-azul h1 { color: white; margin: 5px 0 0 0; font-size: 40px; font-weight: bold; }
    
    .card-dark {
        background-color: #2b2b2b;
        padding: 15px;
        border-radius: 8px;
        margin-bottom: 10px;
        border: 1px solid #3d3d3d;
    }
    .card-dark h4 { color: #ffffff; font-size: 14px; font-weight: bold; margin-bottom: 10px;}
    .card-dark p { color: #dddddd; font-size: 13px; margin-bottom: 5px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;}
    
    .item-title { color: #f1c40f; font-weight: bold; font-size: 15px; margin-bottom: 10px; }
    .item-linha { font-family: 'Courier New', Courier, monospace; font-size: 13px; color: #cccccc; margin-bottom: 6px; margin-left: 10px; }
    
    div.row-widget.stRadio > div { flex-direction: row; justify-content: center; background-color: #1e1e1e; padding: 10px; border-radius: 10px; }
    
    .titulo-bloco { margin-top: 15px; margin-bottom: 5px; color: #f1c40f; font-weight: bold; font-size: 13px; text-transform: uppercase; letter-spacing: 1px; border-bottom: 1px solid #333; padding-bottom: 5px; }
</style>
""", unsafe_allow_html=True)

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
# INTERFACE PRINCIPAL
# ==========================================
st.markdown("<h2 style='text-align: center; color: #f1c40f; margin-bottom: 5px;'>🛒 Projetista Virtual - AçoNobre</h2>", unsafe_allow_html=True)

col_esq, col_dir = st.columns([1.5, 2.5], gap="large")

with col_esq:
    st.markdown("<h4 style='margin-bottom:0px;'>📦 DADOS DO PEDIDO</h4>", unsafe_allow_html=True)
    inp_pedido_nome = st.text_input("Nº do Pedido", placeholder="Ex: 6885", label_visibility="collapsed", key="pedido_global")
    
    aba_param, aba_comp, aba_livre = st.tabs(["📐 Catálogo Paramétrico", "🧩 Módulo Composto", "☕ Peça Avulsa"])
    
    # ------------------ CATÁLOGO PARAMÉTRICO ------------------
    with aba_param:
        cp_prod, cp_base = st.columns(2, gap="small")
        tampo_nome = cp_prod.selectbox("Produto Principal", list(mapa_tampo.keys()), key="sel_prod_prin")
        cod_tampo = mapa_tampo[tampo_nome]
        
        if "ESTANTE" in cod_tampo or cod_tampo == "PRAT_PAREDE":
            base_nome = "Contraventamento"; cod_base = "CONTRAVENTAMENTO"
        else:
            base_nome = cp_base.selectbox("Tipo de Base", list(mapa_base.keys()), key="sel_tipo_base")
            cod_base = mapa_base[base_nome]
            
        st.markdown("<div class='titulo-bloco'>GEOMETRIA</div>", unsafe_allow_html=True)
        c1, c2, c3, c4, c5 = st.columns(5, gap="small")
        qtd = c1.number_input("QTD", min_value=1, value=1, key="num_qtd_param")
        comp = c2.number_input("C(mm)", min_value=100.0, value=1200.0, step=50.0, key="num_comp_param")
        larg = c3.number_input("L(mm)", min_value=100.0, value=600.0, step=50.0, key="num_larg_param")
        
        alt = 0.0; planos = 0
        if cod_tampo == "PRAT_PAREDE":
            pass 
        elif "ESTANTE" in cod_tampo:
            alt = c4.number_input("A(mm)", min_value=0.0, value=1800.0, step=50.0, key="num_alt_param_est")
            planos = c5.number_input("Planos", min_value=1, value=4, key="num_plan_param")
        else:
            alt = c4.number_input("A(mm)", min_value=0.0, value=900.0, step=50.0, key="num_alt_param_mesa")
            
        # ================= NOVO DESIGN EM LINHA =================
        st.markdown("<div class='titulo-bloco'>ESTRUTURA E REFORÇOS</div>", unsafe_allow_html=True)
        
        # --- LINHA 1: TAMPO / PLANO ---
        lbl_tampo = "Plano" if "ESTANTE" in cod_tampo else "Principal" if cod_tampo == "PRAT_PAREDE" else "Tampo"
        has_ref_tampo = cod_tampo != "ESTANTE_GRADEADA"
        
        if has_ref_tampo:
            c_mt, c_et, c_qr, c_mr, c_er = st.columns([1.5, 1, 1, 1.2, 1], gap="small")
        else:
            c_mt, c_et = st.columns([1.5, 1], gap="small")
            
        mat_tampo = c_mt.selectbox(lbl_tampo, ["INOX 304", "INOX 430", "INOX 201"], key="sel_mat_tampo")
        esp_tampo = c_et.selectbox("Esp. (T)" if cod_tampo != "PRAT_PAREDE" else "Esp.", ["Esp. Padrão", "0.8", "1.0", "1.2", "1.5", "2.0"], key="sel_esp_tampo")
        
        mat_ref = "INOX 430"; esp_ref = "0.8"; qtd_ref = "Padrão"
        if has_ref_tampo:
            qtd_ref = c_qr.selectbox("Qtd Ref.", ["Padrão", "0", "1", "2", "3", "4", "5", "6"], key="sel_qtd_ref1")
            idx_ref1 = ["INOX 430", "INOX 304", "INOX 201"].index(mat_tampo) if mat_tampo in ["INOX 430", "INOX 304", "INOX 201"] else 0
            mat_ref = c_mr.selectbox("Mat. Ref.", ["430", "304", "201"], index=idx_ref1, key="sel_mat_ref1")
            esp_ref = c_er.selectbox("Esp. Ref.", ["0.6", "0.8", "1.0", "1.2", "1.5"], index=1, key="sel_esp_ref1")
            mat_ref = f"INOX {mat_ref}"
        else:
            qtd_ref = "0" # Garante que não gere reforço pra gradeada

        # --- LINHA 2: BASE / PRATELEIRA ---
        mat_base = "INOX 430"; esp_base = "Esp. Padrão"
        mat_ref_prat = "INOX 430"; esp_ref_prat = "0.8"; qtd_ref_prat = "Padrão"
        
        if cod_tampo != "PRAT_PAREDE" and cod_base != "CONTRAVENTAMENTO" and "ESTANTE" not in cod_tampo:
            lbl_base = "Prateleira" if "PRAT" in cod_base else "Base"
            has_ref_prat = "PRAT" in cod_base
            
            try: idx_base = ["INOX 304", "INOX 430", "INOX 201"].index(mat_tampo)
            except: idx_base = 1
            
            if has_ref_prat:
                c_mb, c_eb, c_qrp, c_mrp, c_erp = st.columns([1.5, 1, 1, 1.2, 1], gap="small")
            else:
                c_mb, c_eb = st.columns([1.5, 1], gap="small")
                
            mat_base = c_mb.selectbox(lbl_base, ["INOX 304", "INOX 430", "INOX 201"], index=idx_base, key="sel_mat_base")
            esp_base = c_eb.selectbox("Esp. (B)", ["Esp. Padrão", "0.8", "1.0", "1.2", "1.5", "2.0"], key="sel_esp_base")
            
            if has_ref_prat:
                qtd_ref_prat = c_qrp.selectbox("Qtd Ref. Prt", ["Padrão", "0", "1", "2", "3", "4", "5", "6"], key="sel_qtd_ref2")
                idx_ref2 = ["INOX 430", "INOX 304", "INOX 201"].index(mat_base) if mat_base in ["INOX 430", "INOX 304", "INOX 201"] else 0
                mat_ref_prat = c_mrp.selectbox("Mat. Ref. Prt", ["430", "304", "201"], index=idx_ref2, key="sel_mat_ref2")
                esp_ref_prat = c_erp.selectbox("Esp. Ref. Prt", ["0.6", "0.8", "1.0", "1.2", "1.5"], index=1, key="sel_esp_ref2")
                mat_ref_prat = f"INOX {mat_ref_prat}"

        st.markdown("<div style='margin-top:10px;'></div>", unsafe_allow_html=True)
        if st.button("➕ Adicionar ao Pedido", use_container_width=True, type="primary", key="btn_add_param"):
            item = {
                "num": len(st.session_state.carrinho) + 1, "tipo": "parametrico", "qtd": int(qtd),
                "comp": float(comp), "larg": float(larg), "alt": float(alt), "planos": int(planos),
                "tampo_nome": tampo_nome, "tampo_cod": cod_tampo, "base_nome": base_nome, "base_cod": cod_base,
                "mat_tampo": mat_tampo, "esp_tampo": esp_tampo, "mat_base": mat_base, "esp_base": esp_base,
                "mat_ref": mat_ref, "esp_ref": esp_ref, "qtd_ref": qtd_ref,
                "mat_ref_prat": mat_ref_prat, "esp_ref_prat": esp_ref_prat, "qtd_ref_prat": qtd_ref_prat
            }
            if "ESTANTE" in cod_tampo: item["desc_carrinho"] = f"{tampo_nome} {planos} Pl ({int(comp)}x{int(larg)}x{int(alt)}) - {mat_tampo}"
            elif cod_tampo == "PRAT_PAREDE": item["desc_carrinho"] = f"{tampo_nome} ({int(comp)}x{int(larg)}) - {mat_tampo}"
            else: 
                if cod_base == "CONTRAVENTAMENTO": item["desc_carrinho"] = f"{tampo_nome} c/ Contra. ({int(comp)}x{int(larg)}x{int(alt)}) - {mat_tampo}"
                else: item["desc_carrinho"] = f"{tampo_nome} c/ {base_nome} ({int(comp)}x{int(larg)}x{int(alt)})"
            
            st.session_state.carrinho.append(item)
            st.rerun()

    # ------------------ ITEM COMPOSTO ------------------
    with aba_comp:
        cm1, cm2 = st.columns([1, 3], gap="small")
        qtd_item_comp = cm1.number_input("QTD Mód", min_value=1, value=1, key="num_qtd_comp")
        nome_item_comp = cm2.text_input("Nome Módulo", placeholder="Ex: Gaveteiro", key="txt_nome_comp")
        
        st.markdown("<div class='titulo-bloco'>PEÇAS DO MÓDULO</div>", unsafe_allow_html=True)
        cc1, cc2 = st.columns([3, 1], gap="small")
        nome_pc = cc1.text_input("Nome da Peça", placeholder="Chapa lateral", key="txt_nome_pc")
        qtd_pc = cc2.number_input("Qtd", min_value=1, value=1, key="num_qtd_pc")
        
        cd1, cd2, cd3, cd4 = st.columns(4, gap="small")
        comp_pc = cd1.number_input("C(mm) ", value=100.0, step=10.0, key="num_c_pc")
        larg_pc = cd2.number_input("L(mm) ", value=100.0, step=10.0, key="num_l_pc")
        mat_pc = cd3.selectbox("Mat", ["304", "430", "201"], key="sel_mat_pc")
        esp_pc = cd4.selectbox("Esp", ["0.6", "0.8", "1.0", "1.2", "1.5", "2.0"], index=2, key="sel_esp_pc")
        
        if st.button("⏬ Inserir Peça", use_container_width=True, key="btn_inserir_pc"):
            if nome_pc:
                st.session_state.pecas_temp_composto.append({
                    "nome": nome_pc, "qtd": int(qtd_pc), "comp": float(comp_pc), "larg": float(larg_pc),
                    "mat": f"INOX {mat_pc}", "esp": float(esp_pc)
                })
                st.rerun()
            
        if st.session_state.pecas_temp_composto:
            df_temp = pd.DataFrame(st.session_state.pecas_temp_composto)
            st.dataframe(df_temp[["qtd", "nome", "mat"]].rename(columns={"qtd":"QTD", "nome":"PEÇA", "mat":"MAT"}), use_container_width=True, hide_index=True)
            
        st.markdown("<div style='margin-top:5px;'></div>", unsafe_allow_html=True)
        if st.button("✅ EMPACOTAR ITEM E ADD", type="primary", use_container_width=True, key="btn_empacotar"):
            if st.session_state.pecas_temp_composto and nome_item_comp:
                item = {
                    "num": len(st.session_state.carrinho) + 1, "tipo": "composto",
                    "nome_item": nome_item_comp, "qtd_item": int(qtd_item_comp), "qtd": int(qtd_item_comp),
                    "pecas": list(st.session_state.pecas_temp_composto), "desc_carrinho": f"🧩 {nome_item_comp}"
                }
                st.session_state.carrinho.append(item)
                st.session_state.pecas_temp_composto = [] 
                st.rerun()

    # ------------------ PEÇA AVULSA ------------------
    with aba_livre:
        ca1, ca2 = st.columns([1, 3], gap="small")
        qtd_livre = ca1.number_input("QTD", min_value=1, value=1, key="num_qtd_livre")
        nome_livre = ca2.text_input("Descrição", placeholder="Ex: Chapa de Proteção", key="txt_nome_livre")
        
        cl1, cl2, cl3, cl4 = st.columns(4, gap="small")
        comp_livre = cl1.number_input("C Plan", value=500.0, step=10.0, key="num_c_livre")
        larg_livre = cl2.number_input("L Plan", value=500.0, step=10.0, key="num_l_livre")
        mat_livre = cl3.selectbox("Mat.", ["304", "430", "201"], key="sel_mat_livre")
        esp_livre = cl4.selectbox("Esp.", ["0.6", "0.8", "1.0", "1.2", "1.5", "2.0"], index=2, key="sel_esp_livre")
        
        st.markdown("<div style='margin-top:10px;'></div>", unsafe_allow_html=True)
        if st.button("➕ Adicionar Peça Solta", type="primary", use_container_width=True, key="btn_add_livre"):
            if nome_livre:
                item = {
                    "num": len(st.session_state.carrinho) + 1, "tipo": "livre", "qtd": int(qtd_livre),
                    "nome_peca": nome_livre, "comp_pl": float(comp_livre), "larg_pl": float(larg_livre),
                    "esp": float(esp_livre), "material": f"INOX {mat_livre}", "desc_carrinho": f"✏️ {nome_livre}"
                }
                st.session_state.carrinho.append(item)
                st.rerun()

    # ------------------ CARRINHO ------------------
    st.markdown("<div class='titulo-bloco' style='margin-top:20px;'>🛒 CARRINHO</div>", unsafe_allow_html=True)
    if st.session_state.carrinho:
        df_carrinho = pd.DataFrame([{"Nº": i['num'], "QTD": i.get('qtd', i.get('qtd_item', 1)), "DESCRIÇÃO": i['desc_carrinho']} for i in st.session_state.carrinho])
        st.dataframe(df_carrinho, hide_index=True, use_container_width=True)
        
        c_btn_limpar, c_btn_gerar = st.columns([1, 2], gap="small")
        if c_btn_limpar.button("❌ Limpar", use_container_width=True, key="btn_limpar_car"):
            st.session_state.carrinho = []
            st.rerun()
            
        gerar_calc = c_btn_gerar.button("🚀 CALCULAR PROJETO", type="primary", use_container_width=True, key="btn_gerar_calc")
    else:
        st.info("O carrinho está vazio.")
        gerar_calc = False


# ==========================================
# 4. COLUNA DIREITA (DASHBOARD)
# ==========================================
with col_dir:
    if not st.session_state.carrinho:
        st.markdown("<div style='margin-top:50px; text-align:center; color:#888;'>Adicione itens ao pedido e clique em <b>CALCULAR PROJETO</b> para visualizar os custos e a engenharia.</div>", unsafe_allow_html=True)
    elif gerar_calc:
        with st.spinner('Engenharia processando...'):
            dados_para_df = []
            pecas_para_nesting_global = [] 
            resumo_geral_chapas = {}
            resumo_geral_tubos = {}
            linhas_resumo_excel = []
            peso_total_pedido = 0.0
            custo_total_geral = 0.0
            
            html_detalhamento = "<h4>📄 DETALHAMENTO DE FABRICAÇÃO</h4>"
            
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

                peso_item_total = 0.0; custo_unit_item_chapa = 0.0; custo_unit_item_tubo = 0.0
                
                html_detalhamento += f"<div class='card-dark'><div class='item-title'>ITEM {item['num']}: {item.get('qtd', item.get('qtd_item',1))}x {item['desc_carrinho']}</div>"
                
                pecas_processadas = []
                molde_reforco = None

                for p in pb_cruas:
                    if p["CÓDIGO"] in ["SAPATA", "PARAFUSO", "CUBA_PADRAO"]: p["MAT_CUSTOM"] = "-"
                    else:
                        is_reforco = any(x in p["DESC"].upper() for x in ["REFORÇO", "REFORCO", "OMEGA"])
                        if is_reforco and "MAIOR" not in p["DESC"].upper(): molde_reforco = p.copy()
                        
                        # Bloqueio de Reforço para Estante Gradeada e Mesa Lisa
                        if item.get("tampo_cod") == "LISA" and is_reforco and "MAIOR" in p["DESC"].upper(): continue
                        if item.get("tampo_cod") == "ESTANTE_GRADEADA" and is_reforco: continue
                        
                        if is_reforco and item["tipo"] == "parametrico":
                            if "MAIOR" in p["DESC"].upper():
                                p["DESC"] = "REFORÇO TRASEIRO DO TAMPO"
                                p["MAT_CUSTOM"] = item["mat_tampo"]; p["ESP"] = 0.8
                            elif "PRAT" in p["DESC"].upper() or "BASE" in p["DESC"].upper():
                                p["DESC"] = "REFORÇO PRATELEIRA"
                                p["MAT_CUSTOM"] = item.get("mat_ref_prat", "INOX 430"); p["ESP"] = float(item.get("esp_ref_prat", 0.8))
                                qtd_r = item.get("qtd_ref_prat", "Padrão")
                                if qtd_r != "Padrão":
                                    if qtd_r == "0": continue
                                    p["QTD"] = int(qtd_r) * (2 if "DUPLA" in item.get("base_cod", "") else 1)
                            elif "PLANO" in p["DESC"].upper() or "ESTANTE" in item.get("tampo_cod",""):
                                p["DESC"] = "REFORÇO PLANO"
                                p["MAT_CUSTOM"] = item.get("mat_ref", "INOX 430"); p["ESP"] = float(item.get("esp_ref", 0.8))
                                qtd_r = item.get("qtd_ref", "Padrão")
                                if qtd_r != "Padrão":
                                    if qtd_r == "0": continue
                                    p["QTD"] = int(qtd_r) * item.get("planos", 4)
                            else:
                                p["DESC"] = "REFORÇO TAMPO"
                                p["MAT_CUSTOM"] = item.get("mat_ref", "INOX 430"); p["ESP"] = float(item.get("esp_ref", 0.8))
                                qtd_r = item.get("qtd_ref", "Padrão")
                                if qtd_r != "Padrão":
                                    if qtd_r == "0": continue
                                    p["QTD"] = int(qtd_r)
                        elif item["tipo"] == "parametrico":
                            e_base = any(x in p["DESC"].upper() for x in ["PRAT", "GRADE", "PERNA", "CONTRA", "TRAVESSA", "COLUNA", "DIVISÓRIA", "SUPORTE"])
                            p["MAT_CUSTOM"] = item["mat_base"] if e_base else item["mat_tampo"]
                            if item.get("tampo_cod") not in ["ESTANTE_LISA", "ESTANTE_GRADEADA", "PRAT_PAREDE"]:
                                esp_tela = item["esp_base"] if e_base else item["esp_tampo"]
                                if ("CHAPA" in p["CÓDIGO"] or (p.get("COMP PL", "-") != "-" and p.get("LARG PL", "-") != "-")) and "TUBO" not in p["CÓDIGO"]:
                                    if esp_tela != "Esp. Padrão":
                                        try: p["ESP"] = float(esp_tela)
                                        except: pass
                                    
                    if p.get("MAT_CUSTOM") != "-" and "TUBO" not in p["CÓDIGO"]:
                        if p.get("COMP PL", "-") != "-" and p.get("LARG PL", "-") != "-":
                            p["PESO UNIT"] = float(p["COMP PL"]) * float(p["LARG PL"]) * p["ESP"] * 0.000008
                        
                    pecas_processadas.append(p)
                
                if item["tipo"] == "parametrico" and "PRAT" in item.get("base_cod", "") and not any(p["DESC"] == "REFORÇO PRATELEIRA" for p in pecas_processadas):
                    if molde_reforco:
                        ref_prat = molde_reforco.copy()
                        ref_prat["DESC"] = "REFORÇO PRATELEIRA"
                        ref_prat["MAT_CUSTOM"] = item.get("mat_ref_prat", "INOX 430")
                        ref_prat["ESP"] = float(item.get("esp_ref_prat", 0.8))
                        qtd_r = item.get("qtd_ref_prat", "Padrão")
                        mult = 2 if "DUPLA" in item.get("base_cod", "") else 1
                        if qtd_r != "Padrão":
                            if qtd_r != "0": 
                                ref_prat["QTD"] = int(qtd_r) * mult
                                ref_prat["PESO UNIT"] = float(ref_prat["COMP PL"]) * float(ref_prat["LARG PL"]) * ref_prat["ESP"] * 0.000008
                                pecas_processadas.append(ref_prat)
                        else:
                            ref_prat["QTD"] = ref_prat.get("QTD", 1) * mult
                            ref_prat["PESO UNIT"] = float(ref_prat["COMP PL"]) * float(ref_prat["LARG PL"]) * ref_prat["ESP"] * 0.000008
                            pecas_processadas.append(ref_prat)

                for p in pecas_processadas:
                    qtd_final = p["QTD"] * item.get("qtd", item.get("qtd_item", 1))
                    custo_peca_un = 0.0
                    mat_peca = p.get("MAT_CUSTOM", "-")
                    
                    if "TUBO" in p["CÓDIGO"]:
                        metros_unit = float(p.get("COMP PL", 0)) / 1000.0
                        metros_tot = metros_unit * qtd_final
                        peso_peca = 0
                        custo_peca_un = metros_unit * custo_tubo.get(p["CÓDIGO"], 0.0)
                        custo_unit_item_tubo += custo_peca_un * p["QTD"]
                        resumo_geral_tubos[p["CÓDIGO"]] = resumo_geral_tubos.get(p["CÓDIGO"], 0) + (float(p.get("COMP PL", 0)) * qtd_final)
                        linha_med = f"{metros_unit:.2f} M un | {metros_tot:.2f} M tot".replace('.',',')
                    
                    elif "CHAPA" in p["CÓDIGO"] or (p.get("COMP PL", "-") != "-" and p.get("LARG PL", "-") != "-"):
                        peso_peca = p.get("PESO UNIT", 0) * qtd_final
                        if p.get("LARG PL", "-") != "-":
                            for _ in range(qtd_final): pecas_para_nesting_global.append({'nome': p['DESC'], 'material': f"{mat_peca} {p.get('ESP','-')}mm", 'comp': float(p["COMP PL"]), 'larg': float(p["LARG PL"])})
                        if p.get("ESP", 0) > 0:
                            custo_peca_un = p.get("PESO UNIT", 0) * custo_chapa.get(mat_peca, {}).get(p["ESP"], 0.0)
                            custo_unit_item_chapa += custo_peca_un * p["QTD"]
                            resumo_geral_chapas[(mat_peca, p["ESP"])] = resumo_geral_chapas.get((mat_peca, p["ESP"]), 0) + peso_peca
                        linha_med = f"{p.get('PESO UNIT', 0):.2f} KG un | {peso_peca:.2f} KG tot".replace('.',',')
                        
                    else:
                        peso_peca = 0
                        linha_med = "- | -"

                    peso_item_total += peso_peca
                    mat_str = f"{mat_peca} " if mat_peca != "-" else ""
                    html_detalhamento += f"<div class='item-linha'>&#x25AB; {qtd_final}x {mat_str}- {p['DESC']} | {linha_med} | Custo un: R$ {custo_peca_un:.2f}</div>".replace('.',',')
                    
                    dados_para_df.append({
                        "ITEM": f"Item {item['num']}", "QTD": qtd_final, "CÓDIGO": p["CÓDIGO"],
                        "DESCRIÇÃO": p["DESC"], "MATERIAL": p.get("MAT_CUSTOM", "-"), "ESP": p.get("ESP", "-"),
                        "MEDIDA (C x L)": f"{p.get('COMP PL','-')} x {p.get('LARG PL','-')}", "PESO TOT (KG)": round(peso_peca, 2)
                    })
                    
                c_tot = (custo_unit_item_chapa + custo_unit_item_tubo) * item.get("qtd", item.get("qtd_item", 1))
                peso_total_pedido += peso_item_total
                html_detalhamento += f"<hr><div class='item-linha' style='color:#fff;'><b>CUSTO DO ITEM: R$ {c_tot:.2f} | Peso: {peso_item_total:.2f} KG</b></div></div>".replace('.',',')

            custo_tot_chapas = sum([peso * custo_chapa.get(mat, {}).get(esp, 0.0) for (mat, esp), peso in resumo_geral_chapas.items()])
            custo_tot_tubos = sum([(c_total / 1000.0) * custo_tubo.get(cod, 0.0) for cod, c_total in resumo_geral_tubos.items()])
            custo_total_geral = custo_tot_chapas + custo_tot_tubos
            
            tab_lista, tab_dash = st.tabs(["✂️ Detalhamento de Corte", "📊 Dashboard Econômico"])
            
            with tab_dash:
                titulo_pedido = f"CUSTO DO PEDIDO {inp_pedido_nome}" if inp_pedido_nome else "CUSTO TOTAL APROXIMADO"
                st.markdown(f"""
                <div class="box-azul">
                    <h4>{titulo_pedido}</h4>
                    <h1>R$ {custo_total_geral:,.2f}</h1>
                </div>
                """.replace(',', 'X').replace('.', ',').replace('X', '.'), unsafe_allow_html=True)
                
                c_dash1, c_dash2 = st.columns(2)
                with c_dash1:
                    if resumo_geral_chapas:
                        html_chapas = "<div class='card-dark'><h4>📦 CHAPAS INOX</h4>"
                        for (mat, esp), peso in sorted(resumo_geral_chapas.items()):
                            prc = custo_chapa.get(mat, {}).get(esp, 0.0)
                            html_chapas += f"<p>• {mat} {esp}mm ➔ {peso:.2f} KG | R$ {peso*prc:.2f}</p>".replace('.',',')
                            linhas_resumo_excel.append({"TIPO": "CHAPA", "MATERIAL": f"{mat} {esp}mm", "QTD": f"{peso:.2f} KG".replace('.', ','), "CUSTO (R$)": round(peso*prc, 2)})
                        html_chapas += f"<hr><p style='font-weight:bold;'>Peso Total: {peso_total_pedido:.2f} KG</p></div>".replace('.',',')
                        st.markdown(html_chapas, unsafe_allow_html=True)
                        
                with c_dash2:
                    if resumo_geral_tubos:
                        html_tubos = "<div class='card-dark'><h4>📎 TUBOS E PERFIS</h4>"
                        for cod, c_total in resumo_geral_tubos.items():
                            mts = c_total / 1000.0; prc = custo_tubo.get(cod, 0.0)
                            html_tubos += f"<p>• {cod} ➔ {mts:.2f} Mts ({(mts/6.0):.1f} br) | R$ {mts*prc:.2f}</p>".replace('.',',')
                            linhas_resumo_excel.append({"TIPO": "TUBO", "MATERIAL": f"{cod}", "QTD": f"{mts:.2f} M ({(mts/6.0):.1f} un)".replace('.', ','), "CUSTO (R$)": round(mts*prc, 2)})
                        html_tubos += "</div>"
                        st.markdown(html_tubos, unsafe_allow_html=True)
                    
                if HAS_NESTING and pecas_para_nesting_global:
                    html_nest = f"<div class='card-dark'><h4>🧩 ESTIMATIVA DE ENCAIXE (3000x1250)</h4>"
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
                        
                        html_nest += f"<p>• {mat_nome}: <b>{chp_usadas} chapa(s)</b>. (Aprov: {aprov:.1f}%)</p>".replace('.',',')
                        dados_para_graficos[mat_nome] = packer
                        
                    html_nest += "</div>"
                    st.markdown(html_nest, unsafe_allow_html=True)
                    
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

            with tab_lista:
                st.markdown(html_detalhamento, unsafe_allow_html=True)
                
                df_resumo = pd.DataFrame(linhas_resumo_excel)
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    pd.DataFrame(dados_para_df).to_excel(writer, sheet_name='Lista_Corte', index=False)
                    if not df_resumo.empty: df_resumo.to_excel(writer, sheet_name='Resumo_Financeiro', index=False)
                
                st.download_button(label="📥 Baixar Excel Completo (PCP e Custos)", data=buffer.getvalue(), file_name=f'PCP_AcoNobre_{inp_pedido_nome}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)
