import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from rectpack import newPacker, PackingMode, PackingBin

from regras_engenharia import calcular_mesa_parametrica, calcular_estante_parametrica, calcular_prateleira_parede

# Configuração da Página Web
st.set_page_config(page_title="Projetista AçoNobre", page_icon="🛒", layout="wide")

# ==========================================
# VARIÁVEIS DE SESSÃO E MAPAS
# ==========================================
if 'carrinho' not in st.session_state: st.session_state.carrinho = []
if 'pecas_temp_composto' not in st.session_state: st.session_state.pecas_temp_composto = []
if 'pecas_para_nesting_global' not in st.session_state: st.session_state.pecas_para_nesting_global = []

if 'custo_chapa' not in st.session_state:
    st.session_state.custo_chapa = {
        "INOX 304": {0.6: 29.01, 0.8: 33.16, 1.0: 30.74, 1.2: 28.66, 1.5: 29.00, 2.0: 30.96},
        "INOX 430": {0.6: 23.32, 0.8: 19.41, 1.0: 18.91, 1.2: 19.42, 1.5: 19.38, 2.0: 19.38},
        "INOX 201": {0.6: 15.00, 0.8: 16.00, 1.0: 16.50, 1.2: 17.00}
    }
if 'custo_tubo' not in st.session_state:
    st.session_state.custo_tubo = {"TUBO-RED-38": 18.47, "TUBO-RED-25": 12.08}

# Tenta carregar do Excel se existir na pasta
try:
    df_chapas = pd.read_excel("CHAPAS.xlsx", sheet_name="CHAPAS")
    st.session_state.custo_chapa["INOX 304"] = dict(zip(df_chapas[df_chapas['Aisi']==304]['Espessura'], df_chapas[df_chapas['Aisi']==304]['Custo última compra']))
    st.session_state.custo_chapa["INOX 430"] = dict(zip(df_chapas[df_chapas['Aisi']==430]['Espessura'], df_chapas[df_chapas['Aisi']==430]['Custo última compra']))
    df_tubos = pd.read_excel("CHAPAS.xlsx", sheet_name="TUBOS")
    for _, r in df_tubos.iterrows():
        if "016.201.020" in str(r['Código']): st.session_state.custo_tubo["TUBO-RED-38"] = float(r['Custo última compra'])
        elif "016.201.014" in str(r['Código']): st.session_state.custo_tubo["TUBO-RED-25"] = float(r['Custo última compra'])
except: pass

mapa_tampo = {"Mesa Lisa": "LISA", "Mesa com Encosto": "ENCOSTO", "Pia com Cuba": "PIA", "Prateleira Parede": "PRAT_PAREDE", "Estante Lisa": "ESTANTE_LISA", "Estante Gradeada": "ESTANTE_GRADEADA"}
mapa_base = {"Contraventamento": "CONTRAVENTAMENTO", "Prat. Lisa": "PRAT_LISA", "Prat. Lisa Dupla": "PRAT_LISA_DUPLA", "Prat. Gradeada": "PRAT_GRADEADA", "Prat. Gradeada Dupla": "PRAT_GRADEADA_DUPLA"}

dim_chapa_x = 3000
dim_chapa_y = 1250

# ==========================================
# INTERFACE PRINCIPAL - PROJETISTA VIRTUAL
# ==========================================
st.title("🛒 Projetista Virtual - AçoNobre")
st.markdown("Crie orçamentos dinâmicos, visualize os custos e gere o mapa de corte em tempo real.")

col_esq, col_dir = st.columns([1, 2])

with col_esq:
    st.subheader("📦 Montar Pedido")
    pedido_num = st.text_input("Nº do Pedido (Ex: 6885)")

    aba_param, aba_comp, aba_livre = st.tabs(["📐 Catálogo Paramétrico", "🧩 Item Composto", "✏️ Peça Avulsa"])
    
    # --- ABA PARAMÉTRICO ---
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

    # --- ABA ITEM COMPOSTO ---
    with aba_comp:
        cc1, cc2 = st.columns([1, 2])
        qtd_item_comp = cc1.number_input("QTD Módulo", min_value=1, value=1)
        nome_item_comp = cc2.text_input("Nome Módulo", placeholder="Ex: Gaveteiro")
        
        st.markdown("**Inserir Peças:**")
        cp1, cp2 = st.columns([2, 1])
        nome_pc = cp1.text_input("Nome da Peça")
        qtd_pc = cp2.number_input("QTD Peça", min_value=1, value=1)
        
        cp3, cp4, cp5, cp6 = st.columns([1, 1, 1, 1])
        comp_pc = cp3.number_input("C(mm) ", value=0.0)
        larg_pc = cp4.number_input("L(mm) ", value=0.0)
        mat_pc = cp5.selectbox("Mat", ["304", "430", "201"])
        esp_pc = cp6.selectbox("Esp", ["0.6", "0.8", "1.0", "1.2", "1.5", "2.0"], index=2)

        if st.button("⏬ Inserir Peça no Módulo", use_container_width=True):
            if nome_pc and comp_pc > 0 and larg_pc > 0:
                st.session_state.pecas_temp_composto.append({
                    "nome": nome_pc, "qtd": qtd_pc, "comp": comp_pc, "larg": larg_pc, 
                    "mat": f"INOX {mat_pc}", "esp": float(esp_pc)
                })
                st.rerun()
        
        if st.session_state.pecas_temp_composto:
            st.dataframe(pd.DataFrame(st.session_state.pecas_temp_composto), use_container_width=True)
            if st.button("✅ EMPACOTAR ITEM E ADD", type="primary", use_container_width=True):
                item = {
                    "num": len(st.session_state.carrinho)+1, "tipo": "composto", 
                    "nome_item": nome_item_comp if nome_item_comp else "Item Composto", 
                    "qtd_item": qtd_item_comp, "pecas": list(st.session_state.pecas_temp_composto),
                    "desc_carrinho": f"🧩 {nome_item_comp if nome_item_comp else 'Módulo Composto'}"
                }
                st.session_state.carrinho.append(item)
                st.session_state.pecas_temp_composto.clear()
                st.rerun()

    # --- ABA PEÇA AVULSA ---
    with aba_livre:
        cl1, cl2 = st.columns([1, 2])
        qtd_livre = cl1.number_input("QTD Avulsa", min_value=1, value=1)
        nome_livre = cl2.text_input("Nome da Peça Avulsa")
        
        cl3, cl4 = st.columns(2)
        comp_livre = cl3.number_input("C Planif (mm)", value=0.0)
        larg_livre = cl4.number_input("L Planif (mm)", value=0.0)
        
        cl5, cl6 = st.columns(2)
        mat_livre = cl5.selectbox("Material", ["INOX 304", "INOX 430", "INOX 201"])
        esp_livre = cl6.selectbox("Espessura", ["0.6", "0.8", "1.0", "1.2", "1.5", "2.0"], index=2)

        if st.button("➕ Adicionar Peça Solta", type="primary", use_container_width=True):
            if nome_livre and comp_livre > 0 and larg_livre > 0:
                item = {
                    "num": len(st.session_state.carrinho)+1, "tipo": "livre", 
                    "qtd": qtd_livre, "nome_peca": nome_livre, 
                    "comp_pl": comp_livre, "larg_pl": larg_livre, 
                    "esp": float(esp_livre), "material": mat_livre,
                    "desc_carrinho": f"✏️ {nome_livre}"
                }
                st.session_state.carrinho.append(item)
                st.rerun()

    st.markdown("### 🛒 Carrinho")
    if st.session_state.carrinho:
        for idx, it in enumerate(st.session_state.carrinho):
            st.info(f"**{it.get('qtd', it.get('qtd_item', 1))}x** {it['desc_carrinho']}")
        if st.button("❌ Limpar Carrinho", use_container_width=True):
            st.session_state.carrinho = []
            st.rerun()
    else:
        st.warning("Carrinho vazio.")


# ----------------- TELA DE RESULTADOS DO PROJETISTA -----------------
with col_dir:
    st.subheader("📊 Painel de Engenharia")
    
    if st.session_state.carrinho:
        aba_lista_pcp, aba_dashboard, aba_nesting = st.tabs(["✂️ Lista de Corte", "📈 Dashboard Financeiro", "🧩 Mapa de Corte (Nesting)"])
        
        dados_para_df = []
        st.session_state.pecas_para_nesting_global = []
        resumo_geral_chapas = {}
        resumo_geral_tubos = {}
        peso_total_pedido = 0.0
        custo_total_geral = 0.0

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
                            if "MAIOR" in p["DESC"].upper(): p["DESC"] = "REFORÇO TRASEIRO DO TAMPO"; is_traseiro = True
                            elif "PRAT" in p["DESC"].upper() or "BASE" in p["DESC"].upper(): p["DESC"] = "REFORÇO PRATELEIRA"; is_prat = True
                            elif "PLANO" in p["DESC"].upper() or "ESTANTE" in item["tampo_cod"]: p["DESC"] = "REFORÇO PLANO"; is_plano = True
                            else: p["DESC"] = "REFORÇO TAMPO"; is_tampo = True
                                
                            if is_traseiro:
                                p["MAT_CUSTOM"] = item["mat_tampo"]
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

            # --- POPULANDO DADOS DO ITEM ---
            qtd_multiplicador = item.get("qtd", item.get("qtd_item", 1))
            peso_item_total = 0.0
            custo_unit_item = 0.0
            linhas_ui_item = []

            for p in pecas_base:
                qtd_final = p["QTD"] * qtd_multiplicador
                peso_total_final = p["PESO UNIT"] * qtd_final
                mat_peca = p.get("MAT_CUSTOM", "-")
                desc_final = p["DESC"]
                custo_peca_un = 0.0
                med = f"{p['COMP PL']}x{p['LARG PL']}" if p['LARG PL'] != "-" else f"{p['COMP PL']}mm" if p['COMP PL'] != "-" else "-"
                esp_str = str(p['ESP']).replace('.', ',') if p['ESP'] > 0 else "-"

                if "CHAPA" in p["CÓDIGO"] or (p["COMP PL"] != "-" and p["LARG PL"] != "-"):
                    if p["LARG PL"] != "-":
                        for _ in range(qtd_final): st.session_state.pecas_para_nesting_global.append({'nome': p['DESC'], 'material': f"{mat_peca} {p['ESP']}mm", 'comp': float(p["COMP PL"]), 'larg': float(p["LARG PL"])})
                    if p["ESP"] > 0:
                        custo_peca_un = p["PESO UNIT"] * st.session_state.custo_chapa.get(mat_peca, {}).get(p["ESP"], 0.0)
                        custo_unit_item += custo_peca_un * p["QTD"]
                        chave = (mat_peca, p["ESP"])
                        resumo_geral_chapas[chave] = resumo_geral_chapas.get(chave, 0) + peso_total_final
                    linha_medida = f"{p['PESO UNIT']:.2f} KG un. | {peso_total_final:.2f} KG tot."
                elif "TUBO" in p["CÓDIGO"]:
                    metros_unit = float(p["COMP PL"]) / 1000.0
                    metros_tot = metros_unit * qtd_final
                    custo_peca_un = metros_unit * st.session_state.custo_tubo.get(p["CÓDIGO"], 0.0)
                    custo_unit_item += custo_peca_un * p["QTD"]
                    resumo_geral_tubos[p["CÓDIGO"]] = resumo_geral_tubos.get(p["CÓDIGO"], 0) + (float(p["COMP PL"]) * qtd_final)
                    linha_medida = f"{metros_unit:.2f} M un.  | {metros_tot:.2f} M tot."
                else:
                    linha_medida = "- | -"

                linhas_ui_item.append(f"▫ **{qtd_final}x {mat_peca} - {desc_final}** | {linha_medida} | R$ {custo_peca_un:.2f}")

                dados_para_df.append({
                    "ITEM": f"Item {item['num']}", "QTD": qtd_final, "CÓDIGO PEÇA": p['CÓDIGO'], 
                    "DESCRIÇÃO": p['DESC'], "MATERIAL": mat_peca, "ESPESSURA": esp_str, 
                    "MEDIDA CORTE": med, "PESO UNIT (KG)": round(p['PESO UNIT'],2), "PESO TOTAL (KG)": round(peso_total_final,2)
                })
                peso_item_total += peso_total_final

            c_tot = custo_unit_item * qtd_multiplicador
            peso_total_pedido += peso_item_total
            
            with aba_lista_pcp:
                with st.expander(f"▶ ITEM {item['num']}: {qtd_multiplicador}x {item['desc_carrinho']} (R$ {c_tot:.2f})", expanded=False):
                    for linha in linhas_ui_item: st.markdown(linha.replace('.', ','))
                    st.info(f"**Custo Unitário:** R$ {custo_unit_item:.2f} | **Custo Total:** R$ {c_tot:.2f} | **Peso Total:** {peso_item_total:.2f} KG".replace('.', ','))

        # --- TOTAIS DASHBOARD ---
        custo_tot_chapas = sum([peso * st.session_state.custo_chapa.get(mat, {}).get(esp, 0.0) for (mat, esp), peso in resumo_geral_chapas.items()])
        custo_tot_tubos = sum([(c_total / 1000.0) * st.session_state.custo_tubo.get(cod, 0.0) for cod, c_total in resumo_geral_tubos.items()])
        custo_total_geral = custo_tot_chapas + custo_tot_tubos

        with aba_dashboard:
            st.success(f"### 💰 Custo Total Fabril: R$ {custo_total_geral:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            g1, g2 = st.columns(2)
            with g1:
                st.markdown("#### 📦 Consumo de Chapas INOX")
                for (mat, esp), peso in sorted(resumo_geral_chapas.items()):
                    prc = st.session_state.custo_chapa.get(mat, {}).get(esp, 0.0)
                    st.write(f"• {mat} {esp}mm ➔ **{peso:.2f} KG** (R$ {prc:.2f}/kg) | Sub: R$ {peso*prc:.2f}".replace('.', ','))
                st.caption(f"Peso Total em Chapas: {peso_total_pedido:.2f} KG".replace('.', ','))
            with g2:
                st.markdown("#### 📏 Consumo de Tubos e Perfis")
                for cod, c_total in resumo_geral_tubos.items():
                    mts = c_total / 1000.0
                    prc = st.session_state.custo_tubo.get(cod, 0.0)
                    st.write(f"• {cod} ➔ **{mts:.2f} Mts** ({(mts/6.0):.1f} barras) | Sub: R$ {mts*prc:.2f}".replace('.', ','))

        # --- MOTOR DE NESTING VISUAL ---
        with aba_nesting:
            if HAS_NESTING and st.session_state.pecas_para_nesting_global:
                st.markdown(f"#### 🧩 Mapa de Corte (Chapa Padrão {dim_chapa_x}x{dim_chapa_y})")
                if st.button("👁️ Gerar Visão de Encaixe", type="primary", use_container_width=True):
                    agrup_nesting = {}
                    for pc in st.session_state.pecas_para_nesting_global:
                        chave = pc['material']
                        if chave not in agrup_nesting: agrup_nesting[chave] = []
                        agrup_nesting[chave].append((pc['comp'], pc['larg']))
                    
                    for mat_nome, dimensoes in agrup_nesting.items():
                        st.markdown(f"**Material:** {mat_nome}")
                        packer = newPacker(mode=PackingMode.Offline, bin_algo=PackingBin.Global, rotation=True)
                        for idx, (c, l) in enumerate(dimensoes): packer.add_rect(width=c + 5, height=l + 5, rid=idx)
                        for _ in range(50): packer.add_bin(width=dim_chapa_x, height=dim_chapa_y)
                        packer.pack()
                        chp_usadas = len(packer)
                        area_pc = sum([c * l for c, l in dimensoes])
                        area_ch = chp_usadas * (dim_chapa_x * dim_chapa_y)
                        aprov = (area_pc / area_ch) * 100 if area_ch > 0 else 0
                        
                        st.info(f"**{chp_usadas} chapa(s)** necessárias. (Aproveitamento: {aprov:.1f}%)")
                        
                        # Plota os gráficos com o matplotlib direto no Streamlit
                        for b_idx, bin in enumerate(packer):
                            fig, ax = plt.subplots(figsize=(10, 4))
                            ax.set_xlim(0, dim_chapa_x); ax.set_ylim(0, dim_chapa_y)
                            ax.set_title(f"Plano de Corte - {mat_nome} | Chapa {b_idx + 1}")
                            ax.add_patch(patches.Rectangle((0, 0), dim_chapa_x, dim_chapa_y, facecolor='#ecf0f1', edgecolor='black'))
                            for rect in bin: ax.add_patch(patches.Rectangle((rect.x, rect.y), rect.width, rect.height, facecolor='#2980b9', edgecolor='white'))
                            plt.gca().set_aspect('equal', adjustable='box'); plt.tight_layout()
                            st.pyplot(fig)
                        st.markdown("---")

        with aba_lista_pcp:
            st.markdown("---")
            df_lista = pd.DataFrame(dados_para_df)
            st.dataframe(df_lista, use_container_width=True, hide_index=True)

            if not df_lista.empty:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_lista.to_excel(writer, index=False, sheet_name='Lista_Corte')
                excel_data = output.getvalue()
                
                st.download_button(
                    label="📥 Baixar Tabela em Excel",
                    data=excel_data,
                    file_name=f"PCP_Pedido_{pedido_num}.xlsx" if pedido_num else "PCP_Orcamento.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
