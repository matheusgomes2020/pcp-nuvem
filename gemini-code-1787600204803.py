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
    except: return 0.0, 0

def calcular_tempo_e_custo_laser(espessura_str, perimetro_mm, entradas):
    try:
        esp = float(str(espessura_str).replace(',', '.'))
        avanco, peck = PARAMETROS_LASER_INOX.get(esp, (5000, 1.5))
        tempo_corte = perimetro_mm / avanco
        tempo_furo = (entradas * peck) / 60.0
        return round(tempo_corte + tempo_furo, 2)
    except: return 0.0

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

            # INTELIGÊNCIA VISUAL: Valores padrão
            alt = 0.0; planos_val = 4; base_nome = "Contraventamento"; cod_base = "CONTRAVENTAMENTO"
            
            # Adaptação dinâmica da tela
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
            # Oculta a escolha de base se for contraventamento ou prat de parede
            if cod_tampo != "PRAT_PAREDE" and cod_base != "CONTRAVENTAMENTO":
                # Sincroniza o índice baseando-se no Tampo
                idx_mat = ["INOX 304", "INOX 430", "INOX 201"].index(mat_tampo)
                c_base1, c_base2 = st.columns(2)
                mat_base = c_base1.selectbox(lbl_base, ["INOX 304", "INOX 430", "INOX 201"], index=idx_mat)
                esp_base = c_base2.selectbox(f"Esp. {lbl_base}", ["Esp. Padrão", "0.8", "1.0", "1.2", "1.5", "2.0"])

            # REFORÇOS DINÂMICOS
            st.markdown("---")
            st.markdown("**Engenharia de Reforços**")
            
            lbl_ref1 = "Ref. Plano" if "ESTANTE" in cod_tampo else "Reforço" if cod_tampo == "PRAT_PAREDE" else "Ref. Tampo"
            
            c_r1, c_r2, c_r3 = st.columns(3)
            qtd_ref = c_r1.selectbox(f"Qtd {lbl_ref1}", ["Padrão", "0", "1", "2", "3", "4", "5", "6"])
            idx_ref1 = ["INOX 430", "INOX 304", "INOX 201"].index(mat_tampo) if mat_tampo in ["INOX 430", "INOX 304", "INOX 201"] else 0
            mat_ref = c_r2.selectbox(f"Mat. {lbl_ref1}", ["INOX 430", "INOX 304", "INOX 201"], index=idx_ref1)
            esp_ref = c_r3.selectbox(f"Esp. {lbl_ref1}", ["0.6", "0.8", "1.0", "1.2", "1.5"], index=1) # Index 1 é "0.8"

            mat_ref_prat = "INOX 430"; esp_ref_prat = "0.8"; qtd_ref_prat = "Padrão"
            # Aparece apenas se tiver prateleira inferior na jogada
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
                
                # Monta a descrição
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
                
                # CHAMA O MOTOR DE GEOMETRIA (Regras_Engenharia)
                if item["tipo"] == "parametrico":
                    if "ESTANTE" in item["tampo_cod"]:
                        tipo_est = "LISA" if "LISA" in item["tampo_cod"] else "GRADEADA"
                        pb_cruas = calcular_estante_parametrica(item["comp"], item["larg"], item["alt"], planos=item.get("planos", 4), tipo=tipo_est, material=item["mat_tampo"])
                    elif item["tampo_cod"] == "PRAT_PAREDE":
                        pb_cruas = calcular_prateleira_parede(item["comp"], item["larg"], item["mat_tampo"])
                    else:
                        pb_cruas = calcular_mesa_parametrica(item["comp"], item["larg"], item["alt"], item["tampo_cod"], item["base_cod"])

                    molde_reforco = None

                    # INJETOR DINÂMICO DE REFORÇOS
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

            # Exibição de Custos e Totais na Tela
            custo_tot_chapas = sum([peso * st.session_state.custo_chapa.get(mat, {}).get(esp, 0.0) for (mat, esp), peso in resumo_geral_chapas.items()])
            custo_tot_tubos = sum([(c_total / 1000.0) * st.session_state.custo_tubo.get(cod, 0.0) for cod, c_total in resumo_geral_tubos.items()])
            custo_total_geral = custo_tot_chapas + custo_tot_tubos

            st.success(f"💰 **Custo Total Fabril:** R$ {custo_total_geral:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            
            # Grids com Resumo
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

            # Botão de Exportar para Excel (Cria o buffer em memória)
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


# ----------------- ABA PCP (ARQUIVOS LOCAIS) -----------------
with tab_pcp:
    st.subheader("Gerador de PCP - Tratamento de Engenharia")
    st.info("⚠️ *Aviso: Como o Streamlit roda no navegador, ele precisa que você copie e cole os caminhos físicos das pastas da rede.*")

    col1, col2 = st.columns(2)
    pasta_pdf = col1.text_input("Pasta Origem (PDFs):", value=r"\\Servidor\aco_nobre\5 - ENG DESENV\00-PROJETOS\NOVA CODIFICAÇÃO\PDF")
    pasta_dxf = col2.text_input("Pasta Origem (DXFs):", value=r"\\Servidor\aco_nobre\5 - ENG DESENV\00-PROJETOS\NOVA CODIFICAÇÃO\DXF")
    pasta_destino = st.text_input("Pasta Destino (Resultados):", placeholder=r"C:\Users\Acn.Gomes\Downloads")

    st.markdown("---")
    st.write("Para habilitar o processamento em lote em nuvem, você deve fazer o Upload da planilha comercial:")
    arquivo_planilha = st.file_uploader("Selecione a Planilha Comercial (.xlsx)", type=['xlsx'])

    st.write("Ou atualizar item avulso:")
    c1, c2, c3, c4 = st.columns(4)
    ped_avulso = c1.text_input("Nº Pedido")
    item_avulso = c2.text_input("Item")
    cod_avulso = c3.text_input("Código Pai")
    qtd_avulso = c4.text_input("Qtd")

    c_opt1, c_opt2, c_opt3 = st.columns(3)
    chk_pdf = c_opt1.checkbox("Unir PDFs e DXFs", value=True)
    chk_conf = c_opt2.checkbox("Gerar Conferência", value=True)
    chk_pcp = c_opt3.checkbox("Atualizar Excel", value=True)

    if st.button("🚀 INICIAR PROCESSAMENTO PCP WEB", type="primary", use_container_width=True):
        st.warning("Motor ativado. No ambiente de nuvem real, a leitura em pastas de rede (\\Servidor) depende de um script rodando localmente na máquina servidora. A interface Web enviará o comando para o motor via rede.")