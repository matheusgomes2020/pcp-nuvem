import math

def calcular_mesa_parametrica(comp, larg, alt, tipo_tampo="LISA", tipo_base="CONTRAVENTAMENTO"):
    """
    Motor de Engenharia AçoNobre.
    Gera a lista de peças planificadas deduzindo as medidas reais.
    Dobra Padrão: Allowance de 104mm para planificação.
    Fator de Peso: Constante 0.000008
    """
    pecas = []
    
    # ==========================================
    # LÓGICA INTELIGENTE DE "ESPESSURA PADRÃO"
    # ==========================================
    if tipo_tampo == "PIA":
        esp_padrao_tampo = 1.2
    elif comp > 1500:
        esp_padrao_tampo = 1.2
    else:
        esp_padrao_tampo = 1.0

    esp_padrao_base = 1.2 if comp > 1500 else 1.0

    # ==========================================
    # 1. TAMPO PRINCIPAL E REFORÇOS
    # ==========================================
    comp_pl = comp + 104
    larg_pl = larg + 104
    
    desc_tampo = "TAMPO MESA LISA"

    if tipo_tampo == "ENCOSTO":
        desc_tampo = "TAMPO MESA C/ ENCOSTO"
        larg_pl += 100 
    elif tipo_tampo == "PIA":
        desc_tampo = "TAMPO PIA C/ CUBA"
        larg_pl += 100 
        pecas.append({"CÓDIGO": "CUBA_PADRAO", "DESC": "CUBA INOX PADRÃO", "QTD": 1, "COMP PL": "-", "LARG PL": "-", "ESP": 0, "PESO UNIT": 3.5})

    pecas.append({
        "CÓDIGO": "CHAPA", "DESC": desc_tampo, "QTD": 1, 
        "COMP PL": comp_pl, "LARG PL": larg_pl, "ESP": esp_padrao_tampo, 
        "PESO UNIT": comp_pl * larg_pl * esp_padrao_tampo * 0.000008
    })
    
    comp_ref_maior = comp - 100
    comp_ref_menor = larg - 100
    larg_ref_pl = 104
    
    pecas.append({"CÓDIGO": "CHAPA", "DESC": "REFORÇO OMEGA MAIOR", "QTD": 2, "COMP PL": comp_ref_maior, "LARG PL": larg_ref_pl, "ESP": 0.8, "PESO UNIT": comp_ref_maior * larg_ref_pl * 0.8 * 0.000008})
    pecas.append({"CÓDIGO": "CHAPA", "DESC": "REFORÇO OMEGA MENOR", "QTD": 2, "COMP PL": comp_ref_menor, "LARG PL": larg_ref_pl, "ESP": 0.8, "PESO UNIT": comp_ref_menor * larg_ref_pl * 0.8 * 0.000008})

    # ==========================================
    # 2. PÉS / ESTRUTURA
    # ==========================================
    alt_tubo = alt - 40 
    pecas.append({"CÓDIGO": "TUBO-RED-38", "DESC": "PÉ TUBO REDONDO 38mm", "QTD": 4, "COMP PL": alt_tubo, "LARG PL": "-", "ESP": 1.2, "MAT_CUSTOM": "INOX 201 ESC", "PESO UNIT": 0})
    pecas.append({"CÓDIGO": "SAPATA", "DESC": "SAPATA NIVELADORA PEAD", "QTD": 4, "COMP PL": "-", "LARG PL": "-", "ESP": 0, "PESO UNIT": 0.1})

    # ==========================================
    # 3. BASE (PRATELEIRA OU CONTRAVENTAMENTO)
    # ==========================================
    qtd_prateleiras = 2 if "DUPLA" in tipo_base else 1

    if tipo_base == "CONTRAVENTAMENTO":
        c_contra_maior = comp - 120
        c_contra_menor = larg - 120
        pecas.append({"CÓDIGO": "TUBO-RED-25", "DESC": "CONTRAVENTAMENTO MAIOR", "QTD": 1, "COMP PL": c_contra_maior, "LARG PL": "-", "ESP": 1.2, "MAT_CUSTOM": "INOX 201 ESC", "PESO UNIT": 0})
        pecas.append({"CÓDIGO": "TUBO-RED-25", "DESC": "CONTRAVENTAMENTO MENOR", "QTD": 2, "COMP PL": c_contra_menor, "LARG PL": "-", "ESP": 1.2, "MAT_CUSTOM": "INOX 201 ESC", "PESO UNIT": 0})
        
    elif tipo_base in ["PRAT_LISA", "PRAT_LISA_DUPLA"]:
        c_prat_pl = (comp - 80) + 104
        l_prat_pl = (larg - 80) + 104
        desc_prat = "PRATELEIRA LISA INFERIOR DUPLA" if qtd_prateleiras == 2 else "PRATELEIRA LISA INFERIOR"
        
        pecas.append({"CÓDIGO": "CHAPA", "DESC": desc_prat, "QTD": 1 * qtd_prateleiras, "COMP PL": c_prat_pl, "LARG PL": l_prat_pl, "ESP": esp_padrao_base, "PESO UNIT": c_prat_pl * l_prat_pl * esp_padrao_base * 0.000008})
        
    elif tipo_base in ["PRAT_GRADEADA", "PRAT_GRADEADA_DUPLA"]:
        c_prat = comp - 80
        l_prat = larg - 80
        qtd_divisorias = math.floor(c_prat / 100)
        
        pecas.append({"CÓDIGO": "CHAPA", "DESC": "QUADRO PRAT. GRADEADA MAIOR", "QTD": 2 * qtd_prateleiras, "COMP PL": c_prat, "LARG PL": 104, "ESP": esp_padrao_base, "PESO UNIT": c_prat * 104 * esp_padrao_base * 0.000008})
        pecas.append({"CÓDIGO": "CHAPA", "DESC": "QUADRO PRAT. GRADEADA MENOR", "QTD": 2 * qtd_prateleiras, "COMP PL": l_prat, "LARG PL": 104, "ESP": esp_padrao_base, "PESO UNIT": l_prat * 104 * esp_padrao_base * 0.000008})
        pecas.append({"CÓDIGO": "CHAPA", "DESC": "PERFIL U DIVISÓRIA GRADE", "QTD": qtd_divisorias * qtd_prateleiras, "COMP PL": l_prat, "LARG PL": 60, "ESP": esp_padrao_base, "PESO UNIT": l_prat * 60 * esp_padrao_base * 0.000008})

    return pecas


def calcular_estante_parametrica(comp, larg, alt, planos=4, tipo="GRADEADA", material="INOX 304"):
    """
    Motor de Engenharia AçoNobre para Estantes.
    """
    pecas = []
    peso_fator = 0.000008

    comp_coluna = alt - 45.6
    pecas.append({"CÓDIGO": "CHAPA", "DESC": "COLUNA DIREITA", "QTD": 2, "COMP PL": comp_coluna, "LARG PL": 103.2, "ESP": 1.2, "MAT_CUSTOM": material})
    pecas.append({"CÓDIGO": "CHAPA", "DESC": "COLUNA ESQUERDA", "QTD": 2, "COMP PL": comp_coluna, "LARG PL": 103.2, "ESP": 1.2, "MAT_CUSTOM": material})

    pecas.append({"CÓDIGO": "TUBO-RED-38", "DESC": "TUBO DA SAPATA (50mm)", "QTD": 4, "COMP PL": 50, "LARG PL": "-", "ESP": 1.2, "MAT_CUSTOM": "INOX 201 ESC"})
    pecas.append({"CÓDIGO": "SAPATA", "DESC": "PE REG. CHAN. 38,1 PEAD", "QTD": 4, "COMP PL": "-", "LARG PL": "-", "ESP": 0, "MAT_CUSTOM": "-"})
    pecas.append({"CÓDIGO": "PARAFUSO", "DESC": "PARAFUSO ALLEN CABEÇA ABAULADA M5X10 - INOX", "QTD": 12 * planos, "COMP PL": "-", "LARG PL": "-", "ESP": 0, "MAT_CUSTOM": "-"})

    if tipo == "LISA":
        c_tampo_f = comp - 3
        l_tampo_f = larg - 2
        c_tampo_pl = c_tampo_f + 104
        l_tampo_pl = l_tampo_f + 129
        pecas.append({"CÓDIGO": "CHAPA", "DESC": "TAMPO ESTANTE LISA", "QTD": planos, "COMP PL": c_tampo_pl, "LARG PL": l_tampo_pl, "ESP": 0.8, "MAT_CUSTOM": material})

        c_ref_pl = larg - 4.4
        l_ref_pl = 134
        qtd_ref_por_plano = max(1, round(comp / 400))
        
        pecas.append({"CÓDIGO": "CHAPA", "DESC": "REFORÇO DO TAMPO (OMEGA)", "QTD": qtd_ref_por_plano * planos, "COMP PL": c_ref_pl, "LARG PL": l_ref_pl, "ESP": 0.8, "MAT_CUSTOM": material})

    elif tipo == "GRADEADA":
        c_viga_pl = comp - 2.4
        pecas.append({"CÓDIGO": "CHAPA", "DESC": "VIGA LATERAL MAIOR PRATELEIRA", "QTD": 2 * planos, "COMP PL": c_viga_pl, "LARG PL": 101.9, "ESP": 0.8, "MAT_CUSTOM": material})

        c_lat_menor_pl = larg - 4.4
        pecas.append({"CÓDIGO": "CHAPA", "DESC": "LATERAL MENOR DA PRATELEIRA", "QTD": 2 * planos, "COMP PL": c_lat_menor_pl, "LARG PL": 152, "ESP": 0.8, "MAT_CUSTOM": material})

        c_div_pl = larg - 4.8
        qtd_div_por_plano = round(comp / 140)  
        pecas.append({"CÓDIGO": "CHAPA", "DESC": "DIVISÓRIAS PRATELEIRA", "QTD": qtd_div_por_plano * planos, "COMP PL": c_div_pl, "LARG PL": 134, "ESP": 0.8, "MAT_CUSTOM": material})

    for p in pecas:
        if "CHAPA" in p["CÓDIGO"] and p["COMP PL"] != "-" and p["LARG PL"] != "-":
            p["PESO UNIT"] = p["COMP PL"] * p["LARG PL"] * p["ESP"] * peso_fator
        else: p["PESO UNIT"] = 0
    return pecas


def calcular_prateleira_parede(comp, larg, material="INOX 304"):
    """
    Motor de Engenharia AçoNobre para Prateleiras de Parede.
    Regra Fixa: Tudo 0.8mm. Quantidade de suportes e reforços calculados por vão livre.
    """
    pecas = []
    peso_fator = 0.000008

    # Lógica de suportes (1 a cada ~700mm)
    qtd_suportes = round(comp / 700) + 1
    if qtd_suportes < 2: qtd_suportes = 2
    
    qtd_esq = 1
    qtd_dir = qtd_suportes - 1

    # ==========================================
    # 1. SUPORTES (MÃOS FRANCESAS)
    # ==========================================
    c_mao = larg + 30.7
    l_mao = 221
    
    pecas.append({"CÓDIGO": "CHAPA", "DESC": "MÃO FRANCESA ESQUERDA", "QTD": qtd_esq, "COMP PL": c_mao, "LARG PL": l_mao, "ESP": 0.8, "MAT_CUSTOM": material})
    pecas.append({"CÓDIGO": "CHAPA", "DESC": "MÃO FRANCESA DIREITA", "QTD": qtd_dir, "COMP PL": c_mao, "LARG PL": l_mao, "ESP": 0.8, "MAT_CUSTOM": material})

    # ==========================================
    # 2. REFORÇO DA PRATELEIRA (Transversal)
    # ==========================================
    # Para cada suporte além das extremidades, vai 1 reforço no meio
    qtd_reforco = max(0, qtd_suportes - 2) 
    
    if qtd_reforco > 0:
        c_ref = larg - 20
        l_ref = 136.6
        pecas.append({"CÓDIGO": "CHAPA", "DESC": "REFORÇO DA PRATELEIRA", "QTD": qtd_reforco, "COMP PL": c_ref, "LARG PL": l_ref, "ESP": 0.8, "MAT_CUSTOM": material})

    # ==========================================
    # 3. TAMPO DA PRATELEIRA
    # ==========================================
    c_prat = comp + 104
    l_prat = larg + 104
    pecas.append({"CÓDIGO": "CHAPA", "DESC": "PRATELEIRA (TAMPO)", "QTD": 1, "COMP PL": c_prat, "LARG PL": l_prat, "ESP": 0.8, "MAT_CUSTOM": material})

    # ==========================================
    # 4. CÁLCULO DE PESO
    # ==========================================
    for p in pecas:
        p["PESO UNIT"] = p["COMP PL"] * p["LARG PL"] * p["ESP"] * peso_fator

    return pecas