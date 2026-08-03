"""
Parâmetros de negócio da análise de reabastecimento.

Tudo que é regra de negócio (e não lógica de cálculo) fica centralizado aqui,
para que uma mudança de política não exija mexer no motor. Os valores atuais
foram confirmados com o usuário em 2026-08-03.
"""

# --- Estoque físico -------------------------------------------------------
# Só estes armazéns contam como estoque disponível para reabastecimento.
# O Empenho (reservado) é considerado JÁ CONSUMIDO -> disponível = Qtde - Empenho.
ARMAZENS_ALVO = {"01", "05", "22"}

# --- Mapa Tipo -> natureza da ação ---------------------------------------
# PA é o ÚNICO tipo produzido; todos os demais são comprados.
TIPO_NATUREZA = {
    "PA": "Produzir",
    "MP": "Comprar",   # Matéria-prima
    "PI": "Comprar",   # Produto intermediário
    "IM": "Comprar",   # Produto importado
    "BN": "Comprar",   # Beneficiamento
    "SV": "Comprar",   # Serviço
    "MO": "Comprar",   # Mão de obra
}

# --- SKUs de revisão manual ----------------------------------------------
# Importados sob contrato anual / embarque bimestral, com validade comercial
# parcialmente consumida antes de entrar no estoque. Quebram a lógica padrão
# de reposição contínua: a ferramenta calcula uma recomendação, mas os marca
# para revisão manual em vez de decidir automaticamente.
# (Registrados no Protheus como Tipo=PA e com Valor unitário 0, por isso não
#  são identificáveis por Tipo nem por valor — fixados por código.)
SKUS_REVISAO_MANUAL = {
    "5973-BR": "Importado (VALO X BRASIL) - contrato anual, embarque bimestral, validade parcial",
    "9121-P2": "Importado (GEMINI EVO BRASIL) - contrato anual, embarque bimestral, validade parcial",
}

# --- Horizonte -----------------------------------------------------------
DIAS_POR_MES = 30            # conversão mês -> dia para cobertura/lead time
MESES_RUN_RATE = 3           # janela para o consumo diário (run-rate de curto prazo)

# --- Cruzamento com a Necessidade nativa do Protheus ---------------------
# Divergência relativa acima disto entre o volume calculado e a Necessidade
# somada do Protheus é sinalizada para revisão (não escondida).
LIMIAR_DIVERGENCIA_NECESSIDADE = 0.30   # 30%

# Divergência absoluta entre o estoque do arquivo físico e o +Estoque[m1]
# do MRP acima disto (em unidades) vira alerta de reconciliação.
LIMIAR_DIVERGENCIA_ESTOQUE = 0.5

# --- Rótulos das linhas de detalhamento do relatório do MRP --------------
# O parser localiza cada série pelo TEXTO do rótulo na coluna A, nunca por
# posição fixa — resiliente a reordenação entre execuções do MRP.
ROTULO_ESTOQUE = "+ Estoque"
ROTULO_ENTRADAS = "+ Entradas"
ROTULO_SAIDAS = "- Saídas"
ROTULO_SAIDA_ESTRUTURA = "- Saída Estrutura"
ROTULO_TRANSF_SAIDA = "- Transf. Saída"
ROTULO_TRANSF_ENTRADA = "- Trasnf. Entrada"  # (grafia do relatório, com erro de digitação)
ROTULO_SALDO_FINAL = "Saldo Final"
ROTULO_NECESSIDADE = "Necessidade"
ROTULO_PRODUTO = "Produto"
ROTULO_PERIODO = "Período"

# Séries mensais que compõem a SAÍDA total do item (consumo).
# Sinal +1 = aumenta a saída; -1 = na verdade é uma entrada (reduz a saída).
COMPONENTES_SAIDA = {
    ROTULO_SAIDAS: +1,
    ROTULO_SAIDA_ESTRUTURA: +1,
    ROTULO_TRANSF_SAIDA: +1,
    ROTULO_TRANSF_ENTRADA: -1,
}
