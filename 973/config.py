# config.py - Configurações do robô

# ======================================
# CONFIGURAÇÕES DA ABA
# ======================================
ABA_NUMERO = "973"  
ABA_TEXTO_BUSCA = f"({ABA_NUMERO})"  # Texto que deve estar no nome da aba

# ======================================
# CONFIGURAÇÕES DAS COLUNAS
# ======================================
COLUNA_INICIAL = "C"  # Primeira coluna a ser preenchida
COLUNA_FINAL = "U"    # Última coluna a ser preenchida

# ======================================
# CONFIGURAÇÕES DE COLUNAS PARA PULAR
# ======================================
COLUNAS_PULAR = ["C", "G" ,"H" , "J" , "N" , "O" , "Q" , "R"]  # Colunas que o robô deve pular (só dar tab, não inserir dados)
TEMPO_TAB_PULAR = 0.1

# ======================================
# CONFIGURAÇÕES DE DIGITAÇÃO - MAIS RÁPIDO
# ======================================
TEMPO_ENTRE_CARACTERES = 0.03  # ALTERADO: era 0.1, agora 0.03 (3x mais rápido)

# ======================================
# OUTRAS CONFIGURAÇÕES
# ======================================
EMPRESA_PADRAO = 200  # Número da empresa para login
INTERVALO_ENTRE_ACOES = 0.7  # ALTERADO: era 0.7, agora 0.5 (mais rápido)
TABS_INICIAIS = 7  # Quantidade de tabs antes de começar a inserir dados

# ======================================
# FUNÇÃO AUXILIAR PARA GERAR LISTA DE COLUNAS
# ======================================
def gerar_lista_colunas(inicio, fim):
    """
    Gera lista de colunas de A até Z, AA até ZZ etc.
    
    Args:
        inicio: Coluna inicial (ex: "C")
        fim: Coluna final (ex: "W")
    
    Returns:
        list: Lista de colunas entre início e fim
    """
    def letra_para_numero(letra):
        """Converte letra da coluna para número"""
        resultado = 0
        for char in letra:
            resultado = resultado * 26 + (ord(char) - ord('A') + 1)
        return resultado
    
    def numero_para_letra(numero):
        """Converte número para letra da coluna"""
        resultado = ""
        while numero > 0:
            numero -= 1
            resultado = chr(numero % 26 + ord('A')) + resultado
            numero //= 26
        return resultado
    
    inicio_num = letra_para_numero(inicio)
    fim_num = letra_para_numero(fim)
    
    return [numero_para_letra(i) for i in range(inicio_num, fim_num + 1)]

# Gerar lista automática de colunas baseada na configuração
COLUNAS_PARA_PREENCHER = gerar_lista_colunas(COLUNA_INICIAL, COLUNA_FINAL)

# Debug das configurações
if __name__ == "__main__":
    print("=== CONFIGURAÇÕES DO ROBÔ ===")
    print(f"Aba a ser utilizada: {ABA_TEXTO_BUSCA}")
    print(f"Colunas: {COLUNA_INICIAL} até {COLUNA_FINAL}")
    print(f"Lista de colunas: {COLUNAS_PARA_PREENCHER}")
    print(f"Total de colunas: {len(COLUNAS_PARA_PREENCHER)}")
    print(f"Colunas para pular: {COLUNAS_PULAR}")
    print(f"Tempo entre caracteres: {TEMPO_ENTRE_CARACTERES}s")