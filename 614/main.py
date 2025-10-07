# main.py - Versão FINAL, COMPLETA E CORRIGIDA

import logging
import sys
import time
import pyautogui
import xlwings as xw
from pathlib import Path

# --- Imports dos outros módulos do projeto ---
from login import configurar_logging
from cesar_modulo import (
    encontrar_proxima_linha_sem_pa_coluna_b,
    ler_planilha_dados_linha_especifica
)
from insercao_dados import (
    navegar_7_tabs,
    inserir_dados_planilha,
    executar_processo_completo_com_insercao
)

# ✅ IMPORTAR CONFIGURAÇÕES
from config import EMPRESA_PADRAO, TABS_INICIAIS

# ==============================================================================
# SEÇÃO DE FUNÇÕES DE CONTROLO E REINÍCIO
# ==============================================================================

def aguardar_janela_informacao_sair():
    """
    Função de segurança que aguarda a janela 'informacao.PNG' sair da tela,
    pressionando Enter para tentar fechá-la.
    """
    try:
        logging.info("🔍 Verificando se a janela de informação precisa ser fechada...")
        
        if not Path("img/informacao.PNG").exists():
            logging.warning("⚠️ Arquivo img/informacao.PNG não encontrado, pulando verificação.")
            return True
        
        tempo_limite = 10  # Tenta por 10 segundos
        tempo_inicio = time.time()
        
        while time.time() - tempo_inicio < tempo_limite:
            try:
                localizacao = pyautogui.locateOnScreen("img/informacao.PNG", confidence=0.7)
                if localizacao:
                    logging.info("📋 Janela de informação ainda visível. Pressionando Enter para fechar...")
                    pyautogui.press("enter")
                    time.sleep(1)
                else:
                    logging.info("✅ Janela de informação não está na tela.")
                    return True
            except pyautogui.ImageNotFoundException:
                logging.info("✅ Janela de informação não está na tela.")
                return True
        
        logging.warning("⚠️ Timeout aguardando janela de informação sair.")
        return True
        
    except Exception as e:
        logging.error(f"❌ Erro ao aguardar janela sair: {e}")
        return True

def obter_dimensoes_tela():
    """ Obtém as dimensões da tela. """
    try:
        size = pyautogui.size()
        logging.info(f"Dimensões da tela: {size.width}x{size.height}")
        return size.width, size.height
    except Exception as e:
        logging.error(f"Erro ao obter dimensões da tela: {e}")
        return 1920, 1080

def executar_sequencia_reinicio():
    """
    Executa a sequência de reinício SIMPLIFICADA para cadastro do próximo item.
    1. Clica no centro da tela para garantir o foco no programa.
    2. Pressiona Home para reposicionar o cursor para a próxima inserção.
    """
    try:
        logging.info("🔄 === INICIANDO SEQUÊNCIA DE REINÍCIO (SIMPLIFICADA) ===")
        
        # Mantemos esta verificação por segurança, para fechar pop-ups antigos.
        aguardar_janela_informacao_sair()

        largura_tela, altura_tela = obter_dimensoes_tela()
        meio_x, meio_y = largura_tela // 2, altura_tela // 2

        # Passo 1: Mover para o centro e clicar para garantir o foco
        logging.info(f"Passo 1: Movendo mouse para o centro ({meio_x}, {meio_y}) e clicando.")
        pyautogui.moveTo(meio_x, meio_y, duration=0.5)
        pyautogui.click()
        time.sleep(1)

        # Passo 2: Pressionar Home
        logging.info("Passo 2: Pressionando Home.")
        pyautogui.press('home')
        time.sleep(0.5)
        
        logging.info("✅ Sequência de reinício SIMPLIFICADA executada com sucesso!")
        return True

    except Exception as e:
        logging.error(f"❌ Erro na sequência de reinício simplificada: {e}")
        return False

def processar_linha_especifica(linha):
    """ 
    Processa uma linha da planilha (a partir da segunda).
    """
    try:
        logging.info(f"📋 PROCESSANDO LINHA {linha}...")
        
        dados = ler_planilha_dados_linha_especifica(linha)
        if not dados: 
            return False
        
        if dados.get('B'):
            logging.info(f"⏭️ Linha {linha} já tem PA na coluna B: {dados.get('B')} - PULANDO")
            return True

        # Inserir dados da planilha
        if not inserir_dados_planilha(dados): 
            return False
        
        logging.info(f"✅ Linha {linha} processada com sucesso!")
        return True
        
    except Exception as e:
        logging.error(f"❌ Erro ao processar linha {linha}: {e}")
        return False

def executar_cadastro_completo():
    """ 
    Função principal que executa o cadastro completo com lógica de erro corrigida. 
    """
    logging.info("🚀 === INICIANDO CADASTRO COMPLETO A PARTIR DO MAIN ===")
    linhas_processadas_sucesso, linhas_com_erro = 0, 0
    
    while True:
        linha_atual = encontrar_proxima_linha_sem_pa_coluna_b()
        if linha_atual is None:
            logging.info("🎉 === TODAS AS LINHAS FORAM PROCESSADAS! ===")
            break
        
        logging.info(f"\n{'='*80}\n📌 PROCESSANDO LINHA {linha_atual} (Sucesso: {linhas_processadas_sucesso}, Falhas: {linhas_com_erro})\n{'='*80}")
        
        # Lógica para a primeira linha da execução
        if linhas_processadas_sucesso == 0 and linhas_com_erro == 0:
            logging.info("🔑 PRIMEIRA LINHA - Executando processo completo com login...")
            if executar_processo_completo_com_insercao(numero_empresa=EMPRESA_PADRAO):
                linhas_processadas_sucesso += 1
            else:
                linhas_com_erro += 1
        
        # Lógica para as linhas seguintes
        else:
            logging.info(f"🔄 LINHA SEGUINTE ({linha_atual}) - Executando sequência de reinício...")
            # Primeiro executa o reinício, e SÓ SE ele funcionar, processa a linha
            if executar_sequencia_reinicio():
                if processar_linha_especifica(linha_atual):
                    linhas_processadas_sucesso += 1
                else:
                    linhas_com_erro += 1
            else:
                logging.error(f"Falha na sequência de reinício para a linha {linha_atual}. Pulando.")
                linhas_com_erro += 1

        logging.info("⏸️ Pausa ")
        time.sleep(1)
    
    logging.info(f"\n{'='*80}\n🏁 === CADASTRO COMPLETO FINALIZADO ===\n   ✅ Sucesso: {linhas_processadas_sucesso}\n   ❌ Falhas: {linhas_com_erro}\n{'='*80}")
    return linhas_com_erro == 0

# ==============================================================================
# PONTO DE ENTRADA PRINCIPAL DO ROBÔ
# ==============================================================================
if __name__ == "__main__":
    configurar_logging()
    try:
        if executar_cadastro_completo():
            logging.info("🎉 === ROBÔ EXECUTADO COM SUCESSO TOTAL! ===")
        else:
            logging.info("⚠️ === ROBÔ CONCLUÍDO, MAS COM REGISTRO DE FALHAS. Verifique os logs. ===")
        
        logging.info("Sistema finalizado. Pressione Ctrl+C para sair.")
        while True: 
            time.sleep(10)
    except KeyboardInterrupt:
        logging.info("\n🛑 Execução interrompida pelo usuário.")
        sys.exit(0)
    except Exception as e:
        logging.error(f"💥 ERRO INESPERADO E FATAL NO MAIN: {e}", exc_info=True)
        sys.exit(1)