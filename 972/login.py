import subprocess
import pyautogui
import time
import logging
import os
import pandas as pd
import sys
from pathlib import Path

# Configuração do pyautogui para maior segurança
pyautogui.PAUSE = 0.5
pyautogui.FAILSAFE = True

def configurar_logging():
    """Configura o sistema de logging"""
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(
                log_dir / f"login_{time.strftime('%Y%m%d_%H%M%S')}.log", 
                encoding='utf-8'  # ADICIONAR ESTA LINHA
            ),
            logging.StreamHandler(sys.stdout)
        ]
    )

def verificar_janela_programa(titulo_janela="", timeout=30):
    """
    Verifica se a janela do programa está aberta
    
    Args:
        titulo_janela: Título ou parte do título da janela
        timeout: Tempo limite em segundos
    
    Returns:
        bool: True se encontrar a janela, False caso contrário
    """
    inicio = time.time()
    while time.time() - inicio < timeout:
        try:
            janelas = pyautogui.getWindowsWithTitle(titulo_janela)
            if janelas:
                janelas[0].activate()  # Ativa a janela
                return True
        except Exception as e:
            logging.warning(f"Erro ao verificar janela: {e}")
        time.sleep(1)
    return False

def executar_login(numero_empresa=None, usuario="rudieri", senha="rohden", 
                  caminho_programa=r"\\192.168.1.250\Programas\SYSTEAM\PROGRAMA\Indust\Engenh.exe"):
    """
    Executa o programa e realiza o login
    
    Args:
        numero_empresa: Número da empresa para login (usa 200 como padrão se None)
        usuario: Nome de usuário para login
        senha: Senha para login
        caminho_programa: Caminho para o executável
    
    Returns:
        bool: True se login foi bem-sucedido, False caso contrário
    """
    try:
        logging.info("Iniciando execução e login no programa")
        
        # Verificar se o arquivo existe
        if not os.path.exists(caminho_programa):
            logging.error(f"Programa não encontrado no caminho: {caminho_programa}")
            return False
        
        # Iniciar o programa
        logging.info(f"Iniciando programa: {caminho_programa}")
        processo = subprocess.Popen(caminho_programa)
        logging.info(f"Programa iniciado com PID: {processo.pid}")
        
        # Aguardar abertura da janela
        logging.info("Aguardando abertura da janela do programa...")
        if not verificar_janela_programa("", timeout=60):
            logging.error("Janela do programa não foi encontrada no tempo limite")
            return False
        
        time.sleep(3)  # Tempo adicional para garantir que a janela está pronta
        
        # Realizar login
        logging.info("Iniciando processo de login...")
        
        # Campo usuário
        logging.info("Inserindo nome de usuário")
        pyautogui.write(usuario)
        time.sleep(1)
        pyautogui.press("tab")
        time.sleep(1)
        
        # Campo senha
        logging.info("Inserindo senha")
        pyautogui.write(senha)
        time.sleep(1)
        pyautogui.press("tab")
        time.sleep(1)
        
        # Campo empresa - ALTERADO PARA 200 COMO PADRÃO
        if numero_empresa is None or (isinstance(numero_empresa, float) and pd.isna(numero_empresa)):
            numero_empresa = 200  # MUDANÇA AQUI: era 100, agora é 200
            logging.info("Usando número de empresa padrão: 200")
        else:
            logging.info(f"Usando número de empresa: {numero_empresa}")
        
        pyautogui.write(str(int(numero_empresa)))
        time.sleep(1)
        
        # Navegar pelos campos e confirmar
        pyautogui.press("tab", presses=3, interval=0.3)
        time.sleep(1)
        pyautogui.press("enter")
        
        logging.info(f"Login realizado com sucesso - Empresa: {numero_empresa}")
        time.sleep(3)  # Aguardar processamento do login
        
        return True
        
    except Exception as e:
        logging.error(f"Erro durante execução do login: {e}")
        return False

def main():
    """Função principal para execução standalone"""
    configurar_logging()
    
    try:
        # Usando empresa 200
        sucesso = executar_login(numero_empresa=200)
        
        if sucesso:
            logging.info("Login executado com sucesso!")
        else:
            logging.error("Falha na execução do login")
            sys.exit(1)
            
    except KeyboardInterrupt:
        logging.info("Execução interrompida pelo usuário")
        sys.exit(0)
    except Exception as e:
        logging.error(f"Erro inesperado: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()