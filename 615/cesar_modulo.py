# cesar_modulo.py
import pyautogui
import time
import logging
import sys
import re
import xlwings as xw
import os
import glob
from pathlib import Path
from login import executar_login, configurar_logging

# ✅ IMPORTAR CONFIGURAÇÕES
from config import ABA_NUMERO, ABA_TEXTO_BUSCA, COLUNAS_PARA_PREENCHER, EMPRESA_PADRAO

# Configuração do pyautogui
pyautogui.PAUSE = 0.5
pyautogui.FAILSAFE = True

def encontrar_planilha_mais_recente():
    """
    Encontra a planilha mais recente na pasta
    
    Returns:
        str ou None: Caminho da planilha mais recente ou None se não encontrar
    """
    try:
        pasta_base = r"\\192.168.1.250\Programas\robo custos\lancar"
        logging.info(f"Procurando planilha mais recente em: {pasta_base}")
        
        if not os.path.exists(pasta_base):
            logging.error(f"Pasta não existe: {pasta_base}")
            return None
        
        # Padrões de arquivos Excel e CSV
        padroes = ['*.xlsx', '*.xls', '*.csv']
        arquivos_encontrados = []
        
        # Buscar arquivos com os padrões
        for padrao in padroes:
            caminho_padrao = os.path.join(pasta_base, padrao)
            arquivos = glob.glob(caminho_padrao)
            arquivos_encontrados.extend(arquivos)
            logging.info(f"Padrão {padrao}: {len(arquivos)} arquivos encontrados")
        
        if not arquivos_encontrados:
            logging.error("Nenhum arquivo Excel/CSV encontrado na pasta")
            return None
        
        logging.info(f"Total de arquivos encontrados: {len(arquivos_encontrados)}")
        
        # Encontrar o arquivo mais recente por data de modificação
        arquivo_mais_recente = None
        data_mais_recente = 0
        
        for arquivo in arquivos_encontrados:
            try:
                stat = os.stat(arquivo)
                data_modificacao = stat.st_mtime
                nome_arquivo = os.path.basename(arquivo)
                
                logging.info(f"Arquivo: {nome_arquivo}")
                logging.info(f"  Data modificação: {time.ctime(data_modificacao)}")
                logging.info(f"  Tamanho: {stat.st_size} bytes")
                
                if data_modificacao > data_mais_recente:
                    data_mais_recente = data_modificacao
                    arquivo_mais_recente = arquivo
                    
            except Exception as e:
                logging.warning(f"Erro ao verificar arquivo {arquivo}: {e}")
        
        if arquivo_mais_recente:
            logging.info(f"Arquivo mais recente selecionado: {os.path.basename(arquivo_mais_recente)}")
            logging.info(f"  Caminho: {arquivo_mais_recente}")
            logging.info(f"  Data: {time.ctime(data_mais_recente)}")
            return arquivo_mais_recente
        else:
            logging.error("Não foi possível determinar arquivo mais recente")
            return None
            
    except Exception as e:
        logging.error(f"Erro ao procurar planilha mais recente: {e}")
        return None

def extrair_numero_do_nome_aba(nome_aba):
    """
    Extrai o número entre parênteses do nome da aba
    Ex: "Porta (614)" retorna "614"
    
    Args:
        nome_aba: Nome da aba para extrair número
        
    Returns:
        str: Número extraído entre parênteses ou None se não encontrar
    """
    try:
        if not nome_aba:
            logging.warning("Nome da aba vazio")
            return None
            
        # Procurar por números entre parênteses
        import re
        padrao = r'\((\d+)\)'  # Busca números entre parênteses
        match = re.search(padrao, str(nome_aba))
        
        if match:
            numero = match.group(1)  # Pega o número dentro dos parênteses
            logging.info(f"Número extraído do nome da aba '{nome_aba}': {numero}")
            return numero
        else:
            logging.warning(f"Nenhum número entre parênteses encontrado no nome da aba: {nome_aba}")
            return None
            
    except Exception as e:
        logging.error(f"Erro ao extrair número do nome da aba: {e}")
        return None

def encontrar_aba_configurada(wb):
    """
    Encontra a aba configurada no config.py
    
    Args:
        wb: Workbook do xlwings
        
    Returns:
        worksheet ou None: Aba encontrada ou None se não encontrar
    """
    try:
        logging.info(f"Procurando aba com {ABA_TEXTO_BUSCA} no nome...")
        
        for sheet in wb.sheets:
            nome_sheet = sheet.name
            logging.info(f"Verificando aba: {nome_sheet}")
            
            # Verificar se contém o texto configurado
            if ABA_TEXTO_BUSCA in str(nome_sheet):
                logging.info(f"✅ Aba configurada encontrada: {nome_sheet}")
                return sheet
        
        logging.error(f"❌ Nenhuma aba com {ABA_TEXTO_BUSCA} encontrada!")
        return None
        
    except Exception as e:
        logging.error(f"Erro ao procurar aba configurada: {e}")
        return None

def encontrar_proxima_linha_sem_pa_coluna_b():
    """
    Encontra a próxima linha que não tem PA na coluna B
    BUSCA NA ABA CONFIGURADA
    
    Returns:
        int ou None: Número da linha sem PA ou None se não encontrar
    """
    app = None
    wb = None
    try:
        # Encontrar a planilha mais recente
        caminho_planilha = encontrar_planilha_mais_recente()
        if not caminho_planilha:
            logging.error("Não foi possível encontrar planilha")
            return None
        
        logging.info(f"Procurando próxima linha sem PA na coluna B (aba {ABA_TEXTO_BUSCA})...")
        
        # Abrir planilha
        app = xw.App(visible=False)
        wb = app.books.open(caminho_planilha)
        
        # BUSCAR A ABA CONFIGURADA
        ws = encontrar_aba_configurada(wb)
        if not ws:
            logging.error(f"Não foi possível encontrar aba configurada {ABA_TEXTO_BUSCA}")
            wb.close()
            app.quit()
            return None
        
        logging.info(f"Usando aba: {ws.name}")
        
        # Começar da linha 2 e procurar até linha 1000
        for linha in range(2, 1001):
            valor_a = ws.range(f'A{linha}').value
            valor_b = ws.range(f'B{linha}').value
            
            # Se linha A está vazia, chegamos ao fim
            if not valor_a:
                logging.info(f"Fim dos dados na linha {linha}")
                break
                
            # Se coluna B está vazia, esta linha precisa ser processada
            if not valor_b:
                logging.info(f"Próxima linha para processar (sem PA na coluna B): {linha}")
                wb.close()
                app.quit()
                return linha
        
        logging.info("Todas as linhas já foram processadas!")
        wb.close()
        app.quit()
        return None
        
    except Exception as e:
        logging.error(f"Erro ao procurar próxima linha sem PA na coluna B: {e}")
        
        # Limpar recursos
        if wb:
            try:
                wb.close()
            except:
                pass
        if app:
            try:
                app.quit()
            except:
                pass
        
        return None

def ler_planilha_dados_linha_especifica(linha_especifica):
    """
    Lê os dados de uma linha específica da planilha mais recente usando xlwings (COM)
    BUSCA A ABA CONFIGURADA e lê as COLUNAS CONFIGURADAS
    
    Args:
        linha_especifica: Número da linha para ler
    
    Returns:
        dict ou None: Dados da planilha ou None se erro
    """
    app = None
    wb = None
    try:
        # Encontrar a planilha mais recente
        caminho_planilha = encontrar_planilha_mais_recente()
        
        if not caminho_planilha:
            logging.error("Não foi possível encontrar planilha")
            return None
        
        logging.info(f"Abrindo planilha via COM: {caminho_planilha}")
        
        # Criar instância do Excel
        app = xw.App(visible=False)
        wb = app.books.open(caminho_planilha)
        
        logging.info(f"Planilha aberta com sucesso!")
        
        # BUSCAR A ABA CONFIGURADA
        ws = encontrar_aba_configurada(wb)
        if not ws:
            logging.error(f"Não foi possível encontrar aba configurada {ABA_TEXTO_BUSCA}")
            wb.close()
            app.quit()
            return None
        
        nome_aba = ws.name
        logging.info(f"Usando aba: {nome_aba}, linha: {linha_especifica}")
        
        # Extrair número do nome da aba
        numero_tipo = extrair_numero_do_nome_aba(nome_aba)
        
        # Verificar se realmente extraiu o número esperado
        if numero_tipo != ABA_NUMERO:
            logging.warning(f"⚠️ Número extraído ({numero_tipo}) não é {ABA_NUMERO}!")
        
        # Dados básicos
        dados = {
            'numero_tipo': numero_tipo,
            'A': ws.range(f'A{linha_especifica}').value,
            'B': ws.range(f'B{linha_especifica}').value,
            'linha_usada': linha_especifica,
            'nome_aba': nome_aba,
            'arquivo': os.path.basename(caminho_planilha)
        }
        
        # Ler as colunas configuradas
        for coluna in COLUNAS_PARA_PREENCHER:
            valor = ws.range(f'{coluna}{linha_especifica}').value
            dados[coluna] = valor
        
        logging.info(f"Dados lidos da aba configurada - linha {linha_especifica}:")
        logging.info(f"  Arquivo: {os.path.basename(caminho_planilha)}")
        logging.info(f"  Nome da aba: {nome_aba}")
        logging.info(f"  Número extraído: {numero_tipo} ({'✅ CORRETO' if numero_tipo == ABA_NUMERO else '❌ INCORRETO'})")
        logging.info(f"  Coluna A (controle): {dados['A']}")
        logging.info(f"  Coluna B (PA existente): {dados['B']}")
        
        logging.info(f"  📋 COLUNAS CONFIGURADAS ({COLUNAS_PARA_PREENCHER[0]} até {COLUNAS_PARA_PREENCHER[-1]}):")
        for coluna in COLUNAS_PARA_PREENCHER:
            logging.info(f"    Coluna {coluna}: {dados[coluna]}")
        
        # Fechar planilha e Excel
        wb.close()
        app.quit()
        
        return dados
        
    except Exception as e:
        logging.error(f"Erro ao ler planilha via COM: {e}")
        
        # Limpar recursos em caso de erro
        if wb:
            try:
                wb.close()
            except:
                pass
        if app:
            try:
                app.quit()
            except:
                pass
        
        return None

def ler_planilha_dados():
    """
    Lê os dados da planilha mais recente usando xlwings (COM) - primeira linha sem PA na coluna B
    
    Returns:
        dict ou None: Dados da planilha ou None se erro
    """
    try:
        # Encontrar próxima linha sem PA na coluna B
        linha_sem_pa = encontrar_proxima_linha_sem_pa_coluna_b()
        if linha_sem_pa is None:
            logging.info("Nenhuma linha sem PA encontrada na coluna B")
            return None
        
        # Ler dados da linha específica
        return ler_planilha_dados_linha_especifica(linha_sem_pa)
        
    except Exception as e:
        logging.error(f"Erro ao ler planilha: {e}")
        return None

def encontrar_e_clicar_imagem(caminho_imagem, timeout=30, confidence=0.8):
    """
    Encontra uma imagem na tela e clica nela
    
    Args:
        caminho_imagem: Caminho para o arquivo de imagem
        timeout: Tempo limite para encontrar a imagem
        confidence: Nível de confiança para correspondência (0.0 a 1.0)
    
    Returns:
        bool: True se encontrou e clicou, False caso contrário
    """
    inicio = time.time()
    caminho_completo = Path(caminho_imagem)
    
    if not caminho_completo.exists():
        logging.error(f"Arquivo de imagem não encontrado: {caminho_imagem}")
        return False
    
    logging.info(f"Procurando imagem: {caminho_imagem}")
    
    while time.time() - inicio < timeout:
        try:
            # Procurar a imagem na tela
            localizacao = pyautogui.locateOnScreen(str(caminho_completo), confidence=confidence)
            
            if localizacao:
                # Calcular o centro da imagem encontrada
                centro = pyautogui.center(localizacao)
                logging.info(f"Imagem encontrada em: {centro}")
                
                # Clicar no centro da imagem
                pyautogui.click(centro)
                logging.info(f"Clique realizado em: {centro}")
                return True
                
        except pyautogui.ImageNotFoundException:
            pass
        except Exception as e:
            logging.warning(f"Erro ao procurar imagem {caminho_imagem}: {e}")
        
        time.sleep(1)
    
    logging.error(f"Imagem não encontrada no tempo limite: {caminho_imagem}")
    return False

def digitar_configurador(texto):
    """
    Digita um texto de forma normal e rápida
    
    Args:
        texto: Texto a ser digitado
    """
    logging.info(f"Digitando texto: {texto}")
    pyautogui.write(texto)

def acessar_configurador():
    """
    Acessa o módulo configurador após o login
    
    Returns:
        bool: True se conseguiu acessar, False caso contrário
    """
    try:
        logging.info("Iniciando acesso ao configurador...")
        
        # Aguardar um tempo para o sistema carregar após login
        time.sleep(2)
        
        # Passo 1: Clicar no menu systeam
        logging.info("Passo 1: Procurando e clicando no menu systeam...")
        if not encontrar_e_clicar_imagem("img/menu systeam.PNG", timeout=60):
            logging.error("Não foi possível encontrar o menu systeam")
            return False
        
        time.sleep(0.5)
        
        # Passo 2: Digitar "configurador"
        logging.info("Passo 2: Digitando 'configurador'...")
        digitar_configurador("configurador")
        time.sleep(0.5)
        
        # Passo 3: Clicar no módulo "item por descricao"
        logging.info("Passo 3: Procurando e clicando no módulo...")
        if not encontrar_e_clicar_imagem("img/modulo item por descricao.PNG", timeout=30):
            logging.error("Não foi possível encontrar o módulo 'item por descricao'")
            return False
        
        logging.info("Configurador acessado com sucesso!")
        time.sleep(1)
        
        return True
        
    except Exception as e:
        logging.error(f"Erro ao acessar configurador: {e}")
        return False

def pesquisar_tipo_da_planilha(numero_extraido):
    """
    Pesquisa o tipo extraído da planilha no módulo
    
    Args:
        numero_extraido: Número extraído do nome da aba da planilha
    
    Returns:
        bool: True se conseguiu pesquisar, False caso contrário
    """
    try:
        logging.info(f"Iniciando pesquisa do tipo: {numero_extraido}")
        
        # Aguardar módulo carregar
        time.sleep(2)
        
        # Passo 1: Clicar no botão de pesquisar tipos
        logging.info("Passo 1: Procurando e clicando no botão pesquisar tipos...")
        if not encontrar_e_clicar_imagem("img/pesquisar tipos 615.PNG", timeout=30):
            logging.error("Não foi possível encontrar o botão pesquisar tipos")
            return False
        
        time.sleep(0.5)
        
        # Passo 2: Digitar o número extraído
        logging.info(f"Passo 2: Digitando '{numero_extraido}'...")
        digitar_configurador(numero_extraido)
        time.sleep(0.5)
        pyautogui.press("enter")
        
        logging.info(f"Pesquisa do tipo {numero_extraido} realizada com sucesso!")
        return True
        
    except Exception as e:
        logging.error(f"Erro ao pesquisar tipo {numero_extraido}: {e}")
        return False

def executar_processo_completo(numero_empresa=None, linha_especifica=None):
    """
    Executa o processo completo: login + acesso ao configurador + pesquisa tipo da planilha
    
    Args:
        numero_empresa: Número da empresa para login
        linha_especifica: Linha específica para processar (None = próxima linha vazia)
    
    Returns:
        tuple: (bool sucesso, str numero_extraido, dict dados_planilha)
    """
    try:
        if numero_empresa is None:
            numero_empresa = EMPRESA_PADRAO
            
        logging.info("Iniciando processo completo...")
        
        # Passo 1: Ler dados da planilha
        if linha_especifica:
            logging.info(f"Lendo dados da linha específica: {linha_especifica}")
            dados = ler_planilha_dados_linha_especifica(linha_especifica)
        else:
            logging.info("Lendo dados da próxima linha sem PA na coluna B...")
            dados = ler_planilha_dados()
            
        if dados is None:
            logging.error("Falha ao ler planilha ou nenhuma linha para processar")
            return False, None, None
        
        # Verificar se já tem PA na coluna B
        if dados.get('B'):
            logging.warning(f"Linha {dados.get('linha_usada')} já tem PA na coluna B: {dados.get('B')}")
            return False, None, dados
        
        # Passo 2: Pegar o número já extraído do nome da aba
        numero_extraido = dados.get('numero_tipo')
        if not numero_extraido:
            logging.error("Falha ao extrair número do nome da aba")
            return False, None, dados
        
        # Passo 3: Realizar login
        logging.info("Executando login...")
        if not executar_login(numero_empresa=numero_empresa):
            logging.error("Falha no login")
            return False, numero_extraido, dados
        
        # Passo 4: Acessar configurador
        logging.info("Acessando configurador...")
        if not acessar_configurador():
            logging.error("Falha ao acessar configurador")
            return False, numero_extraido, dados
        
        # Passo 5: Pesquisar tipo da planilha
        logging.info(f"Pesquisando tipo {numero_extraido}...")
        if not pesquisar_tipo_da_planilha(numero_extraido):
            logging.error(f"Falha ao pesquisar tipo {numero_extraido}")
            return False, numero_extraido, dados
        
        logging.info("Processo completo executado com sucesso!")
        return True, numero_extraido, dados
        
    except Exception as e:
        logging.error(f"Erro no processo completo: {e}")
        return False, None, None

def main():
    """Função principal do módulo César"""
    configurar_logging()
    
    try:
        logging.info("=== INICIANDO MÓDULO CÉSAR ===")
        
        # Verificar se as imagens existem
        imagens_necessarias = [
            "img/menu systeam.PNG",
            "img/modulo item por descricao.PNG",
            "img/pesquisar tipos 615.PNG"
        ]
        
        for imagem in imagens_necessarias:
            if not Path(imagem).exists():
                logging.error(f"Imagem necessária não encontrada: {imagem}")
                logging.error("Certifique-se de que todas as imagens estão na pasta 'img'")
                return False
        
        # Executar processo completo
        sucesso, numero, dados = executar_processo_completo(numero_empresa=EMPRESA_PADRAO)
        
        if sucesso:
            logging.info("=== MÓDULO CÉSAR EXECUTADO COM SUCESSO ===")
            logging.info(f"Arquivo usado: {dados.get('arquivo', 'N/A')}")
            logging.info(f"Linha processada: {dados.get('linha_usada', 'N/A')}")
            logging.info(f"Número pesquisado: {numero}")
            
            # Manter o programa rodando para permitir interação manual
            logging.info("Sistema pronto para uso. Pressione Ctrl+C para sair.")
            try:
                while True:
                    time.sleep(10)
            except KeyboardInterrupt:
                logging.info("Finalizando módulo César...")
        else:
            logging.error("=== FALHA NA EXECUÇÃO DO MÓDULO CÉSAR ===")
            return False
            
    except KeyboardInterrupt:
        logging.info("Execução interrompida pelo usuário")
        return True
    except Exception as e:
        logging.error(f"Erro inesperado no módulo César: {e}")
        return False

if __name__ == "__main__":
    sucesso = main()
    sys.exit(0 if sucesso else 1)