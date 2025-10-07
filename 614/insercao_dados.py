# insercao_dados.py
import pyautogui
import time
import logging
import sys
import re
import xlwings as xw
from pathlib import Path
from cesar_modulo import executar_processo_completo, configurar_logging, encontrar_planilha_mais_recente
import os 
from PIL import Image

# IMPORTAR CONFIGURAÇÕES (INCLUINDO TEMPO_TAB_PULAR)
from config import ABA_NUMERO, ABA_TEXTO_BUSCA, COLUNAS_PARA_PREENCHER, INTERVALO_ENTRE_ACOES, EMPRESA_PADRAO, TABS_INICIAIS, TEMPO_ENTRE_CARACTERES, COLUNAS_PULAR, TEMPO_TAB_PULAR

# Configuração do pyautogui
pyautogui.PAUSE = 0.1
pyautogui.FAILSAFE = True

def digitar_pausadamente(texto):
    """
    Função que digita texto pausadamente, controlada pelo config.py
    
    Args:
        texto: Texto a ser digitado
    """
    if not texto:
        return
    
    texto_str = str(texto)
    logging.info(f"Digitando pausadamente: '{texto_str}' (intervalo: {TEMPO_ENTRE_CARACTERES}s)")
    
    for caracter in texto_str:
        pyautogui.write(caracter)
        time.sleep(TEMPO_ENTRE_CARACTERES)

def deve_pular_coluna(coluna):
    """
    Verifica se a coluna deve ser pulada (apenas dar tab, não inserir dados)
    
    Args:
        coluna: Letra da coluna (ex: "C", "D")
        
    Returns:
        bool: True se deve pular, False se deve inserir dados
    """
    return coluna in COLUNAS_PULAR

def pular_coluna(coluna):
    """
    Pula uma coluna específica (apenas pressiona tab sem inserir dados) - RÁPIDO
    
    Args:
        coluna: Letra da coluna que está sendo pulada
    """
    logging.info(f"Pulando coluna {coluna} (configurada para pular)")
    pyautogui.press("tab")
    time.sleep(TEMPO_TAB_PULAR)  # USA O TEMPO RÁPIDO PARA PULAR

def formatar_valor(valor):
    """
    Formatar valor removendo decimais desnecessários
    
    Args:
        valor: Valor a ser formatado
        
    Returns:
        str: Valor formatado como string
    """
    # Verificar se o valor é 0
    if valor == 0:
        return "0"
    
    if not valor:
        return ""
    
    # Se for float e for um número inteiro (ex: 570.0), converter para int
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    
    return str(valor)

def navegar_7_tabs():
    """
    Navega pressionando Tab conforme configurado
    
    Returns:
        bool: True se navegou com sucesso, False caso contrário
    """
    try:
        logging.info(f"Navegando com {TABS_INICIAIS}x Tab...")
        
        # Aguardar um tempo para garantir que a tela está carregada
        time.sleep(1)
        
        # Pressionar Tab conforme configurado
        for i in range(TABS_INICIAIS):
            pyautogui.press("tab")
            logging.info(f"Tab {i+1}/{TABS_INICIAIS} pressionado")
            time.sleep(0.3)
        
        logging.info(f"Navegação com {TABS_INICIAIS}x Tab concluída!")
        time.sleep(1)
        
        return True
        
    except Exception as e:
        logging.error(f"Erro ao navegar com {TABS_INICIAIS}x Tab: {e}")
        return False

def inserir_dados_planilha_completa(dados):
    """
    Insere os dados nas colunas configuradas sequencialmente COM DIGITAÇÃO PAUSADA
    E com funcionalidade de PULAR COLUNAS ESPECÍFICAS
    
    Args:
        dados: Dicionário com os dados da planilha
    
    Returns:
        bool: True se inseriu com sucesso, False caso contrário
    """
    try:
        logging.info("=== INSERÇÃO SEQUENCIAL CONFIGURADA (DIGITAÇÃO PAUSADA + PULAR COLUNAS) ===")
        logging.info(f"Preenchendo colunas: {COLUNAS_PARA_PREENCHER[0]} até {COLUNAS_PARA_PREENCHER[-1]}")
        logging.info(f"Total de colunas: {len(COLUNAS_PARA_PREENCHER)}")
        logging.info(f"Colunas para pular: {COLUNAS_PULAR}")
        logging.info(f"Tempo entre caracteres: {TEMPO_ENTRE_CARACTERES}s")
        logging.info(f"Tempo para pular colunas: {TEMPO_TAB_PULAR}s")
        
        # Aguardar estar no configurador
        time.sleep(1)
        
        # Inserir cada coluna sequencialmente
        for i, coluna in enumerate(COLUNAS_PARA_PREENCHER):
            
            # VERIFICAR SE DEVE PULAR A COLUNA
            if deve_pular_coluna(coluna):
                pular_coluna(coluna)
            else:
                # Inserir dados normalmente
                valor = formatar_valor(dados.get(coluna, ''))
                logging.info(f"Digitando campo {coluna}: '{valor}'")
                
                # Inserir valor pausadamente se não estiver vazio
                if valor:
                    digitar_pausadamente(valor)  # USA A FUNÇÃO DE DIGITAÇÃO PAUSADA
                
                time.sleep(INTERVALO_ENTRE_ACOES)
                
                # Se não for a última coluna, pressiona tab
                if i < len(COLUNAS_PARA_PREENCHER) - 1:
                    pyautogui.press("tab")
                    time.sleep(INTERVALO_ENTRE_ACOES)
        
        # Pressionar F2 para finalizar
        logging.info("Pressionando F2 para finalizar...")
        pyautogui.press("f2")
        
        # Aguardar um tempo após F2
        logging.info("Aguardando após F2...")
        time.sleep(3)
        
        logging.info("Inserção configurada concluída!")
        return True
        
    except Exception as e:
        logging.error(f"Erro ao inserir dados configurados: {e}")
        return False

def extrair_codigo_pa_do_texto(texto):
    """
    Extrai o código PA do texto, corrigindo erros comuns de OCR (S->5, B->8, T->7)
    e garantindo que o resultado final contém apenas números após 'PA'.
    """
    try:
        if not texto or not texto.strip():
            logging.warning("Texto vazio para extrair PA")
            return None
        
        logging.info(f"Texto original do OCR: {repr(texto)}")
        
        texto_limpo = texto.replace('\n', ' ').replace('\r', ' ')
        
        # --- NOVA REGRA DE CORREÇÃO DE OCR ---
        # Substitui letras que são frequentemente confundidas com números.
        # Adicionei mais algumas correções comuns como O->0 e I->1.
        correcoes = {'S': '5', 'B': '8', 'T': '7', 'O': '0', 'I': '1'}
        texto_corrigido = texto_limpo.upper() # Converte para maiúsculas para pegar 's', 'b', etc.
        for letra, numero in correcoes.items():
            texto_corrigido = texto_corrigido.replace(letra, numero)
        
        logging.info(f"Texto após correções automáticas: {repr(texto_corrigido)}")

        # --- NOVA REGRA DE REGEX: APENAS NÚMEROS ---
        # Voltamos a usar \d para garantir que apenas dígitos são capturados.
        padroes = [
            r'ITEM\s*[\'"]?PA(\d{4,8})[\'"]?',
            r'CADASTRAD[AO]\s+PARA\s+O\s+ITEM\s*[\'"]?PA(\d{4,8})[\'"]?',
            r'PA(\d{4,8})',
            # O padrão BA é mantido para caso o OCR erre o 'P' de 'PA'
            r'BA(\d{4,8})', 
        ]
        
        for i, padrao in enumerate(padroes):
            logging.info(f"Testando padrão numérico {i+1}: {padrao}")
            # Usamos o texto corrigido e em maiúsculas para a busca
            matches = re.findall(padrao, texto_corrigido)
            
            if matches:
                numero_extraido = max(matches, key=len)
                codigo_pa = f"PA{numero_extraido}"
                
                logging.info(f"SUCESSO! PA numérico extraído com padrão {i+1}: {codigo_pa}")
                return codigo_pa
        
        logging.warning("NENHUM código PA numérico encontrado no texto, mesmo após correções.")
        logging.warning(f"Texto original analisado: {texto_limpo[:200]}...")
        return None
        
    except Exception as e:
        logging.error(f"ERRO ao extrair código PA: {e}")
        return None

def extrair_pa_da_janela_informacao(localizacao):
    """
    Extrai o PA da janela com PRÉ-PROCESSAMENTO DE IMAGEM para melhorar o OCR.
    """
    try:
        logging.info(f"Extraindo PA da janela na posição: {localizacao}")
        
        margem_esquerda, margem_direita, margem_cima, margem_baixo = 50, 400, 50, 200
        left = max(0, int(localizacao.left) - margem_esquerda)
        top = max(0, int(localizacao.top) - margem_cima)
        width = int(localizacao.width) + margem_esquerda + margem_direita
        height = int(localizacao.height) + margem_cima + margem_baixo
        
        logging.info(f"Capturando região da tela para OCR...")
        screenshot_janela = pyautogui.screenshot(region=(left, top, width, height))

        # --- NOVO BLOCO DE PRÉ-PROCESSAMENTO DE IMAGEM ---
        logging.info("Iniciando pré-processamento da imagem para melhorar a leitura...")
        
        # Converte a captura de tela para um objeto de imagem do Pillow
        img = Image.frombytes('RGB', screenshot_janela.size, screenshot_janela.tobytes())
        
        # 1. Converte para escala de cinza
        img = img.convert('L')
        
        # 2. Redimensiona a imagem (upscaling) - Tesseract gosta de imagens maiores
        largura, altura = img.size
        img = img.resize((largura * 3, altura * 3), Image.LANCZOS)
        
        # 3. Binarização (Thresholding) - Converte para preto e branco puro
        # Pixels mais escuros que um limiar (180) ficam pretos (0), mais claros ficam brancos (255)
        # Este valor (180) pode ser ajustado se necessário, mas é um bom ponto de partida.
        threshold = 180 
        img = img.point(lambda p: 255 if p > threshold else 0)
        
        # Salva a imagem processada para vermos exatamente o que o OCR está a tentar ler.
        try:
            nome_debug_ocr = f"DEBUG_imagem_para_ocr_{int(time.time())}.png"
            img.save(nome_debug_ocr)
            logging.info(f"Imagem pré-processada para OCR salva como: {nome_debug_ocr}")
        except Exception as e_save:
            logging.warning(f"Não foi possível salvar a imagem de debug do OCR: {e_save}")
        
        # --- FIM DO BLOCO DE PRÉ-PROCESSAMENTO ---

        try:
            import pytesseract
            # Envia a IMAGEM TRATADA para o Pytesseract, não a original
            texto_janela = pytesseract.image_to_string(img, lang='por', config=r'--oem 3 --psm 6')
            logging.info(f"TEXTO EXTRAÍDO DA IMAGEM PROCESSADA:\n{texto_janela}")
            
            codigo_pa = extrair_codigo_pa_do_texto(texto_janela)
            
            pyautogui.press("enter")
            time.sleep(1)
            
            if codigo_pa:
                logging.info(f"PA EXTRAÍDO: {codigo_pa}")
            else:
                logging.warning("PA NÃO foi extraído da região marcada")
            
            return codigo_pa
            
        except ImportError:
            logging.error("A biblioteca 'pytesseract' não está instalada!")
            pyautogui.press("enter")
            return "ERRO_SEM_OCR"
        except Exception as e:
            logging.error(f"Erro durante o processo de OCR: {e}")
            pyautogui.press("enter")
            return "ERRO_OCR"
            
    except Exception as e:
        logging.error(f"Erro ao extrair PA da janela: {e}")
        try:
            pyautogui.press("enter", failsafe=False)
        except:
            pass
        return None

def processar_apos_f2():
    """
    Processa o que acontece após pressionar F2
    FLUXO CORRETO:
    1. Se aparecer img/CONFIRMAR.PNG = item NOVO → clica SIM → espera informacao.PNG → extrai PA
    2. Se aparecer img/informacao.PNG direto = item JÁ CADASTRADO → extrai PA direto
    
    Returns:
        str ou None: Código PA extraído ou None se não encontrou
    """
    try:
        logging.info("=== PROCESSANDO RESULTADO APÓS F2 ===")
        
        # Aguardar um tempo para alguma janela aparecer
        time.sleep(3)
        
        # Procurar por ambas as imagens por até 20 segundos
        tempo_limite = 20
        tempo_inicio = time.time()
        tentativa = 0
        
        while time.time() - tempo_inicio < tempo_limite:
            tentativa += 1
            
            if tentativa % 5 == 0:  # Log a cada 5 tentativas
                logging.info(f"Tentativa {tentativa}: Verificando janelas após F2...")
            
            try:
                # PRIMEIRO: Verificar se apareceu img/CONFIRMAR.PNG (item NOVO)
                try:
                    localizacao_confirmar = pyautogui.locateOnScreen("img/CONFIRMAR.PNG", confidence=0.7)
                    if localizacao_confirmar:
                        logging.info("ITEM NOVO DETECTADO - janela CONFIRMAR.PNG encontrada!")
                        
                        # Clicar em SIM
                        if not clicar_sim():
                            logging.error("Falha ao clicar em SIM")
                            return None
                        
                        # Aguardar processamento
                        logging.info("Aguardando sistema processar após SIM...")
                        time.sleep(5)
                        
                        # Agora aguardar img/informacao.PNG aparecer
                        logging.info("Aguardando janela informacao.PNG aparecer após SIM...")
                        return aguardar_informacao_apos_sim()
                        
                except pyautogui.ImageNotFoundException:
                    pass
                
                # SEGUNDO: Verificar se apareceu img/informacao.PNG direto (item JÁ CADASTRADO)
                try:
                    localizacao_info = pyautogui.locateOnScreen("img/informacao.PNG", confidence=0.7)
                    if localizacao_info:
                        logging.info("ITEM JÁ CADASTRADO - janela informacao.PNG encontrada diretamente!")
                        return extrair_pa_da_janela_informacao(localizacao_info)
                except pyautogui.ImageNotFoundException:
                    pass
                
            except Exception as e:
                logging.warning(f"Erro ao procurar janelas (tentativa {tentativa}): {e}")
            
            time.sleep(1)
        
        logging.warning("Nenhuma janela encontrada após F2 no tempo limite")
        return None
        
    except Exception as e:
        logging.error(f"Erro ao processar após F2: {e}")
        return None

def clicar_sim():
    """
    Clica especificamente no botão SIM quando aparece janela de confirmação
    
    Returns:
        bool: True se clicou com sucesso
    """
    try:
        logging.info("Procurando e clicando em img/SIM.PNG...")
        
        if not Path("img/SIM.PNG").exists():
            logging.error("Arquivo img/SIM.PNG não encontrado!")
            return False
        
        # Procurar e clicar em SIM por até 10 segundos
        tempo_limite = 10
        tempo_inicio = time.time()
        
        while time.time() - tempo_inicio < tempo_limite:
            try:
                localizacao_sim = pyautogui.locateOnScreen("img/SIM.PNG", confidence=0.8)
                if localizacao_sim:
                    centro_sim = pyautogui.center(localizacao_sim)
                    pyautogui.click(centro_sim)
                    logging.info(f"SIM clicado com sucesso na posição: {centro_sim}")
                    return True
            except pyautogui.ImageNotFoundException:
                pass
            time.sleep(0.5)
        
        logging.error("Não foi possível encontrar/clicar em img/SIM.PNG no tempo limite")
        return False
        
    except Exception as e:
        logging.error(f"Erro ao clicar em SIM: {e}")
        return False

def aguardar_informacao_apos_sim():
    """
    Aguarda especificamente a janela img/informacao.PNG aparecer APÓS clicar em SIM
    e extrai o PA
    
    Returns:
        str ou None: Código PA extraído ou None se não encontrou
    """
    try:
        logging.info("Aguardando janela informacao.PNG aparecer após confirmar item novo...")
        
        # Verificar se a imagem existe
        if not Path("img/informacao.PNG").exists():
            logging.error("Arquivo img/informacao.PNG não encontrado!")
            return None
        
        # Aguardar pela janela informacao.PNG por até 30 segundos
        tempo_limite = 30
        tempo_inicio = time.time()
        tentativa = 0
        
        while time.time() - tempo_inicio < tempo_limite:
            tentativa += 1
            try:
                if tentativa % 10 == 0:  # Log a cada 10 tentativas
                    logging.info(f"Tentativa {tentativa}: Aguardando janela informacao.PNG...")
                
                # Procurar a imagem informacao.PNG na tela
                localizacao = pyautogui.locateOnScreen("img/informacao.PNG", confidence=0.7)
                
                if localizacao:
                    logging.info("Janela informacao.PNG encontrada após cadastro!")
                    time.sleep(2)  # Aguardar carregar
                    return extrair_pa_da_janela_informacao(localizacao)
                    
            except pyautogui.ImageNotFoundException:
                pass
            except Exception as e:
                logging.warning(f"Erro ao aguardar janela (tentativa {tentativa}): {e}")
            
            time.sleep(1)
        
        logging.error("Janela informacao.PNG NÃO apareceu no tempo limite após SIM")
        return None
        
    except Exception as e:
        logging.error(f"Erro ao aguardar janela informacao: {e}")
        return None

def salvar_codigo_pa_na_planilha(codigo_pa, linha_processada):
    """
    Salva o código PA na coluna B da aba configurada
    
    Args:
        codigo_pa: Código PA para salvar
        linha_processada: Número da linha que está sendo processada
        
    Returns:
        bool: True se salvou com sucesso
    """
    app = None
    wb = None
    try:
        if not codigo_pa or codigo_pa in ["ERRO_SEM_OCR", "ERRO_OCR"]:
            logging.warning("Código PA inválido, não salvando na planilha")
            return False
        
        if not linha_processada:
            logging.error("Linha processada não informada")
            return False
        
        # IMPORT GENÉRICO
        from cesar_modulo import encontrar_planilha_mais_recente, encontrar_aba_configurada
        caminho_planilha = encontrar_planilha_mais_recente()
        if not caminho_planilha:
            logging.error("Não foi possível encontrar planilha para salvar PA")
            return False
        
        logging.info(f"Abrindo planilha '{os.path.basename(caminho_planilha)}' para salvar o PA...")
        app = xw.App(visible=False)
        wb = app.books.open(caminho_planilha)
        
        # BUSCAR A ABA CONFIGURADA
        ws = encontrar_aba_configurada(wb)
        if not ws:
            logging.error(f"Não foi possível encontrar aba configurada {ABA_TEXTO_BUSCA} para salvar PA")
            wb.close()
            app.quit()
            return False
        
        logging.info(f"Usando aba: {ws.name}")
        logging.info(f"Escrevendo o código '{codigo_pa}' na célula B{linha_processada}...")
        ws.range(f'B{linha_processada}').value = codigo_pa
        
        logging.info("Salvando o ficheiro Excel...")
        wb.save()
        logging.info("Fechando o ficheiro Excel...")
        wb.close()
        logging.info("Encerrando a instância do Excel...")
        app.quit()
        
        logging.info("Aguardando 1 segundo para garantir que o ficheiro foi atualizado na rede...")
        time.sleep(1)
        
        logging.info(f"Código PA salvo na aba configurada - linha {linha_processada}.")
        return True
        
    except Exception as e:
        logging.error(f"Erro ao salvar código PA na planilha: {e}", exc_info=True)
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
        return False

def inserir_dados_planilha(dados):
    """
    Função principal para inserir dados da planilha (atualizada)
    """
    try:
        linha_atual = dados.get('linha_usada')
        if not linha_atual:
            logging.error("Linha atual não encontrada nos dados.")
            return False
        
        logging.info(f"Processando linha {linha_atual}")
        
        # Usar a função de inserção configurada
        if not inserir_dados_planilha_completa(dados):
            return False
        
        # Após F2, tenta obter o código PA
        codigo_pa = processar_apos_f2()
        
        # SE NÃO OBTEVE O CÓDIGO PA
        if not codigo_pa:
            logging.warning("Código PA não foi obtido. Salvando status de erro na planilha.")
            # Tenta salvar a tag de erro para não processar esta linha novamente
            salvar_codigo_pa_na_planilha("ERRO_OCR", linha_atual)
            # Retorna False para que o loop principal saiba que esta linha falhou
            return False

        # SE OBTEVE O CÓDIGO, TENTA SALVAR
        if not salvar_codigo_pa_na_planilha(codigo_pa, linha_atual):
            logging.warning("Falha ao salvar o código PA na planilha.")
            # Também retorna False se o salvamento falhar
            return False
            
        # Apenas retorna True se TUDO deu certo (obteve o PA e salvou)
        return True
        
    except Exception as e:
        logging.error(f"Erro ao inserir dados da planilha: {e}")
        return False

def executar_processo_completo_com_insercao(numero_empresa=None):
    """
    Executa o processo completo: login + configurador + pesquisa + navegação + inserção dados
    
    Args:
        numero_empresa: Número da empresa para login
    
    Returns:
        bool: True se todo o processo foi bem-sucedido
    """
    try:
        if numero_empresa is None:
            numero_empresa = EMPRESA_PADRAO
            
        logging.info("=== INICIANDO PROCESSO COMPLETO COM INSERÇÃO ===")
        
        # Passo 1: Executar processo até acessar o tipo da planilha
        sucesso, numero, dados = executar_processo_completo(numero_empresa=numero_empresa)
        if not sucesso:
            logging.error("Falha no processo inicial (login + configurador + pesquisa)")
            return False
        
        # Aguardar um tempo adicional para garantir que a tela carregou
        time.sleep(2)
        
        # Passo 2: Navegar com tabs configurados
        if not navegar_7_tabs():
            logging.error(f"Falha ao navegar com {TABS_INICIAIS}x Tab")
            return False
        
        # Passo 3: Inserir dados da planilha e verificar PA
        if not inserir_dados_planilha(dados):
            logging.error("Falha ao inserir dados da planilha")
            return False
        
        logging.info("=== PROCESSO COMPLETO COM INSERÇÃO FINALIZADO ===")
        return True
        
    except Exception as e:
        logging.error(f"Erro no processo completo com inserção: {e}")
        return False

def main():
    """Função principal do módulo de inserção de dados"""
    configurar_logging()
    
    try:
        logging.info("=== INICIANDO MÓDULO DE INSERÇÃO DE DADOS ===")
        
        # Verificar se imagens necessárias existem
        imagens_necessarias = [
            "img/informacao.PNG",
            "img/CONFIRMAR.PNG", 
            "img/SIM.PNG"
        ]
        
        for imagem in imagens_necessarias:
            if not Path(imagem).exists():
                logging.warning(f"Arquivo {imagem} não encontrado - funcionalidade pode não funcionar")
        
        # Executar processo completo UMA ÚNICA VEZ
        sucesso = executar_processo_completo_com_insercao(numero_empresa=EMPRESA_PADRAO)
        
        if sucesso:
            logging.info("=== MÓDULO DE INSERÇÃO EXECUTADO COM SUCESSO ===")
            logging.info("Linha processada com sucesso!")
            
            # Manter programa ativo para interação manual
            logging.info("Sistema finalizado. Pressione Ctrl+C para sair.")
            try:
                while True:
                    time.sleep(10)
            except KeyboardInterrupt:
                logging.info("Finalizando módulo de inserção...")
        else:
            logging.error("=== FALHA NA EXECUÇÃO DO MÓDULO DE INSERÇÃO ===")
            return False
            
    except KeyboardInterrupt:
        logging.info("Execução interrompida pelo usuário")
        return True
    except Exception as e:
        logging.error(f"Erro inesperado no módulo de inserção: {e}")
        return False

if __name__ == "__main__":
                                                                                                        
    sys.exit(0 if sucesso else 1)