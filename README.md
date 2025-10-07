📋 README - Robô de Cadastro Automático
📖 Descrição
Robô RPA (Robotic Process Automation) desenvolvido em Python para automatizar o cadastro de itens no sistema SYSTEAM. O robô realiza login, acessa o módulo configurador, insere dados de planilhas Excel/CSV e extrai códigos PA usando OCR.
🎯 Funcionalidades

✅ Login automático no sistema SYSTEAM
✅ Leitura de planilhas Excel/CSV da rede
✅ Acesso automático ao módulo configurador
✅ Pesquisa de tipos baseada na aba da planilha
✅ Inserção automática de dados em campos configuráveis
✅ Extração de códigos PA via OCR (pytesseract)
✅ Atualização automática da planilha com códigos PA
✅ Processamento em lote de múltiplas linhas
✅ Sistema completo de logs
✅ Tratamento de erros e validações

🗂️ Estrutura do Projeto
projeto/
│
├── main.py                  # Orquestrador principal - processa todas as linhas
├── login.py                 # Módulo de login no sistema
├── cesar_modulo.py          # Acesso ao configurador e pesquisa de tipos
├── insercao_dados.py        # Inserção de dados e extração de PA
├── config.py                # Configurações centralizadas
│
├── img/                     # Imagens para reconhecimento visual
│   ├── menu systeam.PNG
│   ├── modulo item por descricao.PNG
│   ├── pesquisar tipos 615.PNG
│   ├── informacao.PNG
│   ├── CONFIRMAR.PNG
│   └── SIM.PNG
│
└── logs/                    # Logs de execução (criado automaticamente)
⚙️ Configuração
Arquivo config.py
Todas as configurações principais estão centralizadas neste arquivo:
python# Configurações da Aba
ABA_NUMERO = "973"              # Número da aba a processar
ABA_TEXTO_BUSCA = "(973)"       # Texto identificador no nome da aba

# Configurações das Colunas
COLUNA_INICIAL = "C"            # Primeira coluna a preencher
COLUNA_FINAL = "U"              # Última coluna a preencher
COLUNAS_PULAR = ["C", "G", "H", "J", "N", "O", "Q", "R"]  # Colunas que só recebem TAB

# Configurações de Velocidade
TEMPO_ENTRE_CARACTERES = 0.03   # Velocidade de digitação
TEMPO_TAB_PULAR = 0.1           # Tempo ao pular colunas
INTERVALO_ENTRE_ACOES = 0.7     # Pausa entre ações

# Outras Configurações
EMPRESA_PADRAO = 200            # Código da empresa para login
TABS_INICIAIS = 7               # Tabs antes de começar a inserir
🔧 Requisitos
Bibliotecas Python
bashpip install pyautogui
pip install xlwings
pip install pytesseract
pip install Pillow
pip install pandas
pip install pathlib
Software Adicional

Tesseract OCR: Necessário para extração de texto das imagens

Download: https://github.com/UB-Mannheim/tesseract/wiki
Instalar e adicionar ao PATH do sistema


Microsoft Excel: Necessário para xlwings funcionar

📁 Planilha
Localização
\\192.168.1.250\Programas\robo custos\lancar\
Formato Esperado

Coluna A: Código de controle (obrigatório)
Coluna B: Código PA (preenchido pelo robô)
Colunas C-U: Dados para inserção (configurável)

Requisitos da Aba

Nome da aba deve conter o número configurado entre parênteses
Exemplo: "Porta (973)" para ABA_NUMERO = "973"

🚀 Execução
Modo Completo (Recomendado)
Processa todas as linhas da planilha automaticamente:
bashpython main.py
Módulos Individuais
Para testes ou execução parcial:
bash# Apenas login
python login.py

# Login + acesso ao configurador + pesquisa
python cesar_modulo.py

# Processo completo de uma linha
python insercao_dados.py
📊 Fluxo de Execução
mermaidgraph TD
    A[Início] --> B[Buscar próxima linha sem PA]
    B --> C{Primeira linha?}
    C -->|Sim| D[Login completo no sistema]
    C -->|Não| E[Sequência de reinício]
    D --> F[Acessar configurador]
    E --> F
    F --> G[Pesquisar tipo da aba]
    G --> H[Navegar 7 tabs]
    H --> I[Inserir dados da planilha]
    I --> J[Pressionar F2]
    J --> K{Janela de confirmação?}
    K -->|Item novo| L[Clicar em SIM]
    K -->|Item existente| M[Janela de informação]
    L --> M
    M --> N[Extrair PA via OCR]
    N --> O[Salvar PA na coluna B]
    O --> P{Mais linhas?}
    P -->|Sim| B
    P -->|Não| Q[Fim]
🔍 Sistema de Logs
Os logs são salvos automaticamente na pasta logs/ com timestamp:
logs/
├── login_20251007_143022.log
├── login_20251007_150133.log
└── ...
Níveis de Log

INFO: Operações normais e progresso
WARNING: Situações inesperadas mas recuperáveis
ERROR: Erros que impedem o processamento

🎨 Imagens Necessárias
O robô usa reconhecimento visual. Certifique-se de que estas imagens estão na pasta img/:
ImagemPropósitomenu systeam.PNGIdentificar menu principalmodulo item por descricao.PNGAcessar módulopesquisar tipos 615.PNGBotão de pesquisainformacao.PNGJanela com código PACONFIRMAR.PNGJanela de confirmaçãoSIM.PNGBotão de confirmação
🛡️ Segurança e Failsafe

FAILSAFE ativado: Mover mouse para canto superior esquerdo interrompe execução
Timeout configurado: Evita loops infinitos
Validações: Verifica arquivos, abas e dados antes de processar
Tratamento de erros: Registra erros e continua processamento quando possível

📝 Tratamento de Erros
Linha com Erro
Se uma linha falhar:

Salva "ERRO_OCR" na coluna B
Registra erro no log
Continua para próxima linha

Coluna B Preenchida

Linhas com PA já cadastrado são automaticamente puladas
Evita reprocessamento desnecessário

🔧 Personalização
Alterar Colunas a Processar
Edite em config.py:
pythonCOLUNA_INICIAL = "C"  # Primeira coluna
COLUNA_FINAL = "U"    # Última coluna
COLUNAS_PULAR = ["C", "G"]  # Colunas que só recebem TAB
Alterar Velocidade
pythonTEMPO_ENTRE_CARACTERES = 0.03  # Diminuir = mais rápido
INTERVALO_ENTRE_ACOES = 0.7    # Pausa entre ações
Alterar Empresa
pythonEMPRESA_PADRAO = 200  # Código da empresa
⚠️ Observações Importantes

Não mover o mouse durante a execução
Não usar o teclado enquanto o robô está executando
Manter a resolução de tela constante (afeta reconhecimento de imagens)
Garantir acesso à rede para ler/escrever planilhas
Tesseract OCR deve estar instalado e no PATH

🐛 Resolução de Problemas
Robô não encontra imagens

Verificar se imagens estão na pasta img/
Ajustar confidence no código (padrão: 0.7-0.8)
Tirar novas screenshots se resolução mudou

OCR não extrai PA corretamente

Verificar instalação do Tesseract
Ajustar threshold de binarização em extrair_pa_da_janela_informacao()
Verificar imagem de debug gerada: DEBUG_imagem_para_ocr_*.png

Planilha não é encontrada

Verificar acesso à rede: \\192.168.1.250\Programas\robo custos\lancar\
Confirmar permissões de leitura/escrita
Verificar formato do arquivo (.xlsx, .xls, .csv)

Login falha

Verificar credenciais em login.py
Confirmar caminho do executável
Verificar número da empresa

📞 Suporte
Para dúvidas ou problemas:

Verificar logs na pasta logs/
Revisar mensagens de erro detalhadas
Testar módulos individualmente

📄 Licença
Projeto interno - Todos os direitos reservados
