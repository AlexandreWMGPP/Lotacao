import pyautogui as ptg
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from pathlib import Path
from PIL import Image, ImageGrab
import pytesseract
import time
from unidecode import unidecode
import os
import io
import sys
import cv2
import json
import re
from machine_utils import get_positions, get_machine_info, get_image_path, get_windows_scale_factor, resize_image, get_popup_coordinate

# Garante que stdout e stderr usem UTF-8
os.environ["PYTHONIOENCODING"] = "utf-8"
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# Configurar o caminho do Tesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
os.environ["TESSDATA_PREFIX"] = r"C:\Program Files\Tesseract-OCR\tessdata"

def find_image_on_screen(image_path, confidence=0.3):
    """ Tenta localizar a imagem na tela com diferentes escalas """
    for scale in [1.0, 1.25, 1.5]:
        try:
            resized_img = resize_image(image_path, scale)
            if resized_img is None:
                continue
            location = ptg.locateOnScreen(resized_img, grayscale=True, confidence=confidence)
            if location:
                return location
        except ptg.ImageNotFoundException:
            pass
    return None

# Definir posição dos elementos de tela
positions = get_positions("Allocation")
machineId = get_machine_info()

def json_error_message(message):
    error_message = {
        "status": "error",
        "message": message
    }
    print(json.dumps(error_message, ensure_ascii=False))
    sys.exit(1)

def sort_spreadsheet_by_column(filePath, ordering_column):
    try:
        # Carregar o arquivo Excel
        workbook = load_workbook(filePath)
        aba_lotacao = workbook.active

        # Ler os dados da planilha na memória
        dados = list(aba_lotacao.iter_rows(values_only=True))

        # Identificar o cabeçalho na segunda linha
        cabecalho = dados[1]
        conteudo = dados[2:]  # Dados começam após o cabeçalho

        # Verificar se a coluna de ordenação existe
        if ordering_column not in cabecalho:
            raise ValueError(f"A coluna '{ordering_column}' não foi encontrada na planilha.")

        # Obter o índice da coluna a ser usada para ordenação
        indice_coluna = cabecalho.index(ordering_column)

        # Ordenar os dados pelo índice da coluna
        conteudo_ordenado = sorted(conteudo, key=lambda row: row[indice_coluna] if row[indice_coluna] else "")

        # Escrever os dados de volta na planilha
        aba_lotacao.delete_rows(3, aba_lotacao.max_row - 2)  # Remove todas as lines de dados, preservando cabeçalho e primeira linha

        for i, linha in enumerate(conteudo_ordenado, start=3):  # Inserir dados a partir da terceira linha
            for j, valor in enumerate(linha, start=1):
                aba_lotacao.cell(row=i, column=j, value=valor)

        # Salvar o arquivo ordenado
        workbook.save(filePath)

    except Exception as e:
          json_error_message(f"Erro na ordenação da Planilha de Lotação: {e}")
    
def click_and_wait(position, duration=0.5):
    ptg.moveTo(position, duration=duration)
    ptg.click()
    time.sleep(duration)

def click_and_write(position, text, duration=0.5):
    ptg.moveTo(position, duration=duration)
    ptg.click(position)
    ptg.write(text.strip())
    time.sleep(duration)

def select_secretary(secretaria):
    # Sempre que quebrar a Secretaria/Setor clico no botão da Secretaria atual para garatir
    # que estamos na pagina principal
    click_and_wait(positions["btnSetorAtivo"])
    time.sleep(0.5)
   
    # Selecioar o orgão/secretaria
    click_and_write(positions["selSetor"], secretaria, 0.5)
    ptg.press("enter")
    time.sleep(0.5)
    
    #click na secretaria selecionada para ir para página de lotação
    click_and_wait(positions["selSecr"])

def grava_mensagem_ocupacao_nao_encontrada_planilha(aba_lotacao, i):
    
    aba_lotacao[f"H{i}"].font = Font(bold = True, color='FF0000')
    aba_lotacao[f"I{i}"].font = Font(bold = True, color='FF0000')
    aba_lotacao[f"H{i}"].value = "NOK"
    aba_lotacao[f"I{i}"].value = "Ocupação do Servidor não Encontrada"  
    ptg.click(positions["fecharOcupacao"])
    ptg.click(positions["fecharNovaLotacao"])    

def grava_mensagem_unidade_nao_encontrada_planilha(aba_lotacao, i):
   
    aba_lotacao[f"H{i}"].font = Font(bold = True, color='FF0000')
    aba_lotacao[f"I{i}"].font = Font(bold = True, color='FF0000')
    aba_lotacao[f"H{i}"].value = "NOK"
    aba_lotacao[f"I{i}"].value = "Unidade de Lotação do Servidor não Encontrada"  
    ptg.click(positions["fecharUnidadeLotacao"], duration = 0.5)
    ptg.click(positions["fecharNovaLotacao"], duration = 0.5)   
                            
def extrair_texto_ocr(imagem_popup):  # Parâmetro agora é a imagem_popup em vez do caminho
    # Certifique-se de que a imagem_popup já esteja em escala de cinza
    if len(imagem_popup.shape) == 3:  # Verifica se a imagem_popup está em RGB ou BGR
        imagem_popup = cv2.cvtColor(imagem_popup, cv2.COLOR_BGR2GRAY)

    # Aplica um filtro Gaussiano para reduzir ruído
    imagem_popup = cv2.GaussianBlur(imagem_popup, (5, 5), 0)

    # Usa a técnica de binarização Otsu para destacar o texto
    _, imagem_popup = cv2.threshold(imagem_popup, 150, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    # Opcional: aumenta a imagem_popup para melhorar a leitura
    largura = int(imagem_popup.shape[1] * 2)  # Duplicando a largura
    altura = int(imagem_popup.shape[0] * 2)   # Duplicando a altura
    imagem_popup = cv2.resize(imagem_popup, (largura, altura), interpolation=cv2.INTER_CUBIC)

    # Configurações do Tesseract
    custom_config = r'--oem 3 --psm 6'

    # Extração de texto
    texto_extraido = pytesseract.image_to_string(imagem_popup, lang="eng", config=custom_config)   

    return texto_extraido.strip()

def extrair_sigla(texto):
    """Extrai siglas válidas de qualquer parte do texto."""
    match = re.search(r"\b([A-Z0-9]{2,})\b", texto)  # Permite letras e números
    return match.group(1) if match else ""  # Retorna a sigla encontrada ou string vazia

def capturar_popup_e_selecionar(texto_desejado, popup_type):
    """ Captura a tela do popup, processa com OCR e clica na opção correta. """
    time.sleep(1)  # Tempo para o popup carregar

     # Obtem as coordenadas do popup com base no tipo
    result = get_popup_coordinate(machineId, popup_type)
    if result is None:
        raise ValueError("Coordenadas do popup {popup_type} não encontradas para maquina {machineId}.")
    else:
        x1, y1, x2, y2, altura_linha_fixa = result

    # Captura a região do popup
    screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
    # Teste Maquina HP
    # screenshot.save(r"C:\Users\seger\Downloads\popup_capture.png")  # Para depuração
    # Teste Maquina AlienWare
    # screenshot.save(r"C:\Users\lucia\Downloads\popup_capture.png")  # Para depuração
    screenshot = np.array(screenshot)    

    # Pré-processamento para OCR
    screenshot_gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)
    screenshot_gray = cv2.GaussianBlur(screenshot_gray, (3, 3), 0)
    screenshot_gray = cv2.threshold(screenshot_gray, 100, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]

    # Teste Maquina HP
    # cv2.imwrite(r"C:\Users\seger\Downloads\screenshot_gray.png", screenshot_gray)
    # Teste Maquina AlienWare
    # cv2.imwrite(r"C:\Users\lucia\Downloads\screenshot_gray.png", screenshot_gray)

    # **Extração de texto por linhas** 
    texto_extraido = extrair_texto_ocr(screenshot_gray)
    
    linhas = [linha.strip() for linha in texto_extraido.split("\n") if linha.strip()]

    if not linhas:
        raise ValueError("Erro: Nenhuma linha detectada no popup.")

    encontrou_texto = False
    index_linha_correta = None  

    for i, linha in enumerate(linhas):
        # Para Unidades, extraímos apenas a sigla
        if popup_type == "unidade":
            texto_processado = extrair_sigla(linha)
            #Marreta para correção posterito
            if texto_processado == 'OSCFV17':
                texto_processado = 'DSCFV17'
        else:
            # Para Ocupação, usamos a linha inteira
            texto_processado = linha  

        # **Verificação exata primeiro**
        if unidecode(texto_desejado.upper()) == unidecode(texto_processado.upper()):
            encontrou_texto = True
            index_linha_correta = i
            break  # Encontramos a opção correta, saímos do loop
    
    if not encontrou_texto:            
        return False

    x_centro = (x1 + x2) // 2  # Centraliza horizontalmente
    y_centro = y1 + (index_linha_correta * altura_linha_fixa) + (altura_linha_fixa // 2)  # Centraliza verticalmente
    
    # Move o cursor para a opção correta e clica
    ptg.moveTo(x_centro, y_centro, duration=0.3)
    ptg.click()
    
    return True
    
def localizar_campo_e_clicar(imagem_campo, texto_para_escrever):
    # Valida se a imagem existe e é um caminho válido
    if not isinstance(imagem_campo, str) or not os.path.isfile(imagem_campo):
        raise ValueError(f"Erro: O caminho '{imagem_campo}' não é válido ou o arquivo não existe.")

    # Captura uma screenshot da tela
    screenshot = ptg.screenshot()
    tela = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

    # Carrega a imagem do campo
    campo = cv2.imread(imagem_campo, cv2.IMREAD_UNCHANGED)
    if campo is None:
        raise ValueError(f"Erro ao carregar a imagem: {imagem_campo}. Verifique se o caminho está correto.")

    # Converte as imagens para escala de cinza
    tela_gray = cv2.cvtColor(tela, cv2.COLOR_BGR2GRAY)
    campo_gray = cv2.cvtColor(campo, cv2.COLOR_BGR2GRAY) if len(campo.shape) == 3 else campo

    # Garante que ambas as imagens têm o mesmo tipo de dados
    tela_gray = tela_gray.astype(np.uint8)
    campo_gray = campo_gray.astype(np.uint8)

    # Faz a correspondência da imagem
    resultado = cv2.matchTemplate(tela_gray, campo_gray, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(resultado)

    # Define o limiar de confiança para encontrar o campo
    limiar_conf = 0.3
    if max_val >= limiar_conf:
        # Encontra o centro do campo
        altura, largura = campo_gray.shape
        centro_x = max_loc[0] + largura // 2
        centro_y = max_loc[1] + altura // 2

        # Move e clica no centro do campo
        ptg.moveTo(centro_x, centro_y, duration=1)
        ptg.click()

        # Escreve o texto
        ptg.write(texto_para_escrever.strip())
    else:
        print(json.dumps({"status": "error", "message": "Campo não encontrado com confiança suficiente."}))

    
def allocateServer(city, filePath, ordering_column):
    
    scale_factor = get_windows_scale_factor()
    
    imgLotacaoJaCadatrada = get_image_path("lotacaoJaCadastrada.png")
    
    imgOcupacaoServidorEncontrada = get_image_path("ocupacaoServidorEncontrada.png")
    imgOcupacaoServidorEncontrada = resize_image(imgOcupacaoServidorEncontrada, scale_factor)
    
    # Carregar a planilha
    workbook = load_workbook(filePath)
    aba_lotacao = workbook.active

    # Obter todas as lines (a partir da linha com cabeçalho)
    lines = list(aba_lotacao.iter_rows(values_only=True))
    
    # Verificar se há registros suficientes
    if len(lines) < 3:
        raise ValueError(f"Error| A planilha não contém dados suficientes para processamento.")

    # Configurações iniciais
    cabecalho = lines[1]  # Considerando que o cabeçalho está na segunda linha
    inicio_dados = 2  # Registros começam na terceira linha (índice 2 em Python)
    
    if ordering_column not in cabecalho:
        raise ValueError(f"Error| A coluna {ordering_column}' não foi encontrada na planilha.")
    
    # Insiro as Novas colunas de Status e Ocorrência
    aba_lotacao["H2"].value = "Status"
    aba_lotacao.column_dimensions["H"].width = 8.57
    aba_lotacao["H2"].font = Font(size=14, bold = True, color='000000')
    aba_lotacao["H2"].alignment = Alignment(horizontal='center', vertical='center')
    aba_lotacao["I2"].value = "Ocorrencia"
    aba_lotacao.column_dimensions["I"].width = 43.43
    aba_lotacao["I2"].font = Font(size=14, bold = True, color='000000')
    aba_lotacao["I2"].alignment = Alignment(horizontal='center', vertical='center')
    
    ind_secr = cabecalho.index(ordering_column)  # Índice da coluna "Sigla"
    contador_total = 0
    contadores_por_secretaria = {}

    secretaria_atual = None
    mensagens_log = []
    
    ptg.hotkey('alt', 'tab')
    time.sleep(1)

    #Obtenho o nome do Patriarca
    patriarca = aba_lotacao["C1"].value
    # Obtenho a Secretaria para inicio do processament
    secretaria_atual = aba_lotacao["A3"].value

    # Obtenho a quantidade de lines da planilha
    qtd_lines = aba_lotacao.max_row
    ptg.PAUSE = 0.3
    
    click_and_write(positions["selPatriarca"], unidecode(patriarca), 0.5)
    ptg.press("enter")
    time.sleep(0.5)

    # Selecioar o orgão/secretaria
    click_and_write(positions["selSetor"], unidecode(secretaria_atual), 0.5)
    ptg.press("enter")
    time.sleep(0.5)
    
    #click na secretaria selecionada para ir para página de lotação
    click_and_wait(positions["selSecr"])

    # Aguardo a carga da janela de Funções de Lotação
    imgfuncoesJanelaLotacao = get_image_path("funcoesJanelaLotacao.png")  
    imgfuncoesJanelaLotacao = resize_image(imgfuncoesJanelaLotacao, scale_factor)

    foundWindow = None
    tentativas = 0
    max_tentativas = 10  # Número máximo de tentativas antes de desistir
    while not foundWindow and tentativas < max_tentativas:
        try:
            foundWindow = ptg.locateOnScreen(imgfuncoesJanelaLotacao, grayscale=True, confidence=0.3)
            
            if foundWindow:
                pass
            else:
                time.sleep(1)
        except ptg.ImageNotFoundException:
            time.sleep(1)
        
        tentativas += 1

    if not foundWindow:
        raise ValueError("Erro: Imagem Funções Janela Lotação não encontrada após múltiplas tentativas!")
              
    # Processar as lines de dados
    cont_linhas = inicio_dados + 1
    for i, linha in enumerate(lines[inicio_dados:], start=inicio_dados + 1):
                
        try:
            sigla_secr = linha[ind_secr]
        
            if sigla_secr != secretaria_atual:
                if secretaria_atual:
                    mensagens_log.append(f"INFO| Total de servidores na secretaria {secretaria_atual}: {contadores_por_secretaria.get(secretaria_atual, 0)}")
                
                secretaria_atual = sigla_secr
                contadores_por_secretaria.setdefault(secretaria_atual, 0)  # Inicializa se não existir

                click_and_wait(positions["fecharNovaLotacao"])

                select_secretary(secretaria_atual)

            if secretaria_atual:
                contadores_por_secretaria[secretaria_atual] = contadores_por_secretaria.get(secretaria_atual, 0) + 1

            contador_total += 1

            #Obtenho os dados do servidor da linha corrente
            nome = linha[2]
            cpf = linha[3]
            ocupacao_cargo = linha[4]
            sigla_setor = linha[5]
            unidadeLotacao = linha[6]
            
            if not nome:
                break
            
            #Clicar no quadro de Lotação
            click_and_wait(positions["qdrLotacao"])
            
            # Aguardo a carga da imagem do botão Nova Lotação
            imgbtnNovaLotacaoCarregado = get_image_path("btnNovaLotacao.png")
            imgbtnNovaLotacaoCarregado = resize_image(imgbtnNovaLotacaoCarregado, scale_factor)
            
            foundWindow = None
            tentativas = 0
            max_tentativas = 10  # Número máximo de tentativas antes de desistir
            while not foundWindow and tentativas < max_tentativas:
                try:
                    foundWindow = ptg.locateOnScreen(imgbtnNovaLotacaoCarregado, grayscale=True, confidence=0.3)
                    
                    if foundWindow:
                        pass
                    else:
                        time.sleep(1)
                except ptg.ImageNotFoundException:
                    time.sleep(1)
                
                tentativas += 1

            if not foundWindow:
                raise ValueError("Erro: Imagem Botão Nova Lotação não encontrada após múltiplas tentativas!")
           
            click_and_wait(positions["btnNovaLotacao"])

            click_and_wait(positions["qdrServidor"])

            remover = "'.','-'"
            if cpf:
                cpfn = ''.join(filter(lambda i: i not in remover, cpf))
            
            click_and_write(positions["cmpCPF"], cpfn, 0.5)
            ptg.press("enter")
            time.sleep(0.5)
            
            # Verifica se não deu erro no CPF
            location = None
            imgCpfValido = get_image_path("cpfValido.png")

            if imgCpfValido:
                location = find_image_on_screen(imgCpfValido)
            else:
                location = None

            if location == None:
                # Quando os campos Ocupação e Unidade não são exibidos identifica um problema
                aba_lotacao[f"H{i}"].font = Font(bold = True, color='FF0000')
                aba_lotacao[f"I{i}"].font = Font(bold = True, color='FF0000')
                aba_lotacao[f"H{i}"].value = "NOK"
                aba_lotacao[f"I{i}"].value = "Erro na validação do CPF"
                click_and_wait(positions["fecharNovaLotacao"]) 
                time.sleep(0.5)
                continue                 

            # Aguardo os campos Ocupação e Unidade serem exibidos
            imgocupacaoUnidadeExibida = get_image_path("ocupacaoUnidadeExibida.png")
            imgocupacaoUnidadeExibida = resize_image(imgocupacaoUnidadeExibida, scale_factor) 

            foundWindow = None
            tentativas = 0
            max_tentativas = 10  # Número máximo de tentativas antes de desistir
            while not foundWindow and tentativas < max_tentativas:
                try:
                    foundWindow = ptg.locateOnScreen(imgocupacaoUnidadeExibida, grayscale=True, confidence=0.3)
                    
                    if foundWindow:
                        pass
                    else:
                        time.sleep(1)
                except ptg.ImageNotFoundException:
                    time.sleep(1)
                
                tentativas += 1

            if not foundWindow:
                raise ValueError("Erro: Imagem Ocupação Unidade Exibida não encontrada após múltiplas tentativas!")
           
            #Quando os campos Ocupação e Unidade são exibidos significa que o CPF é valido
            #Informo o Cargo do Servidor
            click_and_wait(positions["cmpQualOcupacao"])

            try:
                imgcmpSelOcupacao = get_image_path("cmpSelOcupacao.png")
                imgcmpSelOcupacao = resize_image(imgcmpSelOcupacao, scale_factor) 

                localizar_campo_e_clicar(imgcmpSelOcupacao, unidecode(ocupacao_cargo.strip()))
            except ValueError as e:
                print(e)

            time.sleep(0.5)
            ptg.press("enter")
            time.sleep(0.5)
            
            #verifico se a ocupação informada foi encontrada
            try:
                location = ptg.locateOnScreen(imgOcupacaoServidorEncontrada, grayscale=True, confidence=0.3) 
                pass
            except ptg.ImageNotFoundException:
                grava_mensagem_ocupacao_nao_encontrada_planilha(aba_lotacao, i)
                continue    

            if not capturar_popup_e_selecionar(ocupacao_cargo, "ocupacao"):
                grava_mensagem_ocupacao_nao_encontrada_planilha(aba_lotacao, i)
                continue
            
            time.sleep(0.5)
            
            #Informo a Unidade de Lotação
            click_and_wait(positions["cmpQuallUnidade"])

            # Localizar Unidade Lotação

            try:
                imgcmpSelUnidLotacao = get_image_path("cmpSelUnidLotacao.png")
                imgcmpSelUnidLotacao = resize_image(imgcmpSelUnidLotacao, scale_factor)

                localizar_campo_e_clicar(imgcmpSelUnidLotacao, unidecode(sigla_setor.strip()))
            except ValueError as e:
                print(e)
                        
            ptg.press("enter")
            time.sleep(0.5)
            
            imgUnidadeLotacaoEncontrada = get_image_path("unidadeLotacaoEncontrada.png")
            imgUnidadeLotacaoEncontrada = resize_image(imgUnidadeLotacaoEncontrada, scale_factor)
            #verifico se a Unidade de Lotação informada foi encontrada
            try:
                location = ptg.locateOnScreen(imgUnidadeLotacaoEncontrada, grayscale=True, confidence=0.3) 
                pass       
            except ptg.ImageNotFoundException:
                grava_mensagem_unidade_nao_encontrada_planilha(aba_lotacao, i)
                continue
            
            ## capturar_popup_e_selecionar(unidecode(sigla_setor.strip()))
            if not capturar_popup_e_selecionar(sigla_setor, "unidade"):
                grava_mensagem_unidade_nao_encontrada_planilha(aba_lotacao, i)
                continue
                
            time.sleep(0.5)

            click_and_wait(positions["btnCriarLotacao"])
            time.sleep(0.5)
            
            location = None
  
            if imgLotacaoJaCadatrada:
                location = find_image_on_screen(imgLotacaoJaCadatrada,confidence=0.9)

            if location:
                aba_lotacao[f"H{i}"].font = Font(bold = True, color='FF0000') 
                aba_lotacao[f"I{i}"].font = Font(bold = True, color='FF0000')
                aba_lotacao[f"H{i}"].value = "NOK"
                aba_lotacao[f"I{i}"].value = "Lotação do Servidor Já Cadastrada"
                click_and_wait(positions["btnFecharAlerta"])
                time.sleep(1)
                click_and_wait(positions["fecharNovaLotacao"])
                continue 
            else:    
                aba_lotacao[f"H{i}"].font = Font(bold = True, color='006400')
                aba_lotacao[f"I{i}"].font = Font(bold = True, color='006400')
                aba_lotacao[f"H{i}"].value = "OK"
                aba_lotacao[f"I{i}"].value = "Lotação Cadastrada com Sucesso!"

            time.sleep(0.5)

            cont_linhas += 1
            
            if cont_linhas > len(lines):
                break
        
        except Exception as e:
            mensagens_log.append(f"ERROR| Erro ao processar a linha {i}: {e}")
        
    # Finalizar contagem
    if secretaria_atual:
        mensagens_log.append(f"INFO| Total de servidores na secretaria {secretaria_atual}: {contadores_por_secretaria.get(secretaria_atual, 0)}")
    mensagens_log.append(f"INFO| Total geral de servidores processados: {contador_total}")

    workbook.save(filePath)

     # Volto para pagina principal
    click_and_wait(positions["btnSetorAtivo"])
    time.sleep(0.5)

    # Retorna para tela do sistema
    imgretornoTelaAutomatizacaoLotacao = get_image_path("retornoTelaAutomatizacaoLotacao.png")
    imgretornoTelaAutomatizacaoLotacao = resize_image(imgretornoTelaAutomatizacaoLotacao, scale_factor)

    ptg.hotkey('alt', 'tab')
    
    return contador_total, contadores_por_secretaria, mensagens_log
 
def save_log_to_file(log, output_path):
    try:
        with open(output_path, 'w', encoding='utf-8') as log_file:
            json.dump(log, log_file, ensure_ascii=False, indent=4)
    except Exception as e:
        print(json.dumps({"status": "error", "message": f"Não foi possível salvar o log no arquivo: {e}"}))

if __name__ == "__main__":
    if len(sys.argv) != 3:
        error_message = "Número incorreto de argumentos. Uso: python automaticServerAllocation.py <município> <pasta>"
        print(json.dumps({"status": "error", "message": error_message}))
        sys.exit(1)

    city = sys.argv[1]
    filePath = sys.argv[2]

    if not os.path.exists(filePath):
        error_message = "A pasta fornecida não existe."
        print(json.dumps({"status": "error", "message": error_message}))
        sys.exit(1)
    
    # Ordenar a Planilha de Lotação pela Coluna "Órgão/Secretaria"
    ordering_column = "Sigla"
    sort_spreadsheet_by_column(filePath, ordering_column)
   
    quantidade, contadores, log = allocateServer(city, filePath, ordering_column)
    
    log_path = os.path.normpath(os.path.join(os.path.dirname(filePath), f"log_{city}.json"))
    output = {
        "status": "success",
        "quantidade": quantidade,
        "contadores": contadores,
        "log": log,
        "log_path": log_path  # Adiciona o caminho do log no JSON
    }
    
    save_log_to_file(log, output["log_path"])  # Salva o log no arquivo

    try:
        resultado = {"status": "sucesso", "mensagem": "Processamento concluído"}
        print(json.dumps(resultado))
        # Serializando corretamente para JSON
        json_output = json.dumps(output, ensure_ascii=False)

        print(json_output, flush=True)  # Esta é a saída que o chamador irá capturar
    except TypeError as e:
        error_output = {"error": f"Erro ao serializar JSON: {str(e)}"}
        print(json.dumps(error_output, ensure_ascii=False), flush=True)
        print("Erro ao serializar JSON:", str(e), flush=True)
    except Exception as e:
       error_output = {"error": f"Erro ao serializar JSON: {str(e)}"}
       print(json.dumps(error_output, ensure_ascii=False), flush=True)
       print("Erro ao serializar JSON:", str(e), flush=True)

