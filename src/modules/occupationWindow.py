import pandas as pd 
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import os
import subprocess
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import unicodedata
import json
import datetime

# Caminho do arquivo JSON de municípios
JSON_PATH = os.path.join(os.getcwd(), "src", "config", "municipios.json")

# Função para carregar municípios do JSON
def carregar_municipios():
    try:
        with open(JSON_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar municípios: {e}")
        return []

# Lista global de municípios carregados
municipios_ES = carregar_municipios()

def iniciar_tela(root):
    """Configura a interface de ocupações na janela passada como parâmetro."""
    frame = tk.Frame(root, padx=20, pady=20, relief="groove", borderwidth=3)
    frame.pack(padx=10, pady=10, fill="both", expand=True)

    def limpar_campos():
        entry_pasta.config(state="normal")
        entry_pasta.delete(0, tk.END)
        entry_pasta.config(state="disabled")

        text_planilhas.config(state="normal")
        text_planilhas.delete("1.0", tk.END)
        text_planilhas.config(state="disabled")

        entry_qtd_ocupacoes.config(state="normal")
        entry_qtd_ocupacoes.delete(0, tk.END)
        entry_qtd_ocupacoes.config(state="disabled")

        entry_arquivo_gerado.config(state="normal")
        entry_arquivo_gerado.delete(0, tk.END)
        entry_arquivo_gerado.config(state="disabled")
        
        if not combo_municipio.get().strip():
            combo_municipio['values'] = municipios_ES  
            combo_municipio.set("")
            desabilitar_campos()

    def habilitar_campos(event=None):
        municipio = combo_municipio.get().strip()
        if municipio in municipios_ES:
            botao_pasta.config(state="normal")
            botao_gerar_planilha.config(state="normal")
            botao_arquivo_gerado.config(state="normal")
            botao_carregar_ocupacoes.config(state="normal")
        else:
            desabilitar_campos()

    def desabilitar_campos():
        entry_pasta.config(state="disabled")
        botao_pasta.config(state="disabled")
        text_planilhas.config(state="disabled")
        botao_gerar_planilha.config(state="disabled")
        entry_qtd_ocupacoes.config(state="disabled")
        entry_arquivo_gerado.config(state="disabled")
        botao_arquivo_gerado.config(state="disabled")
        botao_carregar_ocupacoes.config(state="disabled")

    # Função para filtrar municípios
    def filtrar_municipios(event):
        """Filtra os municípios com base no texto digitado."""
        texto_digitado = combo_municipio.get().strip().lower()

        if not texto_digitado:
            combo_municipio['values'] = municipios_ES  
            desabilitar_campos()
            return
        
        if len(texto_digitado) >= 3:
            lista_filtrada = [mun for mun in municipios_ES if texto_digitado in mun.lower()]
            combo_municipio['values'] = lista_filtrada
            if lista_filtrada:
                combo_municipio.event_generate('<Down>')

    # Adicionando Label e Combobox para Município
    tk.Label(frame, text="Município", anchor="w").pack(fill="x")
    combo_municipio = ttk.Combobox(frame, values=municipios_ES, width=101)
    combo_municipio.pack(fill="x", pady=(0, 10))
    combo_municipio.bind('<KeyRelease>', filtrar_municipios)
    combo_municipio.bind('<<ComboboxSelected>>', habilitar_campos)

    # Campo Pasta Planilha(s) Recebida(s)
    tk.Label(frame, text="Pasta Planilha(s) Recebida(s)", anchor="w").pack(fill="x")
    frame_pasta = tk.Frame(frame)
    frame_pasta.pack(fill="x", pady=(0, 10))

    entry_pasta = tk.Entry(frame_pasta, state="disabled", width=104)
    entry_pasta.pack(side="left", anchor="e")

    botao_pasta = tk.Button(frame_pasta, text="Selecionar", state="disabled", width=12, 
                            command=lambda: selecionar_pasta(entry_pasta, combo_municipio, text_planilhas))
    botao_pasta.pack(side="left", padx=5)

    tk.Label(frame, text="Planilhas Recebidas").pack(side="top", anchor="w")
    text_planilhas = tk.Text(frame, height=10, state="disabled", width=103)
    text_planilhas.pack(side="top", anchor="w", fill="x", pady=(0, 10))

    # Alinhando os botões
    button_frame = tk.Frame(frame)
    button_frame.pack(fill="x", pady=(10, 0))

    botao_gerar_planilha = tk.Button(button_frame, text="Gerar Planilha de Ocupações", state="disabled", width=25, 
                                     command=lambda: dados_gerar_planilha(combo_municipio, entry_pasta, 
                                                                          text_planilhas, entry_qtd_ocupacoes, 
                                                                          entry_arquivo_gerado))
    botao_gerar_planilha.pack(side="right", pady=(0, 5))

    separator = tk.Frame(frame, height=2, bd=1, relief="sunken")
    separator.pack(fill="x", pady=10)
       
    # Frame principal para os campos e rótulos
    frame_campos = tk.Frame(frame)
    frame_campos.pack(fill="x", pady=(0, 10))
    
    tk.Label(frame_campos, text="Planilha de Ocupações Gerada").grid(row=0, column=0, padx=(0, 10), sticky="w")
    entry_arquivo_gerado = tk.Entry(frame_campos, state="disabled", width=94)
    entry_arquivo_gerado.grid(row=1, column=0, padx=(0, 10), pady=(0, 5), sticky="w")

    tk.Label(frame_campos, text="Quantidade de Ocupações").grid(row=0, column=1, padx=(10, 0), sticky="e")
    entry_qtd_ocupacoes = tk.Entry(frame_campos, state="disabled", width=23)
    entry_qtd_ocupacoes.grid(row=1, column=1, padx=(10, 0), pady=(0, 5), sticky="e")

    botao_arquivo_gerado = tk.Button(frame_campos, text="Selecionar", state="disabled", width=12,
                                     command=lambda: selecionar_arquivo(entry_arquivo_gerado, combo_municipio, entry_qtd_ocupacoes))
    botao_arquivo_gerado.grid(row=2, column=0, pady=(10, 0), sticky="w")
   
    botao_carregar_ocupacoes = tk.Button(frame_campos, text="Carregar Ocupações", state="disabled", width=20,
                                         command=lambda: carregar_ocupacoes(combo_municipio, entry_arquivo_gerado))
    botao_carregar_ocupacoes.grid(row=2, column=1, pady=(5, 0), sticky="e")

    botao_fechar = tk.Button(frame, text="Fechar", command=root.destroy, width=10)
    botao_fechar.pack(side="right", padx=5, pady=(5, 0))
  
    # Desabilitar os campos inicialmente
    desabilitar_campos()

def filtrar_municipios(event, entry_municipio):
    """Filtra os municípios com base no texto digitado."""
    texto_digitado = entry_municipio.get().lower()
    lista_filtrada = [mun for mun in municipios_ES if texto_digitado in mun.lower()]
    entry_municipio['values'] = lista_filtrada
    entry_municipio.event_generate('<Down>')  # Abre o dropdown automaticamente


def selecionar_pasta(entry_pasta, combo_municipio, text_planilhas):
    municipio = combo_municipio.get().strip()
    base_path = os.getcwd() + r"\lotacaoMun\planilhasRecebidas"

    if municipio:
        pasta = os.path.join(base_path, f"mun{municipio}")
        if not os.path.exists(pasta):
            messagebox.showerror("Erro - Pasta não encontrada.", f"A pasta com a planilha de Lotação para o município {municipio} não existe:\n {pasta}\n\n" + 
                                   "Verifique se a planilha já foi enviada e crie a pasta com a planilha de Lotação dentro.")
            return

        entry_pasta.config(state="normal")
        entry_pasta.delete(0, tk.END)
        entry_pasta.insert(0, pasta)
        entry_pasta.config(state="disabled")

        arquivos = [f for f in os.listdir(pasta) if f.endswith(".csv") or f.endswith(".xlsx") or f.endswith(".xls")]
        if not arquivos:
            messagebox.showerror("Erro - Arquivo não encontrado.", f"A pasta para conter a planilha de Lotação para o município {municipio} existe, mas está vazia:\n {pasta}\n\n" + 
                                   "Verifique se a planilha já foi enviada e a coloque dentro da pasta de Planilhas Recebidas do municipio.")
            return

        # Seleciona apenas as planilhas com as colunas da planilha de modelo
        colunas_planilha_modelo = ["Sigla", "Secretaria/Órgão", "Nome do Servidor", "CPF", "Ocupação/Cargo", "Sigla Setor", "Setor de Lotação"]

        # Lista para armazenar os arquivos válidos
        arquivos_validos = []

        for arquivo in arquivos:
            caminho_arquivo = os.path.join(pasta, arquivo)
    
            try:
                # Carregar arquivo dependendo da extensão
                if arquivo.endswith(".csv"):
                    df = pd.read_csv(caminho_arquivo, sep=';', encoding='ISO-8859-1', header=1)  # Cabeçalho na segunda linha
                else:  # Para arquivos Excel
                    df = pd.read_excel(caminho_arquivo, header=1)  # Cabeçalho na segunda linha
                
                # Verifica se todas as colunas estão presentes
                if all(col in df.columns for col in colunas_planilha_modelo):
                    arquivos_validos.append(arquivo)     
                    
            except Exception as e:
                print(f"Erro ao processar o arquivo {arquivo}: {e}")

        if not arquivos_validos:
            messagebox.showerror("Erro - Planilha de Lotação do municipio não encontrada.", 
                                 f"A pasta para conter a planilha de Lotação para o município {municipio} existe, mas não foi encontrada " +
                                 f"nenhuma planilha no modelo 'Lotação-Modelo.xlsx' que foi enviada.\n {pasta}\n\n" + 
                                 "Verifique se a planilha já foi enviada e a coloque dentro da pasta de Planilhas Recebidas do municipio.")
            return
        
        text_planilhas.config(state="normal")
        text_planilhas.delete("1.0", tk.END)
        for arquivo in arquivos_validos:
            text_planilhas.insert(tk.END, arquivo + "\n")
        text_planilhas.config(state="disabled")
    else:
        messagebox.showerror("Erro - Seleção de município", "Selecione um município antes de escolher a pasta.")

def dados_gerar_planilha(combo_municipio, entry_pasta, text_planilhas, entry_qtd_ocupacoes, entry_arquivo_gerado):
    # Obtendo o município selecionado
    municipio = combo_municipio.get().strip()
    # Verificar se o município está definido
    if not municipio:
        messagebox.showerror("Erro - Campo obrigatório", "O preenchimento do campo município é obrigatorio, selecione um município para gerar a Planilha de Ocupações.")
        return

    if  municipio not in municipios_ES:
        messagebox.showerror("Erro - Município inválido - Município não cadastrado", "Selecione um município da lista.")
        return

    # Obtendo o caminho da pasta
    caminho_arquivos = entry_pasta.get().strip()
    if not os.path.exists(caminho_arquivos):
        messagebox.showerror("Erro - Pasta não encontrada", f"O caminho especificado com a planilha enviada pelo município não encontrado.\n\n {caminho_arquivos}")
        return
    
    # Obtendo os arquivos válidos listados
    arquivos_validos = text_planilhas.get("1.0", tk.END).strip().split("\n")
    arquivos_validos = [os.path.join(caminho_arquivos, arquivo) for arquivo in arquivos_validos if arquivo]  # Filtrar linhas vazias
    if not arquivos_validos:
        messagebox.showerror("Erro - Arquivo nã encontrado", 
                             f"Nenhum arquivo válido foi encontrado na pasta enviada pelo município.\n\n {caminho_arquivos}")
        return
    
    # Caminho do modelo
    caminho_modelo = os.getcwd() + r"\lotacaoMun\Lotacao-Modelo.xlsx"
    if not os.path.exists(caminho_modelo):
        messagebox.showerror("Erro - Arquivo não encontrado.",
                             f"O modelo de planilha não foi encontrado:\n\n {caminho_modelo}")
        return
   
    # Diretório de saída
    caminho_saida = os.getcwd() + fr"\lotacaoMun\planilhasLotacao\mun{municipio}"
    if not os.path.exists(caminho_saida):
        os.makedirs(caminho_saida)
    
    # Nome do arquivo de saída
    nome_saida = f"lotacao_mun{municipio}.xlsx"
    caminho_completo_saida = os.path.join(caminho_saida, nome_saida)

    try:
        # Copiar o modelo
        wb = load_workbook(caminho_modelo)
        ws = wb.active

        # Adicionar o nome do município na célula B1
        ws.cell(row=1, column=3, value=municipio)

        # Combinar os dados de todos os arquivos válidos
        dados_combinados = []
        for arquivo in arquivos_validos:
            df = pd.read_excel(arquivo, usecols="A:G")  # Lê apenas as colunas A até G
            dados_combinados.append(df)

        # Concatenar todos os dados
        df_final = pd.concat(dados_combinados, ignore_index=True)

        qtd_linhas_antes = len(df_final)
        print(f"Quantidade de linhas antes de remover duplicidades: {qtd_linhas_antes}")
        
        # Remover duplicidades
        df_final = df_final.drop_duplicates()

         # Exibir quantidade de linhas após a remoção de duplicidades
        qtd_linhas_depois = len(df_final)
        print(f"Quantidade de linhas após remover duplicidades: {qtd_linhas_depois}")
        
        # Escrever os dados na planilha a partir da segunda linha
        for r_idx, row in enumerate(df_final.itertuples(index=False, name=None), start=2):  # Começa na linha 2
            for c_idx, value in enumerate(row, start=1):  # Começa na coluna 1
                if c_idx > 7:  # Garante que apenas colunas até G sejam copiadas
                    break
                ws.cell(row=r_idx, column=c_idx, value=value)

        # Ajustar largura das colunas automaticamente (opcional)
        for col in ws.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    max_length = max(max_length, len(str(cell.value)))
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = max_length + 2

        # Salvar a nova planilha
        wb.save(caminho_completo_saida)

        print(f"Planilha gerada com sucesso em:\n{caminho_completo_saida}")
    except Exception as e:
        messagebox.showerror("Erro no processamento", 
                             f"Ocorreu um erro ao processar os dados: {e}")
   
    qtd_ocupacoes = 0 
    arquivo_gerado = None

    qtd_ocupacoes, arquivo_gerado = gerar_planilha(municipio, nome_saida, caminho_saida)
    
    # Atualiza os campos na interface
    entry_qtd_ocupacoes.config(state="normal")
    entry_qtd_ocupacoes.delete(0, tk.END)
    entry_qtd_ocupacoes.insert(0, str(qtd_ocupacoes))
    entry_qtd_ocupacoes.config(state="disabled")

    entry_arquivo_gerado .config(state="normal")
    entry_arquivo_gerado .delete(0, tk.END)
    entry_arquivo_gerado .insert(0, arquivo_gerado)
    entry_arquivo_gerado .config(state="disabled")

    messagebox.showinfo("Sucesso", "Planilha gerada com sucesso!")

# Botão para chamar o programa createCityOccupationsSpreadsheet.py que gera a Planilha de Ocupações
def gerar_planilha(municipio, planilha_lotacao, pasta_planilha ):
    try:
        funGerarPlanilha = os.getcwd() + r"\src\modules\createCityOccupationsSpreadsheet.py"

        result = subprocess.run(
            ["python", funGerarPlanilha, municipio, planilha_lotacao, pasta_planilha],
            check=True,
            stdout=subprocess.PIPE,
            text=True,
            encoding='utf-8'  # Força a codificação para UTF-8
        )

        # Processa a saída do script executado
        lines = result.stdout.strip().splitlines()

        # Variáveis para armazenar resultados
        qtd_ocupacoes = None
        arquivo_gerado = None

         # Processa cada linha para encontrar as informações relevantes
        for line in lines:
            line = normalize_string(line)  # Normaliza a linha antes de verificar
            
            if line.startswith("Quantidade de Ocupacoes:"):
                qtd_ocupacoes = int(line.split(":", 1)[1].strip())
            elif line.startswith("Caminho do Arquivo Gerado:"):
                arquivo_gerado = line.split(":", 1)[1].strip()
            elif "Erro:" in line:
                raise ValueError(f"Erro detectado na saída do script: {line}")

        # Valida se os dados foram encontrados
        if not qtd_ocupacoes or not arquivo_gerado:
            raise ValueError("Erro: Não foi possível encontrar os dados na saída do script.")

        return qtd_ocupacoes, arquivo_gerado
    
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Erro no processo de geração da planilha", 
                             f"Ocorreu um erro no processo de geração da Planilha de Ocupações: {e}")
    except Exception as e:
        messagebox.showerror("Erro na geração da planilha.", f"Ocorreu um erro ao gerar a planilha: {e}")

# Função para normalizar strings
def normalize_string(s):
    return unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode('ascii')

# Função para selecionar a planilha de ocupações
def selecionar_arquivo(entry_arquivo_gerado, combo_municipio, entry_qtd_ocupacoes):
    municipio = combo_municipio.get().strip()
    base_path = os.getcwd() +  r"\lotacaoMun\planilhasOcupacoes"
    
    if not municipio:
        messagebox.showerror("Erro - Seleção de município", "Selecione um município antes de seleciona a Planilha de Ocupações.")
        return
    
    # Construir o caminho da pasta com base no município selecionado
    municipio_folder = os.path.join(base_path, f"mun{municipio}")

    # Verifica se a pasta existe
    if not os.path.exists(municipio_folder):
        messagebox.showerror("Erro - Pasta não encontrada",
            f"""A pasta com a Planilha de Ocupações para o município {municipio} ainda não foi criada: 
               \n{municipio_folder}.\n\n"""
               "Certifique-se que o processo 'Gerar Planilha de Ocupações seja executado antes."
        )  
        return
    else:
        start_path = municipio_folder

    full_path = filedialog.askopenfilename(
        initialdir = start_path,
        title="Selecione o arquivo de entrada",
        filetypes=[("Arquivos Excel XLSX", "*.xlsx"), ("Todos os Arquivos", "*.*")]
        )
    
    if full_path:
        # Pegar o nome e o caminho do arquivo
        file_path, file_name = os.path.split(full_path)
        
        entry_arquivo_gerado.config(state="normal")
        entry_arquivo_gerado.delete(0,tk.END)
        entry_arquivo_gerado.insert(0,full_path)
        entry_arquivo_gerado.config(state="disabled")

        # Limpar o campo "Quantidade de Ocupações"
        entry_qtd_ocupacoes.config(state="normal")
        entry_qtd_ocupacoes.delete(0, tk.END)

        try:
            # Abrir a planilha para contar as linhas (excluindo o cabeçalho)
            data = pd.read_excel(full_path)
            qtd_linhas = len(data)
            entry_qtd_ocupacoes.insert(0, str(qtd_linhas))
        except Exception as e:
            messagebox.showerror("Erro - Carga da planilha de ocupações", f"Erro ao carregar a planilha: {e}")

        entry_qtd_ocupacoes.config(state="disabled")

def carregar_ocupacoes(combo_municipio, pasta_planilha_ocupacoes):
    
    municipio = combo_municipio.get().strip()
    planilha_ocupacoes = os.path.basename(pasta_planilha_ocupacoes.get().strip())
    
    if not municipio:
        messagebox.showerror("Erro - Seleção de Município Obrigatório", "Município ainda não foi informado para efetuar a Carga das Ocupaçoes.")
        return
    
    if not planilha_ocupacoes:
        messagebox.showerror("Erro - Planilha de Ocupações não encontrada", "A Planilha de Ocupações ainda não foi gerada para efetuar a carga.")
        return
    
    resposta = messagebox.askquestion(
        "Confirmação",
        f"Antes de efetuar a carga das ocupações do município {municipio}, a Planilha de Ocupações, {planilha_ocupacoes}, "
         "precisa ter as inconsistências corrigidas.\n\n"
        "Você já validou a planilha e corrigiu as inconsistências?"
    )

    if resposta == 'yes':
        try:
     
            #   
            funCarregarOcupacoes = os.getcwd() + r"\src\modules\loadCityOccupations.py"
            caminho_arquivo = pasta_planilha_ocupacoes.get().strip()
            result =subprocess.run(
                ["python", funCarregarOcupacoes, caminho_arquivo],
                check=True,
                stdout=subprocess.PIPE,
                text=True
            )

            lines = result.stdout.strip().splitlines()  # Divide a saída em linhas
            for line in lines:
                if line.startswith("SUCCESS"):
                    output = line  # Pega apenas a linha correta
                    break

            if output.startswith("SUCCESS"):
                _, total_read, total_found, total_inserted = output.split("|")
                
                # Carregar log JSON
                log_file_path = os.path.join(os.path.dirname(caminho_arquivo), "occupation_log.json")
                now = datetime.datetime.now().strftime("%d/%m/%Y - %H:%M:%S")
                
                message = (
                    f"Ocupaçoes: Carga de Ocupações do município {municipio} - em {now}\n"
                    f"Status: Realizada com Sucesso!!!\n"
                    f"Total de Ocupações Lidas: {total_read}\n"
                    f"Total de Ocupações já Cadastradas: {total_found}\n"
                    f"Total de Ocupações Novas Cadastradas: {total_inserted}\n\n"
                    f"Log gravado em {log_file_path}"
                )
                
                messagebox.showinfo("Sucesso", message)

            elif output.startswith("ERROR"):
                _, descricao = output.split("|")
                messagebox.showerror("Erro", f"Houve um erro na carga de ocupações para o município {municipio}.\nErro - {descricao}")

        except subprocess.CalledProcessError as e:
            messagebox.showerror("Erro na carga de ocupações", f"Erro ao executar o processo de Carga da Ocupações: {e}")
        except Exception as e:
            messagebox.showerror("Erro inesperado", f"Ocorreu um erro inesperado: {e}")

