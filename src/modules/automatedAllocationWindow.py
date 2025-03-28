import os
import json
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from pathlib import Path
import subprocess
import traceback

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

def openServerAllocationWindow(root):
    """Configura a interface para Lotação de Servidores."""
    frame = tk.Frame(root, padx=20, pady=20, relief="groove", borderwidth=3)
    frame.pack(padx=10, pady=10, fill="both", expand=True)

    # Função para desabilitar os campos
    def desabilitar_campos():
        """Desativa todos os campos e botões da interface."""
        entry_planilha.config(state="disabled")
        btn_selecionar_planilha.config(state="disabled")
        btn_lotar_servidores.config(state="disabled")

    # Função para habilitar os campos
    def habilitar_campos(event=None):
        """Habilita os campos quando um município válido é selecionado."""
        municipio = combo_municipio.get().strip()
        if municipio in municipios_ES:
            entry_planilha.config(state="normal")
            btn_selecionar_planilha.config(state="normal")
            btn_lotar_servidores.config(state="normal")
        else:
            desabilitar_campos()

    # Função para limpar campos e restaurar a lista de municípios
    def limpar_campos():
        """Limpa os campos e redefine a lista de municípios no combobox."""
        entry_planilha.config(state="normal")
        entry_planilha.delete(0, tk.END)
        entry_planilha.config(state="disabled")

        if not combo_municipio.get().strip():
            combo_municipio['values'] = municipios_ES  
            combo_municipio.set("")
            desabilitar_campos()

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

    # Interface para seleção do município
    tk.Label(frame, text="Município", anchor="w").pack(fill="x")
    combo_municipio = ttk.Combobox(frame, values=municipios_ES, width=101)
    combo_municipio.pack(fill="x", pady=(0, 10))
    combo_municipio.bind('<KeyRelease>', filtrar_municipios)
    combo_municipio.bind('<<ComboboxSelected>>', habilitar_campos)

    # Campo para seleção da planilha
    tk.Label(frame, text="Planilha para Lotação de Servidores", anchor="w").pack(fill="x")
    frame_pasta = tk.Frame(frame)
    frame_pasta.pack(fill="x", pady=(0, 10))

    entry_planilha = tk.Entry(frame_pasta, state="disabled", width=104)
    entry_planilha.pack(side="left", anchor="e")

    btn_selecionar_planilha = tk.Button(frame_pasta, text="Selecionar", state="disabled", width=15, 
                            command=lambda: selecionar_planilha(entry_planilha, combo_municipio))
    btn_selecionar_planilha.pack(side="left", padx=5)

    # Botões principais
    frame_campos = tk.Frame(frame)
    frame_campos.pack(fill="x", pady=(0, 10))

    btn_lotar_servidores = tk.Button(frame_campos, text="Lotar Servidores", state="disabled", width=15,
                                     command=lambda: lotar_servidores(combo_municipio, entry_planilha))
    btn_lotar_servidores.grid(row=2, column=2, pady=(5, 0), sticky="e")

    botao_fechar = tk.Button(frame, text="Fechar", command=root.destroy, width=15)
    botao_fechar.pack(side="right", padx=5, pady=(5, 0))

    desabilitar_campos()

def selecionar_planilha(entry_planilha, combo_municipio):
    """Seleciona a planilha de lotação do município informado."""
    municipio = combo_municipio.get().strip()

    if not municipio:
        messagebox.showerror("Erro", "Selecione um município antes de escolher a planilha.")
        return
    
    base_path = os.path.join(os.getcwd(), "lotacaoMun", "planilhasLotacao")
    municipio_folder = os.path.join(base_path, f"mun{municipio}")

    if not os.path.exists(municipio_folder):
        messagebox.showerror("Erro", f"A pasta para {municipio} não existe: {municipio_folder}")
        return

    full_path = filedialog.askopenfilename(
        initialdir=municipio_folder,
        title="Selecione a Planilha para Lotação de Servidores",
        filetypes=[("Arquivos Excel XLSX", "*.xlsx"), ("Todos os Arquivos", "*.*")]
    )

    if full_path:
        entry_planilha.config(state="normal")
        entry_planilha.delete(0, tk.END)
        entry_planilha.insert(0, full_path)
        entry_planilha.config(state="disabled")

def lotar_servidores(combo_municipio, entry_planilha):
    """Executa o script de lotação de servidores."""
    municipio = combo_municipio.get().strip()
    caminho_arquivo = entry_planilha.get().strip()

    if not municipio or not caminho_arquivo:
        messagebox.showerror("Erro", "Município ou arquivo não informado.")
        return

    resposta = messagebox.askquestion(
        "Confirmação",
        f"Confirma a lotação de servidores para {municipio}?\n\n"
    )

    script_lotacao = os.path.join(os.getcwd(), "src", "modules", "automaticServerAllocation.py")
    if not os.path.exists(script_lotacao):
        messagebox.showerror("Erro", f"Script de lotação não encontrado: {script_lotacao}")
        return

    if resposta == 'yes':
        try:
            result = subprocess.run(
                ["python", script_lotacao, municipio, caminho_arquivo],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding='utf-8'
            )

            if result.returncode != 0:  # Se o subprocesso falhar, exibe erro
                print(f"Erro no subprocesso: {result.stderr.strip()}")
                messagebox.showerror("Erro", f"Erro ao executar lotação: {result.stderr.strip()}")
                return

            output = result.stdout.strip()

            if not output:  # Se não houver saída, evita erro na conversão
                output = "{}"
                print("⚠ Aviso: O subprocesso retornou saída vazia!")

            # Divide a saída em linhas e tenta pegar a última que seja um JSON válido
            output_lines = output.split("\n")
            json_lines = [line.strip() for line in output_lines if line.strip().startswith("{")]

            # Pega a última linha JSON válida
            json_string = json_lines[-1] if json_lines else "{}"

            try:
                output = json.loads(json_string)  # Converte para dicionário
            except json.JSONDecodeError as e:
                print(f"Erro ao decodificar JSON: {e}")
                output = {"status": "error", "mensagem": "Saída do subprocesso não é um JSON válido"}

            if output.get("status") == "success":
                log_path = output.get("log_path", "Caminho do log não informado")
                contadores = output.get("contadores", {})  # Dicionário de servidores por secretaria
                quantidade_total = output.get("quantidade", 0)

                # Formatar os dados por secretaria
                contadores_texto = "\n".join([f"{secretaria}: {qtd}" for secretaria, qtd in contadores.items()])

                # Exibir mensagem
                messagebox.showinfo(
                    "Sucesso",
                    f"Lotação concluída com sucesso para o município {municipio}.\n\n"
                    f"Servidores por secretaria:\n{contadores_texto}\n\n"
                    f"Total de servidores lotados: {quantidade_total}\n\n"
                    f"Log salvo em: {log_path}"
                )

        except Exception as e:
            messagebox.showerror("Erro", f"Erro inesperado: {traceback.format_exc()}")


