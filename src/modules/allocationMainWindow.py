import tkinter as tk
import os
from PIL import Image, ImageTk # necessário para trabalhar com imagens
from machine_utils import get_image_path
import occupationWindow  # Importa o módulo occupationWindow
import automatedAllocationWindow # Importa a funçao de abertura da tela de Lotação de Servidores

def centralizar_janela(janela, largura, altura):
    """Centraliza a janela em relação à tela principal."""
    root.update_idletasks()
    root_width = root.winfo_width()
    root_height = root.winfo_height()
    root_x = root.winfo_x()
    root_y = root.winfo_y()
    pos_x = root_x + (root_width - largura) // 2
    pos_y = root_y + (root_height - altura) // 2
    janela.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

def abrir_tela_carga_ocupacoes():
    modal = tk.Toplevel(root)
    modal.title("Carga de Ocupações dos Municípios")
    centralizar_janela(modal, 800, 570)
    modal.transient(root)
    modal.grab_set()
    occupationWindow.iniciar_tela(modal)

def chamar_lotacao_servidores():
    modal = tk.Toplevel(root)
    modal.title("Lotação Automatizada de Servidores")
    centralizar_janela(modal, 800, 300)
    modal.transient(root)
    modal.grab_set()
    automatedAllocationWindow.openServerAllocationWindow(modal)

def carregar_logo():
    """Tenta carregar a logo e exibir na tela"""
    logo_path = get_image_path("Logo_GPP_Azul_Hor.png")
    if not os.path.exists(logo_path):
        print("Aviso: Imagem da logo não encontrada.")
        return None
    try:
        logo_image = Image.open(logo_path)
        logo_image = logo_image.resize((400, 150), Image.Resampling.LANCZOS)
        return ImageTk.PhotoImage(logo_image)
    except Exception as e:
        print(f"Erro ao carregar a logo: {e}")
        return None

def criar_menu():
    """Cria a barra de menus da aplicação."""
    menu_bar = tk.Menu(root)
    menu_opcoes = tk.Menu(menu_bar, tearoff=0)
    menu_opcoes.add_command(label="Carga de Ocupações", command=abrir_tela_carga_ocupacoes)
    menu_opcoes.add_command(label="Lotação de Servidores", command=chamar_lotacao_servidores)
    menu_opcoes.add_separator()
    menu_opcoes.add_command(label="Sair", command=root.quit)
    menu_bar.add_cascade(label="Opções", menu=menu_opcoes)
    root.config(menu=menu_bar)

def criar_interface():
    """Cria a interface gráfica principal."""
    global frame_principal
    frame_principal = tk.Frame(root, padx=20, pady=20, relief="groove", borderwidth=3)
    frame_principal.pack(fill="both", expand=True, padx=10, pady=10)
    tk.Label(frame_principal, text="Lotação Automatizada de Servidores", font=("NEC", 20), fg="#1670BC", pady=25).pack()
    logo_photo = carregar_logo()
    if logo_photo:
        logo_label = tk.Label(frame_principal, image=logo_photo)
        logo_label.image = logo_photo
        logo_label.place(relx=0.5, rely=0.5, anchor="center")
    tk.Label(frame_principal, text="Sistema Automatizado de Lotação de Servidores - Versão 1.2", font=("Arial", 10)).pack(side="bottom", pady=10)

# Inicializa a janela principal
root = tk.Tk()
root.title("Sistema Automatizado de Lotação de Servidores - Tela Principal")
root.geometry("800x600")
root.state('zoomed')

criar_menu()
criar_interface()

root.mainloop()
