import ctypes
import os
import socket
import json
import cv2

# Definir caminho base das imagens e do JSON de coordenadas
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_IMAGE_PATH = os.path.join(BASE_DIR, "..", "utils", "images")
COORD_FILE = ''

def get_real_screen_resolution():
    """Obtém a resolução real do monitor sem interferência do escalonamento de DPI."""
    class DevMode(ctypes.Structure):
        _fields_ = [
            ("dmDeviceName", ctypes.c_wchar * 32),
            ("dmSpecVersion", ctypes.c_ushort),
            ("dmDriverVersion", ctypes.c_ushort),
            ("dmSize", ctypes.c_ushort),
            ("dmDriverExtra", ctypes.c_ushort),
            ("dmFields", ctypes.c_ulong),
            ("dmOrientation", ctypes.c_short),
            ("dmPaperSize", ctypes.c_short),
            ("dmPaperLength", ctypes.c_short),
            ("dmPaperWidth", ctypes.c_short),
            ("dmScale", ctypes.c_short),
            ("dmCopies", ctypes.c_short),
            ("dmDefaultSource", ctypes.c_short),
            ("dmPrintQuality", ctypes.c_short),
            ("dmColor", ctypes.c_short),
            ("dmDuplex", ctypes.c_short),
            ("dmYResolution", ctypes.c_short),
            ("dmTTOption", ctypes.c_short),
            ("dmCollate", ctypes.c_short),
            ("dmFormName", ctypes.c_wchar * 32),
            ("dmLogPixels", ctypes.c_short),
            ("dmBitsPerPel", ctypes.c_ulong),
            ("dmPelsWidth", ctypes.c_ulong),
            ("dmPelsHeight", ctypes.c_ulong),
            ("dmDisplayFlags", ctypes.c_ulong),
            ("dmDisplayFrequency", ctypes.c_ulong),
            ("dmICMMethod", ctypes.c_ulong),
            ("dmICMIntent", ctypes.c_ulong),
            ("dmMediaType", ctypes.c_ulong),
            ("dmDitherType", ctypes.c_ulong),
            ("dmReserved1", ctypes.c_ulong),
            ("dmReserved2", ctypes.c_ulong),
            ("dmPanningWidth", ctypes.c_ulong),
            ("dmPanningHeight", ctypes.c_ulong),
        ]

    devmode = DevMode()
    #devmode.dmSize = ctypes.sizeof(DevMode)
    
    if ctypes.windll.user32.EnumDisplaySettingsW(None, -1, ctypes.byref(devmode)):
        return devmode.dmPelsWidth, devmode.dmPelsHeight
    else:
        return None, None

def get_machine_info():
    """Obtém o nome da máquina e a resolução da tela formatada corretamente."""
    hostname = socket.gethostname()  # Nome da máquina

    largura_real, altura_real = get_real_screen_resolution()
    #largura_real, altura_real = 2560,1600
    if not largura_real or not altura_real:
       raise ValueError("Não foi possível obter a resolução real da tela.")
    
    # Ajuste para máquinas específicas
    if hostname == "SEGER-AlienWare":
        hostname = f"MACHINE-ALIENWARE_{hostname}"
    elif hostname == "DESKTOP-ULU74A2":
        hostname = f"MACHINE-HP_{hostname}"
        
    return f"{hostname}_RT{largura_real}X{altura_real}"

def load_coordinates(coordinateType):
    
    if coordinateType == "Occupation":
        COORD_FILE = os.path.join(BASE_DIR, "..","config", "occupation_coordinates.json")
    else:
        COORD_FILE = os.path.join(BASE_DIR, "..", "config", "allocation_coordinates.json")

    """Carrega o JSON contendo as coordenadas de cada máquina."""
    try:
        with open(COORD_FILE, "r", encoding="utf-8") as file:
            return json.load(file)
    except FileNotFoundError:
        print(f"⚠️ Arquivo de coordenadas não encontrado: {COORD_FILE}")
        return {}
    except json.JSONDecodeError as e:
        print(f"⚠️ Erro ao carregar JSON: {e}")
        return {}

def get_positions(coordinateType):
    """Obtém as coordenadas para a máquina atual."""
    machine_id = get_machine_info()  # Obtém o nome formatado
    coordinates = load_coordinates(coordinateType)  # Carrega o JSON

    return coordinates.get(machine_id, {})  # Retorna {} se a máquina não for encontrada

def get_image_path(image_name):
    """Retorna o caminho correto da imagem para a máquina atual."""
    machine_name = get_machine_info()
    machine_image_path = os.path.join(BASE_IMAGE_PATH, machine_name, image_name)

    if os.path.exists(machine_image_path):
        return machine_image_path
    else:
        print(f"⚠️ Imagem '{image_name}' não encontrada para {machine_name}. Usando 'default'.")
        return os.path.join(BASE_IMAGE_PATH, "default", image_name)

def get_windows_scale_factor():
    """ Obtém o fator de escala do Windows """
    hdc = ctypes.windll.user32.GetDC(0)
    dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)  # 88 = LOGPIXELSX
    ctypes.windll.user32.ReleaseDC(0, hdc)
    return dpi / 96.0

def resize_image(image_path, scale_factor):
    """ Redimensiona a imagem baseada na escala do Windows """
    img = cv2.imread(image_path)
    if img is None:
        print(f"Erro: Imagem não encontrada em {image_path}")
        return None
    
    width = int(img.shape[1] * scale_factor)
    height = int(img.shape[0] * scale_factor)
    resized = cv2.resize(img, (width, height), interpolation=cv2.INTER_LINEAR)
    
    temp_image_path = "resized_temp.png"
    cv2.imwrite(temp_image_path, resized)
    return temp_image_path

def get_popup_coordinate(machineId, popup_type):
    # Mapeamento de resoluções para coordenadas
    COORDENADAS_POPUP = {
    "MACHINE-HP_DESKTOP-ULU74A2_RT1366X768": {"ocupacao": (311, 269, 1000, 690, 45), "unidade": (311, 307, 1000, 585, 25)},
    "MACHINE-ALIENWARE_SEGER-AlienWare_RT2560X1600": {"ocupacao": (720, 515, 1750, 1000, 40), "unidade": (717, 543, 1750, 1000, 55)},
    "MACHINE-ALIENWARE_SEGER-AlienWare_RT1920X1080": {"ocupacao": (590, 350, 1270, 700, 45), "unidade": (590, 357, 1260, 780, 55)}
    }

    if machineId not in COORDENADAS_POPUP:
        return None
    
    # Obtém as coordenadas corretas para a resolução atual
    x1, y1, x2, y2, altura_linha_fixa = COORDENADAS_POPUP[machineId][popup_type]

    return x1, y1, x2, y2, altura_linha_fixa


