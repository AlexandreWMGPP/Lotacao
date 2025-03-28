import pandas as pd
import os
import io
import sys
from openpyxl import load_workbook

# Garante que stdout e stderr usem UTF-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# Função para carregar a planilha
def uploadFile(fullPath):
    # Verifica a extensão do arquivo
    extension = os.path.splitext(fullPath)[1].lower()
    
    try:
        if extension == '.csv':
            # Carrega o arquivo .csv
            dataAllocation = pd.read_csv(fullPath, sep=';', encoding='ISO-8859-1', skiprows=1)
        elif extension in ['.xlsx', '.xls']:
            # Carrega o arquivo Excel
            dataAllocation = pd.read_excel(fullPath, skiprows=1)
        else:
            raise ValueError("Extensão de arquivo não suportada. Use .csv, .xlsx ou .xls")
        
        if dataAllocation.empty:
            raise ValueError("O arquivo está vazio ou não contém dados.")

        ocupacao_col = next((col for col in dataAllocation.columns if "Ocupação" in col or "Cargo" in col), None)
        
        # Renomeia a coluna "Cargo" para "Ocupações"
        if ocupacao_col:
            dataAllocation = dataAllocation.rename(columns={ocupacao_col: 'Ocupações'})
        else:
            raise KeyError("A coluna de ocupação não foi encontrada no arquivo.")
        
        # **Remover espaços extras das ocupações**
        dataAllocation['Ocupações'] = dataAllocation['Ocupações'].fillna("").astype(str).str.strip()
        
        return dataAllocation
    except FileNotFoundError:
        raise FileNotFoundError(f"Erro: O arquivo não foi encontrado no caminho especificado: {fullPath}")
    except Exception as e:
        raise Exception(f"Erro ao carregar o arquivo: {e}")

# Função para criar uma nova pasta
def createCityOccupationsFolder(fullPath, city):

    try:
        # Verifica se "planilhasRecebidas" está no caminho
        if "planilhasLotacao" not in fullPath:
            raise ValueError("O caminho especificado não contém 'planilhasLotacao'. Verifique o fullPath.")

        basePath = os.path.abspath(os.path.join(fullPath, "..", "..", ".."))

        # Define o novo caminho com a estrutura desejada
        newPath = os.path.join(basePath, "planilhasOcupacoes", f"mun{city}")

        # Cria a pasta se ela não existir
        os.makedirs(newPath, exist_ok=True)

        return newPath
    except Exception as e:
        raise RuntimeError(f"Erro ao criar a pasta de ocupações: {e}")

# Gerar a planilha com as funções do Município
def createOccupationsSpreadsheet(dataAllocation, occupationsFolderPath, city):

    # Conta o número de ocorrências de cada ocupação
    occupationsCity = (
        dataAllocation['Ocupações']
        .value_counts()
        .reset_index()
    )

    # elimine linhas onde "Ocupações" está vazia:
    occupationsCity = dataAllocation.loc[dataAllocation['Ocupações'].str.strip() != ""]
    occupationsCity = occupationsCity['Ocupações'].value_counts().reset_index()
    # Renomeia as colunas corretamente
    occupationsCity.columns = ['Ocupações', 'Quantidade']

    # Ordena o DataFrame pelo nome da ocupação
    occupationsCity = occupationsCity.sort_values(by='Ocupações')
 
    # Salva a planilha no disco
    occupationsPath = os.path.join(occupationsFolderPath, f"ocupacoes_{city}.xlsx")
    occupationsCity.to_excel(occupationsPath, index=False, sheet_name="Ocupacoes")

     # Ajusta automaticamente o tamanho das colunas
    adjust_column_width(occupationsPath, "Ocupacoes")

    # Devolve a quantidade de Ocupações inseridas 
    return len(occupationsCity), occupationsPath

def adjust_column_width(file_path, sheet_name):
    """
    Ajusta automaticamente o tamanho das colunas em uma planilha do Excel.
    """
    try:
        # Carrega a planilha gerada
        workbook = load_workbook(file_path)
        sheet = workbook[sheet_name]

        # Ajusta o tamanho de cada coluna com base no conteúdo
        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter  # Obtém a letra da coluna
            for cell in column:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            # Ajusta o tamanho da coluna (adiciona um pequeno espaçamento)
            adjusted_width = max_length + 2
            sheet.column_dimensions[column_letter].width = adjusted_width

        # Salva as alterações na planilha
        workbook.save(file_path)
    except Exception as e:
        print(f"Erro ao ajustar tamanho das colunas ({sheet_name} em {file_path}): {e}")

# Programa principal
def main(city, spreadsheet, filePath):
    
    fullPath = os.path.normpath(os.path.join(filePath, spreadsheet))

    # Carregar a planilha
    dataAllocation = uploadFile(fullPath)

    if dataAllocation is not None:        
        occupationsFolderPath = createCityOccupationsFolder(fullPath, city)
        qtd_ocupacoes, occupationsPath = createOccupationsSpreadsheet(dataAllocation, occupationsFolderPath, city)    
        
        # Devolve a quantidade de Ocupações e o caminho do arquivo
        print(f"Quantidade de Ocupações: {qtd_ocupacoes}")
        print(f"Caminho do Arquivo Gerado: {occupationsPath}")

        # Devolve a quantidade de Ocupações e o caminho do arquivo
        return qtd_ocupacoes, occupationsPath
    else:
        raise ValueError("Não foi possível carregar os Dados de Lotação.")

if __name__ == "__main__":

    if len(sys.argv) != 4:
        print("Erro: Uso correto: python createCityOccupationsSpreadsheet.py <Município> <Arquivo> <Caminho>")
    else:
        city, spreadsheet, filePath = sys.argv[1], sys.argv[2], sys.argv[3]
        try:
            qtd, path = main(city, spreadsheet, filePath)
            # Saída formatada explicitamente para subprocess
            print(f"Quantidade de Ocupações: {qtd}")
            print(f"Caminho do Arquivo Gerado: {path}")
        except Exception as e:
            # Saída para erros
            print(f"Erro: {e}")
            sys.exit(1)  # Sinaliza erro para o subprocess


