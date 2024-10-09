from pathlib import Path
import pandas as pd
import os
import warnings
warnings.simplefilter("ignore")

# Define o caminho da pasta onde estão os arquivos Excel
pasta_atual = Path(__file__).parent
# Inicializa um DataFrame vazio para armazenar todos os dados
df_unificado = pd.DataFrame()
# Define o caminho do arquivo de saída
arquivo_saida = pasta_atual / 'resultado' / 'arquivo_saida.xlsx'
# Função para lidar com erros ao processar os arquivos

def processa_arquivos_excel():
    for arquivo in os.listdir(pasta_atual):
        try:
            # Verifica se o arquivo é uma planilha Excel (.xlsx ou .xls)
            if arquivo.endswith('.xlsx') or arquivo.endswith('.xls'):
                caminho_arquivo = os.path.join(pasta_atual, arquivo)                
                # Tenta ler o arquivo Excel
                try:
                    df = pd.read_excel(caminho_arquivo)
                except Exception as e:
                    print(f"Erro ao ler o arquivo {arquivo}: {e}")
                    continue
                # Concatena o DataFrame atual ao DataFrame unificado
                global df_unificado
                df_unificado = pd.concat([df_unificado, df], ignore_index=True)                
        except Exception as e:
            print(f"Erro ao processar o arquivo {arquivo}: {e}")
            continue

# Função principal que executa o processo
def main():
    try:
        # Processa todos os arquivos Excel
        processa_arquivos_excel()
        # Identifica duplicatas
        duplicatas = df_unificado[df_unificado.duplicated(keep=False)]
        # Se houver duplicatas, exibe o número das linhas duplicadas
        if not duplicatas.empty:
            # Obtém os índices das linhas duplicadas
            linhas_duplicadas = duplicatas.index + 1  # Adiciona 1 para que as linhas sejam baseadas em 1
            print("Linhas duplicadas encontradas nos arquivos:")
            print(linhas_duplicadas.tolist())
        # Remove duplicatas do DataFrame unificado
        df_unificado.drop_duplicates(inplace=True)
        # Salva o DataFrame unificado em um arquivo Excel
        try:
            df_unificado.to_excel(arquivo_saida, index=False)
            print(f"Arquivo de saída salvo com sucesso em: {arquivo_saida}")
        except Exception as e:
            print(f"Erro ao salvar o arquivo Excel: {e}")
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")
# Executa a função principal
if __name__ == "__main__":
    main()