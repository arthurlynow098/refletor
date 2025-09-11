import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog
import os
import json
from openpyxl.utils import column_index_from_string

# --- CONFIGURAÇÃO DE MODELOS ---
# Adicione ou ajuste as regras para cada modelo de sinalizador aqui
MODEL_RULES = {
    "40": {"row_start": 1, "row_end": 20, "col_start": "A", "col_end": "AD", "use_l411": True},
    "47": {"row_start": 2, "row_end": 23, "col_start": "A", "col_end": "AG", "use_l411": False},
    "54": {"row_start": 2, "row_end": 25, "col_start": "A", "col_end": "AL", "use_l411": False},
    # Adicione outros modelos conforme necessário. Ex:
    # "61": {"row_start": 2, "row_end": 27, "col_start": "A", "col_end": "AP", "use_l411": False}
}

COLOR_MAP = {'AZ': 'blue', 'AB': 'yellow', 'BR': 'white', 'VM': 'red', 'RG': 'green'}
SVG_COLOR_MAP = {'blue': '#0070c0', 'yellow': '#ffc000', 'white': '#ffffff', 'red': '#ff0000', 'green': '#00b050'}

# --- FUNÇÕES DE INTERAÇÃO ---
def selecionar_arquivo_excel():
    """Abre uma janela gráfica para o usuário selecionar um arquivo Excel."""
    root = tk.Tk()
    root.withdraw()
    filetypes = (('Arquivos Excel', '*.xlsx *.xlsm'), ('Todos os arquivos', '*.*'))
    return filedialog.askopenfilename(title='Selecione o arquivo Excel com o layout', filetypes=filetypes)

def escolher_modelo():
    """Abre uma janela para o usuário digitar o modelo desejado."""
    root = tk.Tk()
    root.withdraw()
    modelos = ", ".join(MODEL_RULES.keys())
    modelo = simpledialog.askstring("Modelo do Sinalizador", f"Informe o modelo ({modelos}):")
    if modelo and modelo in MODEL_RULES:
        return modelo
    print(f"Modelo inválido ou não informado. Modelos válidos: {modelos}")
    return None

def encontrar_planilha_pol(caminho_arquivo):
    """Encontra a primeira planilha que CONTÉM 'POL' no nome."""
    try:
        xl = pd.ExcelFile(caminho_arquivo, engine='openpyxl')
        for nome_planilha in xl.sheet_names:
            if "POL" in nome_planilha.strip().upper():
                return nome_planilha
        print("ERRO: Nenhuma planilha com 'POL' no nome foi encontrada.")
        return None
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return None

# --- FUNÇÕES DE PROCESSAMENTO ---
def extract_module_info(cell_content):
    """
    Extrai informações de cor e quantidade APENAS de células que contêm um código de cor conhecido.
    Ignora números isolados como '61'.
    """
    if not isinstance(cell_content, str):
        return None
    
    content_upper = cell_content.upper().strip()
    
    found_color_code = None
    for code in COLOR_MAP.keys():
        if code in content_upper:
            found_color_code = code
            break
            
    if not found_color_code:
        return None

    color = COLOR_MAP[found_color_code]
    
    quantity = 1
    match_num = re.match(r'^(\d+)', content_upper)
    if match_num:
        quantity = int(match_num.group(1))

    return {
        'color': color,
        'quantity': quantity,
        'original_text': cell_content
    }

def find_module_cells(df, use_l411):
    """Encontra células que contêm padrões de módulos, respeitando a regra 'use_l411'."""
    module_cells = []
    
    for r in range(df.shape[0]):
        for c in range(df.shape[1]):
            cell_value = df.iat[r, c]
            if pd.isna(cell_value):
                continue
            
            if use_l411:
                if not isinstance(cell_value, str) or 'L411' not in cell_value.upper():
                    continue
            
            module_info = extract_module_info(cell_value)
            if module_info:
                module_cells.append({
                    'row': r,
                    'col': c,
                    'color': module_info['color'],
                    'quantity': module_info['quantity'],
                    'content': module_info['original_text']
                })
                print(f"Célula Válida Encontrada: '{cell_value}' -> {module_info['quantity']}x {module_info['color']}")
    
    return module_cells

def recortar_dataframe(df, modelo):
    """Recorta o DataFrame para a área de interesse definida nas regras do modelo."""
    regra = MODEL_RULES.get(modelo)
    if not regra:
        raise ValueError(f"Modelo '{modelo}' não está configurado!")
    
    row_start = regra["row_start"] - 1
    row_end = regra["row_end"]
    col_start = column_index_from_string(regra["col_start"]) - 1
    col_end = column_index_from_string(regra["col_end"])
    
    print(f"Recortando planilha para o modelo '{modelo}': Linhas {row_start+1}-{row_end}, Colunas {regra['col_start']}-{regra['col_end']}")
    return df.iloc[row_start:row_end, col_start:col_end], regra["use_l411"]

def processar_layout_excel(caminho_arquivo, nome_planilha, modelo):
    """Função principal que orquestra a leitura, recorte e extração de dados."""
    try:
        df = pd.read_excel(caminho_arquivo, sheet_name=nome_planilha, header=None, engine='openpyxl')
        df_recortado, use_l411 = recortar_dataframe(df, modelo)
        
        print("Procurando células com padrões de módulos na área recortada...")
        module_cells = find_module_cells(df_recortado, use_l411)
        
        if not module_cells:
            print("ERRO CRÍTICO: Nenhuma célula com padrão de módulo e cor foi encontrada na área definida.")
            return {}, {}
        
        module_cells_sorted = sorted(module_cells, key=lambda x: (x['row'], x['col']))
        
        layout_final = {}
        module_counter = 1
        
        for cell in module_cells_sorted:
            for i in range(cell['quantity']):
                # *** NOVA VERIFICAÇÃO DE LIMITE ***
                if module_counter > 60:
                    print("AVISO: Limite de 40 módulos atingido. A lista foi truncada.")
                    break
                
                layout_final[str(module_counter)] = cell['color']
                module_counter += 1
            
            # Se o limite foi atingido no loop interno, para o loop externo também
            if module_counter > 40:
                break
        
        print(f"\nProcessamento concluído. Total de módulos sequenciais gerados: {len(layout_final)}")
        return layout_final, module_cells_sorted
        
    except Exception as e:
        print(f"Ocorreu um erro inesperado ao processar o layout: {e}")
        import traceback
        traceback.print_exc()
        return None, None

# --- FUNÇÕES DE EXPORTAÇÃO (sem alterações) ---
def gerar_arquivo_txt(dados_ordenados, caminho_saida):
    try:
        with open(caminho_saida, 'w', encoding='utf-8') as f:
            f.write("--- Layout de Cores dos Módulos ---\n\n")
            for modulo, cor in dados_ordenados.items():
                f.write(f"Módulo {modulo}: {cor.capitalize()}\n")
        print(f"SUCESSO! Arquivo TXT salvo em: '{caminho_saida}'")
    except Exception as e:
        print(f"ERRO ao salvar o arquivo TXT: {e}")

def gerar_arquivo_svg(dados_ordenados, module_cells, caminho_saida):
    try:
        if not module_cells:
            print("AVISO: Nenhuma célula de módulo para gerar SVG.")
            return
            
        CELL_SIZE = 40
        PADDING = 20
        
        min_row = min(cell['row'] for cell in module_cells)
        max_row = max(cell['row'] for cell in module_cells)
        min_col = min(cell['col'] for cell in module_cells)
        max_col = max(cell['col'] for cell in module_cells)

        svg_width = (max_col - min_col + 1) * CELL_SIZE + 2 * PADDING
        svg_height = (max_row - min_row + 1) * CELL_SIZE + 2 * PADDING

        svg_content = f'<svg width="{svg_width}" height="{svg_height}" xmlns="http://www.w3.org/2000/svg" style="background-color:#f0f0f0;">\n'
        svg_content += f' 	<style>.text {{ font: bold {int(CELL_SIZE/2.5)}px sans-serif; text-anchor: middle; dominant-baseline: middle; fill: black; }} .text-white {{ fill: white; }}</style>\n'

        module_counter = 1
        for cell in module_cells:
            if module_counter > 40: # Garante que o SVG também não desenhe mais de 40
                break

            x = (cell['col'] - min_col) * CELL_SIZE + PADDING
            y = (cell['row'] - min_row) * CELL_SIZE + PADDING
            fill_color = SVG_COLOR_MAP.get(cell['color'], '#cccccc')
            text_class = "text-white" if cell['color'] in ['blue', 'red', 'green'] else "text"
            
            svg_content += f' 	<rect x="{x}" y="{y}" width="{CELL_SIZE}" height="{CELL_SIZE}" fill="{fill_color}" stroke="#333" stroke-width="1.5" rx="3"/>\n'
            
            # Ajusta a lógica para não ultrapassar 40 no texto do SVG
            num_to_add = cell['quantity']
            end_range = module_counter + num_to_add - 1
            if end_range > 40:
                end_range = 40
            
            if num_to_add == 1 or module_counter == end_range:
                 svg_content += f' 	<text x="{x + CELL_SIZE/2}" y="{y + CELL_SIZE/2}" class="text {text_class}">{module_counter}</text>\n'
            else:
                svg_content += f' 	<text x="{x + CELL_SIZE/2}" y="{y + CELL_SIZE/2}" class="text {text_class}">{module_counter}-{end_range}</text>\n'
            
            module_counter += num_to_add

        svg_content += '</svg>'
        
        with open(caminho_saida, 'w', encoding='utf-8') as f:
            f.write(svg_content)
        print(f"SUCESSO! Arquivo SVG salvo em: '{caminho_saida}'")
        
    except Exception as e:
        print(f"ERRO ao gerar o arquivo SVG: {e}")

# --- EXECUÇÃO PRINCIPAL ---
if __name__ == "__main__":
    arquivo_excel = selecionar_arquivo_excel()
    if not arquivo_excel:
        print("Nenhum arquivo selecionado. Encerrando.")
    else:
        modelo = escolher_modelo()
        if not modelo:
            input("\nEncerrando. Pressione Enter para fechar...")
            exit()

        planilha_alvo = encontrar_planilha_pol(arquivo_excel)
        if not planilha_alvo:
            input("\nEncerrando. Pressione Enter para fechar...")
            exit()

        print(f"\nProcessando arquivo: {os.path.basename(arquivo_excel)}")
        dados_layout, module_cells = processar_layout_excel(arquivo_excel, planilha_alvo, modelo)

        if dados_layout:
            print(f"\n--- Layout de Cores Extraído com Sucesso ({len(dados_layout)} módulos) ---")
            print(json.dumps(dados_layout, indent=4))

            nome_base_saida = os.path.splitext(os.path.basename(arquivo_excel))[0]
            caminho_diretorio_saida = os.path.dirname(arquivo_excel)
            
            caminho_saida_json = os.path.join(caminho_diretorio_saida, f"{nome_base_saida}_layout.json")
            caminho_saida_txt = os.path.join(caminho_diretorio_saida, f"{nome_base_saida}_layout.txt")
            caminho_saida_svg = os.path.join(caminho_diretorio_saida, f"{nome_base_saida}_layout.svg")

            with open(caminho_saida_json, 'w', encoding='utf-8') as f:
                json.dump(dados_layout, f, indent=4)
            print(f"\nSUCESSO! Layout JSON salvo em: '{caminho_saida_json}'")

            gerar_arquivo_txt(dados_layout, caminho_saida_txt)
            gerar_arquivo_svg(dados_layout, module_cells, caminho_saida_svg)
        else:
            print("\nFalha na extração do layout. Nenhum arquivo foi gerado.")

    input("\nPressione Enter para fechar...")