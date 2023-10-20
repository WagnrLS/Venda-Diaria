import pandas as pd
from openpyxl import Workbook
import tkinter as tk
from tkinter import ttk

# Função para gerar o relatório
def generate_report():
    superior = combo_superior.get()  # Obter o valor Superior selecionado no ComboBox

    # Filtrar o DataFrame de vendedores pelo Superior selecionado
    vendedores_df = pd.read_csv('VENDEDOR.CSV', sep=';')

    # Filtrar o DataFrame de pedidos pelo Superior selecionado
    filtered_pedidos_df = merged_df[merged_df['Superior'] == superior]

    # Calcular a soma dos valores da coluna Quant CX agrupados por Nome do Produto
    result = filtered_pedidos_df.groupby(['Nome', 'Vendedor'])['Quant CX'].sum().unstack(fill_value=0)

    # Criar um novo arquivo Excel
    with pd.ExcelWriter(f'Relatorio_Superior_{superior}.xlsx', engine='openpyxl') as writer:
        result.to_excel(writer, sheet_name='Soma_Pedidos')

    workbook = writer.book
    worksheet = writer.sheets['Soma_Pedidos']

    # Defina o nome da primeira coluna como "Nome"
    worksheet['A1'] = 'Nome'

    workbook.save(f'Relatorio_Superior_{superior}.xlsx')

# Carregar o arquivo TERC.CSV
terc_df = pd.read_csv('TERC.CSV', sep=';')

# Carregar o arquivo PEDIDOS.CSV
pedidos_df = pd.read_csv('PEDIDOS.CSV', sep=';')

# Mesclar os DataFrames usando o código como chave
merged_df = pedidos_df.merge(terc_df, left_on='Cod Red', right_on='Codigo', how='left')

# Carregar o arquivo VENDEDOR.CSV
vendedores_df = pd.read_csv('VENDEDOR.CSV', sep=';')

# Criar uma janela tkinter
root = tk.Tk()
root.title("Relatório por Superior")

# Label explicativo
label = tk.Label(root, text="Selecione o Superior:")
label.pack()

# ComboBox (Combobox) para seleção do Superior
superiores = vendedores_df['Superior'].unique()
combo_superior = ttk.Combobox(root, values=superiores)
combo_superior.pack()

# Botão para gerar o relatório
generate_button = tk.Button(root, text="Gerar Relatório", command=generate_report)
generate_button.pack()

# Iniciar o loop principal da interface gráfica
root.mainloop()
