import tkinter as tk
from tkinter import messagebox
import pandas as pd
from datetime import datetime

# Função para salvar os dados no Excel
def save_to_excel(data):
    # Definindo o caminho do arquivo Excel
    file_path = r'C:\Projects\formulario_visitas\visits_latam.xlsx'

    # Verifica se o arquivo Excel já existe
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
    except FileNotFoundError:
        # Se não existir, cria um novo DataFrame com os cabeçalhos
        df = pd.DataFrame(columns=[
            'Date_Visit', 'Day_Visit', 'Responsible_RGD', 'Code_CC', 
            'Contact_CC', 'Customer_Mining', 'Site', 'Activity', 
            'Summary', 'City', 'CC', 'Status_Visit', 'BL'])

    # Adiciona os dados à tabela
    df = df.append(data, ignore_index=True)

    # Salva no arquivo Excel
    df.to_excel(file_path, index=False, engine='openpyxl')

# Função chamada ao clicar no botão de "Salvar"
def submit_data():
    # Obtendo a data e o dia da semana atuais
    current_date = datetime.now()
    date_visit = current_date.strftime('%d/%m/%Y')  # Formato DD/MM/YYYY
    day_visit = current_date.strftime('%A')  # Nome do dia (ex: Monday, Tuesday, etc.)

    # Preenchendo os campos com valores fixos e os obtidos da data
    data = {
        'Date_Visit': date_visit,
        'Day_Visit': day_visit,
        'Responsible_RGD': entry_responsible_rgd.get(),
        'Code_CC': 'ZXBRC',  # Valor fixo
        'Contact_CC': entry_contact_cc.get(),
        'Customer_Mining': entry_customer_mining.get(),
        'Site': entry_site.get(),
        'Activity': entry_activity.get(),
        'Summary': entry_summary.get(),
        'City': entry_city.get(),
        'CC': entry_cc.get(),
        'Status_Visit': entry_status_visit.get(),  # Este campo será preenchido pelo usuário
        'BL': entry_bl.get()
    }

    # Validação para garantir que Status_Visit tenha valores válidos
    status_valid_values = ['Done', 'Schedule', 'Canceled']
    status_visit = data['Status_Visit']
    if status_visit not in status_valid_values:
        messagebox.showerror("Erro", f"O status '{status_visit}' não é válido. Deve ser um dos seguintes: {', '.join(status_valid_values)}.")
        return  # Não prossegue com a salvamento dos dados se o status for inválido

    # Converte os dados para o formato de DataFrame e salva no Excel
    save_to_excel(data)

    # Exibe uma mensagem de sucesso
    messagebox.showinfo("Sucesso", "Dados salvos com sucesso!")

    # Limpa os campos do formulário
    entry_responsible_rgd.delete(0, tk.END)
    entry_contact_cc.delete(0, tk.END)
    entry_customer_mining.delete(0, tk.END)
    entry_site.delete(0, tk.END)
    entry_activity.delete(0, tk.END)
    entry_summary.delete(0, tk.END)
    entry_city.delete(0, tk.END)
    entry_cc.delete(0, tk.END)
    entry_status_visit.delete(0, tk.END)
    entry_bl.delete(0, tk.END)

# Configuração da janela principal
window = tk.Tk()
window.title("Formulário de Inserção de Dados")

# Layout do formulário
tk.Label(window, text="Date Visit (DD/MM/YYYY)").grid(row=0, column=0)
tk.Label(window, text="Day Visit").grid(row=1, column=0)
tk.Label(window, text="Responsible RGD").grid(row=2, column=0)
tk.Label(window, text="Code CC").grid(row=3, column=0)
tk.Label(window, text="Contact CC").grid(row=4, column=0)
tk.Label(window, text="Customer Mining").grid(row=5, column=0)
tk.Label(window, text="Site").grid(row=6, column=0)
tk.Label(window, text="Activity").grid(row=7, column=0)
tk.Label(window, text="Summary").grid(row=8, column=0)
tk.Label(window, text="City").grid(row=9, column=0)
tk.Label(window, text="CC").grid(row=10, column=0)
tk.Label(window, text="Status Visit").grid(row=11, column=0)
tk.Label(window, text="BL").grid(row=12, column=0)

entry_responsible_rgd = tk.Entry(window)
entry_contact_cc = tk.Entry(window)
entry_customer_mining = tk.Entry(window)
entry_site = tk.Entry(window)
entry_activity = tk.Entry(window)
entry_summary = tk.Entry(window)
entry_city = tk.Entry(window)
entry_cc = tk.Entry(window)
entry_status_visit = tk.Entry(window)
entry_bl = tk.Entry(window)

# Posicionamento dos campos de entrada
entry_responsible_rgd.grid(row=2, column=1)
entry_contact_cc.grid(row=4, column=1)
entry_customer_mining.grid(row=5, column=1)
entry_site.grid(row=6, column=1)
entry_activity.grid(row=7, column=1)
entry_summary.grid(row=8, column=1)
entry_city.grid(row=9, column=1)
entry_cc.grid(row=10, column=1)
entry_status_visit.grid(row=11, column=1)
entry_bl.grid(row=12, column=1)

# Botão de envio
submit_button = tk.Button(window, text="Salvar Dados", command=submit_data)
submit_button.grid(row=13, column=1)

# Executa a interface gráfica
window.mainloop()
