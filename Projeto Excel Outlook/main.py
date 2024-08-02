from datetime import *
from openpyxl import load_workbook as load_wb
import win32com.client as win32

# -=-=-=--=--=--=--=--=--=--=--=--=--=-=--=--Importar--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-

outlook = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')

inbox = outlook.GetDefaultFolder(6)  # Ele vai pegar o a caixa n 6 do outlook, no caso o inbox

emails = inbox.Items
emails = emails.Restrict("[MessageClass]='IPM.Note'")  # Filtro para pegar apenas emails.

data = datetime(2024, 8, 1)

data = data.strftime('%d/%m/%Y')

dataf = f"[ReceivedTime] >= '{data}'"

emails = emails.Restrict(dataf)
n_mensagens = emails.Count

print(f'Seu numero de emails no inbox: {n_mensagens}')

# -=-=-=--=--=--=--=--=--=--=--=--=--=--=--=--Exportar--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-

wb = load_wb(r'C:\Users\jvpic\OneDrive\Área de Trabalho\Trabalhos\Book3.xlsx')
ws = wb['Sheet1']

if ws['A2']:
    ws.insert_rows(2)

ws['A2'] = str(data) + ':'
ws['B2'] = str(n_mensagens)

wb.save(r'C:\Users\jvpic\OneDrive\Área de Trabalho\Trabalhos\Book3.xlsx')
