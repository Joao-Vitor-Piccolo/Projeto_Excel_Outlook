from datetime import datetime
import win32com.client as win32

# -=-=-=--=--=--=--=--=--=--=--=--=--=-=--=--Importar--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-

outlook = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')

inbox = outlook.GetDefaultFolder(6)  # Ele vai pegar o a caixa n 6 do outlook, no caso o inbox

emails = inbox.Items
emails = emails.Restrict("[MessageClass]='IPM.Note'")  # Filtro para pegar apenas emails.

data = datetime(2024, 8, 1)
data = data.strftime('%d/%m/%Y')
data = f"[ReceivedTime] >= '{data}'"

emails = emails.Restrict(data)
n_mensagens = emails.Count

print(f'Seu numero de emails no inbox: {n_mensagens}')

# -=-=-=--=--=--=--=--=--=--=--=--=--=--=--=--Exportar--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-

