import pandas as pd
import win32com.client as win32  

outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)
clientes = pd.read_excel('./clientes.xlsx')

for index, cliente in clientes.iterrows():
    nome = cliente['nome']
    endereco_email = cliente['email']
    
    email.To = f"{nome} <{endereco_email}>"
email.Subject = " Coleção Nova!!"
email.HTMLBody =f"""
<p>Lançamento Coleção Primavera Verão</p>
<p> Muitas novidades em Bijuterias e Roupas Femininas!</p>
<p> </p>
"""

anexo="C:/xampp/htdocs/python/automatização de tarefas/email marketing/"
email.Attachments.Add(anexo)
email.Send()
print("Enviado")