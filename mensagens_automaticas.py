#Fazendo as importações
import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
from os import path
import pywhatkit as pt
import pandas as pd
import time , keyboard, openpyxl
from datetime import datetime






lista = []
#Função para registrar em lista
def registro():
    contato = caixa.get()
    telefone = caixa_numero.get()
    mensage = caixa_mensagem.get()
    dados_ = (contato, telefone, mensage)
    lista.append(dados_)

    tabela.insert('', tk.END , values= dados_)

    caixa(0, tk.END)
    caixa_numero(0, tk.END)
    caixa_mensagem(0, tk.END)

def edição():
    selecionado = tabela.selection()
    if selecionado:
        indice = tabela.index(selecionado)
        lista.pop(indice)
        tabela.delete(selecionado)
    

# Salvando tudo antes de fechar
def salvar_todos_dados():
    caminho = path.join(path.dirname(path.abspath(__file__)), 'planilha.xlsx')
    planilha = openpyxl.Workbook()
    planilha_ativa = planilha.active
    planilha_ativa.append(['Nome do Contato', 'Número do Contato', 'Mensagem'])
    for dados in lista:
        planilha_ativa.append(dados)
    planilha.save(caminho)

    contato = caixa.get()
    telefone = caixa_numero.get()
    mensage = caixa_mensagem.get()
    dados_ = (contato, telefone, mensage)
    lista.append(dados_)

    tabela.insert('', tk.END , values= dados_)

    caixa.delete(0, tk.END)
    caixa_numero.delete(0, tk.END)
    caixa_mensagem.delete(0, tk.END)

#Função para enviar as mensagens

def envio():

    caminho = path.join(path.dirname(path.abspath(__file__)), 'planilha.xlsx')
    contatos = pd.read_excel(caminho)
    if contatos.empty:
        tempo.configure(text = 'Nenhum dado cadastrado')
    else:
        tempo.configure(text = 'Aguarde 2 minutos')
        time.sleep(5)
        for i , mensagem in enumerate(contatos['Mensagem']):
            numero = contatos.loc[i, 'Número do Contato']
            pt.sendwhatmsg(f'+{numero}', f'{mensagem}', datetime.now().hour ,datetime.now().minute + 1)
            time.sleep(10)
            keyboard.press_and_release('ctrl + w')
    



janela = ctk.CTk()
#Configurações da janela
janela.title('App de mensagens')
janela.geometry('700x500')


#Entrys para receber o texto

caixa = ctk.CTkEntry(janela,width=300, placeholder_text='Insira o nome do contato', fg_color='white', bg_color='transparent',border_color='gray',text_color='black')
caixa.place(x = 0, y= 10)
caixa_numero = ctk.CTkEntry(janela,width=300, placeholder_text='Insira o número do contato',fg_color='white',bg_color='transparent',border_color='gray',text_color='black')
caixa_numero.place(x=0, y=50)
caixa_mensagem = ctk.CTkEntry(janela,width=300 ,placeholder_text='Insira a mensagem para contato',fg_color='white',bg_color='transparent',border_color='gray', text_color='black' )
caixa_mensagem.place(x=0 , y=90)

#Botão que salva os dados inseridos na caixa de entrada
btn = ctk.CTkButton(janela, text='Registrar', width=100, command= salvar_todos_dados).place(x= 310, y=10)
btn2 = ctk.CTkButton(janela,text = 'Enviar mensagens',width=100, command= envio).place(x = 420, y = 10)
btn3 = ctk.CTkButton(janela, text='Excluir linha', command= edição, width=100 ).place(x= 310, y = 50)

#Caixas de Texto onde serão Armazenados os dados dos campos
frame1 = ctk.CTkFrame(master= janela, width=650, height= 350, fg_color='gray').place(x = 10, y = 130 )
tabela = ttk.Treeview(janela, columns= ('Nomes' , 'Números' , 'Mensagens'), show='headings',)
tabela.heading('Nomes', text='Nome do Contato')
tabela.heading('Números', text='Número do Contato')
tabela.heading('Mensagens', text='Mensagem')
tabela.place(x = 30, y =160 )
tempo = ctk.CTkLabel(janela, text= "")
tempo.place(x = 420, y = 50)

lbl = ctk.CTkLabel(janela, text= "Confirme se os numeros estão no formato: 5500900000000 ", bg_color='gray', text_color='black').place(x = 30, y = 400)

#puxando do banco de dados
def carregar_dados():
    caminho = path.join(path.dirname(path.abspath(__file__)), 'planilha.xlsx')
    if path.exists(caminho):
        planilha = openpyxl.load_workbook(caminho)
        planilha_ativa = planilha.active
        for row in planilha_ativa.iter_rows(min_row=2, values_only=True):
            lista.append(row)
            tabela.insert('', tk.END, values=row)


carregar_dados()

def fechando():
    salvar_todos_dados()
    janela.destroy()

janela.protocol('WM_DELETE_WINDOW', fechando)
        
janela.mainloop()