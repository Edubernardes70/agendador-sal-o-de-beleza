from tkinter import*
from tkinter import messagebox
from tkinter import ttk
import tkinter as tk
from PIL import Image, ImageTk
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
from tkinter import messagebox, filedialog  # Importe filedialog aqui


janela = Tk()
janela.attributes('-fullscreen', True)
janela.title('AGENDAMENTO SALÃO')
#--------------------------------------------------------------------------------------
cliente=Label(text="Cliente", font="Arial 12 bold", fg= "#000000")
cliente.grid(row=2, column=0, stick="W")
campoDigitadoCliente = Entry(font="Arial 12")
campoDigitadoCliente.grid(row=2, column=1, stick= "W")

data=Label(text='Data Agendamento', font="Arial 12 bold", fg="#000000")
data.grid(row=2, column=2, stick="W")
campoDigitadoData = Entry (font="Arial 12")
campoDigitadoData.grid(row=2, column=3, stick="W")

procedimento=Label(text="Procedimento", font="Arial 12 bold", fg="#000000")
procedimento.grid(row=2, column=4, stick="W")
campoDigitadoprocedimento = Entry (font=" Arial 12")
campoDigitadoprocedimento.grid (row=2, column=5, stick= "W")

profissional = Label (text= "Profissional", font="Arial 12 bold", fg="#000000")
profissional.grid (row=2, column=6, stick = "W")
campoDigitadoprofissional = Entry (font="Arial 12")
campoDigitadoprofissional.grid (row=2, column=7, stick="W")

valor = Label (text= "Valor", font="Arial 12 bold", fg="#000000")
valor.grid(row=2, column= 8, stick= "W")
campoDigitadovalor = Entry (font="Arial 12")
campoDigitadovalor.grid(row=2, column=9, stick="W")

#--------------------------------------------------------------------------------------
#BOTÕES TELA
def addItemTela():
    #ERRO AO NÃO ADICIONAR ALGUM ITEM OBRIGATÓRIO
    if str(campoDigitadoCliente.get())=="":
        janela2 = Toplevel()
        janela2.title("ERRO!")
        janela2.geometry("200x70")
        largura_janela = 200
        altura_janela = 70
        largura_tela = janela2.winfo_screenwidth()
        altura_tela = janela2.winfo_screenheight()
        x = (largura_tela - largura_janela) // 2
        y = (altura_tela - altura_janela) // 2
        janela2.geometry(f"{largura_janela}x{altura_janela}+{x}+{y}")
        janela2.config(bg="#C0C0C0")
        lbl_mensagem = Label(janela2, text="ATENÇÃO!!!\n Informe o nome do (a) cliente", font="Arial 10 bold", fg="#1C1C1C",bg="#C0C0C0")
        lbl_mensagem.pack()

        botao_fechar = Button(janela2, text="Fechar", command=janela2.destroy)
        botao_fechar.pack()

    if str(campoDigitadovalor.get()) == "":
            janela2 = Toplevel()
            janela2.title("ERRO!")
            janela2.geometry("200x70")
            largura_janela = 200
            altura_janela = 70
            largura_tela = janela2.winfo_screenwidth()
            altura_tela = janela2.winfo_screenheight()
            x = (largura_tela - largura_janela) // 2
            y = (altura_tela - altura_janela) // 2
            janela2.geometry(f"{largura_janela}x{altura_janela}+{x}+{y}")
            janela2.config(bg="#C0C0C0")
            lbl_mensagem = Label(janela2, text="ATENÇÃO!!!\n Informe o valor", font="Arial 10 bold",
                                 fg="#1C1C1C", bg="#C0C0C0")
            lbl_mensagem.pack()

            botao_fechar = Button(janela2, text="Fechar", command=janela2.destroy)
            botao_fechar.pack()

    else:


    #INSERÇÃO DE DADOS
        dados.insert("", "end",
                     values=(str(campoDigitadoCliente.get()),
                             str(campoDigitadoData.get()),
                             str(campoDigitadoprocedimento.get()),
                             str(campoDigitadoprofissional.get()),
                             str(campoDigitadovalor.get())))
        #Apaga dados depois cadastro
        campoDigitadoData.delete(0,"end")
        campoDigitadoprocedimento.delete(0, "end")
        campoDigitadoprofissional.delete(0, "end")
        campoDigitadovalor.delete(0, "end")
        campoDigitadoCliente.delete(0, "end")


app_img_cadastra = Image.open('D:\\PYTHON\\Cadastro_salao\\add.png')
app_img_cadastra = app_img_cadastra.resize((35,35))
app_img_cadastro = ImageTk.PhotoImage(app_img_cadastra)
app_cadastro = Button (janela, image = app_img_cadastro, text=" CADASTRAR",
                      width=150,height=45, compound=LEFT, relief= "raised",
                      font=('Arial 12 bold'), bg="#A9A9A9",  fg='#000000', command=addItemTela)
app_cadastro.place(x=10, y=250)

#-------------------------------------------------------------------------
#--------------------------------------------------------------------------
def deletar_item_tela():
    itemSelecionado = dados.selection()
    for item in itemSelecionado:
        dados.delete(item)

app_img_deleta = Image.open('D:\\PYTHON\\Cadastro_salao\\del.png')
app_img_deleta = app_img_deleta.resize((35,35))
app_img_deletando = ImageTk.PhotoImage(app_img_deleta)
app_deletando = Button(janela, image = app_img_deletando, text=" DELETAR",
                      width=150,height=45, compound=LEFT, relief= "raised",
                      font=('Arial 12 bold'), bg="#A9A9A9",  fg='#000000', command=deletar_item_tela)
app_deletando.place(x=10, y=310)

#---------------------------------------------------------------------------------------
def alterarItemTela():
    itemSelecionado = dados.selection()[0]
    dados.item(itemSelecionado,
               values=(str(campoDigitadoCliente.get()),
                       str(campoDigitadoData.get()),
                       str(campoDigitadoprocedimento.get()),
                       str(campoDigitadoprofissional.get()),
                       str(campoDigitadovalor.get())))

app_img_altera= Image.open('D:\\PYTHON\\Cadastro_salao\\altera.png')
app_img_altera = app_img_altera.resize((35,35))
app_img_alterado = ImageTk.PhotoImage(app_img_altera)
app_img_alteras = Button(janela, image = app_img_alterado, text="  ALTERAR",
                      width=150,height=45, compound=LEFT, relief= "raised",
                      font=('Arial 12 bold'), bg="#A9A9A9",  fg='#000000',command=alterarItemTela)
app_img_alteras.place(x=180, y=250)
#-----------------------------------------------------------------------------------------------------
def exporta_excel():
    file_path = filedialog.asksaveasfilename(defaultextension="xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])
    if file_path:
        #Cria arquivo do excel
        workbook = Workbook()
        planilha = workbook.active
        dados_dataframe = pd.DataFrame(columns=["Cliente", "Data", "Procedimento", "Profissional", "Valor"])
        dados_dataframe.loc[0]=["", "", "", "", ""]
        #Começa preencher o arquivo com os dados da tela a partir da segunda linha
        row_index =1
        for item in dados.get_children():
            item_values=dados.item(item, "values")
            dados_dataframe.loc[row_index]=item_values
            row_index +=1
        #escreve na planilha
        for row in dataframe_to_rows(dados_dataframe,index=False,header=True):
            planilha.append(row)

        #salva no local desejado
        workbook.save(file_path)
        messagebox.showinfo("Exportar para excel", "Dados exportados com sucesso!")

app_img_salvar= Image.open('D:\\PYTHON\\Cadastro_salao\\excel.png')
app_img_salvar = app_img_salvar.resize((35,35))
app_img_salva = ImageTk.PhotoImage(app_img_salvar)
app_img_save = Button(janela, image = app_img_salva, text="  SALVAR",
                      width=150,height=45, compound=LEFT, relief= "raised",
                      font=('Arial 12 bold'), bg="#A9A9A9",  fg='#000000',command=exporta_excel)
app_img_save.place(x=180, y=310)
#--------------------------------------------------------------------------------------------------
def carregar_excel():
    tipo_de_arquivo = [("Excel files", "*.xlsx; *xls")]
    nome_do_arquivo = filedialog.askopenfilename(title="Selecione o arquivo", filetypes=tipo_de_arquivo)
    if nome_do_arquivo:
            try:
                # Lê o arquivo Excel na biblioteca pandas e armazena o conteúdo em um DataFrame
                df = pd.read_excel(nome_do_arquivo)
                # Limpa a Treeview, removendo todas as linhas existentes
                for item in dados.get_children():
                    dados.delete(item)
                # Preenche a Treeview com os dados do DataFrame

                for _, row in df.iterrows():
                    dados.insert("", "end", values=row.tolist())
            except Exception as e:
                messagebox.showerror("Erro",f"Não foi possivel abrir o arquivo {e}" )
            # Exibe uma mensagem de erro se não for possível abrir o arquivo

app_img_abrir= Image.open('D:\\PYTHON\\Cadastro_salao\\abrir.png')
app_img_abrir = app_img_abrir.resize((35,35))
app_img_abre = ImageTk.PhotoImage(app_img_abrir)
app_img_abr= Button(janela, image = app_img_abre, text="  ABRIR",
                      width=150,height=45, compound=LEFT, relief= "raised",
                      font=('Arial 12 bold'), bg="#A9A9A9",  fg='#000000',command=carregar_excel)
app_img_abr.place(x=350, y=250)

#-----------------------------------------------------------------------------------------------------
def sairtela():
    janela.destroy()
app_img_sair = Image.open('D:\\PYTHON\\Cadastro_salao\\sair.png')
app_img_sair = app_img_sair.resize((35,35))
app_img_sai = ImageTk.PhotoImage(app_img_sair)
appsai= Button(janela, image = app_img_sai, text="  SAIR",
                      width=150,height=45, compound=LEFT, relief= "raised",
                      font=('Arial 12 bold'), bg="#A9A9A9",  fg='#000000',command=sairtela)
appsai.place(x=350, y=310)


#----------------------------------------------------------------------------------------
dados = ttk.Treeview(janela, column=(1,2,3,4,5), show="headings")
dados.column("1", anchor=CENTER)
dados.heading("1", text="Cliente")

dados.column("2", anchor=CENTER)
dados.heading("2", text="Data")

dados.column("3", anchor=CENTER)
dados.heading("3", text="Procedimento")

dados.column("4", anchor=CENTER)
dados.heading("4", text="Profissional")

dados.column("5", anchor=CENTER)
dados.heading("5", text="Valor")

#-------------------------------------------------------------------------------------
#TELA
dados.grid(row=3, column=0, columnspan=10, stick="NSEW")
#ESTILO
estilo = ttk.Style()
estilo.theme_use("alt")
estilo.configure(".", font="Arial 12")
#-------------------------------------------------------------------------------------
#CRÉDITOS CRIADOR
label_criador = Label(janela, width=19, height= 4, text='Desenvolvido por:\n @EduardoBernardes ', font=('Arial 9'), fg='#000000')
label_criador.place(x= 50, y=640)

#IMAGEM SALÃO

imagem_salao = Image.open('D:\\PYTHON\\Cadastro_salao\\salao.png')
imagem_salao = imagem_salao.resize((550,440))
imagem_tk = ImageTk.PhotoImage(imagem_salao)
imagem_label = Label(image=imagem_tk)
imagem_label.image = imagem_tk
imagem_label.place(x=700, y= 250)

#LABEL DADOS
label_salao = Label(janela, width=25, height=5, text="SALÃO ALIZA\n Funcionamento : Seg. a Sex\n das 8h ás 19h\nTel.:(35) 98888-8888",font=("Arial 18"), fg=	"#000000")
label_salao.place (x=50, y=450)

janela.mainloop()
