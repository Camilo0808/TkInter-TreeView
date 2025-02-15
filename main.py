import tkinter
import tkinter.messagebox, tkinter.ttk
import openpyxl

import pyodbc
import pandas as pd
import pyautogui as pg
import os
from dotenv import load_dotenv
load_dotenv()


global conexao, cursor

driver = os.getenv('DRIVER')
server = os.getenv('SERVER')
database = os.getenv('DATABASE')
uid = os.getenv('UID')
pwd = os.getenv('KEY')
# print(driver, server, database, uid, pwd)
print(f"Driver={driver};Server={server};Database={database};MultipleActiveResultsets=True;Uid={uid};Pwd={pwd}")
conexao = pyodbc.connect(f"Driver={driver};Server={server};Database={database};MultipleActiveResultsets=True;Uid={uid};Pwd={pwd}")
cursor = conexao.cursor()

# pg.alert('Conectado')

janela_visualizar = tkinter.Tk()
janela_visualizar.title("Cadastros")
# janela_visualizar.geometry("1350x660")

# Obter a resolução da tela
largura_tela = janela_visualizar.winfo_screenwidth()
altura_tela = janela_visualizar.winfo_screenheight()

# Definir o tamanho da janela para o maior possível, mas com um limite de 1350x660
# Ajusta para a resolução da tela (com margens de 10px)
janela_visualizar.geometry(f"{largura_tela - 10}x{altura_tela - 10}")

janela_visualizar.resizable(True, True)
style = tkinter.ttk.Style()
style.theme_use('default')


def query_database():

    global count
    cursor.execute("SELECT * FROM Cadastros")
    records = cursor.fetchall()

    count = 0

    for record in records:
        if count % 2 == 0:                                                                           
            my_tree.insert(
                parent='', 
                index='end',
                iid=count, 
                text='',
                tags=('evenrow'),
                values=(
                    record[0],
                    record[1],
                    record[2],
                    record[3],
                    record[4],
                    record[5],
                    record[6],
                    record[7],
                    record[8],
                    record[9],
                    record[10],
                    record[11],
                    record[12],
                    record[13],
                    record[14],
                    record[15],
                    record[16],
                    record[17],
                    record[18],
                    record[19],
                    record[20],
                    record[21]
                )
            )
        else:
            my_tree.insert(
                parent='', 
                index='end', 
                iid=count, 
                text='',
                tags=('oddrow'),
                values=(
                    record[0],
                    record[1],
                    record[2],
                    record[3],
                    record[4],
                    record[5],
                    record[6],
                    record[7],
                    record[8],
                    record[9],
                    record[10],
                    record[11],
                    record[12],
                    record[13],
                    record[14],
                    record[15],
                    record[16],
                    record[17],
                    record[18],
                    record[19],
                    record[20],
                    record[21]
                )
            )

        count += 1
        cursor.commit()

style.configure("Treeview", background="D3D3D3", foreground="black", rowheight=35, fieldbackground="D3D3D3")
style.map('Treeview', background=[('selected', '#347083')])

tree_frame1 = tkinter.Frame(janela_visualizar)
tree_frame1.pack(padx=10, pady=0)

tree_scroll1 = tkinter.Scrollbar(tree_frame1, orient='horizontal')
tree_scroll1.pack(side=tkinter.BOTTOM, fill=tkinter.X)

tree_scroll = tkinter.Scrollbar(tree_frame1, orient='vertical')
tree_scroll.pack(side=tkinter.RIGHT, fill=tkinter.Y)

my_tree =tkinter.ttk.Treeview(tree_frame1, yscrollcommand=tree_scroll.set, xscrollcommand=tree_scroll1.set, selectmode="extended")
my_tree.pack(fill="both", expand=True)

tree_scroll1.config(command=my_tree.xview)
tree_scroll.config(command=my_tree.yview)

my_tree['columns'] = (
    "ID", 
    "CPF CNPJ", 
    "Razão Social", 
    "Nome Fantasia", 
    "Gerente", 
    "Franqueado", 
    "C.Franqueado", 
    "Captador", 
    "C.Captador", 
    "Diretor", 
    "Status", 
    "Situação", 
    "Email", 
    "Celular", 
    "Mídia",
    "Data 1º contato", 
    "Data de Ativação", 
    "Data Comitê", 
    "Id-Sales Force", 
    "Faturamento", 
    "Tipo de operação", 
    "Desenho operacional"
)

my_tree.column('#0', width=0, stretch=tkinter.NO)
my_tree.column("ID", anchor=tkinter.W, width=30)
my_tree.column("CPF CNPJ", anchor=tkinter.W, width=110)
my_tree.column("Razão Social", anchor=tkinter.W, width=230)
my_tree.column("Nome Fantasia", anchor=tkinter.CENTER, width=200)
my_tree.column("Gerente", anchor=tkinter.CENTER, width=140)
my_tree.column("Franqueado", anchor=tkinter.CENTER, width=140)
my_tree.column("C.Franqueado", anchor=tkinter.CENTER, width=100)
my_tree.column("Captador", anchor=tkinter.CENTER, width=140)
my_tree.column("C.Captador", anchor=tkinter.CENTER, width=100)
my_tree.column("Diretor", anchor=tkinter.CENTER, width=140)
my_tree.column("Status", anchor=tkinter.CENTER, width=140)
my_tree.column("Situação", anchor=tkinter.CENTER, width=140)
my_tree.column("Email", anchor=tkinter.CENTER, width=140)
my_tree.column("Celular", anchor=tkinter.CENTER, width=140)
my_tree.column("Mídia", anchor=tkinter.CENTER, width=140)
my_tree.column("Data 1º contato", anchor=tkinter.CENTER, width=140)
my_tree.column("Data de Ativação", anchor=tkinter.CENTER, width=140)
my_tree.column("Data Comitê", anchor=tkinter.CENTER, width=140)
my_tree.column("Id-Sales Force", anchor=tkinter.CENTER, width=140)
my_tree.column("Faturamento", anchor=tkinter.CENTER, width=140)
my_tree.column("Tipo de operação", anchor=tkinter.CENTER, width=140)
my_tree.column("Desenho operacional", anchor=tkinter.CENTER, width=140)

my_tree.heading("#0", text="", anchor=tkinter.W)
my_tree.heading("#1", text="ID", anchor=tkinter.W)
my_tree.heading("#2", text="CPF CNPJ", anchor=tkinter.W)
my_tree.heading("#3", text="Razão Social", anchor=tkinter.W)
my_tree.heading("#4", text="Nome Fantasia", anchor=tkinter.CENTER)
my_tree.heading("#5", text="Gerente", anchor=tkinter.CENTER)
my_tree.heading("#6", text="Franqueado", anchor=tkinter.CENTER)
my_tree.heading("#7", text="C.Franqueado", anchor=tkinter.CENTER)
my_tree.heading("#8", text="Captador", anchor=tkinter.CENTER)
my_tree.heading("#9", text="C.Captador", anchor=tkinter.CENTER)
my_tree.heading("#10", text="Diretor", anchor=tkinter.CENTER)
my_tree.heading("#11", text="Status", anchor=tkinter.CENTER)
my_tree.heading("#12", text="Situação", anchor=tkinter.CENTER)
my_tree.heading("#13", text="Email", anchor=tkinter.CENTER)
my_tree.heading("#14", text="Celular", anchor=tkinter.CENTER)
my_tree.heading("#15", text="Mídia", anchor=tkinter.CENTER)
my_tree.heading("#16", text="Data 1º contato", anchor=tkinter.CENTER)
my_tree.heading("#17", text="Data de ativação", anchor=tkinter.CENTER)
my_tree.heading("#18", text="Data Comitê", anchor=tkinter.CENTER)
my_tree.heading("#19", text="Id-Sales Force", anchor=tkinter.CENTER)
my_tree.heading("#20", text="Faturamento", anchor=tkinter.CENTER)
my_tree.heading("#21", text="Tipo de operação", anchor=tkinter.CENTER)
my_tree.heading("#22", text="Desenho Operacional", anchor=tkinter.CENTER)

my_tree.tag_configure('oddrow', background="white")
my_tree.tag_configure('evenrow', background="lightblue")

# Add data to the screen
data_frame = tkinter.LabelFrame(janela_visualizar, text="Campos")
data_frame.pack(fill="x", expand="yes", padx=20)

ID_label = tkinter.Label(data_frame, text="ID")
ID_label.grid(row=0, column=0, padx=10, pady=10)
ID_entry = tkinter.Entry(data_frame)
ID_entry.grid(row=0, column=1, padx=10, pady=10)

CNPJ_label = tkinter.Label(data_frame, text="CPF CNPJ")
CNPJ_label.grid(row=0, column=0, padx=10, pady=10)
CNPJ_entry = tkinter.Entry(data_frame)
CNPJ_entry.grid(row=0, column=1, padx=10, pady=10)

Razao_label = tkinter.Label(data_frame, text="Razão Social")
Razao_label.grid(row=0, column=2, padx=10, pady=10)
Razao_entry = tkinter.Entry(data_frame)
Razao_entry.grid(row=0, column=3, padx=10, pady=10)

nomef_label = tkinter.Label(data_frame, text="Nome Fantasia")
nomef_label.grid(row=0, column=4, padx=10, pady=10)
nomef_entry = tkinter.Entry(data_frame)
nomef_entry.grid(row=0, column=5, padx=10, pady=10)

gerente_label = tkinter.Label(data_frame, text="Gerente")
gerente_label.grid(row=0, column=6, padx=10, pady=10)
gerente_entry = tkinter.Entry(data_frame)
gerente_entry.grid(row=0, column=7, padx=10, pady=10)

franqueado_label = tkinter.Label(data_frame, text="Franqueado")
franqueado_label.grid(row=0, column=8, padx=10, pady=10)
franqueado_entry = tkinter.Entry(data_frame)
franqueado_entry.grid(row=0, column=9, padx=10, pady=10)

cf_label = tkinter.Label(data_frame, text="Comissão Franqueado")
cf_label.grid(row=1, column=0, padx=10, pady=10)
cf_entry = tkinter.Entry(data_frame)
cf_entry.grid(row=1, column=1, padx=10, pady=10)

captador_label = tkinter.Label(data_frame, text="Captador")
captador_label.grid(row=1, column=2, padx=10, pady=10)
captador_entry = tkinter.Entry(data_frame)
captador_entry.grid(row=1, column=3, padx=10, pady=10)

cc_label = tkinter.Label(data_frame, text="C.Captador")
cc_label.grid(row=1, column=4, padx=10, pady=10)
cc_entry = tkinter.Entry(data_frame)
cc_entry.grid(row=1, column=5, padx=10, pady=10)

diretor_label = tkinter.Label(data_frame, text="Diretor")
diretor_label.grid(row=1, column=6, padx=10, pady=10)
diretor_entry = tkinter.Entry(data_frame)
diretor_entry.grid(row=1, column=7, padx=10, pady=10)

status_label = tkinter.Label(data_frame, text="Status")
status_label.grid(row=1, column=8, padx=10, pady=10)
status_entry = tkinter.Entry(data_frame)
status_entry.grid(row=1, column=9, padx=10, pady=10)

situacao_label = tkinter.Label(data_frame, text="Situacão")
situacao_label.grid(row=2, column=0, padx=10, pady=10)
situacao_entry = tkinter.Entry(data_frame)
situacao_entry.grid(row=2, column=1, padx=10, pady=10)

email_label = tkinter.Label(data_frame, text="Email")
email_label.grid(row=2, column=2, padx=10, pady=10)
email_entry = tkinter.Entry(data_frame)
email_entry.grid(row=2, column=3, padx=10, pady=10)

celular_label = tkinter.Label(data_frame, text="Celular")
celular_label.grid(row=2, column=4, padx=10, pady=10)
celular_entry = tkinter.Entry(data_frame)
celular_entry.grid(row=2, column=5, padx=10, pady=10)

midia_label = tkinter.Label(data_frame, text="Mídia")
midia_label.grid(row=2, column=6, padx=10, pady=10)
midia_entry = tkinter.Entry(data_frame)
midia_entry.grid(row=2, column=7, padx=10, pady=10)
    
data1c_label = tkinter.Label(data_frame, text="Data 1º contato")
data1c_label.grid(row=2, column=8, padx=10, pady=10)
data1c_entry = tkinter.Entry(data_frame)
data1c_entry.grid(row=2, column=9, padx=10, pady=10)
    
DataAtiv_label = tkinter.Label(data_frame, text="Data de Ativação")
DataAtiv_label.grid(row=3, column=0, padx=10, pady=10)
DataAtiv_entry = tkinter.Entry(data_frame)
DataAtiv_entry.grid(row=3, column=1, padx=10, pady=10)
    
DataComite_label = tkinter.Label(data_frame, text="Data Comitê")
DataComite_label.grid(row=3, column=2, padx=10, pady=10)
DataComite_entry = tkinter.Entry(data_frame)
DataComite_entry.grid(row=3, column=3, padx=10, pady=10)
    
IdSf_label = tkinter.Label(data_frame, text="Id-Sales Force")
IdSf_label.grid(row=3, column=4, padx=10, pady=10)
IdSf_entry = tkinter.Entry(data_frame)
IdSf_entry.grid(row=3, column=5, padx=10, pady=10)
    
Faturamento_label = tkinter.Label(data_frame, text="Faturamento")
Faturamento_label.grid(row=3, column=6, padx=10, pady=10)
Faturamento_entry = tkinter.Entry(data_frame)
Faturamento_entry.grid(row=3, column=7, padx=10, pady=10)
    
TipoOp_label = tkinter.Label(data_frame, text="Tipo de operação")
TipoOp_label.grid(row=3, column=8, padx=10, pady=10)
TipoOP_entry = tkinter.Entry(data_frame)
TipoOP_entry.grid(row=3, column=9, padx=10, pady=10)
    
DO_label = tkinter.Label(data_frame, text="Desenho operacional")
DO_label.grid(row=4, column=0, padx=10, pady=10)
DO_entry = tkinter.Entry(data_frame)
DO_entry.grid(row=4, column=1, padx=10, pady=10)

button_frame = tkinter.LabelFrame(janela_visualizar, text="")
button_frame.pack(fill="x", expand="yes", padx=20)


def clear_entryBoxes():
    # Clear entry boxes
    ID_entry.delete(0, tkinter.END)
    CNPJ_entry.delete(0, tkinter.END)
    Razao_entry.delete(0, tkinter.END)
    nomef_entry.delete(0, tkinter.END)
    gerente_entry.delete(0, tkinter.END)
    franqueado_entry.delete(0, tkinter.END)
    cf_entry.delete(0, tkinter.END)
    captador_entry.delete(0, tkinter.END)
    cc_entry.delete(0, tkinter.END)
    diretor_entry.delete(0, tkinter.END)
    status_entry.delete(0, tkinter.END)
    situacao_entry.delete(0, tkinter.END)
    email_entry.delete(0, tkinter.END)
    celular_entry.delete(0, tkinter.END)
    midia_entry.delete(0, tkinter.END)
    data1c_entry.delete(0, tkinter.END)
    DataAtiv_entry.delete(0, tkinter.END)
    DataComite_entry.delete(0, tkinter.END)
    IdSf_entry.delete(0, tkinter.END)
    Faturamento_entry.delete(0, tkinter.END)
    TipoOP_entry.delete(0, tkinter.END)
    DO_entry.delete(0, tkinter.END)


def select_record(e):

    clear_entryBoxes()
    # Grab Record Number
    selected = my_tree.focus()
    # Grab Record Value
    values = my_tree.item(selected, 'values')

    # Output to entry boxes
    ID_entry.insert(0, values[0])
    CNPJ_entry.insert(0, values[1])
    Razao_entry.insert(0, values[2])
    nomef_entry.insert(0, values[3])
    gerente_entry.insert(0, values[4])
    franqueado_entry.insert(0, values[5])
    cf_entry.insert(0, values[6])
    captador_entry.insert(0, values[7])
    cc_entry.insert(0, values[8])
    diretor_entry.insert(0, values[9])
    status_entry.insert(0, values[10])
    situacao_entry.insert(0, values[11])
    email_entry.insert(0, values[12])
    celular_entry.insert(0, values[13])
    midia_entry.insert(0, values[14])
    data1c_entry.insert(0, values[15])
    DataAtiv_entry.insert(0, values[16])
    DataComite_entry.insert(0, values[17])
    IdSf_entry.insert(0, values[18])
    Faturamento_entry.insert(0, values[19])
    TipoOP_entry.insert(0, values[20])
    DO_entry.insert(0, values[21])


def delete_selected():

    delete = "DELETE FROM Cadastros WHERE ID="+ ID_entry.get()
    current_selection = my_tree.selection()[0]
    my_tree.delete(current_selection)
    cursor = conexao.cursor()

    cursor.execute(delete)
    cursor.commit()
    clear_entryBoxes()
    tkinter.messagebox.showinfo(title="Python 3", message="CADASTRO DELETADO!")

def update_record():
    cursor = conexao.cursor()
    selected = my_tree.focus()
    my_tree.item(
        selected, 
        text="",
        values=(
            ID_entry.get(),
            CNPJ_entry.get(),
            Razao_entry.get(),
            nomef_entry.get(),
            gerente_entry.get(),
            franqueado_entry.get(),
            cf_entry.get(),
            captador_entry.get(),
            cc_entry.get(),
            diretor_entry.get(),
            status_entry.get(),
            situacao_entry.get(),
            email_entry.get(),
            celular_entry.get(),
            midia_entry.get(),
            data1c_entry.get(),
            DataAtiv_entry.get(),
            DataComite_entry.get(),
            IdSf_entry.get(),
            Faturamento_entry.get(),
            TipoOP_entry.get(),
            DO_entry.get()
        )
    )

    cursor.execute(
        f"""UPDATE Cadastros

        SET
        "CPF/CNPJ" = '{CNPJ_entry.get()}',
        "Razão Social" = '{ Razao_entry.get()}',
        "Nome Fantasia" = '{ nomef_entry.get()}',
        Gerente = '{ gerente_entry.get()}',
        Franqueado = '{ franqueado_entry.get()}',
        "Comissão Franqueado" = '{ cf_entry.get()}',
        Captador = '{ captador_entry.get()}',
        "Comissão Captador" = '{ cc_entry.get()}',
        Diretor = '{ diretor_entry.get()}',
        "Status" = '{ status_entry.get()}',
        Situação = '{ situacao_entry.get()}',
        "E-mail" = '{ email_entry.get()}',
        Celular = '{ celular_entry.get()}',
        Mídia = '{ midia_entry.get()}',
        "Data 1º Contato" = '{ data1c_entry.get()}',
        "Data de Ativação" = '{ DataAtiv_entry.get()}',
        "Data Comitê" = '{ DataComite_entry.get()}',
        "Id-salesforce" = '{ IdSf_entry.get()}',
        Faturamento = '{ Faturamento_entry.get()}',
        "Tipo de operação" = '{TipoOP_entry.get()}',
        "Desenho operacional" = '{ DO_entry.get()}'
        WHERE ID = '{ID_entry.get()}'
        """
    )

    cursor.commit()
    clear_entryBoxes()
    tkinter.messagebox.showinfo(title="Python 3", message="Cadastro Alterado!")


def Cadastro():

    for CPF in ('bar','foo'): 
        cadastroC = f"""SELECT "CPF/CNPJ" FROM Cadastros WHERE "CPF/CNPJ" = '{CNPJ_entry.get()}'"""
        cursor.execute(cadastroC)
        data = cursor.fetchall()
    if len(data)==0:
        comando = f"""INSERT INTO Cadastros(ID, "CPF/CNPJ", "Razão Social", "Nome Fantasia", Gerente, Franqueado, "Comissão Franqueado", Captador, "Comissão Captador", Diretor, "Status", Situação, "E-mail", Celular, Mídia, "Data 1º contato", "Data de Ativação", "Data Comitê", "Id-salesforce", Faturamento, "Tipo de operação", "Desenho operacional")
        VALUES(next value for scadastro,'{CNPJ_entry.get()}', '{Razao_entry.get()}', '{nomef_entry.get()}', '{gerente_entry.get()}', '{franqueado_entry.get()}', '{cf_entry.get()}', '{captador_entry.get()}','{cc_entry.get()}','{diretor_entry.get()}','{status_entry.get()}','{situacao_entry.get()}','{email_entry.get()}','{celular_entry.get()}','{midia_entry.get()}','{data1c_entry.get()}','{DataAtiv_entry.get()}','{DataComite_entry.get()}','{IdSf_entry.get()}','{Faturamento_entry.get()}','{TipoOP_entry.get()}','{DO_entry.get()}')"""
        cursor.execute(comando)
        cursor.commit()
        clear_entryBoxes()
        my_tree.delete(*my_tree.get_children())
        query_database()
        tkinter.messagebox.showinfo(title="Python3", message="CADASTRO EFETUADO!")
    else:
        pg.alert('CNPJ já cadastrado: %s'%CNPJ_entry.get())


def gerar_excel():
    comando2 =  f"""SELECT * FROM Cadastros"""
    consultar_tabela = pd.read_sql_query(comando2, conexao)
    ler_tabela = pd.DataFrame(consultar_tabela)
    ler_tabela.to_excel(r'Cadastros.xlsx',index=False)
    tkinter.messagebox.showinfo(title="Microsoft Excel", message="Planilha Gerada")


def search():    
    my_tree.delete(*my_tree.get_children())
    cursor.execute(f"""SELECT * FROM Cadastros WHERE "CPF/CNPJ" ='{CNPJ_entry.get()}' """)
    records = cursor.fetchall()
    count = 0

    for record in records:
        if count % 2 == 0:                                                                           
                my_tree.insert(
                    parent='',
                    index='end',
                    iid=count,
                    text='',
                    tags=('evenrow'),
                    values=(
                        record[0],
                        record[1],
                        record[2],
                        record[3],
                        record[4],
                        record[5],
                        record[6],
                        record[7],
                        record[8],
                        record[9],
                        record[10],
                        record[11],
                        record[12],
                        record[13],
                        record[14],
                        record[15],
                        record[16],
                        record[17],
                        record[18],
                        record[19],
                        record[20],
                        record[21]
                    )
                )
        else:
            my_tree.insert(
                parent='',
                index='end',
                iid=count, 
                text='',
                tags=('oddrow'),
                values=(
                    record[0],
                    record[1],
                    record[2],
                    record[3],
                    record[4],
                    record[5],
                    record[6],
                    record[7],
                    record[8],
                    record[9],
                    record[10],
                    record[11],
                    record[12],
                    record[13],
                    record[14],
                    record[15],
                    record[16],
                    record[17],
                    record[18],
                    record[19],
                    record[20],
                    record[21]
                )
            )
        count += 1

    cursor.commit()


def limpar():
    my_tree.delete(*my_tree.get_children())
    clear_entryBoxes()
    query_database()

buttonExcel = tkinter.Button(button_frame, text="Exportar Excel", command=gerar_excel)
buttonExcel.grid(row=0, column=5, padx=10, pady=10)
    
buttonADD = tkinter.Button(button_frame, text="Cadastrar", command=Cadastro)
buttonADD.grid(row=0, column=4, padx=10, pady=10)
    
buttonUPDATE = tkinter.Button(button_frame, text="Alterar selecionado", command=update_record)
buttonUPDATE.grid(row=0, column=3, padx=10, pady=10)
        
select_record_button = tkinter.Button(button_frame, text="Selecionar Registro", command=select_record)
select_record_button.grid(row=0, column=0, padx=10, pady=10)
    
buttonDELETE = tkinter.Button(button_frame, text="Apagar selecionado", command=delete_selected)
buttonDELETE.grid(row=0, column=1, padx=10, pady=10)

buttonSearch = tkinter.Button(button_frame, text="Filtrar", command=search)
buttonSearch.grid(row=0, column=6, padx=10, pady=10)

clear_button = tkinter.Button(button_frame, text="Limpar Filtros", command=limpar)
clear_button.grid(row=0, column=7, padx=10, pady=10)

sair_button = tkinter.Button(button_frame, text="Sair", command=janela_visualizar.destroy)
sair_button.grid(row=0, column=8, padx=10, pady=10)


query_database()
my_tree.bind("<ButtonRelease-1>", select_record) 
janela_visualizar.mainloop()