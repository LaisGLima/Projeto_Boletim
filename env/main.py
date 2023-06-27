import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from crud import AppBD
from openpyxl import Workbook
import io
import os

class AppBDInterface(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sistema de notas e cadastro")
        self.geometry("800x600")
        self.resizable(False, False)
        self.app_bd = AppBD()
        self.create_tabs()
        self.center_window()

    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry('{}x{}+{}+{}'.format(width, height, x, y))

    def create_tabs(self):
        tab_control = ttk.Notebook(self)

        alunos_tab = ttk.Frame(tab_control)
        self.create_alunos_tab(alunos_tab)
        tab_control.add(alunos_tab, text="Alunos")

        notas_tab = ttk.Frame(tab_control)
        self.create_notas_tab(notas_tab)
        tab_control.add(notas_tab, text="Notas")

        tab_control.pack(expand=True, fill="both")

# ALUNOS   
    def listar_alunos(self):
        alunos = self.app_bd.select_Aluno()
        self.alunos_treeview.delete(*self.alunos_treeview.get_children())
        for aluno in alunos:
            self.alunos_treeview.insert("", "end", values=aluno)

    def cadastrar_aluno(self):
        cpf = self.cpf_entry.get()
        nome = self.nome_entry.get()
        data_nascimento = self.data_nascimento_entry.get()
        sexo = self.sexo_combobox.get()

        if cpf and nome and data_nascimento and sexo:
            self.app_bd.insert_Aluno(cpf, nome, data_nascimento, sexo)
            self.listar_alunos()
        else:
            messagebox.showwarning("Erro", "Por favor, preencha todos os campos corretamente.")

    def atualizar_aluno(self):
        cpf = self.cpf_entry.get()
        nome = self.nome_entry.get()
        data_nascimento = self.data_nascimento_entry.get()
        sexo = self.sexo_combobox.get()

        if cpf and nome and data_nascimento and sexo:
            self.app_bd.update_Aluno(cpf, nome, data_nascimento, sexo)
            self.listar_alunos()
            self.listar_boletim()
        else:
            messagebox.showwarning("Erro", "Por favor, preencha todos os campos corretamente.")

    def excluir_aluno(self):
        cpf = self.cpf_entry.get()

        if cpf:
            self.app_bd.delete_Aluno(cpf)
            self.listar_alunos()
            self.listar_boletim()
        else:
            messagebox.showwarning("Erro", "Por favor, insira o CPF do aluno.")

    def create_alunos_tab(self, tab):
        tab.columnconfigure(0, weight=1)
        tab.columnconfigure(1, weight=1)

        title_frame = ttk.Frame(tab)
        title_frame.grid(row=0, column=0, columnspan=5, pady=10)

        title_label = tk.Label(title_frame, text="CADASTRE OS ALUNOS:", font=("Arial", 18, "bold"))
        title_label.pack()

        cpf_label = tk.Label(tab, text="CPF:", font=("Arial", 12))
        cpf_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.cpf_entry = tk.Entry(tab, font=("Arial", 12))
        self.cpf_entry.insert(0, '000.000.000-00')
        self.cpf_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        nome_label = tk.Label(tab, text="Nome:", font=("Arial", 12))
        nome_label.grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.nome_entry = tk.Entry(tab, font=("Arial", 12))
        self.nome_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        data_nascimento_label = tk.Label(tab, text="Data de Nascimento:", font=("Arial", 12))
        data_nascimento_label.grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.data_nascimento_entry = tk.Entry(tab, font=("Arial", 12))
        self.data_nascimento_entry.insert(0, "dd/mm/aaaa")
        self.data_nascimento_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        sexo_label = tk.Label(tab, text="Sexo:", font=("Arial", 12))
        sexo_label.grid(row=4, column=0, padx=5, pady=5, sticky="e")
        sexo = ["Masculino", "Feminino"]
        self.sexo_combobox = ttk.Combobox(tab, values=sexo, font=("Arial", 11))
        self.sexo_combobox.grid(row=4, column=1, padx=5, pady=5, sticky="w")

        space_frame = ttk.Frame(tab)
        space_frame.grid(row=5, column=0, columnspan=3, pady=5)

        atualizar_button = tk.Button(tab, text="Atualizar Aluno", command=self.atualizar_aluno, font=("Arial", 12))
        atualizar_button.grid(row=6, column=0, padx=40, pady=5, sticky="w")

        cadastrar_button = tk.Button(tab, text="Cadastrar Aluno", command=self.cadastrar_aluno, font=("Arial", 12))
        cadastrar_button.grid(row=6, column=1, padx=5, pady=5, sticky="w")

        excluir_button = tk.Button(tab, text=" Excluir  Aluno ", command=self.excluir_aluno, font=("Arial", 12))
        excluir_button.grid(row=6, column=2, padx=5, pady=5, sticky="w")

        space_frame = ttk.Frame(tab)
        space_frame.grid(row=7, column=0, columnspan=3, pady=10)

        scrollbar = ttk.Scrollbar(tab)
        scrollbar.grid(row=8, column=3, sticky="ns")

        self.alunos_treeview = ttk.Treeview(
            tab,
            columns=("cpf", "nome", "data_nascimento", "sexo"),
            show="headings",
            yscrollcommand=scrollbar.set
        )
        self.alunos_treeview.heading("cpf", text="CPF")
        self.alunos_treeview.heading("nome", text="Aluno")
        self.alunos_treeview.heading("data_nascimento", text="Data de Nascimento")
        self.alunos_treeview.heading("sexo", text="Sexo")

        self.alunos_treeview.column("cpf", width=200, anchor="center")
        self.alunos_treeview.column("nome", width=200, anchor="center")
        self.alunos_treeview.column("data_nascimento", width=200, anchor="center")
        self.alunos_treeview.column("sexo", width=155, anchor="center")

        self.alunos_treeview.grid(row=8, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")

        scrollbar.config(command=self.alunos_treeview.yview)

        exportar_button = tk.Button(tab, text="Baixar Planilha", command=self.exportar_excel, font=("Arial", 12))
        exportar_button.grid(row=9, column=0, columnspan=3, padx=5, pady=5)

        self.listar_alunos()

# NOTAS
    def cadastrar_nota(self):
        id_Aluno = self.id_Aluno_entry.get()
        AV1 = self.AV1_entry.get()
        AV2 = self.AV2_entry.get()
        disciplina = self.disciplina_combobox.get()

        if id_Aluno and disciplina and AV1 and AV2:
            self.app_bd.insert_Notas(id_Aluno, disciplina, AV1, AV2)
            self.listar_boletim()
        else:
            messagebox.showwarning("Erro", "Por favor, preencha todos os campos corretamente.")

    def atualizar_nota(self):
        id_Aluno = self.id_Aluno_entry.get()
        AV1 = self.AV1_entry.get()
        AV2 = self.AV2_entry.get()
        disciplina = self.disciplina_combobox.get()

        if id_Aluno and disciplina and AV1 and AV2:
            try:
                AV1 = float(AV1)
                AV2 = float(AV2)
                self.app_bd.update_Notas(id_Aluno, disciplina, AV1, AV2)
                self.listar_boletim()

            except ValueError:
                messagebox.showerror("Erro", "Por favor, insira valores numéricos válidos para AV1 e AV2.")
        else:
            messagebox.showwarning("Erro", "Por favor, preencha todos os campos corretamente.")

    def excluir_nota(self):
        id_Aluno = self.id_Aluno_entry.get()
        disciplina = self.disciplina_combobox.get()

        if id_Aluno and disciplina:
            self.app_bd.delete_Notas(id_Aluno, disciplina)
            self.listar_boletim()
        else:
            messagebox.showwarning("Erro", "Por favor, insira o CPF e disciplina do aluno.")

    def create_notas_tab(self, tab):
        tab.columnconfigure(0, weight=1)
        tab.columnconfigure(1, weight=1)

        title_frame = ttk.Frame(tab)
        title_frame.grid(row=0, column=0, columnspan=5, pady=10)
        title_label = tk.Label(title_frame, text="CADASTRE AS NOTAS:", font=("Arial", 18, "bold"))
        title_label.pack()
        
        id_Aluno_label = tk.Label(tab, text="CPF:", font=("Arial", 12))
        id_Aluno_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.id_Aluno_entry = tk.Entry(tab, font=("Arial", 12))
        self.id_Aluno_entry.insert(0, '000.000.000-00')
        self.id_Aluno_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        disciplina_label = tk.Label(tab, text="Disciplina:", font=("Arial", 12))
        disciplina_label.grid(row=2, column=0, padx=5, pady=5, sticky="e")
        disciplina = ['Biologia', 'Filosofia', 'Física', 'Geografia', 'Inglês', 'História', 'Literatura', 'Matemática', 'Português', 'Química', 'Sociologia']
        self.disciplina_combobox = ttk.Combobox(tab, values=disciplina, font=("Arial", 11))
        self.disciplina_combobox.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        AV1_label = tk.Label(tab, text="AV1:", font=("Arial", 12))
        AV1_label.grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.AV1_entry = tk.Entry(tab, font=("Arial", 12))
        self.AV1_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        AV2_label = tk.Label(tab, text="AV2:", font=("Arial", 12))
        AV2_label.grid(row=4, column=0, padx=5, pady=5, sticky="e")
        self.AV2_entry = tk.Entry(tab, font=("Arial", 12))
        self.AV2_entry.grid(row=4, column=1, padx=5, pady=5, sticky="w")
        
        space_frame = ttk.Frame(tab)
        space_frame.grid(row=5, column=0, columnspan=3, pady=10)

        atualizar_button = tk.Button(tab, text="Atualizar Nota", command=self.atualizar_nota, font=("Arial", 12))
        atualizar_button.grid(row=6, column=0, padx=40, pady=5, sticky="w")

        cadastrar_button = tk.Button(tab, text="Cadastrar Nota", command=self.cadastrar_nota, font=("Arial", 12))
        cadastrar_button.grid(row=6, column=1, padx=5, pady=5, sticky="w")

        excluir_button = tk.Button(tab, text="  Excluir  Nota  ", command=self.excluir_nota, font=("Arial", 12))
        excluir_button.grid(row=6, column=2, padx=5, pady=5, sticky="w")

        space_frame = ttk.Frame(tab)
        space_frame.grid(row=7, column=0, columnspan=3, pady=10)

        scrollbar = ttk.Scrollbar(tab)
        scrollbar.grid(row=8, column=3, sticky="ns")

        self.boletim_treeview = ttk.Treeview(
            tab,
            columns=("nome", "disciplina", "AV1", "AV2", "media", "status"),
            show="headings",
            yscrollcommand=scrollbar.set
        )
        self.boletim_treeview.heading("nome", text="Nome")
        self.boletim_treeview.heading("disciplina", text="Disciplina")
        self.boletim_treeview.heading("AV1", text="AV1")
        self.boletim_treeview.heading("AV2", text="AV2")
        self.boletim_treeview.heading("media", text="média")
        self.boletim_treeview.heading("status", text="Status")

        self.boletim_treeview.column("nome", width=104, anchor="center")
        self.boletim_treeview.column("disciplina", width=104, anchor="center")
        self.boletim_treeview.column("AV1", width=90, anchor="center")
        self.boletim_treeview.column("AV2", width=90, anchor="center")
        self.boletim_treeview.column("media", width=90, anchor="center")
        self.boletim_treeview.column("status", width=104, anchor="center")

        self.boletim_treeview.grid(row=8, column=0, columnspan=3, padx=5, pady=5, sticky="nsew")
        scrollbar.config(command=self.boletim_treeview.yview)
        
        exportar_button = tk.Button(tab, text="Baixar Planilha", command=lambda: [self.exportar_excel(), messagebox.showinfo("Dowload Concluído", "Os dados foram exportados com sucesso.")], font=("Arial", 12))
        exportar_button.grid(row=9, column=0, columnspan=3, padx=5, pady=5)

        self.listar_boletim()

# BOLETIM
    def listar_boletim(self):
        app_bd = AppBD()
        self.boletim_treeview.delete(*self.boletim_treeview.get_children())
        boletim = app_bd.select_Boletim()

        for row in boletim:
            self.boletim_treeview.insert("", "end", values=row)
        app_bd.connection.close()

# EXPORTAR PARA EXCEL
    def exportar_excel(self):
            alunos = self.app_bd.select_Aluno()
            boletim = self.app_bd.select_Boletim()

            wb = Workbook()
            ws_alunos = wb.active
            ws_alunos.title = "Alunos"
            ws_alunos.append(["CPF", "Nome", "Data de Nascimento", "Sexo"])
            for aluno in alunos:
                ws_alunos.append(aluno)

            ws_boletim = wb.create_sheet(title="Boletim")
            ws_boletim.append(["Nome", "Disciplina", "AV1", "AV2", "Média", "Status"])
            for linha in boletim:
                ws_boletim.append(linha)
                
            buffer = io.BytesIO()
            wb.save(buffer)
            buffer.seek(0)
            destino = os.path.expanduser("~/Downloads/alunos_boletim.xlsx")
            with open(destino, "wb") as file:
                file.write(buffer.getvalue())

if __name__ == "__main__":
    app = AppBDInterface()
    app.mainloop()
