from typing import Optional, Tuple, Union
import customtkinter as ctk
from tkinter import *
from tkinter import messagebox
import pathlib
from openpyxl import Workbook, workbook
import openpyxl, xlrd
from tkinter import ttk, filedialog

# Setando a aparencia do sistema
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.layout_config()
        self.apperence()
        self.todo_sistema()

    def layout_config(self):
        self.title("Sistema de Tratamento de Razão - Equipe Contabil Novonor")
        self.geometry("700x500")

    def apperence(self):
        self.lb_apm = ctk.CTkLabel(
            self, text="Tema", bg_color="transparent", text_color=["#000", "#fff"]
        ).place(x=50, y=440)
        self.opt_apm = ctk.CTkOptionMenu(
            self, values=["Light", "Dark", "System"], command=self.change_apm
        ).place(x=50, y=465)

    def change_apm(self, nova_aparencia):
        ctk.set_appearance_mode(nova_aparencia)
    
    def select_file(self):
        # Abre uma janela de diálogo para selecionar um arquivo
        file_path = filedialog.askopenfilename(
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        # Verifica se um arquivo foi selecionado
        if file_path:
            print("Arquivo selecionado:", file_path)
             # Atualiza o texto do label com o caminho do arquivo selecionado
            self.selected_file_label = ctk.CTkLabel(
                self, text="Arquivo selecionado: " + file_path, font=("Century Gothic", 12)
            ).place(x=50, y=150)
            # Você pode armazenar o caminho do arquivo em uma variável aqui, se necessário.
    

    def todo_sistema(self):
        frame = ctk.CTkFrame(
            self,
            width=700,
            height=50,
            corner_radius=0,
            bg_color="teal",
            fg_color="teal",
        ).place(x=0, y=10)
        title = ctk.CTkLabel(
            frame,
            text="Sistema de Tratamento de Razão - Novonor",
            font=("Century Gothic", 24),
            text_color="#fff",
            bg_color="teal",
        ).place(x=50, y=20)
        span = ctk.CTkLabel(
            self,
            text="Por Favor, selecione o arquivo para tratamento",
            font=("Century Gothic", 16),
            text_color=["#000", "#fff"],
        ).place(x=50, y=70)

        # Adicionar o botão de seleção de arquivo
        self.btn_select_file = ctk.CTkButton(
            self, text="Selecionar Arquivo", command=self.select_file
        ).place(x=50, y=110)

        # Adicionar o botão para executar o tratamento do arquivo
        self.btn_execute = ctk.CTkButton(
            self, text="Executar Tratamento".upper(), 
            command= self.submit,
            fg_color="#151",
            hover_color="#131",
        ).place(x=520, y=465)
        self.btn_execute = ctk.CTkButton(
            self,
            text="Limpar Campos".upper(),
            command=self.clear,
            fg_color="#555",
            hover_color="#333",
        ).place(x=355, y=465)   

        # Adicionar o label para exibir o caminho do arquivo selecionado
        self.selected_file_label = ctk.CTkLabel(
            self, text="Nenhum arquivo selecionado.", font=("Century Gothic", 12)
        ).place(x=50, y=150)


    #Funções
    def submit(self):
        # Verifica se um arquivo foi selecionado
        if hasattr(self, 'file_path') and self.file_path:
            print("Executando tratamento para o arquivo:", self.file_path)
            # Aqui você pode chamar o script de tratamento passando o caminho do arquivo self.file_path
            # Chamar o script de tratamento passando o caminho do arquivo
            #ScriptTratamento.tratar_arquivo(self.file_path)
        else:
            messagebox.showerror(
                "Erro",
                "Nenhum arquivo selecionado. Selecione um arquivo antes de executar o tratamento.",
            )
    
    def clear(self):
        if hasattr(self, "selected_file_label"):
            self.selected_file_label.config(text="Nenhum arquivo selecionado.")
        if hasattr(self, "file_path"):
            self.file_path = None
            


if __name__ == "__main__":
    app = App()
    app.mainloop()
