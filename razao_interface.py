import tkinter as tk
from tkinter import ttk
from tkinter import ttk, filedialog

def on_click():
    arquivo = filedialog.askopenfilename(title="Selecione um arquivo", filetypes=(("Arquivos de texto", "*.txt"), ("Todos os arquivos", "*.*")))
    if arquivo:
        label.config(text="Arquivo selecionado: " + arquivo)
    else:
        label.config(text="Nenhum arquivo selecionado.")

janela = tk.Tk()
janela.title("Interface de Tratamento de Razão")

# Definir o tamanho da janela (largura x altura)
janela.geometry("800x600")

# Criando um botão personalizado usando o módulo ttk
estilo = ttk.Style()
estilo.configure('Meu.TButton', foreground='teal', font=('Helvetica', 12))

botao_personalizado = ttk.Button(janela, text="Clique Aqui", style='Meu.TButton', command=on_click)
botao_personalizado.pack(pady=20)

# Label para mostrar a mensagem quando o botão for clicado
label = ttk.Label(janela, text="")
label.pack()

janela.mainloop()
