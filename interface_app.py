import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import os
from logica_processamento import processar_e_transformar_planilha


class AppLimpezaPlanilha:
    def __init__(self, master):
        self.master = master

        master.title("Igreja Alicerce")
        master.geometry("480x250")
        master.config(padx=10, pady=10)
        caminho_imagem = "icone_igreja.png"

        try:
            if os.path.exists(caminho_imagem):
                img_pil = Image.open(caminho_imagem)
                img_pil_icon = img_pil.resize((32, 32), Image.Resampling.LANCZOS)
                self.img_tk = ImageTk.PhotoImage(img_pil_icon)
                master.iconphoto(True, self.img_tk)

            else:
                self.img_tk = None

        except Exception as e:
            self.img_tk = None

        self.caminho_arquivo = tk.StringVar()
        self.caminho_arquivo.set("Nenhum arquivo selecionado.")
        tk.Label(master,
                 text="Selecione o Arquivo XLSX (ContaAzul)",
                 font=("Arial", 12, "bold")).pack(pady=10)
        self.label_caminho = tk.Label(master,
                                      textvariable=self.caminho_arquivo,
                                      bg="white",
                                      relief="groove",
                                      width=60,
                                      anchor="w")
        self.label_caminho.pack(pady=5)
        self.btn_selecionar = tk.Button(master,
                                        text="Buscar Arquivo",
                                        command=self.abrir_dialogo_arquivo)
        self.btn_selecionar.pack(pady=10)
        self.btn_processar = tk.Button(master,
                                       text="▶️ Transformar",
                                       command=self.iniciar_processamento,
                                       bg="#007ACC",
                                       fg="white",
                                       font=("Arial", 10, "bold"))
        self.btn_processar.pack(pady=15)

    def abrir_dialogo_arquivo(self):
        """Abre o diálogo de seleção de arquivo e atualiza o campo."""
        arquivo_selecionado = filedialog.askopenfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel/CSV", "*.xlsx *.csv"),
                       ("Todos os arquivos", "*.*")]
        )
        if arquivo_selecionado:
            self.caminho_arquivo.set(arquivo_selecionado)

    def iniciar_processamento(self):
        caminho = self.caminho_arquivo.get()

        if "Nenhum arquivo selecionado." in caminho:
            messagebox.showwarning("Atenção", "Por favor, selecione um arquivo primeiro.")
            return

        self.btn_processar.config(state=tk.DISABLED, text="Processando...")

        resultado = processar_e_transformar_planilha(caminho)

        self.btn_processar.config(state=tk.NORMAL, text="▶️ Transformar p/ Formato Longo e FFill")

        if "Sucesso" in resultado:
            messagebox.showinfo("Concluído", resultado)
        else:
            messagebox.showerror("Erro", resultado)


if __name__ == "__main__":
    root = tk.Tk()
    app = AppLimpezaPlanilha(root)
    root.mainloop()