import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk
from PIL import Image, ImageTk
import threading
import os
import Splitter
import Merge

class Aplicacao(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Aplicação")
        self.geometry("500x500")  
        self.resizable(False, False)
        self.caminho_arquivo = None
        self.num_linhas = 0
        self.arquivos_selecionados = []  # Lista para armazenar os arquivos selecionados

        # Configurar o estilo
        self.estilo = ttk.Style(self)
        self.estilo.theme_use('clam')
        self.configure(bg='#151c43')
        self.estilo.configure('.', background='#151c43', foreground='white', fieldbackground='#151c43')
        self.estilo.configure('TLabel', background='#151c43', foreground='white')
        self.estilo.configure('TButton', background='#151c43', foreground='white')
        self.estilo.configure('TEntry', fieldbackground='white', foreground='black')

        # Frames
        self.frame_inicial = ttk.Frame(self)
        self.frame_splitter = ttk.Frame(self)
        self.frame_divisao = ttk.Frame(self)
        self.frame_merge = ttk.Frame(self)

        # Variável para armazenar a imagem
        self.logo_image = None

        # Inicializar a interface
        self.mostrar_frame_inicial()

    def limpar_frames(self):
        for widget in self.winfo_children():
            widget.destroy()

    def mostrar_frame_inicial(self):
        self.limpar_frames()
        self.frame_inicial = ttk.Frame(self)
        self.frame_inicial.pack(fill="both", expand=True)

        # Carregar a imagem
        try:
            image = Image.open("logo.jpg")

            # Manter a proporção da imagem ao redimensionar
            largura_desejada = 400  # Ajuste conforme necessário
            proporcao = largura_desejada / image.width
            altura_nova = int(image.height * proporcao)
            image = image.resize((largura_desejada, altura_nova), self.get_resample_method())

            self.logo_image = ImageTk.PhotoImage(image)
            label_imagem = ttk.Label(self.frame_inicial, image=self.logo_image)
            label_imagem.pack(pady=10)
        except Exception as e:
            print(f"Erro ao carregar a imagem: {e}")
            # Se ocorrer um erro, podemos exibir um label com texto padrão
            label_imagem = ttk.Label(self.frame_inicial, text="Imagem não disponível")
            label_imagem.pack(pady=10)

        label_opcao = ttk.Label(self.frame_inicial, text="Qual opção deseja:")
        label_opcao.pack(pady=10)

        botao_splitter = ttk.Button(self.frame_inicial, text="Splitter", command=self.mostrar_frame_splitter)
        botao_splitter.pack(pady=5)

        botao_merge = ttk.Button(self.frame_inicial, text="Merge", command=self.mostrar_frame_merge)
        botao_merge.pack(pady=5)

        # Mensagem na parte inferior
        label_autor = ttk.Label(self.frame_inicial, text="Criado por LEOGRUP")
        label_autor.pack(side=tk.BOTTOM, pady=10)

    def get_resample_method(self):
        # Tentar usar Image.LANCZOS ou Image.ANTIALIAS dependendo da versão do Pillow
        try:
            resample = Image.LANCZOS
        except AttributeError:
            resample = Image.ANTIALIAS
        return resample

    def mostrar_frame_splitter(self):
        self.limpar_frames()
        self.frame_splitter = ttk.Frame(self)
        self.frame_splitter.pack(fill="both", expand=True)

        botao_voltar = ttk.Button(self.frame_splitter, text="Voltar", command=self.mostrar_frame_inicial)
        botao_voltar.pack(anchor='nw', pady=10, padx=10)

        botao_carregar = ttk.Button(self.frame_splitter, text="Carregar Arquivo", command=self.carregar_arquivo)
        botao_carregar.pack(pady=20)

    def carregar_arquivo(self):
        caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
        if caminho_arquivo:
            caminho_arquivo = os.path.normpath(caminho_arquivo)
            print(f"Caminho do arquivo selecionado: {caminho_arquivo}")
            if os.path.isfile(caminho_arquivo):
                try:
                    num_linhas = Splitter.obter_num_linhas(caminho_arquivo)
                    print(f"Número de linhas obtido: {num_linhas}")
                    if num_linhas is not None:
                        self.caminho_arquivo = caminho_arquivo
                        self.num_linhas = num_linhas
                        self.mostrar_frame_divisao()
                    else:
                        messagebox.showerror("Erro", "Não foi possível ler o arquivo.")
                except Exception as e:
                    print(f"Erro ao obter o número de linhas: {e}")
                    messagebox.showerror("Erro", f"Ocorreu um erro ao ler o arquivo: {e}")
            else:
                messagebox.showerror("Erro", "Arquivo não encontrado ou caminho inválido.")
        else:
            messagebox.showwarning("Aviso", "Nenhum arquivo selecionado.")

    def mostrar_frame_divisao(self):
        self.limpar_frames()
        self.frame_divisao = ttk.Frame(self)
        self.frame_divisao.pack(fill="both", expand=True)

        botao_voltar = ttk.Button(self.frame_divisao, text="Voltar", command=self.mostrar_frame_splitter)
        botao_voltar.pack(anchor='nw', pady=10, padx=10)

        label_linhas = ttk.Label(self.frame_divisao, text=f"Quantidade de Linhas: {self.num_linhas}")
        label_linhas.pack(pady=10)

        frame_entrada = ttk.Frame(self.frame_divisao)
        frame_entrada.pack(pady=5)

        label_partes = ttk.Label(frame_entrada, text="Em quantas partes deseja dividir:")
        label_partes.pack(side=tk.LEFT)

        entrada_partes = ttk.Entry(frame_entrada, width=10)
        entrada_partes.pack(side=tk.LEFT, padx=5)

        label_resultado = ttk.Label(self.frame_divisao, text="")
        label_resultado.pack(pady=5)

        # Barra de progresso
        self.progress_bar_splitter = ttk.Progressbar(self.frame_divisao, orient='horizontal', length=300, mode='determinate')
        self.progress_bar_splitter.pack(pady=10)
        self.progress_bar_splitter['value'] = 0

        def calcular():
            try:
                partes = int(entrada_partes.get())
                if partes <= 0 or partes > self.num_linhas:
                    messagebox.showerror("Erro", "O número de partes deve ser maior que zero e menor ou igual ao número de linhas.")
                    return
                linhas_por_parte = self.num_linhas // partes
                resto = self.num_linhas % partes
                label_resultado.config(text=f"Quantidade de linhas por parte: {linhas_por_parte} linhas\n"
                                            f"Última parte terá {linhas_por_parte + resto} linhas se houver resto.")
            except ValueError:
                messagebox.showerror("Erro", "Por favor, insira um número inteiro válido.")

        botao_calcular = ttk.Button(self.frame_divisao, text="Calcular", command=calcular)
        botao_calcular.pack(pady=5)

        botao_dividir = ttk.Button(self.frame_divisao, text="Dividir arquivo", command=lambda: self.dividir(entrada_partes))
        botao_dividir.pack(pady=10)

    def dividir(self, entrada_partes):
        try:
            partes = int(entrada_partes.get())
            if partes <= 0 or partes > self.num_linhas:
                messagebox.showerror("Erro", "O número de partes deve ser maior que zero e menor ou igual ao número de linhas.")
                return

            # Executar em uma thread separada
            thread = threading.Thread(target=self.thread_dividir, args=(partes,))
            thread.start()
        except ValueError:
            messagebox.showerror("Erro", "Por favor, insira um número inteiro válido.")

    def thread_dividir(self, partes):
        try:
            Splitter.dividir_arquivo(self.caminho_arquivo, partes, self.atualizar_progresso_splitter)
            messagebox.showinfo("Concluído", "Arquivo dividido com sucesso!")
            self.mostrar_frame_inicial()
        except Exception as e:
            print(f"Erro ao dividir o arquivo: {e}")
            messagebox.showerror("Erro", f"Ocorreu um erro ao dividir o arquivo: {e}")

    def atualizar_progresso_splitter(self, valor):
        self.progress_bar_splitter['value'] = valor
        self.update_idletasks()

    def mostrar_frame_merge(self):
        self.limpar_frames()
        self.frame_merge = ttk.Frame(self)
        self.frame_merge.pack(fill="both", expand=True)

        botao_voltar = ttk.Button(self.frame_merge, text="Voltar", command=self.mostrar_frame_inicial)
        botao_voltar.pack(anchor='nw', pady=10, padx=10)

        label_info = ttk.Label(self.frame_merge, text="Selecione os arquivos para mesclar:")
        label_info.pack(pady=10)

        botao_selecionar = ttk.Button(self.frame_merge, text="Selecionar Arquivos", command=self.selecionar_arquivos)
        botao_selecionar.pack(pady=5)

        # Caixa de texto para exibir os arquivos selecionados
        self.texto_arquivos = scrolledtext.ScrolledText(self.frame_merge, width=60, height=10, state='disabled', bg='#ffffff')
        self.texto_arquivos.pack(pady=10)

        # Botões "Unir Arquivos" e "Limpar Arquivos"
        botao_frame = ttk.Frame(self.frame_merge)
        botao_frame.pack(pady=5)

        botao_unir = ttk.Button(botao_frame, text="Unir Arquivos", command=self.unir_arquivos)
        botao_unir.pack(side=tk.LEFT, padx=5)

        botao_limpar = ttk.Button(botao_frame, text="Limpar Arquivos", command=self.limpar_arquivos)
        botao_limpar.pack(side=tk.LEFT, padx=5)

        # Barra de progresso
        self.progress_bar_merge = ttk.Progressbar(self.frame_merge, orient='horizontal', length=300, mode='determinate')
        self.progress_bar_merge.pack(pady=10)
        self.progress_bar_merge['value'] = 0

    def selecionar_arquivos(self):
        arquivos = filedialog.askopenfilenames(filetypes=[("Arquivos Excel", "*.xlsx")])
        if arquivos:
            self.arquivos_selecionados = [os.path.normpath(caminho) for caminho in arquivos]
            print(f"Arquivos selecionados: {self.arquivos_selecionados}")
            # Atualizar a caixa de texto com os nomes dos arquivos selecionados
            self.texto_arquivos.config(state='normal')
            self.texto_arquivos.delete('1.0', tk.END)
            for arquivo in self.arquivos_selecionados:
                self.texto_arquivos.insert(tk.END, f"{arquivo}\n")
            self.texto_arquivos.config(state='disabled')
        else:
            messagebox.showwarning("Aviso", "Nenhum arquivo selecionado.")

    def limpar_arquivos(self):
        self.arquivos_selecionados = []
        self.texto_arquivos.config(state='normal')
        self.texto_arquivos.delete('1.0', tk.END)
        self.texto_arquivos.config(state='disabled')
        messagebox.showinfo("Informação", "Lista de arquivos limpa.")

    def unir_arquivos(self):
        if not self.arquivos_selecionados:
            messagebox.showwarning("Aviso", "Nenhum arquivo selecionado para mesclar.")
            return
        try:
            # Executar em uma thread separada
            thread = threading.Thread(target=self.thread_unir_arquivos)
            thread.start()
        except Exception as e:
            print(f"Erro ao iniciar a thread de mesclagem: {e}")
            messagebox.showerror("Erro", f"Ocorreu um erro ao iniciar a mesclagem: {e}")

    def thread_unir_arquivos(self):
        try:
            # Definir o caminho de destino como "Resultado_Merge.xlsx" na mesma pasta do primeiro arquivo
            primeiro_arquivo = self.arquivos_selecionados[0]
            pasta_destino = os.path.dirname(primeiro_arquivo)
            destino = os.path.join(pasta_destino, "Resultado_Merge.xlsx")

            Merge.mesclar_arquivos(self.arquivos_selecionados, destino, self.atualizar_progresso_merge)
            messagebox.showinfo("Concluído", f"Arquivos mesclados com sucesso!\nArquivo salvo em: {destino}")
            self.limpar_arquivos()
            self.mostrar_frame_inicial()
        except Exception as e:
            print(f"Erro ao mesclar os arquivos: {e}")
            messagebox.showerror("Erro", f"Ocorreu um erro ao mesclar os arquivos: {e}")

    def atualizar_progresso_merge(self, valor):
        self.progress_bar_merge['value'] = valor
        self.update_idletasks()

if __name__ == "__main__":
    app = Aplicacao()
    app.mainloop()
