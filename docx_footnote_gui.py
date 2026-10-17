#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Interface Gráfica (GUI) para o Footnote Injector by Jair Lima.
Desenvolvida utilizando CustomTkinter para oferecer um visual moderno e premium.
Permite selecionar arquivos de entrada, saída, dicionário personalizado e executar o injetor com logs em tempo real.
"""

import os
import sys
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox

from docx_footnote_injector import FootnoteInjector, carregar_dicionario

# Configurações globais de estilo
ctk.set_appearance_mode("System")  # Segue o tema do Windows (Escuro/Claro)
ctk.set_default_color_theme("blue") # Tema de cores padrão


def resource_path(relative_path):
    """Obtém o caminho absoluto para o recurso, compatível com PyInstaller e desenvolvimento."""
    try:
        # PyInstaller cria uma pasta temporária e armazena o caminho em _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)



class FootnoteInjectorGUI(ctk.CTk):
    """Janela principal do aplicativo Footnote Injector."""

    def __init__(self):
        super().__init__()

        # Configurações da Janela
        self.title("Footnote Injector by Jair Lima")
        self.geometry("700x600")
        self.minsize(600, 500)

        # Variáveis de Controle
        self.input_docx_path = ctk.StringVar()
        self.dict_path = ctk.StringVar()
        self.output_docx_path = ctk.StringVar()
        self.only_first_occ = ctk.BooleanVar(value=True)

        # Caminho padrão do dicionário (embutido no executável)
        self.default_dict_path = resource_path("dicionario_padrao.json")

        self.setup_ui()

        # Inicializa o hook de drag and drop nativo do Windows
        try:
            import windnd
            windnd.hook_dropfiles(self, func=self.on_file_drop)
        except Exception as e:
            print(f"Erro ao inicializar o drag and drop: {e}")


    def setup_ui(self):
        """Monta a estrutura de widgets da interface."""
        # Grid layout principal (1 coluna, múltiplas linhas)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(5, weight=1) # O log de progresso expandirá na vertical

        # 1. Cabeçalho (Header)
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        header_frame.grid_columnconfigure(0, weight=1)

        title_label = ctk.CTkLabel(
            header_frame, 
            text="Footnote Injector by Jair Lima", 
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.grid(row=0, column=0, sticky="w")

        subtitle_label = ctk.CTkLabel(
            header_frame, 
            text="Injeção automatizada e inteligente de notas de rodapé explicativas em documentos Word.",
            font=ctk.CTkFont(size=13),
            text_color="gray"
        )
        subtitle_label.grid(row=1, column=0, sticky="w", pady=(5, 0))

        # 2. Frame de Entrada do Documento Word
        input_frame = ctk.CTkFrame(self)
        input_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        input_frame.grid_columnconfigure(1, weight=1)

        input_label = ctk.CTkLabel(input_frame, text="Documento Word (.docx):", font=ctk.CTkFont(weight="bold"))
        input_label.grid(row=0, column=0, padx=15, pady=10, sticky="w")

        self.input_entry = ctk.CTkEntry(
            input_frame, 
            textvariable=self.input_docx_path, 
            placeholder_text="Selecione o arquivo de entrada..."
        )
        self.input_entry.grid(row=0, column=1, padx=(5, 10), pady=10, sticky="ew")

        input_btn = ctk.CTkButton(
            input_frame, 
            text="Buscar", 
            width=100, 
            command=self.browse_input_docx
        )
        input_btn.grid(row=0, column=2, padx=(0, 15), pady=10)

        # 3. Frame do Dicionário
        dict_frame = ctk.CTkFrame(self)
        dict_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        dict_frame.grid_columnconfigure(1, weight=1)

        dict_label = ctk.CTkLabel(dict_frame, text="Dicionário (JSON/CSV):", font=ctk.CTkFont(weight="bold"))
        dict_label.grid(row=0, column=0, padx=15, pady=10, sticky="w")

        self.dict_entry = ctk.CTkEntry(
            dict_frame, 
            textvariable=self.dict_path, 
            placeholder_text="Usar padrão de tecnologia embutido (deixe em branco)..."
        )
        self.dict_entry.grid(row=0, column=1, padx=(5, 10), pady=10, sticky="ew")

        dict_btn = ctk.CTkButton(
            dict_frame, 
            text="Buscar", 
            width=100, 
            command=self.browse_dict
        )
        dict_btn.grid(row=0, column=2, padx=(0, 15), pady=10)

        # 4. Frame de Saída do Documento
        output_frame = ctk.CTkFrame(self)
        output_frame.grid(row=3, column=0, padx=20, pady=10, sticky="ew")
        output_frame.grid_columnconfigure(1, weight=1)

        output_label = ctk.CTkLabel(output_frame, text="Documento de Saída:", font=ctk.CTkFont(weight="bold"))
        output_label.grid(row=0, column=0, padx=15, pady=10, sticky="w")

        self.output_entry = ctk.CTkEntry(
            output_frame, 
            textvariable=self.output_docx_path, 
            placeholder_text="Caminho para salvar o documento resultante..."
        )
        self.output_entry.grid(row=0, column=1, padx=(5, 10), pady=10, sticky="ew")

        output_btn = ctk.CTkButton(
            output_frame, 
            text="Salvar Como", 
            width=100, 
            command=self.browse_output_docx
        )
        output_btn.grid(row=0, column=2, padx=(0, 15), pady=10)

        # 5. Opções Extras e Ações
        options_frame = ctk.CTkFrame(self, fg_color="transparent")
        options_frame.grid(row=4, column=0, padx=20, pady=10, sticky="ew")
        options_frame.grid_columnconfigure(0, weight=1)

        first_occ_chk = ctk.CTkCheckBox(
            options_frame, 
            text="Injetar nota apenas na primeira ocorrência de cada termo", 
            variable=self.only_first_occ
        )
        first_occ_chk.grid(row=0, column=0, sticky="w", padx=10, pady=5)

        self.run_btn = ctk.CTkButton(
            options_frame, 
            text="INJETAR NOTAS DE RODAPÉ", 
            font=ctk.CTkFont(size=14, weight="bold"),
            height=40,
            command=self.start_injection_thread
        )
        self.run_btn.grid(row=0, column=1, padx=10, pady=5, sticky="e")

        # 6. Painel de Log de Progresso
        log_frame = ctk.CTkFrame(self)
        log_frame.grid(row=5, column=0, padx=20, pady=(10, 20), sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(1, weight=1)

        log_title = ctk.CTkLabel(log_frame, text="Log de Status e Progresso:", font=ctk.CTkFont(weight="bold"))
        log_title.grid(row=0, column=0, padx=15, pady=(10, 5), sticky="w")

        self.log_textbox = ctk.CTkTextbox(log_frame, activate_scrollbars=True)
        self.log_textbox.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="nsew")
        self.log_textbox.configure(state="disabled")

        # Log inicial informando status
        self.log_message("[*] Aplicativo inicializado. Aguardando arquivos para iniciar o processamento.\n")

    def log_message(self, message):
        """Adiciona uma mensagem de log na caixa de texto na thread principal."""
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", message + "\n")
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")

    def browse_input_docx(self):
        """Abre seletor para escolher o arquivo DOCX de entrada."""
        path = filedialog.askopenfilename(
            title="Selecione o arquivo Word de entrada",
            filetypes=[("Arquivos Word", "*.docx")]
        )
        if path:
            self.input_docx_path.set(path)
            # Auto-sugere o nome de saída se não estiver definido
            if not self.output_docx_path.get():
                dir_name, file_name = os.path.split(path)
                name_parts = os.path.splitext(file_name)
                suggested_output = os.path.join(
                    dir_name, 
                    f"{name_parts[0]}_com_notas{name_parts[1]}"
                )
                self.output_docx_path.set(suggested_output)
            self.log_message(f"[+] Arquivo de entrada selecionado: {os.path.basename(path)}")

    def browse_dict(self):
        """Abre seletor para escolher o arquivo de dicionário."""
        path = filedialog.askopenfilename(
            title="Selecione o arquivo do Dicionário",
            filetypes=[("Dicionário (JSON/CSV)", "*.json;*.csv"), ("JSON", "*.json"), ("CSV", "*.csv")]
        )
        if path:
            self.dict_path.set(path)
            self.log_message(f"[+] Dicionário selecionado: {os.path.basename(path)}")

    def browse_output_docx(self):
        """Abre seletor para salvar o arquivo de saída."""
        path = filedialog.asksaveasfilename(
            title="Salvar documento resultante como",
            defaultextension=".docx",
            filetypes=[("Arquivos Word", "*.docx")]
        )
        if path:
            self.output_docx_path.set(path)
            self.log_message(f"[+] Arquivo de saída definido: {os.path.basename(path)}")

    def on_file_drop(self, files):
        """Callback acionado ao arrastar e soltar arquivos na janela."""
        try:
            # Decodifica os bytes usando a codificação nativa do sistema de arquivos do Windows
            decoded_files = [f.decode(sys.getfilesystemencoding(), errors='ignore') for f in files]
            if not decoded_files:
                return

            docx_files = [f for f in decoded_files if f.lower().endswith('.docx')]
            dict_files = [f for f in decoded_files if f.lower().endswith(('.json', '.csv'))]

            # Processa arquivo DOCX arrastado
            if docx_files:
                path = docx_files[0]
                self.input_docx_path.set(path)
                self.log_message(f"[+] Arquivo Word arrastado: {os.path.basename(path)}")
                
                # Auto-sugere o caminho de saída
                if not self.output_docx_path.get():
                    dir_name, file_name = os.path.split(path)
                    name_parts = os.path.splitext(file_name)
                    suggested_output = os.path.join(
                        dir_name, 
                        f"{name_parts[0]}_com_notas{name_parts[1]}"
                    )
                    self.output_docx_path.set(suggested_output)

            # Processa dicionário arrastado
            if dict_files:
                path = dict_files[0]
                self.dict_path.set(path)
                self.log_message(f"[+] Dicionário arrastado: {os.path.basename(path)}")

        except Exception as e:
            self.log_message(f"[-] Erro ao processar o arquivo solto: {e}")

    def start_injection_thread(self):
        """Inicia a injeção em uma thread secundária para evitar travamento da GUI."""
        input_file = self.input_docx_path.get()
        output_file = self.output_docx_path.get()
        dictionary_file = self.dict_path.get()

        if not input_file or not os.path.exists(input_file):
            messagebox.showerror("Erro de Entrada", "Por favor, selecione um arquivo Word (.docx) de entrada válido.")
            return

        if not output_file:
            messagebox.showerror("Erro de Saída", "Por favor, defina um caminho para salvar o documento resultante.")
            return

        # O dicionário agora é opcional na interface (se vazio, usa o padrão embutido)
        if dictionary_file and not os.path.exists(dictionary_file):
            messagebox.showerror("Erro de Dicionário", "O arquivo de dicionário especificado não foi localizado.")
            return

        # Desabilita o botão para evitar múltiplos cliques
        self.run_btn.configure(state="disabled", text="PROCESSANDO...")
        
        # Dispara thread
        thread = threading.Thread(target=self.run_injection)
        thread.daemon = True
        thread.start()

    def run_injection(self):
        """Rotina da thread de injeção."""
        input_file = self.input_docx_path.get()
        output_file = self.output_docx_path.get()
        dictionary_file = self.dict_path.get()
        only_first = self.only_first_occ.get()

        try:
            if not dictionary_file:
                self.log_message("[*] Dicionário não especificado. Carregando dicionário de tecnologia embutido...")
                dict_path = resource_path("dicionario_padrao.json")
                dict_data = carregar_dicionario(dict_path)
            else:
                self.log_message(f"\n[*] Lendo dicionário a partir de: {dictionary_file}")
                dict_data = carregar_dicionario(dictionary_file)

            self.log_message(f"[*] Iniciando injeção em: {os.path.basename(input_file)}")
            self.log_message(f"[*] Processamento apenas em primeira ocorrência: {only_first}")

            injector = FootnoteInjector(dict_data)
            injector.inject(input_file, output_file, only_first_occurrence=only_first)

            self.log_message(f"[+] Injeção concluída com sucesso!")
            self.log_message(f"[+] Novo arquivo gerado em: {output_file}")
            self.log_message(f"[+] Total de notas de rodapé inseridas no documento: {injector.next_footnote_id - 1}")

            # Gera o relatório físico
            report_path = os.path.splitext(output_file)[0] + "_relatorio.txt"
            injector.gerar_relatorio(report_path)

            # Exibe o resumo do relatório na caixa de logs da GUI
            self.log_message("\n[+] Resumo dos Termos Explicados:")
            self.log_message("-" * 45)
            self.log_message(f"    {'Termo':<25} | {'Insercoes':<10}")
            self.log_message("-" * 45)
            for term, count in sorted(injector.insertion_report.items()):
                self.log_message(f"    {term:<25} | {count:<10}")
            self.log_message("-" * 45)
            self.log_message(f"[+] Relatório detalhado salvo em: {os.path.basename(report_path)}")

            messagebox.showinfo("Sucesso", f"O arquivo foi processado com sucesso!\nTotal de notas adicionadas: {injector.next_footnote_id - 1}\n\nRelatório salvo ao lado do arquivo de saída.")

        except Exception as e:
            self.log_message(f"[-] Erro crítico durante o processamento: {e}")
            messagebox.showerror("Erro no Processamento", f"Ocorreu um erro ao processar o arquivo:\n{e}")

        finally:
            # Reabilita o botão na thread principal
            self.run_btn.configure(state="normal", text="INJETAR NOTAS DE RODAPÉ")


def main():
    app = FootnoteInjectorGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
