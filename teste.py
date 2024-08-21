# -*- coding: utf-8 -*-
"""
Created on Tue Aug 20 18:11:02 2024

@author: bird
"""

import PyPDF2
import re
import openpyxl
import tkinter as tk
from tkinter import filedialog, messagebox
import ctypes

class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Extractor")
        self.root.geometry("800x600")  # Aumenta o tamanho da janela
        self.create_widgets()

    def create_widgets(self):
        self.label = tk.Label(self.root, text="Selecione o arquivo PDF para começar", font=("Arial", 14))
        self.label.pack(pady=20)

        self.select_button = tk.Button(self.root, text="Selecionar PDF", command=self.select_pdf, font=("Arial", 12), width=20, height=2, bd=4)
        self.select_button.pack(pady=10)

        self.text_area = tk.Text(self.root, height=15, width=80, font=("Arial", 12))
        self.text_area.pack(pady=20)

        self.status_label = tk.Label(self.root, text="", font=("Arial", 14))
        self.status_label.pack(pady=10)

        self.button_frame = tk.Frame(self.root)
        self.button_frame.pack(pady=5)

    def select_pdf(self):
        file_path = filedialog.askopenfilename(title="Selecione o arquivo PDF", filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.text_area.delete(1.0, tk.END)
            self.text_area.insert(tk.END, f"Arquivo selecionado: {file_path}\n")
            text = self.extract_info_from_pdf(file_path)
            if text:
                self.text_area.insert(tk.END, "Texto extraído do PDF:\n")
                self.text_area.insert(tk.END, text)
                self.ask_questions(text)
            else:
                messagebox.showerror("Erro", "Falha ao extrair texto do PDF.")

    def extract_info_from_pdf(self, file_path):
        try:
            with open(file_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                text = ""
                for page_num in range(len(reader.pages)):
                    page = reader.pages[page_num]
                    text += page.extract_text()
            return text
        except Exception as e:
            print(f"Erro ao ler o PDF: {e}")
            return None

    def ask_questions(self, text):
        self.info = self.get_specific_info(text)
        if not self.info:
            messagebox.showerror("Erro", "Informações necessárias não encontradas no PDF.")
            return

        self.status_label.config(text="Pergunta 1: Está PENDENTE ou PAGO?")
        self.create_buttons(["PENDENTE", "PAGO"], self.handle_status)

    def handle_status(self, status):
        self.info.append(status)
        self.status_label.config(text="Pergunta 2: Meio de Integralização (TRANSFERENCIA ou BOLETO)?")
        self.create_buttons(["TRANSFERENCIA", "BOLETO"], self.handle_meio_integralizacao)

    def handle_meio_integralizacao(self, meio_integralizacao):
        self.info.append(meio_integralizacao)
        self.status_label.config(text="Pergunta 3: Originador (FG ou FC)?")
        self.create_buttons(["FG", "FC"], self.handle_originador)

    def handle_originador(self, originador):
        self.info.append(originador)
        self.save_info()

    def create_buttons(self, options, command):
        for widget in self.button_frame.winfo_children():
            widget.destroy()

        for option in options:
            button = tk.Button(self.button_frame, text=option, command=lambda opt=option: command(opt), font=("Arial", 12), width=20, height=2, bd=4)
            button.pack(side=tk.LEFT, padx=5)

    def get_specific_info(self, text):
        try:
            numero_nota = re.search(r'TERMO DE EMISSÃO NO (\d+)', text).group(1)
            data = re.search(r'Data de Emissão\s*([\d/]+)', text).group(1)
            nome_emitente = re.search(r'Razão Social:\s*(.+?)\s*CNPJ/MF:', text, re.DOTALL).group(1).strip()
            cnpj = re.search(r'CNPJ/MF:\s*([\d./-]+)', text).group(1)
            credor = re.search(r'CREDOR\s*Razão Social:\s*(.+?)\s*CNPJ/MF:', text, re.DOTALL).group(1).strip()
            emissao = re.search(r'Valor Total da Emissão:\s*R\$ ([\d,.]+)', text).group(1)
            custo_nota = re.search(r'Custo da Emissão:\s*R\$ ([\d,.]+)', text).group(1)
            escrituracao_nota = re.search(r'Taxa de\s*Implantação/Remuneração:\s*\[X\] Única\s*R\$ ([\d,.]+)', text).group(1)

            custo_nota = re.sub(r'[^\d,]', '', custo_nota).replace(',', '.')
            escrituracao_nota = re.sub(r'[^\d,]', '', escrituracao_nota).replace(',', '.').strip('.')

            custo_nota_float = float(custo_nota)
            escrituracao_nota_float = float(escrituracao_nota)

            valor_estruturacao = custo_nota_float - escrituracao_nota_float

            custo_nota = f"{custo_nota_float:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            escrituracao_nota = f"{escrituracao_nota_float:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            valor_estruturacao = f"{valor_estruturacao:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

            estruturacao_extra = ""

            return [numero_nota, data, nome_emitente, cnpj, credor, emissao, custo_nota, escrituracao_nota, valor_estruturacao, estruturacao_extra]
        except AttributeError:
            return None

    def save_info(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Salvar arquivo como")
        if not save_path:
            print("Nenhum local de salvamento selecionado.")
            return

        workbook = openpyxl.Workbook()
        sheet = workbook.active
        headers = ["Número da Nota", "Data", "Nome do Emitente", "CNPJ", "Credor", "Emissão", "Custo da Nota", "Escrituração da Nota", "Valor da Estruturação", "Estruturação Extra", "Status", "Meio de Integralização", "Originador"]
        sheet.append(headers)
        sheet.append(self.info)
        workbook.save(save_path)

        self.status_label.config(text=f"Informações extraídas e salvas em '{save_path}'.")

        self.create_ok_button()

    def create_ok_button(self):
        for widget in self.button_frame.winfo_children():
            widget.destroy()

        ok_button = tk.Button(self.button_frame, text="OK", command=self.root.quit, font=("Arial", 12), width=20, height=2, bd=4)
        ok_button.pack(side=tk.LEFT, padx=5)

if __name__ == "__main__":
    # Fechar o console
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

    root = tk.Tk()
    app = PDFExtractorApp(root)
    root.mainloop()
