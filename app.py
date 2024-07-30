import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QPushButton, QFileDialog, QTableWidget, QTableWidgetItem, QVBoxLayout, QTextEdit
import pandas as pd
import re

# Definir as funções de extração e processamento
def extrair_cnpj(descricao):
    pos_cnpj = descricao.find("CNPJ")
    if pos_cnpj != -1:
        start_pos = pos_cnpj + len("CNPJ")
        while start_pos < len(descricao) and not descricao[start_pos].isdigit():
            start_pos += 1
        end_pos = start_pos
        while end_pos < len(descricao) and descricao[end_pos] not in [' ', '']:
            end_pos += 1
        cnpj = descricao[start_pos:end_pos]
        return cnpj.strip()
    return None

def extrair_razao_social(descricao):
    padrao = r'Razão Social\s*:\s*([\s\S]*?)(?=\s*Número do Pedido|$)'
    resultado = re.search(padrao, str(descricao), re.DOTALL)
    if resultado:
        return resultado.group(1).strip()
    return None

def extrair_numero_pedido(descricao):
    padrao = r'Número do Pedido\D*(\d{10})'
    resultado = re.search(padrao, descricao)
    if resultado:
        return resultado.group(1)
    return None

def extrair_nota_fiscal(descricao):
    padrao = r'Nota Fiscal\D*(\d+)[^\d]*Valor da Nota Fiscal'
    resultado = re.search(padrao, descricao)
    if resultado:
        return resultado.group(1).strip()
    return None

def extrair_valor_nf(descricao):
    padrao = r'Valor da Nota Fiscal\D*(\d[\d\.,]+)[^\d]*Data de Emissão NF'
    resultado = re.search(padrao, descricao)
    if resultado:
        valor_nf = resultado.group(1).strip()
        valor_nf = re.sub(r'[^\d\.,]', '', valor_nf)  # Extrair apenas dígitos, ".", e ","
        return valor_nf
    return None

def extrair_data_emissao_nf(descricao):
    padrao = r'Data de Emissão NF\D*([\d-]+)[^\d]*'
    resultado = re.search(padrao, descricao)
    if resultado:
        return resultado.group(1).strip()
    return None

def processar_arquivo_excel(file_path, output_elem):
    try:
        output_elem.append("Iniciando importação do arquivo Excel...\n")
        df = pd.read_excel(file_path)  # Não especificamos a planilha aqui
        output_elem.append("Arquivo importado com sucesso!\n")

        output_elem.append("Iniciando processamento do arquivo...\n")
        # Aplicar as funções de extração às respectivas colunas do DataFrame
        df['Cnpj'] = df['Descrição'].apply(extrair_cnpj)
        df['Razão Social'] = df['Descrição'].apply(extrair_razao_social)
        df['Numero de pedido'] = df['Descrição'].apply(extrair_numero_pedido)
        df['Nota Fiscal'] = df['Descrição'].apply(extrair_nota_fiscal)
        df['Valor'] = df['Descrição'].apply(extrair_valor_nf)
        df['Data de Emissão NF'] = df['Descrição'].apply(extrair_data_emissao_nf)

        output_elem.append("Processamento concluído com sucesso! O arquivo já pode ser salvo\n")

        return df
    except Exception as e:
        output_elem.append(f"Erro ao processar arquivo Excel: {e}\n")
        return None

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Importar e Processar Arquivo Excel")
        self.setGeometry(100, 100, 800, 600)

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout()

        # Botão para selecionar arquivo
        self.file_path = None
        self.btn_select_file = QPushButton("Selecionar Arquivo")
        self.btn_select_file.clicked.connect(self.select_file)
        layout.addWidget(self.btn_select_file)

        # Botão "Ok"
        self.btn_ok = QPushButton("Ok", enabled=False)
        self.btn_ok.clicked.connect(self.process_file)
        layout.addWidget(self.btn_ok)

        # Tabela de resultados
        self.table = QTableWidget()
        layout.addWidget(self.table)

        # Botão "Salvar como"
        self.btn_save_as = QPushButton("Salvar como", enabled=False)
        self.btn_save_as.clicked.connect(self.save_as)
        layout.addWidget(self.btn_save_as)

        # Output para mensagens de status
        self.output = QTextEdit()
        layout.addWidget(self.output)

        central_widget.setLayout(layout)

        self.df = None  # DataFrame para armazenar os dados processados

    def select_file(self):
        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter("Excel Files (*.xlsx)")
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        if file_dialog.exec_():
            self.file_path = file_dialog.selectedFiles()[0]
            self.btn_ok.setEnabled(self.file_path.lower().endswith('.xlsx'))

    def process_file(self):
        if self.file_path and self.file_path.lower().endswith('.xlsx'):
            self.output.append("Processando arquivo...\n")
            self.df = processar_arquivo_excel(self.file_path, self.output)
            if self.df is not None:
                self.btn_save_as.setEnabled(True)  # Habilitar o botão "Salvar como"

    def save_as(self):
        if self.df is not None:
            file_dialog = QFileDialog(self)
            file_dialog.setNameFilter("Excel Files (*.xlsx)")
            file_dialog.setAcceptMode(QFileDialog.AcceptSave)
            if file_dialog.exec_():
                save_path = file_dialog.selectedFiles()[0]
                try:
                    self.df.to_excel(save_path, index=False)
                    self.output.append(f"Arquivo salvo com sucesso em:\n{save_path}\n")
                except Exception as e:
                    self.output.append(f"Erro ao salvar arquivo Excel:\n{e}\n")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())
