Projeto de Interface Gráfica para Processamento de Arquivos Excel
Este projeto consiste em uma aplicação com interface gráfica desenvolvida em Python, destinada a realizar o upload de arquivos Excel, extrair informações específicas de uma célula e adicioná-las a novas colunas. O objetivo é facilitar a análise de dados no ecossistema de indicadores.

Funcionalidades
Interface Gráfica (GUI): Desenvolvida com PyQt5 para facilitar a interação do usuário.
Upload de Arquivos Excel: Permite a seleção de arquivos Excel para processamento.
Extração de Informações: Extrai dados específicos de descrições de pedidos e insere essas informações em novas colunas.
Análise de Dados: Os dados extraídos podem ser analisados e exportados para novos arquivos Excel.
Mensagens de Status: Informações detalhadas sobre o progresso e o status do processamento são exibidas na interface.
Tecnologias Utilizadas
Python 3.x: Linguagem principal utilizada no desenvolvimento.
PyQt5: Utilizada para a criação da interface gráfica.
pandas: Utilizada para manipulação e análise de dados.
re: Biblioteca de expressões regulares utilizada para extração de padrões de texto.
openpyxl: Utilizada para operações com arquivos Excel.
Estrutura do Projeto
O projeto é composto pelos seguintes componentes principais:

main.py: Contém a definição da interface gráfica e a lógica principal do aplicativo.
extract_functions.py: Contém as funções para extração de dados das descrições dos pedidos.
Instalação
Clone o repositório:

bash
Copiar código
git clone https://github.com/seu_usuario/seu_repositorio.git
cd seu_repositorio
Crie e ative um ambiente virtual (opcional, mas recomendado):

No Windows:
bash
Copiar código
python -m venv venv
.\venv\Scripts\activate
No macOS/Linux:
bash
Copiar código
python3 -m venv venv
source venv/bin/activate
Instale as dependências:

bash
Copiar código
pip install -r requirements.txt
Uso
Executar o Aplicativo
Inicie a aplicação:

bash
Copiar código
python main.py
Selecione um arquivo Excel:

Clique no botão "Selecionar Arquivo" e escolha um arquivo Excel (.xlsx).
Processar o Arquivo:

Após selecionar o arquivo, clique no botão "Ok" para iniciar o processamento. As informações extraídas serão exibidas na tabela da interface.
Salvar o Arquivo Processado:

Clique no botão "Salvar como" para salvar os dados processados em um novo arquivo Excel.
Estrutura do Código
python
Copiar código
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QPushButton, QFileDialog, QTableWidget, QTableWidgetItem, QVBoxLayout, QTextEdit
import pandas as pd
import re

# Funções de extração e processamento (detalhadas anteriormente)

class MainWindow(QMainWindow):
    def __init__(self):
        # Definição da interface gráfica e lógica principal

    def select_file(self):
        # Lógica para seleção de arquivo

    def process_file(self):
        # Lógica para processar o arquivo selecionado

    def save_as(self):
        # Lógica para salvar o arquivo processado

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())
Funções de Extração
As funções de extração são responsáveis por identificar e extrair informações específicas das descrições dos pedidos.

extrair_cnpj(descricao): Extrai o CNPJ da descrição.
extrair_razao_social(descricao): Extrai a razão social da descrição.
extrair_numero_pedido(descricao): Extrai o número do pedido da descrição.
extrair_nota_fiscal(descricao): Extrai o número da nota fiscal da descrição.
extrair_valor_nf(descricao): Extrai o valor da nota fiscal da descrição.
extrair_data_emissao_nf(descricao): Extrai a data de emissão da nota fiscal da descrição.
Contribuição
Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou enviar pull requests com melhorias, correções de bugs ou novas funcionalidades.
