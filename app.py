import sys
import pandas as pd
import openpyxl
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QGridLayout, QMessageBox
from PyQt5.QtCore import Qt
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
'''
sys: Usado para manipular o sistema, como encerrar a aplicação.

pandas: Biblioteca para manipulação de dados, usada aqui para criar um DataFrame a salvar em Excel.

PyQt5: Biblioteca para criar interfacez gráficas

smtplib: Usado para enviar e-mails.

email.mime.text e email.mime.multipart: Usados para criar e formatar e-mails.
'''

class CadastroAlunos(QWidget):
    def __init__(self):
        super().__init__()

        self.alunos = []
        self.initUI()
        '''
        CadastroAlunos: Classe principal que herde de QWidget, que é a base para todas as interfaces gráficas no PyQt5

Init: Método construtor que inicializa a lista alunos e chama initUI para configurar a interface
        '''

    def initUI(self):
        self.setGeometry(100, 100, 400, 200)
        self.setWindowTitle('Cadastro de Alunos')

        layout = QGridLayout()
        self.setLayout(layout)

        label_nome = QLabel('Nome do aluno:')
        self.input_nome = QLineEdit()
        layout.addWidget(label_nome, 0, 0)
        layout.addWidget(self.input_nome, 0, 1)

        label_turma = QLabel('Turma:')
        self.input_turma = QLineEdit()
        layout.addWidget(label_turma, 1, 0)
        layout.addWidget(self.input_turma, 1, 1)

        label_email = QLabel('Email:')
        self.input_email = QLineEdit()
        layout.addWidget(label_email, 2, 0)
        layout.addWidget(self.input_email, 2, 1)

        botao_adicionar = QPushButton('Adicionar aluno')
        botao_adicionar.clicked.connect(self.adicionar_aluno)
        layout.addWidget(botao_adicionar, 3, 0, 1, 2)

        botao_cadastrar = QPushButton('Cadastrar e enviar emails')
        botao_cadastrar.clicked.connect(self.cadastrar_e_enviar_emails)
        layout.addWidget(botao_cadastrar, 4, 0, 1, 2)

        self.remetente = "concursopilargames@gmail.com"
        self.senha = "@PilarMaturana2"
        '''
        setGeometry: Define a posição e o tamanho da janela.

setWindowTitle: Define o título da janela.

QGridLayout: Layout que organiza os widgets em uma grade.

QLabel: Cria rótulos para os campos de entrada.

QLineEdit: Cria campos de entrada de texto.

QPushButton: Cria botões que executam ações quando clicados.

clicked.connect: Conecta o evento de clique do botão a um método.

remetente e senha: Configurações do e-mail do remetente.
       
        '''

    def adicionar_aluno(self):
        nome = self.input_nome.text()
        turma = self.input_turma.text()
        email = self.input_email.text()

        if nome and turma and email:
                self.alunos.append({'Nome': nome, 'Turma': turma, 'Email': email})
                self.input_nome.clear()
                self.input_turma.clear()
                self.input_email.clear()
        else:
                QMessageBox.warning(self, 'Erro, Preencha todos os campo!')
        '''
            text(): Obtém o texto dos campos de entrada.

append: Adiciona um dicionário com os dados do aluno à lista alunos.

clear: Limpa os campos de entrada após adicionar o aluno.

QMessageBox.warning: Exibe uma mensagem de erro se algum campo estiver vazio.
           
            '''

    def cadastrar_e_enviar_emails(self):
        if self.alunos:
            df = pd.DataFrame(self.alunos)
            df.to_excel("cadastro_aluno.xlsx", index=False)

            for aluno in self.alunos:
                self.enviar_email(aluno['Email'])

            QMessageBox.information(self, 'Sucesso', 'Cadastro completo e email enviados!')
        else:
            QMessageBox.warning(self, 'Erro', 'Nenhum aluno cadastrado!')
            '''
            pd.DataFrame: Cria um DataFrame com a lista de alunos.

to_excel: Salva o DataFrame em um arquivo Excel.

enviar_email: Chama o método para enviar e-mail para cada aluno.

QMessageBox.information: Exibe uma mensagem de sucesso.

            '''

    def enviar_email(self, destinatario):
        msg = MIMEMultipart()
        msg['From'] = self.remetente
        msg['To'] = destinatario
        msg['Subject'] = "Confirmação de Inscrição - Concurso de Jogos"

        corpo = "Olá, sua inscrição no concurso de jogos foi confirmada!"
        msg.attach(MIMEText(corpo, 'plain'))

        servidor = smtplib.SMTP('smtp.gmail.com', 587)
        servidor.starttls()
        servidor.login(self.remetente, self.senha)
        texto = msg.as_string()
        servidor.sendmail(self.remetente, destinatario, texto)
        servidor.quit()
        '''
        MIMEMultipart: Cria uma mensagem de e-mail com múltiplas partes.

MIMEText: Adiciona o corpo do e-mail.

smtplib.SMTP: Conecta ao servidor SMTP do Gmail.

starttls: Inicia uma conexão segura.

login: Faz login no servidor de e-mail.

sendmail: Envia o e-mail.

quit: Fecha a conexão com o servidor.
       
        '''

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = CadastroAlunos()
    ex.show()
    sys.exit(app.exec_())
    '''
    QApplication: Cria a aplicação PyQt5.

CadastroAlunos: Instancia a classe principal.

show: Exibe a janela.

app.exec_: Inicia o loop de eventos da aplicação.

sys.exit: Encerra a aplicação corretamente.
'''