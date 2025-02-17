import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QGridLayout, QMessageBox
from PyQt5.QtCore import Qt
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

class CadastroAlunos(QWidget):
    def __init__(self):
        super().__init__()

        self.alunos = []
        self.initUI()

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

    def adicionar_aluno(self):
        nome = self.input_nome.text()
        turma = self.input_turma.text()
        email = self.input_email.text()

        if nome and turma and email:
            self.alunos.apped({'Nome': nome, 'Turma': turma, 'Email': email})
            self.input_nome.clear()
            self.input_nome.clear()
            self.input_turma.clear()
        else:
            QMessageBox.warning(self, 'Erro', 'Preencha todos os campos!')

    def cadastrar_e_enviar_emails(self):
        if self.alunos:
            df = pd.DataFrame(self.alunos)
            df.to_excel("cadastro_aluno.xlsx", index=False)

            for aluno in self.alunos:
                self.enviar_email(aluno['Email'])

            QMessageBox.information(self, 'Sucesso', 'Cadastro completo e email enviados!')
        else:
            QMessageBox.warning(self, 'Erro', 'Nenhum aluno cadastrado!')

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

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = CadastroAlunos()
    ex.show()
    sys.exit(app.exec_())