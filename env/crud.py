import sqlite3
import os
from datetime import datetime

class AppBD:
    def __init__(self):
        self.abrirConexao()

    def abrirConexao(self):
        try:
            caminho_banco_dados = "env"
            caminho_banco_dados = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Data_Base.db")
            self.connection = sqlite3.connect(caminho_banco_dados)
            self.create_tables() 
        except sqlite3.Error as error:
            print("Falha ao se conectar ao Banco de dados", error)

    def create_tables(self):
        create_aluno_table_query = """
        CREATE TABLE IF NOT EXISTS Aluno(
            cpf TEXT NOT NULL PRIMARY KEY,
            nome TEXT NOT NULL,
            data_nascimento DATE NOT NULL,
            sexo TEXT CHECK(sexo IN ('Feminino', 'Masculino')) NOT NULL
            );
        """
        create_disciplina_table_query = """
        CREATE TABLE IF NOT EXISTS Disciplina(
            disciplina TEXT CHECK(disciplina IN ('Biologia', 'Filosofia', 'Física', 'Geografia', 'Inglês', 'História', 'Literatura', 'Matemática', 'Português', 'Química', 'Sociologia')) PRIMARY KEY 
            );
        """
        create_notas_table_query = """
        CREATE TABLE IF NOT EXISTS Notas(
            id_Aluno TEXT NOT NULL,
            disciplina TEXT CHECK(disciplina IN ('Biologia', 'Filosofia', 'Física', 'Geografia', 'Inglês', 'História', 'Literatura', 'Matemática', 'Português', 'Química', 'Sociologia')),
            AV1 REAL NOT NULL,
            AV2 REAL NOT NULL,
            media REAL NOT NULL,            
            FOREIGN KEY(id_Aluno) REFERENCES Aluno(cpf) ON DELETE CASCADE,
            FOREIGN KEY(disciplina) REFERENCES Disciplina(disciplina)
            );
        """
        try:
            cursor = self.connection.cursor()
            cursor.execute(create_aluno_table_query)
            cursor.execute(create_disciplina_table_query)
            cursor.execute(create_notas_table_query)
            self.connection.commit()
        except sqlite3.Error as error:
            print("Falha ao criar tabelas", error)
        finally:
            if self.connection:
                cursor.close()

# ALUNOS
    def insert_Aluno(self, cpf, nome, data_nascimento, sexo):
        self.abrirConexao()
        insert_query = "INSERT INTO Aluno (cpf, nome, data_nascimento, sexo) VALUES (?, ?, ?, ?)"
        try:
            cursor = self.connection.cursor()
            data_nascimento = datetime.strptime(data_nascimento, "%d/%m/%Y").date()
            cursor.execute(insert_query, (cpf, nome, data_nascimento, sexo))
            self.connection.commit()
            print("Aluno cadastrado com sucesso")
        except sqlite3.Error as error:
            print("Falha ao cadastrar Aluno", error)
        finally:
            if self.connection:
                cursor.close()
                self.connection.close()
                print("A conexão com o SQLite foi fechada.")

    def select_Aluno(self):
        self.abrirConexao()
        select_query = "SELECT * FROM Aluno"
        Alunos = []
        try:
            cursor = self.connection.cursor()
            cursor.execute(select_query)
            Alunos = cursor.fetchall()
        except sqlite3.Error as error:
            print("Falha ao retornar alunos", error)
        finally:
            if self.connection:
                self.connection.close()
                print("A conexão com o SQLite foi fechada.")
        return Alunos

    def update_Aluno(self, cpf, nome, data_nascimento, sexo):
        self.abrirConexao()
        update_query = "UPDATE Aluno SET nome = ?, data_nascimento = ?, sexo = ? WHERE cpf = ?"
        try:
            cursor = self.connection.cursor()
            cursor.execute(update_query, (nome, data_nascimento, sexo, cpf))
            self.connection.commit()
            print("Aluno atualizado com sucesso")
        except sqlite3.Error as error:
            print("Falha ao atualizar o aluno", error)
        finally:
            if self.connection:
                cursor.close()
                self.connection.close()
                print("A conexão com o SQLite foi fechada.")

    def delete_Aluno(self, cpf):
        self.abrirConexao()
        delete_notas_query = "DELETE FROM Notas WHERE id_Aluno = ?"
        delete_aluno_query = "DELETE FROM Aluno WHERE cpf = ?"
        try:
            cursor = self.connection.cursor()
            cursor.execute(delete_notas_query, (cpf,))
            cursor.execute(delete_aluno_query, (cpf,))
            self.connection.commit()
            print("Aluno e suas notas foram deletados com sucesso")
        except sqlite3.Error as error:
            print("Falha ao deletar aluno e suas notas", error)
        finally:
            if self.connection:
                cursor.close()
                self.connection.close()
                print("A conexão foi fechada")

# NOTAS/DISCIPLINAS
    def insert_Notas(self, id_Aluno, disciplina, AV1, AV2):
        self.abrirConexao()
        AV1 = float(AV1)  
        AV2 = float(AV2)  
        media = (AV1 + AV2) / 2

        select_disciplina_query = "SELECT disciplina FROM Disciplina WHERE disciplina = ? COLLATE NOCASE"
        cursor = self.connection.cursor()
        cursor.execute(select_disciplina_query, (disciplina,))
        existing_discipline = cursor.fetchone()

        if existing_discipline:
            print("A disciplina já existe.")

        else:
            insert_disciplina_query = "INSERT INTO Disciplina (disciplina) VALUES (?)"
            try:
                cursor.execute(insert_disciplina_query, (disciplina,))
                self.connection.commit()
                print("Disciplina inserida com sucesso")
            except sqlite3.Error as error:
                print("Falha ao inserir disciplina:", error)
 
        insert_query = "INSERT INTO Notas (id_Aluno, disciplina, AV1, AV2, media) VALUES (?, ?, ?, ?, ?)"

        select_query = "SELECT disciplina FROM Notas WHERE id_Aluno = ? AND disciplina = ?"
        cursor = self.connection.cursor()
        cursor.execute(select_query, (id_Aluno, disciplina))
        existing_discipline_N = cursor.fetchone()

        if existing_discipline_N:
            print("Já existe uma entrada para essa disciplina.")
        else:        
            try:
                cursor = self.connection.cursor()
                cursor.execute(insert_query, (id_Aluno, disciplina, AV1, AV2, media))
                self.connection.commit()
                print("Notas inseridas com sucesso")
            except sqlite3.Error as error:
                print("Falha ao inserir notas", error)
            finally:
                if self.connection:
                    cursor.close()
                    self.connection.close()
                    print("A conexão com o SQLite foi fechada.")

    def update_Notas(self, cpf, disciplina, AV1, AV2):
        self.abrirConexao()
        media = (AV1 + AV2) / 2 
        update_query = "UPDATE Notas SET AV1 = ?, AV2 = ?, media = ? WHERE id_Aluno = ? AND disciplina = ?"
        try:
            cursor = self.connection.cursor()
            cursor.execute(update_query, (AV1, AV2, media, cpf, disciplina))
            self.connection.commit()
            print("Notas atualizadas com sucesso")
        except sqlite3.Error as error:
            print("Falha ao atualizar as notas:", error)
        finally:
            try:
                if self.connection:
                    self.connection.close()
                    print("A conexão com o SQLite foi fechada.")
            except sqlite3.Error as error:
                print("Erro ao fechar a conexão com o SQLite:", error)


    def delete_Notas(self, cpf, disciplina):
        self.abrirConexao()
        delete_query = "DELETE FROM Notas WHERE id_Aluno = ? AND disciplina = ?"

        try:
            cursor = self.connection.cursor()
            cursor.execute(delete_query, (cpf, disciplina))
            self.connection.commit()
            print("Notas deletadas com sucesso")
        except sqlite3.Error as error:
            print("Falha ao deletar notas", error)
        finally:
            if self.connection:
                cursor.close()
                self.connection.close()
                print("A conexão foi fechada")

# BOLETIM
    def select_Boletim(self):
        self.abrirConexao()
        select_query = """SELECT Aluno.nome AS Nome,
                            Notas.disciplina AS disciplina,
                            Notas.AV1 AS AV1,
                            Notas.AV2 AS AV2,
                            Notas.media AS Média,
                            CASE
                                WHEN Notas.media >= 6.0 THEN 'Aprovado'
                                ELSE 'Reprovado'
                            END AS Status
                    FROM Notas
                    INNER JOIN Aluno ON Notas.id_Aluno = Aluno.cpf
                    INNER JOIN Disciplina ON Notas.disciplina = Disciplina.disciplina
                    ORDER BY Aluno.cpf"""

        boletim = []
        try:
            cursor = self.connection.cursor()
            cursor.execute(select_query)
            boletim = cursor.fetchall()

        except sqlite3.Error as error:
            print("Falha no boletim", error)

        finally:
            if self.connection:
                self.connection.close()
                print("A conexão com o SQLite foi fechada.")

        return boletim
    

