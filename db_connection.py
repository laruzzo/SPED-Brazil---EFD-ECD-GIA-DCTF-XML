import mysql.connector

class DBConnection:
    def __init__(self):
        self.connection = None
        self.host = "localhost"
        self.dbname = "fiscal_db"
        self.user = "root"
        self.password = "password"

    def connect_to_database(self):
        try:
            self.connection = mysql.connector.connect(
                host=self.host,
                database=self.dbname,
                user=self.user,
                password=self.password
            )
            print("Conexão com o banco de dados estabelecida com sucesso!")
        except mysql.connector.Error as err:
            print(f"Erro ao conectar ao banco de dados: {err}")

    def disconnect_from_database(self):
        if self.connection and self.connection.is_connected():
            self.connection.close()
            print("Conexão com o banco de dados encerrada.")

    def execute_query(self, query, params=None):
        if not self.connection or not self.connection.is_connected():
            print("Erro: Não há conexão ativa com o banco de dados.")
            return None
        try:
            cursor = self.connection.cursor()
            cursor.execute(query, params)
            self.connection.commit()
            return cursor
        except mysql.connector.Error as err:
            print(f"Erro ao executar a query: {err}")
            return None

    def fetch_all(self, query, params=None):
        if not self.connection or not self.connection.is_connected():
            print("Erro: Não há conexão ativa com o banco de dados.")
            return None
        try:
            cursor = self.connection.cursor(dictionary=True)
            cursor.execute(query, params)
            result = cursor.fetchall()
            cursor.close()
            return result
        except mysql.connector.Error as err:
            print(f"Erro ao buscar dados: {err}")
            return None



