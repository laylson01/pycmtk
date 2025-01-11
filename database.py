import sqlite3

class Database:
    def __init__(self, db_name):
        self.conn = sqlite3.connect(db_name)
        self.cursor = self.conn.cursor()
        self.create_table()

    def create_table(self):
        self.cursor.execute(""" 
            CREATE TABLE IF NOT EXISTS clientes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT,
                cpf TEXT,
                rg TEXT,
                data_nascimento TEXT,
                sexo TEXT,
                telefone TEXT,
                endereco TEXT,
                pis_nis TEXT,
                nip TEXT,
                cei TEXT,
                rgp TEXT,
                email TEXT,
                data_inicio_atividade TEXT,
                titulo_eleitor TEXT
            )
        """)
        self.conn.commit()

    def insert_cliente(self, cliente_data):
        self.cursor.execute("""
            INSERT INTO clientes (nome, cpf, rg, data_nascimento, sexo, telefone, endereco, pis_nis, nip, cei, rgp, email, data_inicio_atividade, titulo_eleitor)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, cliente_data)
        self.conn.commit()

    def close(self):
        self.conn.close()
