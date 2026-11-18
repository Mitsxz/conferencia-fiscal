import sqlite3
from config import Config

def get_db_connection():
    conn = sqlite3.connect(Config.DATABASE_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS conferencias (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            titulo TEXT NOT NULL,
            arquivo_livro_original TEXT NOT NULL,
            arquivo_livro_salvo TEXT NOT NULL,
            arquivos_planilhas_json TEXT NOT NULL,
            config_cruzamento_json TEXT NOT NULL,
            tolerancia_valor REAL DEFAULT 0.01,
            total_analisados INTEGER DEFAULT 0,
            total_correspondencias INTEGER DEFAULT 0,
            total_divergencias INTEGER DEFAULT 0,
            total_ausentes_planilha INTEGER DEFAULT 0,
            total_ausentes_livro INTEGER DEFAULT 0,
            total_duplicidades INTEGER DEFAULT 0,
            percentual_correspondencia REAL DEFAULT 0.0,
            status TEXT DEFAULT 'CONCLUIDO',
            criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS conferencia_itens (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            conferencia_id INTEGER NOT NULL,
            chave_identificacao TEXT NOT NULL,
            categoria_resultado TEXT NOT NULL,
            tipo_divergencia TEXT,
            dados_livro_json TEXT,
            dados_planilha_json TEXT,
            diferencas_json TEXT,
            status_revisao TEXT DEFAULT 'PENDENTE',
            observacao TEXT,
            FOREIGN KEY (conferencia_id) REFERENCES conferencias (id) ON DELETE CASCADE
        )
    ''')

    conn.commit()
    conn.close()
