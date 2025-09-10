# app.py
from flask import Flask, render_template, request, redirect, url_for, session, send_file
import sqlite3
import pandas as pd
from werkzeug.security import generate_password_hash, check_password_hash
import os

app = Flask(__name__)
app.secret_key = "chave_secreta"  # troque para algo seguro

DB_NAME = "funcionarios.db"


# ------------------ Banco de Dados ------------------
def importar_planilha_para_db(excel_path="headcount 2.xlsx"):
    df = pd.read_excel(excel_path, sheet_name="Planilha1")

    # Renomear colunas para bater com os nomes do banco (sem espaços/especiais)
    df = df.rename(columns={
        "FUNCIONÁRIO(A):": "funcionario",
        "Matricula": "matricula",
        "Admissão": "admissao",
        "CPF": "cpf",
        "RG": "rg",
        "Nascimento": "nascimento",
        "CARGO:": "cargo",
        "PERFIL": "perfil",
        "Interno/Volante": "interno_volante",
        "LOCALIDADE": "localidade",
        "FILA": "fila",
        "FIELD": "field",
        "Email Corporativo": "email_corporativo",
        "Cel Corporativo": "cel_corporativo",
        "Cel Pessoal": "cel_pessoal",
        "Municipio de Moradia": "municipio",
        "Bairro": "bairro",
        "Endereço": "endereco",
        "Veiculo": "veiculo",
        "Placa": "placa",
        "Supervisor": "supervisor"
    })

    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()

    # Inserir os dados se a tabela estiver vazia
    c.execute("SELECT COUNT(*) FROM funcionarios")
    if c.fetchone()[0] == 0:
        df.to_sql("funcionarios", conn, if_exists="append", index=False)
        print("✅ Dados importados da planilha para o banco.")
    else:
        print("⚠️ Tabela de funcionários já possui dados, não foi necessário importar.")

    conn.close()

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()

    # Tabela de login
    c.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL
        )
    """)

    # Tabela de funcionários (baseada na planilha1)
    c.execute("""
        CREATE TABLE IF NOT EXISTS funcionarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            funcionario TEXT,
            matricula TEXT,
            admissao TEXT,
            cpf TEXT,
            rg TEXT,
            nascimento TEXT,
            cargo TEXT,
            perfil TEXT,
            interno_volante TEXT,
            localidade TEXT,
            fila TEXT,
            field TEXT,
            email_corporativo TEXT,
            cel_corporativo TEXT,
            cel_pessoal TEXT,
            municipio TEXT,
            bairro TEXT,
            endereco TEXT,
            veiculo TEXT,
            placa TEXT,
            supervisor TEXT
        )
    """)

    # Usuário admin padrão
    c.execute("SELECT * FROM usuarios WHERE username=?", ("admin",))
    if not c.fetchone():
        c.execute("INSERT INTO usuarios (username, password_hash) VALUES (?, ?)",
                  ("admin", generate_password_hash("admin")))
    conn.commit()
    conn.close()


# ------------------ Autenticação ------------------
@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]

        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute("SELECT * FROM usuarios WHERE username=?", (username,))
        user = c.fetchone()
        conn.close()

        if user and check_password_hash(user[2], password):
            session["user"] = username
            return redirect(url_for("listar_funcionarios"))
        else:
            return render_template("login.html", error="Usuário ou senha inválidos.")

    return render_template("login.html")


@app.route("/logout")
def logout():
    session.pop("user", None)
    return redirect(url_for("login"))


# ------------------ CRUD Funcionários ------------------
@app.route("/funcionarios")
def listar_funcionarios():
    if "user" not in session:
        return redirect(url_for("login"))

    filtro = request.args.get("filtro", "")
    valor = request.args.get("valor", "")

    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()

    if filtro and valor:
        query = f"SELECT * FROM funcionarios WHERE {filtro} LIKE ?"
        c.execute(query, (f"%{valor}%",))
    else:
        c.execute("SELECT * FROM funcionarios")

    funcionarios = c.fetchall()
    colunas = [desc[0] for desc in c.description]

    conn.close()

    return render_template("listar.html", funcionarios=funcionarios, colunas=colunas)


@app.route("/funcionarios/add", methods=["GET", "POST"])
def adicionar_funcionario():
    if "user" not in session:
        return redirect(url_for("login"))

    if request.method == "POST":
        dados = [request.form.get(campo) for campo in request.form.keys()]
        conn = sqlite3.connect(DB_NAME)
        c = conn.cursor()
        c.execute(f"""
            INSERT INTO funcionarios ({",".join(request.form.keys())})
            VALUES ({",".join(['?']*len(dados))})
        """, dados)
        conn.commit()
        conn.close()
        return redirect(url_for("listar_funcionarios"))

    return render_template("form.html", acao="Adicionar", funcionario={})


@app.route("/funcionarios/edit/<int:id>", methods=["GET", "POST"])
def editar_funcionario(id):
    if "user" not in session:
        return redirect(url_for("login"))

    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()

    if request.method == "POST":
        dados = [request.form.get(campo) for campo in request.form.keys()]
        campos = ",".join([f"{campo}=?" for campo in request.form.keys()])
        c.execute(f"UPDATE funcionarios SET {campos} WHERE id=?", (*dados, id))
        conn.commit()
        conn.close()
        return redirect(url_for("listar_funcionarios"))

    c.execute("SELECT * FROM funcionarios WHERE id=?", (id,))
    funcionario = c.fetchone()
    colunas = [desc[0] for desc in c.description]
    conn.close()

    return render_template("form.html", acao="Editar", funcionario=dict(zip(colunas, funcionario)))


@app.route("/funcionarios/delete/<int:id>")
def deletar_funcionario(id):
    if "user" not in session:
        return redirect(url_for("login"))

    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("DELETE FROM funcionarios WHERE id=?", (id,))
    conn.commit()
    conn.close()
    return redirect(url_for("listar_funcionarios"))


# ------------------ Exportar Excel ------------------
@app.route("/funcionarios/export")
def exportar_funcionarios():
    if "user" not in session:
        return redirect(url_for("login"))

    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query("SELECT * FROM funcionarios", conn)
    conn.close()

    output = "funcionarios_filtrados.xlsx"
    df.to_excel(output, index=False)

    return send_file(output, as_attachment=True)


if __name__ == "__main__":
    if not os.path.exists(DB_NAME):
        init_db()
        importar_planilha_para_db("headcount 2.xlsx")
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

