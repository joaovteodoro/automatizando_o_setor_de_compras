from main import app
from flask import render_template

#rotas
@app.route("/") #decorator (dá uma funcionalidade extra para a função abaixo dele)
def homepage():
    return render_template("homepage.html")


@app.route("/execucoes") #decorator (dá uma funcionalidade extra para a função abaixo dele)
def execucoes():
    return "Meu site no Flask"