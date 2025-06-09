from main import app
from manipulacao_de_planilha import informacoes_gerais

#rotas
@app.route("/") #decorator (dá uma funcionalidade extra para a função abaixo dele)
def mostrar_informacoes():
    return informacoes_gerais()

