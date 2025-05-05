from flask import Flask


app = Flask(__name__)


from routes import *
#ocorre apenas se executar o main
if __name__ == "__main__":
    app.run()