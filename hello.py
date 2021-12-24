from flask import Flask
myapp = Flask(__name__)

@myapp.route("/")
def hello():
    return "Hello Flask, on Azure App Service for Linux!!! ??? $$$ ==============>"

@myapp.route("/test")
def test():
    return "function test."
