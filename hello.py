from flask import Flask

import os
import Opt_func

myapp = Flask(__name__)
if os.name == 'nt':
    mainpath = 'd:\\opt_web'
    ip = Opt_func.GetIP()
else:
    mainpath = '/aidata'
    ip = ''

@myapp.route("/")
def hello():
    return 'Hello World! ' +  mainpath + 'ip: ' + ip

@myapp.route("/test")
def test():
    return "function test."