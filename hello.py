from flask import Flask, jsonify, make_response

import os
import Opt_func
from flask_httpauth import HTTPBasicAuth

myapp = Flask(__name__)
if os.name == 'nt':
    mainpath = 'd:\\opt_web'
    ip = Opt_func.GetIP()
else:
    mainpath = '/aidata'
    ip = ''

auth = HTTPBasicAuth()

@auth.get_password
def get_password(username):
    if username == 'dw': return 'nthu'
    return None

@auth.error_handler
def unauthorized():
    return make_response(jsonify({'error': 'Unauthorized access'}), 401)

@myapp.route("/")
def hello():
    return 'Hello World! ' +  mainpath + 'ip: ' + ip

@myapp.route("/test")
@auth.login_required
def test():
    return 'function test.'

@myapp.route('/TS/v0.1/Task',  methods = ['GET'])
@auth.login_required
def get_tasks():
    s, ret = Opt_func.DirList(mainpath, '', 'Tasks')
    return jsonify(s)