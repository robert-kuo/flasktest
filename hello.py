from flask import Flask
from flask import request, abort, jsonify, make_response

import os
import Opt_func
from flask_httpauth import HTTPBasicAuth

myapp = Flask(__name__)
if os.name == 'nt':
    mainpath = 'd:\\opt_web'
    ip = Opt_func.GetIP()
else:
    mainpath = '/home'
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

@myapp.route('/TS/v0.1/Task',  methods = ['GET'])
@auth.login_required
def get_tasks():
    s, ret = Opt_func.DirList(mainpath, '', 'Tasks')
    return jsonify(s)

@myapp.route('/TS/v0.1/Task', methods = ['POST', 'PUT'])
@auth.login_required
def create_task():
    taskname = request.form['Name']
    if request.method == 'POST':
        s = Opt_func.Get_TaskName(mainpath, taskname)
    else:
        s = taskname
    print(s)
    if not os.path.isdir(os.path.join(mainpath, s)): abort(404)
    json_task = Opt_func.Create_TaskConfig(mainpath, s)
    cfile = request.files['Calendar']
    ret = Opt_func.savefile(mainpath, s, cfile, '', 'Calendar')
    if ret == 201 or ret == 205:
        sfile = request.files['Setting']
        ret = Opt_func.savefile(mainpath, s, sfile, '', 'Setting')
        if ret != 201 and ret != 205: abort(ret)
    return jsonify(json_task), ret

@myapp.route('/TS/v0.1/Task/<string:taskname>',  methods = ['GET'])
def get_task(taskname):
    s, ret = Opt_func.OpenJsonFile(mainpath, taskname, 'TaskConfig.json', 'Task')
    if ret != 200: abort(ret)
    return jsonify(s), ret

@app.route('/TS/v0.1/Task/<string:taskname>',  methods = ['DELETE'])
@auth.login_required
def delete_task(taskname):
    ret = Opt_func.DeleteTask(mainpath, taskname, '')
    if ret != 200: abort(ret)
    return '', ret