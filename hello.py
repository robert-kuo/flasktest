from flask import Flask
from flask import request, abort, jsonify, make_response, redirect

import os  #, subprocess
import Opt_func
from flask_httpauth import HTTPBasicAuth

myapp = Flask(__name__)
if os.name == 'nt':
    mainpath = 'd:\\opt_web'
    ip = Opt_func.GetIP()
else:
    mainpath = '/aidata/DIPS'
    ip = ''

# subprocess.Popen(['python', 'teststart.py'])

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
    return 'TS Service... ' +  mainpath + ' ip: ' + ip

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

@myapp.route('/TS/v0.1/Task/<string:taskname>',  methods = ['DELETE'])
@auth.login_required
def delete_task(taskname):
    ret = Opt_func.DeleteTask(mainpath, taskname, '')
    if ret != 200: abort(ret)
    return '', ret

@myapp.route('/TS/v0.1/Task/<string:taskname>/Evaluate',  methods = ['GET'])
@auth.login_required
def Evaluate(taskname):
    return redirect('https://reportwebservice.azurewebsites.net/RS/v0.1/' + taskname + '/Evaluate', code=302)

@myapp.route('/TS/v0.1/Task/<string:taskname>/EVR',  methods = ['GET'])
@auth.login_required
def Download_EVR(taskname):
    sfile = os.path.basename(Opt_func.EVRFile(mainpath, taskname))
    spath = os.path.join(mainpath, taskname)
    if not os.path.isfile(os.path.join(spath, sfile)): abort(404)
    result = Opt_func.Download_EXCELFile(spath, sfile)
    return result

@myapp.route('/TS/v0.1/Task/<string:taskname>/<string:dirname>',  methods = ['GET'])
@auth.login_required
def Dataset_Files(taskname, dirname):
    s, ret = Opt_func.FileList(mainpath, taskname, dirname)
    if ret != 200: abort(ret)
    return jsonify(s), ret

@myapp.route('/TS/v0.1/Task/<string:taskname>/StageLearning',  methods = ['GET'])
@auth.login_required
def Stage_Dirs(taskname):
    s, ret = Opt_func.DirList(os.path.join(mainpath, taskname), 'StageLearning', taskname)
    if ret != 200: abort(ret)
    return jsonify(s), ret

@myapp.route('/TS/v0.1/Task/<string:taskname>/StageLearning/NEW',  methods = ['POST'])
@auth.login_required
def NewStage(taskname):
    stagename = Opt_func.NewStage(mainpath, taskname)
    if Opt_func.StageisProcessing(mainpath, taskname, stagename):
        ret = 403
    else:
        os.makedirs(os.path.join(os.path.join(mainpath, taskname), stagename))
        sfile = request.files['StageParameter']
        ret = Opt_func.savefile(mainpath, os.path.join(taskname, stagename), sfile, '', 'none')
    if ret != 200 and ret != 201: abort(ret)
    return stagename, 201

@myapp.route('/TS/v0.1/Task/<string:taskname>/<string:dirname>',  methods = ['DELETE'])
@auth.login_required
def delete_stage(taskname, dirname):
    ret = Opt_func.DeleteTask(mainpath, taskname, dirname)
    if ret != 200: abort(ret)
    return '', ret

@myapp.route('/TS/v0.1/Task/<string:taskname>/<string:stagename>/StageParameter',  methods = ['GET'])
@auth.login_required
def GetStageParameter(taskname, stagename):
    s, ret = Opt_func.OpenJsonFile(mainpath, os.path.join(taskname, stagename), 'StageParameter.json', 'Trial Stage')
    if ret != 200: abort(ret)
    return jsonify(s), ret

@myapp.route('/TS/v0.1/Task/<string:taskname>/<string:stagename>/StageParameter',  methods = ['POST', 'PUT'])
@auth.login_required
def UploadtageParameter(taskname, stagename):
    if not os.path.isdir(os.path.join(os.path.join(mainpath, taskname), stagename)):
        ret = 404
    else:
        sfile = request.files['StageParameter']
        if Opt_func.StageisProcessing(mainpath, taskname, stagename):
            ret = 403
        else:
            ret = Opt_func.savefile(mainpath, os.path.join(taskname, stagename), sfile, '', 'none')
    if ret != 200 and ret !=201: abort(ret)
    return '', ret

@myapp.route('/TS/v0.1/Task/<string:taskname>/<string:stagename>/RUN',  methods = ['GET'])
@auth.login_required
def RunStage(taskname, stagename):
    return redirect('https://stagelearningservice.azurewebsites.net/SL/v0.1/' + taskname + '/' + stagename + '/RUN', code=302)

@myapp.route('/TS/v0.1/Task/<string:taskname>/<string:stagename>/STOP',  methods = ['GET'])
@auth.login_required
def StopStage(taskname, stagename):
    return redirect('https://stagelearningservice.azurewebsites.net/SL/v0.1/' + taskname + '/' + stagename + '/STOP', code=302)