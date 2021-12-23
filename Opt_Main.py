from flask import Flask
from flask import request, abort, jsonify, make_response, session, send_file, redirect

#from flask_ngrok import run_with_ngrok
from flask_httpauth import HTTPBasicAuth
import os, datetime as dt
import mimetypes, pathlib
import Opt_func
import shutil

app = Flask(__name__)
#run_with_ngrok(app)
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

@app.before_request
def make_session_permanent():
    session.permanent = True
    app.permanent_session_lifetime = dt.timedelta(minutes=3)

@auth.error_handler
def unauthorized():
    return make_response(jsonify({'error': 'Unauthorized access'}), 401)

def Download_File(sfile):
    sExt = pathlib.Path(sfile).suffix
    fname = os.path.basename(sfile)
    if sExt.lower() == '.csv':
        mtype= 'text/csv'
    else:
        mtype = mimetypes.suffix_map[sExt]
    result = send_file(sfile, mimetype=mtype, attachment_filename=fname, conditional=False)
    result.headers['x-filename'] = fname
    result.headers["Access-Control-Expose-Headers"] = 'x-filename'
    return result

def DIPSName(pathname):
    pathname = pathname.strip()
    pathname = pathname.replace('$', '\\')
    if pathname == '':
        pathname = mainpath
    else:
        pathname = mainpath + '\\' + pathname
    return pathname


# === route ===
@app.route('/')
def hello():
    return "Hello World!"

@app.route('/TS/v0.1/Task',  methods = ['GET'])
@auth.login_required
def get_tasks():
    s, ret = Opt_func.DirList(mainpath, '', 'Tasks')
    return jsonify(s)

@app.route('/TS/v0.1/Task', methods = ['POST', 'PUT'])
@auth.login_required
def create_task():
    taskname = request.form['Name']
    print(taskname)
    if request.method == 'POST':
        s = Opt_func.Get_TaskName(mainpath, taskname)
    else:
        s = taskname
    print(s)
    if not os.path.isdir(mainpath + '\\' + s): abort(404)
    json_task = Opt_func.Create_TaskConfig(mainpath, s)
    cfile = request.files['Calendar']
    ret = Opt_func.savefile(mainpath, s, cfile, '', 'Calendar')
    if ret == 201 or ret == 205:
        sfile = request.files['Setting']
        ret = Opt_func.savefile(mainpath, s, sfile, '', 'Setting')
        if ret != 201 and ret != 205: abort(ret)
    return jsonify(json_task), ret

@app.route('/TS/v0.1/Task/<string:taskname>',  methods = ['GET'])
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

@app.route('/TS/v0.1/Task/<string:taskname>/Evaluate',  methods = ['GET'])
@auth.login_required
def Evaluate(taskname):
    return redirect('http://' + ip + ':5001/RS/v0.1/' + taskname + '/Evaluate', code=302)

@app.route('/TS/v0.1/Task/<string:taskname>/EVR',  methods = ['GET'])
@auth.login_required
def Download_EVR(taskname):
    sfile = os.path.basename(Opt_func.EVRFile(mainpath, taskname))
    spath = mainpath + '\\' + taskname + '\\'
    if not os.path.isfile(spath + sfile): abort(404)
    result = Opt_func.Download_EXCELFile(spath, sfile)
    return result
    #return send_from_directory(mainpath + '\\' + taskname, sfile, as_attachment=True)

@app.route('/TS/v0.1/Task/<string:taskname>/<string:dirname>',  methods = ['GET'])
@auth.login_required
def Dataset_Files(taskname, dirname):
    s, ret = Opt_func.FileList(mainpath, taskname, dirname)
    if ret != 200: abort(ret)
    return jsonify(s), ret

@app.route('/TS/v0.1/Task/<string:taskname>/StageLearning',  methods = ['GET'])
@auth.login_required
def Stage_Dirs(taskname):
    s, ret = Opt_func.DirList(mainpath + '\\' + taskname, 'StageLearning', taskname)
    if ret != 200: abort(ret)
    return jsonify(s), ret

@app.route('/TS/v0.1/Task/<string:taskname>/StageLearning/NEW',  methods = ['POST'])
@auth.login_required
def NewStage(taskname):
    stagename = Opt_func.NewStage(mainpath, taskname)
    if Opt_func.StageisProcessing(mainpath, taskname, stagename):
        ret = 403
    else:
        os.makedirs(mainpath + '\\' + taskname + '\\' + stagename)
        sfile = request.files['StageParameter']
        ret = Opt_func.savefile(mainpath, taskname + '\\' + stagename, sfile, '', 'none')
    if ret != 200 and ret != 201: abort(ret)
    return stagename, 201

@app.route('/TS/v0.1/Task/<string:taskname>/<string:dirname>',  methods = ['DELETE'])
@auth.login_required
def delete_stage(taskname, dirname):
    ret = Opt_func.DeleteTask(mainpath, taskname, dirname)
    if ret != 200: abort(ret)
    return '', ret

@app.route('/TS/v0.1/Task/<string:taskname>/<string:stagename>/StageParameter',  methods = ['GET'])
@auth.login_required
def GetStageParameter(taskname, stagename):
    s, ret = Opt_func.OpenJsonFile(mainpath, taskname + '\\' + stagename, 'StageParameter.json', 'Trial Stage')
    if ret != 200: abort(ret)
    return jsonify(s), ret

@app.route('/TS/v0.1/Task/<string:taskname>/<string:stagename>/StageParameter',  methods = ['POST', 'PUT'])
@auth.login_required
def UploadtageParameter(taskname, stagename):
    if not os.path.isdir(mainpath + '\\' + taskname + '\\' + stagename):
        ret = 404
    else:
        sfile = request.files['StageParameter']
        if Opt_func.StageisProcessing(mainpath, taskname, stagename):
            ret = 403
        else:
            ret = Opt_func.savefile(mainpath, taskname + '\\' + stagename, sfile, '', 'none')
    if ret != 200 and ret !=201: abort(ret)
    return '', ret

@app.route('/TS/v0.1/Task/<string:taskname>/<string:stagename>/RUN',  methods = ['GET'])
@auth.login_required
def RunStage(taskname, stagename):
    return redirect('http://' + ip + ':5002/SL/v0.1/' + taskname + '/' + stagename + '/RUN', code=302)

@app.route('/TS/v0.1/Task/<string:taskname>/<string:stagename>/STOP',  methods = ['GET'])
@auth.login_required
def StopStage(taskname, stagename):
    return redirect('http://' + ip + ':5002/SL/v0.1/' + taskname + '/' + stagename + '/STOP', code=302)

@app.route('/TS/v0.1/Task/<string:taskname>/<string:stagename>/Record',  methods = ['GET'])
@auth.login_required
def Download_Record(taskname, stagename):
    sfile = 'StageRecord.xlsx'
    spath = mainpath + '\\' + taskname + '\\' + stagename + '\\'
    if not os.path.isfile(spath + sfile): abort(404)
    result = Opt_func.Download_EXCELFile(spath, sfile)
    return result

@app.route('/TS/v0.1/Task/<string:taskname>/<string:stagename>/Results',  methods = ['GET'])
@auth.login_required
def Results_Dir(taskname, stagename):
    s, ret = Opt_func.DirList(mainpath + '\\' + taskname + '\\' + stagename, '', stagename)
    if ret != 200: abort(ret)
    return jsonify(s), ret

@app.route('/TS/v0.1/Task/<string:taskname>/<string:stagename>/Reports/<string:resultname>',  methods = ['GET'])
@auth.login_required
def Download_PSRReport(taskname, stagename, resultname):
    return redirect('http://' + ip + ':5001/RS/v0.1/' + taskname + '/' + stagename + '/' + resultname, code=302)

@app.route('/TS/v0.1/Directory/DIPS/<string:pathname>',  methods = ['GET'])
@auth.login_required
def DIPS_DirList(pathname):
    pathname = DIPSName(pathname)
    if os.path.isfile(pathname):
        ret = Download_File(pathname)
    else:
        if not os.path.isdir(pathname): abort(404)
        ret = jsonify(Opt_func.path_to_dict(pathname))
    return ret

@app.route('/TS/v0.1/Directory/DIPS/<string:pathname>',  methods = ['POST'])
@auth.login_required
def DIPS_CreateFolderandfile(pathname):
    pathname = DIPSName(pathname)
    if not os.path.isdir(pathname):
        s  = os.path.normpath(pathname)
        newfolder = os.path.basename(s)
        pname = os.path.normpath(s[:-len(newfolder)])
        if not os.path.isdir(pname): abort(404)
        os.makedirs(pname + '\\' + newfolder)
        ret = 200
    else:
        if 'file' not in request.files: abort(404)
        dfile = request.files['file']
        ret = Opt_func.savefile(pathname, '', dfile, '', 'none')
    return '', ret

@app.route('/TS/v0.1/Directory/DIPS/<string:pathname>',  methods = ['DELETE'])
@auth.login_required
def DIPS_DeleteFolder(pathname):
    pathname = DIPSName(pathname)
    if not os.path.isdir(pathname): abort(404)
    lstfiles = os.listdir(pathname)
    for xfile in lstfiles:
        if os.path.isfile(pathname + '\\' + xfile):
            print('file', xfile)
            os.remove(pathname + '\\' + xfile)
        else:
            print('folder', xfile)
            shutil.rmtree(pathname + '\\' + xfile)
    return '', 200

@app.route('/TS/v0.1/Directory/DIPS/<string:pathname>',  methods = ['PUT'])
@auth.login_required
def DIPS_UpdateFolder(pathname):
    ret = 200
    pathname = DIPSName(pathname)
    if os.path.isdir(pathname):
        if 'name' not in request.form: abort(404)
        os.rename(pathname, mainpath + '\\' + request.form['name'])
    else:
        if os.path.isfile(pathname):
            if 'file' not in request.files: abort(404)
            dfile = request.files['file']
            ret = Opt_func.savefile(os.path.dirname(pathname), '', dfile, os.path.basename(pathname), 'none')
        else:
            abort(404)
    return '', ret


if __name__ == "__main__" and os.name == 'nt':
    app.secret_key = 'super secret key'
    app.config['SESSION_TYPE'] = 'filesystem'
    app.run(host=ip, port=5000)