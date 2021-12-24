import os
import datetime as dt
import json
import pathlib
import socket

def GetIP():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    s.connect(("8.8.8.8", 80))
    ret = s.getsockname()[0]
    s.close()
    return ret

def DirList(dirpath, dirFilter, label):
    lst_folder = []
    if os.path.isdir(dirpath):
        ret = 200
        lst_dir = os.listdir(dirpath)
        for x in lst_dir:
            if os.path.isdir(os.path.join(dirpath, x)):
                if dirFilter == '':
                    lst_folder.append(x)
                elif x[:len(dirFilter)] == dirFilter:
                    lst_folder.append(x)
    else:
        ret = 404
    return {label: lst_folder}, ret

def Get_TaskName(mainpath, taskname):
    i = 0
    ret = taskname
    s = os.path.join(mainpath, ret)
    while os.path.isdir(s):
        i += 1
        ret = taskname + '_' + '{0:03d}'.format(i)
        s = os.path.join(mainpath, ret)
    os.makedirs(s)
    return ret

def Create_TaskConfig(mainpath, taskname):
    s1 = {'Name': taskname, 'CreatedTime': dt.datetime.strftime(dt.datetime.now(),  '%Y-%m-%d %H:%M:%S'), 'Status': '', 'StatusNotes': ''}
    s2 = {'Preprocess': 'Yes', 'Manipulation': 'None', 'Lambda': 0, 'SplitQuantity': 0}
    s = {'Task': s1, 'Order': s2}
    with open(os.path.join(os.path.join(mainpath, taskname), 'TaskConfig.json'), 'w') as fn_json:
        json.dump(s, fn_json, ensure_ascii=False)
    return s

def savefile(mainpath, taskname, file, newfilename, attribvalue):
    filename = newfilename if newfilename != '' else file.filename
    if taskname == '':
        sfile = os.path.join(mainpath, filename)
    else:
        sfile = os.path.join(os.path.join(mainpath, taskname), filename)
    file.save(sfile)
    foldername = os.path.dirname(sfile)[len(mainpath) + 1:]
    if attribvalue == 'none':
        ret = 201
    else:
        ret = TaskConfig_SaveFileAttrib(mainpath, taskname, filename, attribvalue, foldername)
    return ret

def TaskConfig_SaveFileAttrib(mainpath, taskname, filename, attrib, foldername):
    fn = open(os.path.join(os.path.join(mainpath, taskname), 'TaskConfig.json'), 'r')
    json_data = json.load(fn)
    fn.close()
    lst_file = [] if 'Files' not in json_data else json_data['Files']
    if attrib != 'Calendar' and attrib != 'Setting' and attrib != 'Evaluation' and attrib != 'Dataset': attrib = ''
    ret = 201
    for x in lst_file:
        if x['Attribute'] == 'Calendar' and attrib == 'Calendar':
            x['Attribute'] = ''
            ret = 205
        elif x['Attribute'] == 'Setting' and attrib == 'Setting':
            x['Attribute'] = ''
            ret = 205
        elif x['Attribute'] == 'Evaluation' and attrib == 'Evaluation':
            x['Attribute'] = ''
            ret = 205
        elif x['Attribute'] == 'Dataset' and attrib == 'Dataset':
            x['Attribute'] = ''
            ret = 205
        else:
            continue
    fname = pathlib.Path(os.path.join(os.path.join(mainpath, taskname), filename))
    ttime = dt.datetime.strftime(dt.datetime.fromtimestamp(fname.stat().st_mtime),  '%Y-%m-%d %H:%M:%S')
    lst_file.append({'Attribute': attrib, 'FileName': filename, 'FolderName': foldername, 'ModifiedTime': ttime})
    json_data['Files'] = lst_file
    with open(os.path.join(os.path.join(mainpath, taskname), 'TaskConfig.json'), 'w') as fn_json:
        json.dump(json_data, fn_json, ensure_ascii=False)
    return ret
