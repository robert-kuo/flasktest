import os
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