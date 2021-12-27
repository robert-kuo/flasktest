import datetime as dt
import time

sfile = '/aidata/DIPS/datax.txt'
while True:
    fn = open(sfile, 'a')
    s = dt.datetime.strftime(dt.datetime.now(),  '%Y-%m-%d %H:%M:%S') + '\n'
    fn.writelines(s)
    fn.close()
    time.sleep(5)
