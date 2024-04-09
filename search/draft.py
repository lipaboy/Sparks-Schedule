import psutil
from subprocess import Popen

from psutil import AccessDenied

plist = psutil.process_iter()
for i in plist:
    try:
        if 'excel' in i.name().lower():
            print( i.name(), i.pid, i.cmdline(), i.open_files() )
            if any('PoolTime' in path.path for path in i.open_files()):
                i.kill()
    except AccessDenied:
        print("'%s' Process is not allowing us to view the CPU Usage!" + i.name())