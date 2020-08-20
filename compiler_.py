from subprocess import Popen, PIPE
import os, sys
from utility_functions import *
sys.path.insert(1, os.pardir)

if __name__ == "__main__":
    make_dir(os.path.join(os.getcwd(), 'dist'))

    dir_venv_64 = os.path.join(os.getcwd(), 'venv', 'Scripts')
    file_name_py = 'Entry.py'
    file_name_exe = 'Entry.exe'
    file_to_compile = path_find(file_name_py, os.getcwd())
    file_to_compile = file_to_compile.replace('\\', '/')
    icon_name = os.path.join(os.getcwd(), 'Cyberduck.ico').replace('\\', '/')

    install_command = f'pyinstaller.exe -F --hidden-import=xlrd  {file_to_compile}'
    dir_loc = os.path.join(os.getcwd(), 'dist')
    print(install_command)
    subprocess_cmd(f'cd {dir_venv_64} & {install_command} & cd dist & copy {file_name_exe} {dir_loc}')
    packaging(file_name_exe, 'Cookies_objection', 'driver')
    os.startfile('dist')