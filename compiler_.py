from subprocess import Popen, PIPE
import os, sys
from utility_functions import subprocess_cmd, path_find

sys.path.insert(1, os.pardir)

from os.path import basename
from zipfile import ZipFile
from utility_functions import make_dir


def packaging(filename, *bindings):
    zipname = r'dist\Objections.zip'
    with ZipFile(zipname, 'w') as zipObj:
        if os.path.exists(os.path.join('dist',filename)):
            zipObj.write(os.path.join('dist', filename), basename(filename))
        for binding in bindings:
            for folderName, subfolders, filenames in os.walk(binding):
                for filename in filenames:
                    filePath = os.path.join(folderName, filename)
                    print(os.path.join(binding, basename(filePath)))
                    zipObj.write(filePath, os.path.join(basename(folderName), basename(filePath)))
        print(f'패키징을 완료하였습니다. {zipname}')


if __name__ == "__main__":
    make_dir(os.path.join(os.getcwd(), 'dist'))
    dir_py_32 = r'C:\Users\glovis-laptop\AppData\Local\Programs\Python\Python37-32\Scripts'.replace('\\', '/')
    dir_py_64 = r'C:\Users\glovis-laptop\AppData\Local\Programs\Python\Python37\Scripts'.replace('\\', '/')
    dir_venv_64 = r'D:\devs\Objections\venv\Scripts'

    file_name_py = 'Entry.py'
    file_name_exe = 'Entry.exe'
    file_to_compile = path_find(file_name_py, r'D:/devs/Objections')
    file_to_compile = file_to_compile.replace('\\', '/')
    icon_name = r'D:\devs\Objections\Cyberduck.ico'.replace('\\', '/')


    install_command = f'pyinstaller.exe -F --hidden-import=xlrd --icon={icon_name} {file_to_compile}'
    dir_loc = os.path.join(os.getcwd(), 'dist')
    print(install_command)
    subprocess_cmd(f'cd {dir_venv_64} & {install_command} & cd dist & copy {file_name_exe} {dir_loc}')

    packaging(file_name_exe, 'Cookies_objection', 'driver')
