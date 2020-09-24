import os
from utils.functions import make_dir, packaging, path_find, subprocess_cmd
from utils.config import SCRIPTS_DIR

if __name__ == "__main__":
    make_dir(os.path.join(os.getcwd(), 'dist'))

    file_name_py = 'Entry.py'
    file_name_exe = 'Entry.exe'
    file_to_compile = path_find(file_name_py, os.getcwd())
    file_to_compile = file_to_compile.replace('\\', '/')
    icon_name = "Cyberduck.ico"

    freeze_command_with_icon = f'pyinstaller.exe --onefile --hidden-import=xlrd --icon={icon_name} {file_to_compile}'
    freeze_command_without_icon = f'pyinstaller.exe --onefile --hidden-import=xlrd {file_to_compile}'
    dist_dir = os.path.join(os.getcwd(), 'dist')

    subprocess_cmd(f'cd {SCRIPTS_DIR} & {freeze_command_without_icon} & cd dist & '
                   f'copy {file_name_exe} {dist_dir} & copy {file_name_exe} {os.getcwd()}')

    packaging(file_name_exe, 'Cookies_objection', 'driver')
    os.startfile('dist')