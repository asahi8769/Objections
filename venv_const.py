from compiler_ import subprocess_cmd

dir_venv_64 = r'D:\devs\Objections\venv\Scripts'


def install(lib):
    return f'pip --trusted-host pypi.python.org --trusted-host files.pythonhosted.org --trusted-host pypi.org install {lib}'

# subprocess_cmd(f'cd {dir_venv_64}  & {install("pyinstaller")} & {install("xlrd")} & {install("selenium")} & '
#                f'{install("pandas")} & {install("pyautogui")}')
# subprocess_cmd(f'cd {dir_venv_64} & {install("pyperclip")}')
