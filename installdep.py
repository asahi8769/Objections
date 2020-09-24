from utils.functions import subprocess_cmd, install
import os

dir_venv_64 = os.path.join(os.getcwd(), 'venv', 'Scripts')
subprocess_cmd(f'cd {dir_venv_64} & {install("pyinstaller")}')
subprocess_cmd(f'cd {dir_venv_64} & {install("xlrd")}')
subprocess_cmd(f'cd {dir_venv_64} & {install("selenium")}')
subprocess_cmd(f'cd {dir_venv_64} & {install("pandas")}')
subprocess_cmd(f'cd {dir_venv_64} & {install("pyautogui")}')
subprocess_cmd(f'cd {dir_venv_64} & {install("pyperclip")}')
