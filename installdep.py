from utils.functions import subprocess_cmd, install
from utils.config import BASE_PYTHON, VENV_64_DIR, SCRIPTS_DIR


subprocess_cmd(rf'cd {BASE_PYTHON} && python -m venv {VENV_64_DIR} && cd {SCRIPTS_DIR} && activate')
subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("pyinstaller")}')
subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("xlrd")}')
subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("selenium")}')
subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("pandas")}')
subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("pyautogui")}')
subprocess_cmd(f'cd {SCRIPTS_DIR} & {install("pyperclip")}')
