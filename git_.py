from subprocess import Popen, PIPE
from datetime import datetime
import os


def subprocess_cmd(command):
    print(command)
    try :
        process = Popen(command, stdout=PIPE, shell=True, universal_newlines=True)
        proc_stdout = process.communicate()[0].strip()
    except Exception as e:
        process = Popen(command, stdout=PIPE, shell=True, universal_newlines=False)
        proc_stdout = process.communicate()[0].strip()
    print(proc_stdout)


def old_ver_directory():
    try:
        os.mkdir(os.path.join(os.getcwd(), 'old'))
    except Exception as e:
        pass
    dir = os.path.join(os.getcwd(), 'old', f'old_ver_{datetime.now().strftime("%Y%m%d_%H%M%S")}')
    os.mkdir(dir)
    return dir


class GitCommandLines():
    def __init__(self):
        self.repository = r'https://github.com/asahi8769/Objections.git'
        subprocess_cmd (f'git config --global user.name Ilhee Lee')
        subprocess_cmd(f'git config --global user.email asahi8769@gmail.com')

    def push_rep(self):
        self.clone_rep()
        subprocess_cmd (f'git init')
        subprocess_cmd (f'git add .')
        subprocess_cmd (f'git config --global http.sslVerify false')
        # subprocess_cmd (f'git rm -r --cached images.zip')
        # subprocess_cmd (f'git rm -r --cached .idea/')
        # subprocess_cmd (f'git rm -r --cached __pycache__/')
        subprocess_cmd (f'git commit -m "Add all my files"')
        subprocess_cmd (f'git remote add origin {self.repository}')
        subprocess_cmd (f'git push --force origin master')
        subprocess_cmd(f'git remote remove origin')

    def clone_rep(self):
        dir = os.path.relpath(old_ver_directory(), os.getcwd())
        subprocess_cmd(f'git clone {self.repository[:-4]} {dir}')

    def history(self):
        subprocess_cmd(f'git log ')


if __name__ == "__main__":
    GitCommandLines().push_rep()