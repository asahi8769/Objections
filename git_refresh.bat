@ECHO ON
title git Start

cd /D %~dp0

%~dp0\venv\Scripts\python.exe git_refresh.py

cmd.exe