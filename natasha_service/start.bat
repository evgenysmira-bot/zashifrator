@echo off
REM Запуск локального NLP-сервиса для Data Hedgehog.
REM Слушает 127.0.0.1:18765 — наружу не доступен.
cd /d "%~dp0"
python service.py
