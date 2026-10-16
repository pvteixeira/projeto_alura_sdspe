@echo off
:: ========================================================================
::  USM-Install Launcher - Execução com Elevação de Administrador (UAC)
:: ========================================================================
title USM-Install Launcher

:: Muda para o diretório onde o arquivo .bat reside
cd /d "%~dp0"

echo [USM-Install] Solicitando privilegios de Administrador...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Start-Process powershell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"\"%~dp0USM-Install.ps1\"\"' -Verb RunAs"

if %ERRORLEVEL% NEQ 0 (
    echo [ERRO] Ocorreu uma falha ou a elevacao de Administrador foi recusada.
    pause
)
