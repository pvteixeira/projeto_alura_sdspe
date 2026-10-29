@echo off
:: ========================================================================
::  USM-Install Launcher - Execução com Elevação de Administrador (UAC)
:: ========================================================================
title USM-Install Launcher

:: Muda para o diretório onde o arquivo .bat reside
cd /d "%~dp0"

echo [USM-Install] Solicitando privilegios de Administrador...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$p = '%~dp0USM-Install.ps1'; if (Test-Path $p) { $b = [System.IO.File]::ReadAllBytes($p); if ($b.Length -ge 3 -and -not ($b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF)) { [System.IO.File]::WriteAllBytes($p, ([byte[]](0xEF,0xBB,0xBF) + $b)) } }; Start-Process powershell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File \"\"%~dp0USM-Install.ps1\"\"' -Verb RunAs"

if %ERRORLEVEL% NEQ 0 (
    echo [ERRO] Ocorreu uma falha ou a elevacao de Administrador foi recusada.
    pause
)
