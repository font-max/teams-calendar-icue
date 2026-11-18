@echo off
title Teams Calendar - Uninstall
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup.ps1" -Uninstall %*
