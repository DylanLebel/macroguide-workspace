@echo off
REM Double-click this to deploy MacroGuide files from C:\AllMacros to Y:\
title Deploy Macro Guide
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Deploy-MacroGuide.ps1"
