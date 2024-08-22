@echo off
:: Lines starting with two colons :: are comments

:: Sets the execution policy for the current process to allow the script to run, then runs it
powershell -NoProfile -ExecutionPolicy Bypass -File "Super_God_Mode.ps1"

:: Alternatively, you could run this command yourself in a powershell window before running the script:   		Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
:: Which also would just allow scripts for that current PowerShell session, then you can run the script with: 	.\Super_God_Mode.ps1
