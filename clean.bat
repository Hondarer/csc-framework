@echo off
echo === Clean Build Output ===

if exist "bin" (
    echo Cleaning bin directory...
    del /Q "bin\*.*" 2>nul
    echo Build output cleaned.
) else (
    echo No bin directory found.
)

echo Clean completed.
