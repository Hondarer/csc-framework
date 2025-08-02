@echo off
echo === Clean Build Output ===

set CLEANED=false

if exist "bin\debug" (
    echo Cleaning bin\debug directory...
    rmdir /S /Q "bin\debug" 2>nul
    set CLEANED=true
)

if exist "bin\release" (
    echo Cleaning bin\release directory...
    rmdir /S /Q "bin\release" 2>nul
    set CLEANED=true
)

if "%CLEANED%"=="true" (
    echo Build output cleaned.
) else (
    echo No build output directories found.
)

echo Clean completed.
