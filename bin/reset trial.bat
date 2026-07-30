@echo off
setlocal

:: --- verifica daca ruleaza ca Administrator, altfel se re-lanseaza elevat ---
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Se re-lanseaza ca Administrator...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo.
echo Resetare trial "Applied Genetics"
echo ================================

echo.
echo [1/2] Sterg cheia de registry...
reg delete "HKCU\Software\VB and VBA Program Settings\WinSys" /f
if %errorlevel%==0 (
    echo     OK - registry curatat.
) else (
    echo     Nu exista sau a esuat ^(posibil deja curat^).
)

echo.
echo [2/2] Sterg fisierul de verificare ^(System32 SI SysWOW64, din cauza redirectionarii WOW64^)...

for %%D in ("%WINDIR%\System32\sys_exp.dll" "%WINDIR%\SysWOW64\sys_exp.dll") do (
    if exist %%D (
        attrib -r -h -s %%D
        del /f /q %%D
        if exist %%D (
            echo     EROARE: %%D nu a putut fi sters.
        ) else (
            echo     OK - sters: %%D
        )
    ) else (
        echo     Nu exista: %%D - ok.
    )
)

echo.
echo Gata. Poti porni din nou programul.
echo.
pause
