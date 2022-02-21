:parse
if "%~1"=="" goto main
if "%~1"=="clean" goto clean
shift
goto parse
:endparse
REM endparse

:main
echo off
for /f %%a in ('dir /b "*.py"') do (
    if exist %%a.done (
     for /f "skip=2 tokens=*" %%b in ('dir /b /TW /OD "%%a*"') do if "%%a" == "%%b" (
         (call python %%a && echo > "%%a.done" ) || goto exiterror
     ) else echo make: Nothing to be done for %%a
    ) else (call python %%a && echo > "%%a.done" ) || goto exiterror
)
goto end

:clean
echo off

(for /f %%a in ('dir /b "*.py.log"'  ) do echo "DELETE %%a" && del "%%a" ) 2> nul
(for /f %%a in ('dir /b "*.py.log.*"') do echo "DELETE %%a" && del "%%a" ) 2> nul
(for /f %%a in ('dir /b "*.py.done"' ) do echo "DELETE %%a" && del "%%a" ) 2> nul
(for /f %%a in ('dir /b "*~"'        ) do echo "DELETE %%a" && del "%%a" ) 2> nul
(for /f %%a in ('dir /b "%TMP%\??-demo-revision-*.*"') do echo "DELETE %TMP%\%%a" && del "%TMP%\%%a" ) 2> nul
goto end

:exiterror
echo on
exit /b 1

:end
echo on
