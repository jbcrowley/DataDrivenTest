@echo off

REM create timestamp for filename
For /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set mydate=%%c.%%a.%%b)
For /f "tokens=1-3 delims=:. " %%a in ("%TIME%") do (set mytime=%%a.%%b.%%c)
set resultsFilename=DataDrivenTest_%mydate%_%mytime%

REM set paths
set nunit=%~dp0packages\NUnit.ConsoleRunner.3.7.0\tools\nunit3-console.exe
set dll=%~dp0DataDrivenTest\bin\x86\Debug\DataDrivenTest.dll
set results=%~dp0Results\%resultsFilename%.xml

REM %nunit% %dll% --explore
%nunit% %dll% --where "class==DataDrivenTest.UnitTest1" --workers:4 --result:%results%