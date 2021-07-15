@echo off
color 0a

::===================================================================================================
::						Setup
::===================================================================================================
set NGMode=1
set LogMode=1
set EchoMode=1
set DirCount=1
set DayCount=7

:ParameterLoop
if "%~1"=="-ng" set NGMode=0
if "%~1"=="-log" set LogMode=0
if "%~1"=="-s" set EchoMode=0
shift
if not "%~1"=="" goto ParameterLoop


:: -------------------------- Localization & Exact File Location Formation --------------------------
call :sGetCurrentDate
if not %LogMode%==0 echo %YYYY% - %MM%- %DD% ( %dow%) >> "D:\Log_Removal.txt"
call :sCreateFolderList
call :sDecDD
call :sSetFileLoc


::===================================================================================================
::						Main Program
::===================================================================================================

:: ------------------------------------------ File Removal ------------------------------------------
:sDeleteFile
call :sGetCurrentDate
rd /s /q %File% && set Status=Succeed || set Status=Failed
if not %EchoMode%==0 echo %File%	-%Status%
if not %LogMode%==0 echo (%HH%:%Min%:%Sec%) %File%	-%Status% >> "D:/SpeedX_Log.txt"
call :sNextFile
goto sDeleteFile

:: ------------------------------------------ End Sequence ------------------------------------------
:TheEnd
del /q "D:\List.txt"
exit


::===================================================================================================
::					User Define Function
::===================================================================================================


::							File

:: ------------------------------- Creating List Of File & Folders --------------------------------
:sCreateFolderList
if %DirCount% equ 1 ( set Directory=D:\MNecoo\log\RoboC_ST2\)
if %DirCount% equ 2 ( set Directory=D:\MNecoo2\log\RoboD_ST2\)
dir /a:d /b %Directory% > "D:\List.txt"
echo END >> "D:/List.txt"
exit /b

:sRefreshFolderList
for /f "skip=1 delims=*" %%a in (D:\List.txt) do ( echo %%a >> D:\Temp.txt )
del /q "D:\List.txt"
ren D:\Temp.txt List.txt
set /p Folder=< "D:\List.txt"
exit /b

:sSetFileLoc
set /p Folder=< "D:\List.txt"
if %NGMode%==0 ( if %Folder%==NG ( call :sRefreshFolderList ) )
set File=%Directory%%Folder%\%sY%\%sY%_%sM%_%sD%
set File=%File: =%
exit /b

:: -------------------------------------------- Go Next --------------------------------------------
:sNextFile
set /a DayCount-=1
if %DayCount% gtr 0 ( call :sIncDD & call :sSetFileLoc ) else ( call :sNextFolderInList )
exit/b

:sNextFolderInList
set DayCount=7
call :sDecDD
call :sRefreshFolderList
call :sSetFileLoc
if %Folder% equ END call :sNextList
exit /b

:sNextList
if %DirCount% equ 2 ( goto TheEnd )
set DirCount=2
call :sCreateFolderList
call :sDecDD
call :sSetFileLoc
exit /b
::							File

::							Get Local Time

:: ------------------------------- Getting Local Time Into Variables --------------------------------
:sGetCurrentDate
for /f "tokens=2 delims==*" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "YYYY=%dt:~0,4%"
set "YY=%dt:~2,2%"
set "MM=%dt:~4,2%"
set "DD=%dt:~6,2%"
set "HH=%dt:~8,2%"
set "Min=%dt:~10,2%"
set "Sec=%dt:~12,2%"
for /f %%i in ('powershell ^(get-date^).DayOfWeek') do set dow=%%i
call :sSetDayDecOffset
call :sSetDecDayThreshold
exit/b
:: ------------------------------------------- Time Formating -----------------------------------------

:sShortenDate
if %sD% lss 10 (set /a "sD=1%sD%%%10")
if %sM% lss 10 (set /a "sM=1%sM%%%10")
exit /b

:sLengthenDate
if %sD% lss 10 set sD=0%sD%
if %sM% lss 10 set sM=0%sM%
exit /b

::							Get Local Time
::							Setting Offset

:sSetIncDayThreshold
if %sM% equ 01 set DayIncThreshold=31
set /a "sR=%sY%%%4"
if %sM% equ 02 ( if %sR%==0 (set DayIncThreshold=29) else (set DayIncThreshold=28))
if %sM% equ 03 set DayIncThreshold=31
if %sM% equ 04 set DayIncThreshold=30
if %sM% equ 05 set DayIncThreshold=31
if %sM% equ 06 set DayIncThreshold=30
if %sM% equ 07 set DayIncThreshold=31
if %sM% equ 08 set DayIncThreshold=31
if %sM% equ 09 set DayIncThreshold=30
if %sM% equ 10 set DayIncThreshold=31
if %sM% equ 11 set DayIncThreshold=30
if %sM% equ 12 set DayIncThreshold=31
exit /b

:sSetDecDayThreshold
if %MM% equ 01 set DayDecThreshold=31
if %MM% equ 02 set DayDecThreshold=31
set /a "sR=%YYYY%%%4"
if %MM% equ 03 ( if %sR%==0 (set DayIncThreshold=29) else (set DayIncThreshold=28))
if %MM% equ 04 set DayDecThreshold=31
if %MM% equ 05 set DayDecThreshold=30
if %MM% equ 06 set DayDecThreshold=31
if %MM% equ 07 set DayDecThreshold=30
if %MM% equ 08 set DayDecThreshold=31
if %MM% equ 09 set DayDecThreshold=31
if %MM% equ 10 set DayDecThreshold=30
if %MM% equ 11 set DayDecThreshold=31
if %MM% equ 12 set DayDecThreshold=30
exit /b

:sSetDayDecOffset
if %dow% equ Monday set offset=7
if %dow% equ Tuesday set offset=8
if %dow% equ Wednesday set offset=9
if %dow% equ Thursday set offset=10
if %dow% equ Friday set offset=11
if %dow% equ Saturday set offset=12
if %dow% equ Sunday set offset=13
exit /b

::							Setting Offset
::							Decrease Time

:sDecDD
set sD=%DD% & set sM=%MM% & set sY=%YYYY%
call :sShortenDate
set /a sD=(%sD%-%offset%)
if %sD% lss 0 ( call :sDecMM )
call :sLengthenDate
exit /b

:sDecMM
set /a sD=(%sD%+%DayDecThreshold%)
set /a sM=(%sM%-1)
if %sM% equ 0 ( call :sDecYY )
exit /b

:sDecYY
set sM=12
set /a sY=(%sY%-1)
exit /b

::							Decrease Time
::							Increase Time

:sIncDD
call :sShortenDate
set /a sD=(%sD%+1)
call :sSetIncDayThreshold
if %sD% gtr %DayIncThreshold% ( call :sIncMM )
call :sLengthenDate
exit /b

:sIncMM
set sD=1
set /a sM=(%sM%+1)
if %sM% gtr 12 ( call :sIncYY )
exit /b

:sIncYY
set sM=1
set File=%Directory%%Folder%\%sY%
set File=%File: =%
rd /s /q %File% && Echo ( %sY% ) Empty folder was removed
set /a sY=(%sY%+1)
exit /b

::							Increase Time