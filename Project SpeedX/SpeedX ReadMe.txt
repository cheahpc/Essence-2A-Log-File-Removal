This version is written in batch file

Ussage: Double Click (SpeedX.bat) to run default setting
Default setting:
1) Log file will be created for each folder removal
2) Console will show latest status of the folder that is being removed
3) NG files will also be removed

Arguments available; *To call the exe using a batch file or from cmd
1) -log		This option will not create a log file
2) -s		This switch will show nothing on the console
3) -ng		This switch will skip the NG file for removal

Command Line Variable
Default
call "%~dp0SpeedoX - Copy.exe"

Without NG files
call "%userprofile%\desktop\SpeedoXY.exe" -ng

Without NG files and without log file
call "%userprofile%\desktop\SpeedoXY.exe" -ng -log

Without log file
call "%userprofile%\desktop\SpeedoXY.exe" -log

Without NG files and without log file and show nothing on console
call "%userprofile%\desktop\SpeedoXY.exe" -ng -log -s

Without log file and show nothing on console
call "%userprofile%\desktop\SpeedoXY.exe" -log -s

Without NG files and show nothing on console
call "%userprofile%\desktop\SpeedoXY.exe" -ng -s