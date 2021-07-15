This version is written with vb.net

Ussage: Double Click (SpeedXY.exe) to run default setting
Default setting:
1) Log file will be created for each folder removal
3) NG files will also be removed

Arguments available; *To call the exe using a batch file or from cmd
1) -log		This option will not create a log file
2) -ng		This switch will skip the NG file for removal


Command Line Variable
Default
call "%userprofile%\desktop\SpeedoXY.exe"

Without NG files
call "%userprofile%\desktop\SpeedoXY.exe" -ng

Without NG files and without log file
call "%userprofile%\desktop\SpeedoXY.exe" -ng -log

Without log file
call "%userprofile%\desktop\SpeedoXY.exe" -log

