' #####------------------------------
' #	Import System
' #####------------------------------
Imports System.IO
Imports System.IO.Directory
Imports System.Windows.Forms
' #####------------------------------
' #	Main Module
' #####------------------------------
Public Module MainModule
	Dim IncludeNG As Boolean = True
	Dim LogMode As Boolean = True
	Dim Path() As String = {"D:\MNecoo\log\RoboC_ST2", "D:\MNecoo2\log\RoboD_ST2"}
	Dim CurrentPath As String

	Dim DaysToFirstDayOfLastWeek As Integer
	Dim DayDecThres, DayIncThres As Integer
	Dim Month(14) As Integer
	Dim sD, sM, sY As String
	Dim LogName As String = "D:\SpeedXY_Log.txt"
	Dim Log As System.IO.StreamWriter

	Dim FolderList(50) As String
	Dim ListCounter As Integer

' #####
' #	Function 	: Main Function	
' #	Description	: Remove all file from last week within range
' #	Argument	: 
' #	Return		:
' #####
	Public Sub Main(ByVal Argument() As String)
		Month = {1,3,5,7,8,10,12,1,2,4,6,8,9,11}
		Log = My.Computer.FileSystem.OpenTextFileWriter( LogName, True )	
		Log.WriteLine(DTime(1) & "-" & DTime(2) & "-" & DTime(3) & " ( " & DateTime.Now.DayOfWeek.ToString() & " )")

		If Argument.Length <> 0 Then
			For ArgCnt As Integer = 0 To UBound(Argument)
				If Argument(ArgCnt) = "-ng" Then
					IncludeNG = False
				End If
				If Argument(ArgCnt) = "-log" Then
					LogMode = False
				End IF
			Next
		End If

		' Loop through each path 
		For Each Location As String In Path
			CreateList(Location)

			If Location = "" Then 
				Exit For
			Else
				' Loop through each list
				For Each Folder As String In FolderList
					If Folder = "" Then
						Exit For
					Else
						SetTime(0) ' Decrease time set starting time
						' Loop through each day
						For i As Integer = 1 To 7
							CurrentPath = Folder & "\" & sY & "\" & sY & "_" & sM & "_" & sD
							DeleteCtrl(Folder)
							SetTime(1) ' Increase Time 
						Next
					End If
					SetTime(0)
				Next
			End If
			SetTime(0)
		Next	
		Log.Close()

		' Ending Message
		CreateObject("WScript.Shell").Popup("All Task Completed.",60,"Automated Log Removal")
		
	End Sub	
' #####------------------------------
' #	UDF - Argument Control
' #####------------------------------
' #####
' #	Function 			: 	Check_NgMode()
' #	Description			:	Check for NG flag
' #	Argument
' 		Str				:	Pass on the Foler Array
' #	UDF Used			:	N/A
' #	Variable Used		:	sD ; sM ; sY ; DayDecThres ; DaysToFirstDayOfLastWeek
' # Variable Affected	:	N/A
' #	Return				:	State of NgMode
' #####
	Function Check_NgMode(ByVal Str As String)
		If ( (IncludeNG = False) And (Str.Substring(Str.Length-2) = "NG")) Then
			Return 1
		Else
			Return 0
		End If
	End Function

' #	Function 			: 	LogCtrl()
' #	Description			:	Check for LogMode flag
' #	Argument
' #		Status			:	Message to write at the end of each log of remval ( Removed / Skipped ... Etc)
' #	UDF Used			:	DTime()
' #	Variable Used		:	LogMode ; CurrentPath
' # Variable Affected	:	N/A
' #	Return				:	N/A
' #####
	Sub LogCtrl(ByVal Status)
		If LogMode = True Then
			If Status <> "" Then
				Log.WriteLine("(" & DTime(4) & ":" & DTime(5) & ":" & DTime(6) & ") " & CurrentPath & "		>> " & Status)
			End IF
		End If
	End Sub

' #	Function 			: 	DeleteCtrl()
' #	Description			:	Delete File base on NgMode and File Exist condition
' #	Argument
' 		Str				:	Pass On the Folder Array To Check Mode
' #	UDF Used			:	LogCtrl() ; Check_NgMode()
' #	Variable Used		:	CurrentPath
' # Variable Affected	:	N/A
' #	Return				:	N/A
' #####
	Sub DeleteCtrl(ByVal Str As String)
		If Check_NgMode(Str) Then
			LogCtrl("Skipped")
		Else
			If Exists(CurrentPath) = True Then
				Delete(CurrentPath,True)
				LogCtrl("Removed")
			Else
				LogCtrl("Not Found")
			End If
			
		End If
	End Sub

' #	Function 			: 	CreateList()
' #	Description			:	Create a list of file in the location (NG, SBxxx ... etc)
' #	Argument
' 		Location		:	Pass On the Location String
' #	UDF Used			:	N/A 
' #	Variable Used		:	ListCounter ; FolderList ; Location
' # Variable Affected	:	ListCounter ; FolderList
' #	Return				:	N/A
' #####
	Sub CreateList(ByVal Location As String)
		ListCounter = 0
		Array.Clear(FolderList , 0 , FolderList.Length)
		For Each Dir As String In GetDirectories(Location)
			FolderList(ListCounter) = Dir
			ListCounter += 1
		Next
	End Sub


' #####------------------------------
' #	UDF - Time & Date Control
' #####------------------------------

' #####
' #	Function 			: 	SetStartTime()
' #	Description			:	Subtract prefix value of day form the day this script is run
' #							so the date is at the first day of last week
' #	Argument			:	N/A
' #	UDF Used			:	DTime()
' #	Variable Used		:	sD ; sM ; sY ; DayDecThres ; DaysToFirstDayOfLastWeek
' # Variable Affected	:	sD ; sM ; sY
' #	Return				:	N/A
' #####
	Sub SetStartTime()
		sD = DTime(3) - DaysToFirstDayOfLastWeek
		If sD < 0 Then 
			sD += DayDecThres
			If DTime(2) = 1 Then
				sM = 12
				sY = DTime(1) - 1
			Else
				sM = DTime(2) - 1
				sY = DTime(1)
			End If
		Else
			sM = DTime(2)
			sY = DTime(1)
		End If

	End Sub

' #####
' #	Function 			: 	SetNextTime()
' #	Description			:	Increase the date by one to set the correct date of the following day
' #	Argument			:	N/A
' #	UDF Used			:	SetDayIncThreshold() ; DTime()
' #	Variable Used		:	sD ; sM ; sY ; DayIncThres
' #	Variable Affected	:	sD ; sM ; sY
' #	Return				:	N/A
' #####
	Sub SetNextTime()
		SetDayIncThreshold()
		sD += 1
		If sD > DayIncThres Then
			sD = 1
			If DTime(2) = 12 Then
				sM = 1
				sY += 1
			Else
				sM +=1
				sY = sY 'No Changes
			End If
		Else
			sM = sM ' No Changes
			sY = sY 'No Changes
		End If
	End Sub

' #####
' #	Function 			: 	SetTime( [X] )
' #	Description			:	To produce desired value of the time by increasing or decreasing
' #							the date based on argument
' #	Argument			:	[X = 0,1,2]
' #	UDF Used			:	SetOffsets() ; SetDayDecThreshold() ; SetStartTime() ; SetNextTime()
' #	Variable Used		:	sD ; sM
' #	Variable Affected	:	sD ; sM ; sY
' #	Return				:	N/A
' #####
	Sub SetTime(ByVal X As Integer)
		SetOffsets() 'Set How Many Day To The 1st Day Of Previous Week
		SetDayDecThreshold() 'Set How Many Days Last Month Have
		If X = 0 Then 'Decrease Time
			SetStartTime()
		Else If X = 1 Then 'Increase Time
			SetNextTime()
		End If
		If sD.Length < 2 Then
			sD = 0 & sD
		End If
		If sM.Length < 2 Then
			sM = 0 & sM
		End If
	End Sub

' #####
' #	Function 			: 	SetDayDecThreshold ()
' #	Description			:	Set the maximum day of last month from the month when this script is run
' #	Argument			:	N/A
' #	UDF Used			:	DTime()
' #	Variable Used		:	DayDecThres ; sY
' #	Variable Affected	:	DayDecThres
' #	Return				:	N/A
' #####
	Sub SetDayDecThreshold()
		DayDecThres = 30
		For i As Integer = 7 To 13
			If Month(i) = DTime(2) Then
				DayDecThres = 31
			End If
		Next
		
		If DTime(2) = 3 Then
			If sY Mod 4 = 0 Then
				DayDecThres = 29
			Else
				DayDecThres = 28
			End If
		End If
	End Sub

' #####
' #	Function 			: 	SetDayIncThreshold ()
' #	Description			:	Set maximum number of day of the month to handle the overflow of
' #							date increment
' #	Argument			:	N/A
' #	UDF Used			:	N/A
' #	Variable Used		:	DayIncThres ; sM ; sY
' #	Variable Affected	:	DayIncThres
' #	Return				:	N/A
' #####
	Sub SetDayIncThreshold()
		DayIncThres = 30
		For i As Integer = 0 To 6
			If Month(i) = sM Then
				DayIncThres = 31
			End If
		Next
		
		If sM = 2 Then
			If sY Mod 4 = 0 Then
				DayIncThres = 29
			Else
				DayIncThres = 28
			End If
		End If
	End Sub

' #####
' #	Function 			: 	SetOffsets ()
' #	Description			:	Set how many day to the first day of last week from the day this
' #							script is run
' #	Argument			:	N/A
' #	UDF Used			:	DTime()
' #	Variable Used		:	DaysToFirstDayOfLastWeek
' #	Variable Affected	:	DaysToFirstDayOfLastWeek
' #	Return				:	N/A
' #####
	Sub SetOffsets()
		Dim x As Integer = 7
		DaysToFirstDayOfLastWeek = 14
		For i As Integer = 1 To 6
			DaysToFirstDayOfLastWeek = x
			If DTime(8) = i Then
				Exit For
			Else
				x+=1
			End If
		Next
	End Sub
	
' #####
' #	Function 			: 	DTime ()
' #	Description			:	Get current date and time from system and return corresponding value
' #	Argument			:	[ X = 1-8 ]
' #	UDF Used			:	N/A
' #	Variable Used		:	N/A
' #	Variable Affected	:	N/A
' #	Return				:	Integer
' #####
	Function DTime( ByVal X As Integer ) As String
		If X = 1 Then
			return DateTime.Now.ToString("yyyy") ' Return Year in full
		Else If X = 2
			return DateTime.Now.ToString("MM") ' Return Month 
		Else If X = 3
			return DateTime.Now.ToString("dd") ' Return Date
		Else If X = 4
			return DateTime.Now.ToString("HH") ' Return Hour
		Else If X = 5
			return DateTime.Now.ToString("mm") ' Return Minutes
		Else If X = 6
			return DateTime.Now.ToString("ss") ' Return Seconds
		Else If X = 7
			return DateTime.Now.ToString("( HH:mm:ss )") ' Return Time in full
		Else If X = 8
			return DateTime.Now.DayOfWeek() ' Return Day of week
		End If
		Return 0
	End Function

End Module