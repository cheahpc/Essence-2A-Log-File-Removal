Imports System
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Public Module MainModule
	Dim Delay As Integer = 800
	Dim PathA As String = "D:\MNecoo\log\RoboC_ST2"
	Dim PathB As String = "D:\MNecoo2\log\RoboD_ST2"
	Const KEYDOWN As Integer = &H0
	Const KEYUP As Integer = &H2

	Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, byVal dwFlags As UInteger, ByVal dwExtracInfo As UInteger)

	Public Sub Main()
		

		Process.Start("Explorer.exe", PathA)
		Process.Start("Explorer.exe", PathB)				

		Threading.Thread.Sleep(Delay)
		SetWindowsLocation("RoboC_ST2", "Left")
		SetWindowsLocation("RoboD_ST2", "Right")

	End Sub
	
	Sub SetWindowsLocation(ByVal Name As String, ByVal Position As String)
		AppActivate(Name)
		If Position = "Left" Then
			keybd_event(CByte(Keys.LWin), 0, KEYDOWN, 0)
			keybd_event(CByte(Keys.Left), 0, KEYDOWN, 0)
			keybd_event(CByte(Keys.Left), 0, KEYUP, 0)
			keybd_event(CByte(Keys.LWin), 0, KEYUP, 0)
		Else if Position = "Right"
			keybd_event(CByte(Keys.LWin), 0, KEYDOWN, 0)
			keybd_event(CByte(Keys.Right), 0, KEYDOWN, 0)
			keybd_event(CByte(Keys.Right), 0, KEYUP, 0)
			keybd_event(CByte(Keys.LWin), 0, KEYUP, 0)
		End If
	End Sub

End Module