Attribute VB_Name = "modWindowsNT"
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Function IsWinNT() As Boolean
Dim myOS As OSVERSIONINFO

myOS.dwOSVersionInfoSize = Len(myOS)
GetVersionEx myOS
IsWinNT = myOS.dwPlatformId = 2
End Function
