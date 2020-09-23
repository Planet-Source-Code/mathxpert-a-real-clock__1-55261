VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   45
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   1740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   45
   ScaleWidth      =   1740
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mMenu 
      Caption         =   "&MyMenu"
      Begin VB.Menu mnuAtomicClock 
         Caption         =   "&Atomic Clock"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMinimize 
         Caption         =   "&Minimize"
      End
      Begin VB.Menu mnuSysTray 
         Caption         =   "&Minimize to System Tray"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mSys 
      Caption         =   "&MySys"
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuExit2 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuAtomicClock_Click()
On Error Resume Next

Unload frmAtomicClock
Load frmAtomicClock

frmAtomicClock.Show 1, frmMain
End Sub

Private Sub mnuExit_Click()
Unload frmMain
End Sub

Private Sub mnuExit2_Click()
Unload frmMain
End Sub

Private Sub mnuMinimize_Click()
If bSysTray Then bSysTray = False
frmMain.WindowState = 1
End Sub

Private Sub mnuRestore_Click()
Dim Result As Long
frmMain.WindowState = 0
Result = SetForegroundWindow(frmMain.hwnd)
frmMain.Show
End Sub

Private Sub mnuSettings_Click()
On Error Resume Next

Unload frmSettings
Load frmSettings

frmSettings.Show 1, frmMain
End Sub

Private Sub mnuSysTray_Click()
If Not bSysTray Then bSysTray = True
frmMain.WindowState = 1
End Sub
