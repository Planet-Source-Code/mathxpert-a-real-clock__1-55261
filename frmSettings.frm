VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3360
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3360
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CheckBox Check4 
      Caption         =   "Always on &top"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CheckBox Check3 
      Caption         =   "&Run minimized to tray on Windows start"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Activate &hourly chime"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Show the second hand"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   2055
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Applied As Boolean, tmpNum As String

Private Sub ChangeApply()
If Not Command3.Enabled Then Command3.Enabled = True
End Sub

Private Function BoolToLong(Expression As Boolean) As Long
On Error Resume Next
If Expression Then BoolToLong = 1 Else BoolToLong = 0
End Function

Private Sub ChangeSettings()
On Error Resume Next

If ShowSecondHand <> Check1.Value Then
    ShowSecondHand = Check1.Value
    frmMain.Line3.Visible = ShowSecondHand
End If

If UseChime <> Check2.Value Then UseChime = Check2.Value

If RunWin <> Check3.Value Then
    RunWin = Check3.Value
    If RunWin Then
        SetRegValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run" & IIf(IsWinNT, "", "Services"), App.Title, AppLocation
    Else
        DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run" & IIf(IsWinNT, "", "Services"), App.Title
    End If
End If

If OnTop <> Check4.Value Then
    OnTop = Check4.Value
    SetWinPos frmMain.hwnd, OnTop
End If
End Sub

Private Sub Check1_Click()
ChangeApply
End Sub

Private Sub Check2_Click()
ChangeApply
End Sub

Private Sub Check3_Click()
ChangeApply
End Sub

Private Sub Check4_Click()
ChangeApply
End Sub

Private Sub Command1_Click()
If Command3.Enabled Then Applied = True
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
ChangeSettings
If Command3.Enabled Then Command3.Enabled = False
If Applied Then Applied = False
End Sub

Private Sub Form_Load()
On Error Resume Next

Top = frmMain.Top + frmMain.Height - Height
Left = frmMain.Left + ((frmMain.Width - Width) / 2)

If Check1.Value <> BoolToLong(ShowSecondHand) Then Check1.Value = BoolToLong(ShowSecondHand)
If Check2.Value <> BoolToLong(UseChime) Then Check2.Value = BoolToLong(UseChime)
If Check3.Value <> BoolToLong(RunWin) Then Check3.Value = BoolToLong(RunWin)
If Check4.Value <> BoolToLong(OnTop) Then Check4.Value = BoolToLong(OnTop)

If Applied Then Applied = False
If Command3.Enabled Then Command3.Enabled = False

SetWinPos hwnd, OnTop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Applied Then ChangeSettings
End Sub
