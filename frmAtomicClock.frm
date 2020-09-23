VERSION 5.00
Begin VB.Form frmAtomicClock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atomic Clock"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4080
   Icon            =   "frmAtomicClock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "&Sync Now"
      Height          =   375
      Left            =   3090
      TabIndex        =   4
      Top             =   1200
      Width           =   885
   End
   Begin VB.Frame Frame2 
      Caption         =   "Schedule"
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   3855
      Begin VB.CommandButton Command6 
         Caption         =   ">"
         Height          =   255
         Left            =   2273
         TabIndex        =   9
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "<"
         Height          =   255
         Left            =   1313
         TabIndex        =   7
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0 hours"
         Height          =   255
         Left            =   1673
         TabIndex        =   8
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Update clock every:"
         Height          =   255
         Left            =   1313
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Synchronization Status"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.Label Label1 
         Caption         =   "The time has not been synchronized yet."
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "N/A"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Next update:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "frmAtomicClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const LB_S = " hour"
Private Const LB_P = " hours"

Dim Applied As Boolean
Dim tmpNum  As String
Dim Opening As Boolean
Dim n       As Single

Private Sub ChangeSettings()
If Offset <> CLng(tmpNum) Then
    Offset = CLng(tmpNum)
    ACDate = DateAdd("h", Offset, Now)
    If Offset = 0 Then
        If DateExists Then DateExists = False
        ProcessDt ""
    Else
        If Not DateExists Then DateExists = True
        ProcessDt DateAdd("h", 1, Now)
    End If
End If
End Sub

Private Sub Command1_Click()
If Command3.Enabled Then If Not Applied Then Applied = True
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

Private Sub Command4_Click()
SyncTime
End Sub

Private Sub Command5_Click()
If Not Command3.Enabled Then Command3.Enabled = True

If Not Command6.Enabled Then Command6.Enabled = True
If Label4 = "1" & LB_S Then
    If Command5.Enabled Then Command5.Enabled = False
Else
    If Not Command5.Enabled Then Command5.Enabled = True
End If

Select Case Label4
    Case "6" & LB_P: tmpNum = "5"
    Case "5" & LB_P: tmpNum = "4"
    Case "4" & LB_P: tmpNum = "3"
    Case "3" & LB_P: tmpNum = "2"
    Case "2" & LB_P: tmpNum = "1"
    Case "1" & LB_S: tmpNum = "0"
End Select

Label4 = tmpNum & IIf(tmpNum = "1", LB_S, LB_P)
End Sub

Private Sub Command6_Click()
If Not Command3.Enabled Then Command3.Enabled = True

If Not Command5.Enabled Then Command5.Enabled = True
If Label4 = "5" & LB_P Then
    If Command6.Enabled Then Command6.Enabled = False
Else
    If Not Command6.Enabled Then Command6.Enabled = True
End If

Select Case Label4
    Case "0" & LB_P: tmpNum = "1"
    Case "1" & LB_S: tmpNum = "2"
    Case "2" & LB_P: tmpNum = "3"
    Case "3" & LB_P: tmpNum = "4"
    Case "4" & LB_P: tmpNum = "5"
    Case "5" & LB_P: tmpNum = "6"
End Select

Label4 = tmpNum & IIf(tmpNum = "1", LB_S, LB_P)
End Sub

Private Sub Form_Load()
On Error Resume Next

Top = frmMain.Top + frmMain.Height - Height
Left = frmMain.Left + ((frmMain.Width - Width) / 2)

If Label1 <> Message Then Label1 = Message
If txtLabel4 <> Offset & IIf(Offset = 1, LB_S, LB_P) Then txtLabel4 = Offset & IIf(Offset = 1, LB_S, LB_P)
Select Case txtLabel4
    Case "0" & LB_P
        If Command5.Enabled Then Command5.Enabled = False
        If Not Command6.Enabled Then Command6.Enabled = True
    Case "6" & LB_P
        If Command6.Enabled Then Command6.Enabled = False
        If Not Command5.Enabled Then Command5.Enabled = True
    Case Else
        If Not Command5.Enabled Then Command5.Enabled = True
        If Not Command6.Enabled Then Command6.Enabled = True
End Select
If Label4 <> txtLabel4 Then Label4 = txtLabel4
If Label5 <> txtLabel5 Then Label5 = txtLabel5

If Applied Then Applied = False
If Command3.Enabled Then Command3.Enabled = False
ProcessMsg Message
ProcessDt IIf(DateExists, ACDate, "")

SetWinPos hwnd, OnTop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Applied Then ChangeSettings
End Sub
