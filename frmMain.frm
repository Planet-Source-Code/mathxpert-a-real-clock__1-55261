VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Clock"
   ClientHeight    =   5880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3720
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3240
      Top             =   5400
   End
   Begin Clock.WebSock WebSock1 
      Left            =   0
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   5400
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2520
      Top             =   4320
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   240
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   720
      Top             =   4320
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   3120
      Top             =   120
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   5400
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2520
      Top             =   5400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000D7FF&
      BorderWidth     =   4
      X1              =   1860
      X2              =   1860
      Y1              =   3695
      Y2              =   5280
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000D7FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1620
      Shape           =   2  'Oval
      Top             =   5040
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   1733
      Shape           =   2  'Oval
      Top             =   1733
      Width           =   255
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   920
      X2              =   1860
      Y1              =   3154
      Y2              =   1860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      X1              =   3204
      X2              =   1860
      Y1              =   1193
      Y2              =   1860
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   6
      X1              =   1041
      X2              =   1860
      Y1              =   1286
      Y2              =   1860
   End
   Begin VB.Line TickLine 
      Index           =   59
      X1              =   1682
      X2              =   1698
      Y1              =   169
      Y2              =   318
   End
   Begin VB.Line TickLine 
      Index           =   58
      X1              =   1507
      X2              =   1538
      Y1              =   197
      Y2              =   344
   End
   Begin VB.Line TickLine 
      Index           =   57
      X1              =   1335
      X2              =   1381
      Y1              =   243
      Y2              =   386
   End
   Begin VB.Line TickLine 
      Index           =   56
      X1              =   1169
      X2              =   1230
      Y1              =   307
      Y2              =   444
   End
   Begin VB.Line TickLine 
      BorderWidth     =   3
      Index           =   55
      X1              =   1010
      X2              =   1085
      Y1              =   388
      Y2              =   518
   End
   Begin VB.Line TickLine 
      Index           =   54
      X1              =   861
      X2              =   949
      Y1              =   485
      Y2              =   606
   End
   Begin VB.Line TickLine 
      Index           =   53
      X1              =   722
      X2              =   823
      Y1              =   597
      Y2              =   708
   End
   Begin VB.Line TickLine 
      Index           =   52
      X1              =   597
      X2              =   708
      Y1              =   722
      Y2              =   823
   End
   Begin VB.Line TickLine 
      Index           =   51
      X1              =   485
      X2              =   606
      Y1              =   861
      Y2              =   949
   End
   Begin VB.Line TickLine 
      BorderWidth     =   3
      Index           =   50
      X1              =   388
      X2              =   518
      Y1              =   1010
      Y2              =   1085
   End
   Begin VB.Line TickLine 
      Index           =   49
      X1              =   307
      X2              =   444
      Y1              =   1169
      Y2              =   1230
   End
   Begin VB.Line TickLine 
      Index           =   48
      X1              =   243
      X2              =   386
      Y1              =   1335
      Y2              =   1381
   End
   Begin VB.Line TickLine 
      Index           =   47
      X1              =   197
      X2              =   344
      Y1              =   1507
      Y2              =   1538
   End
   Begin VB.Line TickLine 
      Index           =   46
      X1              =   169
      X2              =   318
      Y1              =   1682
      Y2              =   1698
   End
   Begin VB.Line TickLine 
      BorderWidth     =   3
      Index           =   45
      X1              =   160
      X2              =   310
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line TickLine 
      Index           =   44
      X1              =   169
      X2              =   318
      Y1              =   2038
      Y2              =   2022
   End
   Begin VB.Line TickLine 
      Index           =   43
      X1              =   197
      X2              =   344
      Y1              =   2213
      Y2              =   2182
   End
   Begin VB.Line TickLine 
      Index           =   42
      X1              =   243
      X2              =   386
      Y1              =   2385
      Y2              =   2339
   End
   Begin VB.Line TickLine 
      Index           =   41
      X1              =   307
      X2              =   444
      Y1              =   2551
      Y2              =   2490
   End
   Begin VB.Line TickLine 
      BorderWidth     =   3
      Index           =   40
      X1              =   388
      X2              =   518
      Y1              =   2710
      Y2              =   2635
   End
   Begin VB.Line TickLine 
      Index           =   39
      X1              =   485
      X2              =   606
      Y1              =   2859
      Y2              =   2771
   End
   Begin VB.Line TickLine 
      Index           =   38
      X1              =   597
      X2              =   708
      Y1              =   2998
      Y2              =   2897
   End
   Begin VB.Line TickLine 
      Index           =   37
      X1              =   722
      X2              =   823
      Y1              =   3123
      Y2              =   3012
   End
   Begin VB.Line TickLine 
      Index           =   36
      X1              =   861
      X2              =   949
      Y1              =   3235
      Y2              =   3114
   End
   Begin VB.Line TickLine 
      BorderWidth     =   3
      Index           =   35
      X1              =   1010
      X2              =   1085
      Y1              =   3332
      Y2              =   3202
   End
   Begin VB.Line TickLine 
      Index           =   34
      X1              =   1169
      X2              =   1230
      Y1              =   3413
      Y2              =   3276
   End
   Begin VB.Line TickLine 
      Index           =   33
      X1              =   1335
      X2              =   1381
      Y1              =   3477
      Y2              =   3334
   End
   Begin VB.Line TickLine 
      Index           =   32
      X1              =   1507
      X2              =   1538
      Y1              =   3523
      Y2              =   3376
   End
   Begin VB.Line TickLine 
      Index           =   31
      X1              =   1682
      X2              =   1698
      Y1              =   3551
      Y2              =   3402
   End
   Begin VB.Line TickLine 
      BorderWidth     =   3
      Index           =   30
      X1              =   1860
      X2              =   1860
      Y1              =   3560
      Y2              =   3410
   End
   Begin VB.Line TickLine 
      Index           =   29
      X1              =   2038
      X2              =   2022
      Y1              =   3551
      Y2              =   3402
   End
   Begin VB.Line TickLine 
      Index           =   28
      X1              =   2213
      X2              =   2182
      Y1              =   3523
      Y2              =   3376
   End
   Begin VB.Line TickLine 
      Index           =   27
      X1              =   2385
      X2              =   2339
      Y1              =   3477
      Y2              =   3334
   End
   Begin VB.Line TickLine 
      Index           =   26
      X1              =   2551
      X2              =   2490
      Y1              =   3413
      Y2              =   3276
   End
   Begin VB.Line TickLine 
      BorderWidth     =   3
      Index           =   25
      X1              =   2710
      X2              =   2635
      Y1              =   3332
      Y2              =   3202
   End
   Begin VB.Line TickLine 
      Index           =   24
      X1              =   2859
      X2              =   2771
      Y1              =   3235
      Y2              =   3114
   End
   Begin VB.Line TickLine 
      Index           =   23
      X1              =   2998
      X2              =   2897
      Y1              =   3123
      Y2              =   3012
   End
   Begin VB.Line TickLine 
      Index           =   22
      X1              =   3123
      X2              =   3012
      Y1              =   2998
      Y2              =   2897
   End
   Begin VB.Line TickLine 
      Index           =   21
      X1              =   3235
      X2              =   3114
      Y1              =   2859
      Y2              =   2771
   End
   Begin VB.Line TickLine 
      BorderWidth     =   3
      Index           =   20
      X1              =   3332
      X2              =   3202
      Y1              =   2710
      Y2              =   2635
   End
   Begin VB.Line TickLine 
      Index           =   19
      X1              =   3413
      X2              =   3276
      Y1              =   2551
      Y2              =   2490
   End
   Begin VB.Line TickLine 
      Index           =   18
      X1              =   3477
      X2              =   3334
      Y1              =   2385
      Y2              =   2339
   End
   Begin VB.Line TickLine 
      Index           =   17
      X1              =   3523
      X2              =   3376
      Y1              =   2213
      Y2              =   2182
   End
   Begin VB.Line TickLine 
      Index           =   16
      X1              =   3551
      X2              =   3402
      Y1              =   2038
      Y2              =   2022
   End
   Begin VB.Line TickLine 
      BorderWidth     =   3
      Index           =   15
      X1              =   3560
      X2              =   3410
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line TickLine 
      Index           =   14
      X1              =   3551
      X2              =   3402
      Y1              =   1682
      Y2              =   1698
   End
   Begin VB.Line TickLine 
      Index           =   13
      X1              =   3523
      X2              =   3376
      Y1              =   1507
      Y2              =   1538
   End
   Begin VB.Line TickLine 
      Index           =   12
      X1              =   3477
      X2              =   3334
      Y1              =   1335
      Y2              =   1381
   End
   Begin VB.Line TickLine 
      Index           =   11
      X1              =   3413
      X2              =   3276
      Y1              =   1169
      Y2              =   1230
   End
   Begin VB.Line TickLine 
      BorderWidth     =   3
      Index           =   10
      X1              =   3332
      X2              =   3202
      Y1              =   1010
      Y2              =   1085
   End
   Begin VB.Line TickLine 
      Index           =   9
      X1              =   3235
      X2              =   3114
      Y1              =   861
      Y2              =   949
   End
   Begin VB.Line TickLine 
      Index           =   8
      X1              =   3123
      X2              =   3012
      Y1              =   722
      Y2              =   823
   End
   Begin VB.Line TickLine 
      Index           =   7
      X1              =   2998
      X2              =   2897
      Y1              =   597
      Y2              =   708
   End
   Begin VB.Line TickLine 
      Index           =   6
      X1              =   2859
      X2              =   2771
      Y1              =   485
      Y2              =   606
   End
   Begin VB.Line TickLine 
      BorderWidth     =   3
      Index           =   5
      X1              =   2710
      X2              =   2635
      Y1              =   388
      Y2              =   518
   End
   Begin VB.Line TickLine 
      Index           =   4
      X1              =   2551
      X2              =   2490
      Y1              =   307
      Y2              =   444
   End
   Begin VB.Line TickLine 
      Index           =   3
      X1              =   2385
      X2              =   2339
      Y1              =   243
      Y2              =   386
   End
   Begin VB.Line TickLine 
      Index           =   2
      X1              =   2213
      X2              =   2182
      Y1              =   197
      Y2              =   344
   End
   Begin VB.Line TickLine 
      Index           =   1
      X1              =   2038
      X2              =   2022
      Y1              =   169
      Y2              =   318
   End
   Begin VB.Line TickLine 
      BorderWidth     =   3
      Index           =   0
      X1              =   1860
      X2              =   1860
      Y1              =   160
      Y2              =   310
   End
   Begin VB.Image Image6 
      Height          =   225
      Index           =   10
      Left            =   2640
      Picture         =   "frmMain.frx":0E42
      Top             =   6240
      Width           =   660
   End
   Begin VB.Image Image6 
      Height          =   225
      Index           =   9
      Left            =   2400
      Picture         =   "frmMain.frx":11FE
      Top             =   6240
      Width           =   660
   End
   Begin VB.Image Image6 
      Height          =   225
      Index           =   8
      Left            =   2160
      Picture         =   "frmMain.frx":15BA
      Top             =   6240
      Width           =   660
   End
   Begin VB.Image Image6 
      Height          =   225
      Index           =   7
      Left            =   1920
      Picture         =   "frmMain.frx":1976
      Top             =   6240
      Width           =   660
   End
   Begin VB.Image Image6 
      Height          =   225
      Index           =   6
      Left            =   1680
      Picture         =   "frmMain.frx":1D32
      Top             =   6240
      Width           =   660
   End
   Begin VB.Image Image6 
      Height          =   225
      Index           =   5
      Left            =   1440
      Picture         =   "frmMain.frx":20EE
      Top             =   6240
      Width           =   660
   End
   Begin VB.Image Image6 
      Height          =   225
      Index           =   4
      Left            =   1200
      Picture         =   "frmMain.frx":24AA
      Top             =   6240
      Width           =   660
   End
   Begin VB.Image Image6 
      Height          =   225
      Index           =   3
      Left            =   960
      Picture         =   "frmMain.frx":2866
      Top             =   6240
      Width           =   660
   End
   Begin VB.Image Image6 
      Height          =   225
      Index           =   2
      Left            =   720
      Picture         =   "frmMain.frx":2C22
      Top             =   6240
      Width           =   660
   End
   Begin VB.Image Image6 
      Height          =   225
      Index           =   1
      Left            =   480
      Picture         =   "frmMain.frx":2FDE
      Top             =   6240
      Width           =   660
   End
   Begin VB.Image Image6 
      Height          =   225
      Index           =   0
      Left            =   240
      Picture         =   "frmMain.frx":339A
      Top             =   6240
      Width           =   660
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   10
      Left            =   2640
      Picture         =   "frmMain.frx":3756
      Top             =   6000
      Width           =   675
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   9
      Left            =   2400
      Picture         =   "frmMain.frx":3B21
      Top             =   6000
      Width           =   675
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   8
      Left            =   2160
      Picture         =   "frmMain.frx":3EEC
      Top             =   6000
      Width           =   675
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   7
      Left            =   1920
      Picture         =   "frmMain.frx":42B7
      Top             =   6000
      Width           =   675
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   6
      Left            =   1680
      Picture         =   "frmMain.frx":4682
      Top             =   6000
      Width           =   675
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   5
      Left            =   1440
      Picture         =   "frmMain.frx":4A4D
      Top             =   6000
      Width           =   675
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   4
      Left            =   1200
      Picture         =   "frmMain.frx":4E18
      Top             =   6000
      Width           =   675
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   3
      Left            =   960
      Picture         =   "frmMain.frx":51E3
      Top             =   6000
      Width           =   675
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   2
      Left            =   720
      Picture         =   "frmMain.frx":55AE
      Top             =   6000
      Width           =   675
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   1
      Left            =   480
      Picture         =   "frmMain.frx":5979
      Top             =   6000
      Width           =   675
   End
   Begin VB.Image Image5 
      Height          =   225
      Index           =   0
      Left            =   240
      Picture         =   "frmMain.frx":5D44
      Top             =   6000
      Width           =   675
   End
   Begin VB.Image Image4 
      Height          =   225
      Left            =   1870
      Picture         =   "frmMain.frx":610F
      ToolTipText     =   "Settings"
      Top             =   5610
      Width           =   660
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   1200
      Picture         =   "frmMain.frx":64CB
      ToolTipText     =   "Atomic Clock"
      Top             =   5610
      Width           =   675
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1050
      TabIndex        =   12
      Top             =   540
      Width           =   495
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   570
      TabIndex        =   11
      Top             =   1020
      Width           =   495
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   570
      TabIndex        =   9
      Top             =   2385
      Width           =   495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1050
      TabIndex        =   8
      Top             =   2865
      Width           =   495
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1620
      TabIndex        =   7
      Top             =   3030
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   2865
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2385
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1020
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   540
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1620
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   3735
      Left            =   113
      Shape           =   3  'Circle
      Top             =   -7
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1935
      Left            =   1193
      Top             =   3668
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00002080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   2295
      Left            =   1073
      Top             =   3233
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00002080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   1073
      Shape           =   2  'Oval
      Top             =   5033
      Width           =   1575
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00002080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3735
      Left            =   -7
      Shape           =   3  'Circle
      Top             =   -7
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (lpszSoundName As Byte, ByVal uFlags As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_MEMORY = &H4

Private Const PI = 3.14159265358979

Dim Ind                 As Single
Dim Ind2                As Integer
Dim Ind3                As Integer
Dim Subt                As Boolean
Dim Subt2               As Boolean
Dim Subt3               As Boolean
Dim Pt                  As POINTAPI
Dim Points(1 To 315)    As POINTAPI
Dim ButtonDown          As Boolean

Dim XP                  As Single
Dim YP                  As Single

Dim XPos                As Single
Dim YPos                As Single

Dim Hr                  As Single
Dim realHr              As Single
Dim Min                 As Single
Dim realMin             As Single
Dim Sec                 As Single
Dim LastHr              As Single
Dim LastMin             As Single
Dim LastSec             As Single

Dim SoundPlayed         As Boolean
Dim b()                 As Byte

Dim DisableMenu         As Boolean

Private Function IsBool(Expression) As Boolean
On Error GoTo hndl
Expression = LCase(CStr(Expression))
IsBool = IsNumeric(Expression)
If IsBool Then Exit Function
IsBool = (Expression = "true" Or Expression = "false")
Exit Function
hndl:
    IsBool = False
End Function

Private Sub SaveSettings()
SaveSetting App.Title, "AtomicClock", "Offset", Offset
SaveSetting App.Title, "Settings", "SecondHand", ShowSecondHand
SaveSetting App.Title, "Settings", "Chime", UseChime
SaveSetting App.Title, "Settings", "OnTop", OnTop
End Sub

Private Sub GetSettings()
Dim tmpStg, tmpStg2, tmpStg3, tmpStg4, tmpStg5

tmpStg = GetSetting(App.Title, "AtomicClock", "Offset", 0)
If IsNumeric(tmpStg) Then
    If tmpStg >= 0 And tmpStg <= 6 Then
        Offset = tmpStg
    Else
        Offset = 0
    End If
Else
    Offset = 0
End If

tmpStg2 = GetSetting(App.Title, "Settings", "SecondHand", True)
If IsBool(tmpStg2) Then ShowSecondHand = tmpStg2 Else ShowSecondHand = True
If Line3.Visible <> ShowSecondHand Then Line3.Visible = ShowSecondHand

tmpStg3 = GetSetting(App.Title, "Settings", "Chime", True)
If IsBool(tmpStg3) Then UseChime = tmpStg3 Else UseChime = True

tmpStg4 = GetRegValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run" & IIf(IsWinNT, "", "Services"), App.Title)
If tmpStg4 <> "" Then
    If tmpStg4 <> AppLocation Then SetRegValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run" & IIf(IsWinNT, "", "Services"), App.Title, AppLocation
End If
RunWin = tmpStg4 <> ""

tmpStg5 = GetSetting(App.Title, "Settings", "OnTop", False)
If IsBool(tmpStg5) Then OnTop = tmpStg5 Else OnTop = False

DateExists = Offset
ProcessDt IIf(DateExists, DateAdd("h", Offset, Now), "")

If DontSync Then DontSync = False
SetWinPos hwnd, OnTop
End Sub

Private Sub GlowButtons(X As Single, Y As Single)
X = ScaleX(X, vbPixels, vbTwips)
Y = ScaleY(Y, vbPixels, vbTwips)

With Image3
    If X >= .Left And X <= .Left + .Width And Y >= .Top And Y <= .Top + .Height Then
        If Subt2 Then Subt2 = False
    Else
        If Not Subt2 Then Subt2 = True
    End If
End With

With Image4
    If X >= .Left And X <= .Left + .Width And Y >= .Top And Y <= .Top + .Height Then
        If Subt3 Then Subt3 = False
    Else
        If Not Subt3 Then Subt3 = True
    End If
End With
End Sub

Private Sub ModifyHands()
If Hr <> Hour(Now) Then Hr = Hour(Now)
If Min <> Minute(Now) Then Min = Minute(Now)
If Sec <> Second(Now) Then Sec = Second(Now)

If realHr <> Hr + Min / 60 Then realHr = Hr + Min / 60
If realMin <> Min + Sec / 60 Then realMin = Min + Sec / 60

If LastSec <> Sec Then
    If Sec > -1 And Sec < 61 Then
        Line3.X1 = 1600 * Cos(PI / 180 * (6 * Sec - 90)) + Line3.X2
        Line3.Y1 = 1600 * Sin(PI / 180 * (6 * Sec - 90)) + Line3.Y2
    End If
End If

If LastMin <> realMin Then
    If realMin > -1 And realMin < 61 Then
        Line1.X1 = 1500 * Cos(PI / 180 * (6 * realMin - 90)) + Line1.X2
        Line1.Y1 = 1500 * Sin(PI / 180 * (6 * realMin - 90)) + Line1.Y2
    End If
End If

If LastHr <> realHr Then
    If realHr > -1 And realHr < 25 Then
        Line2.X1 = 1000 * Cos(PI / 180 * (30 * realHr - 90)) + Line2.X2
        Line2.Y1 = 1000 * Sin(PI / 180 * (30 * realHr - 90)) + Line2.Y2
    End If
End If

If LastHr <> realHr Then LastHr = realHr
If LastMin <> realMin Then LastMin = realMin
If LastSec <> Sec Then LastSec = Sec
End Sub

Private Sub MD(X As Single, Y As Single)
If Not ButtonDown Then ButtonDown = True
If XP <> X Then XP = X
If YP <> Y Then YP = Y
End Sub

Private Sub MM(X As Single, Y As Single)
If ButtonDown Then
    If Left <> Left + X - XP Then Left = Left + X - XP
    If Top <> Top + Y - YP Then Top = Top + Y - YP
End If
End Sub

Private Sub MU(Button As Integer)
If ButtonDown Then ButtonDown = False
If Button = 2 Then PopupMenu frmMenu.mMenu
End Sub

Private Sub ReshapeFormToCustomShape()
Dim hRgn As Long

Points(1).X = 0
Points(1).Y = 119

Points(2).X = 1
Points(2).Y = 107

Points(3).X = 3
Points(3).Y = 96

Points(4).X = 4
Points(4).Y = 92

Points(5).X = 5
Points(5).Y = 88

Points(6).X = 6
Points(6).Y = 85

Points(7).X = 7
Points(7).Y = 83

Points(8).X = 8
Points(8).Y = 80

Points(9).X = 9
Points(9).Y = 77

Points(10).X = 10
Points(10).Y = 75

Points(11).X = 11
Points(11).Y = 74

Points(12).X = 12
Points(12).Y = 71

Points(13).X = 13
Points(13).Y = 69

Points(14).X = 14
Points(14).Y = 66

Points(15).X = 15
Points(15).Y = 64

Points(16).X = 16
Points(16).Y = 63

Points(17).X = 17
Points(17).Y = 61

Points(18).X = 18
Points(18).Y = 60

Points(19).X = 19
Points(19).Y = 58

Points(20).X = 20
Points(20).Y = 56

Points(21).X = 21
Points(21).Y = 55

Points(22).X = 22
Points(22).Y = 53

Points(23).X = 23
Points(23).Y = 52

Points(24).X = 24
Points(24).Y = 51

Points(25).X = 25
Points(25).Y = 50

Points(26).X = 26
Points(26).Y = 49

Points(27).X = 27
Points(27).Y = 47

Points(28).X = 28
Points(28).Y = 46

Points(29).X = 29
Points(29).Y = 45

Points(30).X = 30
Points(30).Y = 44

Points(31).X = 31
Points(31).Y = 43

Points(32).X = 32
Points(32).Y = 41

Points(33).X = 33
Points(33).Y = 40

Points(34).X = 34
Points(34).Y = 39

Points(35).X = 35
Points(35).Y = 38

Points(36).X = 36
Points(36).Y = 36

Points(37).X = 38
Points(37).Y = 35

Points(38).X = 39
Points(38).Y = 34

Points(39).X = 40
Points(39).Y = 33

Points(40).X = 41
Points(40).Y = 32

Points(41).X = 43
Points(41).Y = 31

Points(42).X = 44
Points(42).Y = 30

Points(43).X = 45
Points(43).Y = 29

Points(44).X = 46
Points(44).Y = 28

Points(45).X = 47
Points(45).Y = 27

Points(46).X = 49
Points(46).Y = 26

Points(47).X = 50
Points(47).Y = 25

Points(48).X = 51
Points(48).Y = 24

Points(49).X = 52
Points(49).Y = 23

Points(50).X = 53
Points(50).Y = 22

Points(51).X = 55
Points(51).Y = 21

Points(52).X = 56
Points(52).Y = 20

Points(53).X = 58
Points(53).Y = 19

Points(54).X = 60
Points(54).Y = 18

Points(55).X = 61
Points(55).Y = 17

Points(56).X = 63
Points(56).Y = 16

Points(57).X = 64
Points(57).Y = 15

Points(58).X = 66
Points(58).Y = 14

Points(59).X = 69
Points(59).Y = 13

Points(60).X = 71
Points(60).Y = 12

Points(61).X = 73
Points(61).Y = 11

Points(62).X = 75
Points(62).Y = 10

Points(63).X = 77
Points(63).Y = 9

Points(64).X = 80
Points(64).Y = 8

Points(65).X = 83
Points(65).Y = 7

Points(66).X = 85
Points(66).Y = 6

Points(67).X = 88
Points(67).Y = 5

Points(68).X = 92
Points(68).Y = 4

Points(69).X = 96
Points(69).Y = 3

Points(70).X = 101
Points(70).Y = 2

Points(71).X = 107
Points(71).Y = 1

Points(72).X = 119
Points(72).Y = 0

Points(73).X = 140
Points(73).Y = 1

Points(74).X = 146
Points(74).Y = 2

Points(75).X = 151
Points(75).Y = 3

Points(76).X = 155
Points(76).Y = 4

Points(77).X = 159
Points(77).Y = 5

Points(78).X = 162
Points(78).Y = 6

Points(79).X = 164
Points(79).Y = 7

Points(80).X = 167
Points(80).Y = 8

Points(81).X = 170
Points(81).Y = 9

Points(82).X = 172
Points(82).Y = 10

Points(83).X = 174
Points(83).Y = 11

Points(84).X = 176
Points(84).Y = 12

Points(85).X = 178
Points(85).Y = 13

Points(86).X = 181
Points(86).Y = 14

Points(87).X = 183
Points(87).Y = 15

Points(88).X = 184
Points(88).Y = 16

Points(89).X = 186
Points(89).Y = 17

Points(90).X = 187
Points(90).Y = 18

Points(91).X = 189
Points(91).Y = 19

Points(92).X = 191
Points(92).Y = 20

Points(93).X = 192
Points(93).Y = 21

Points(94).X = 194
Points(94).Y = 22

Points(95).X = 195
Points(95).Y = 23

Points(96).X = 196
Points(96).Y = 24

Points(97).X = 197
Points(97).Y = 25

Points(98).X = 198
Points(98).Y = 26

Points(99).X = 200
Points(99).Y = 27

Points(100).X = 201
Points(100).Y = 28

Points(101).X = 202
Points(101).Y = 29

Points(102).X = 203
Points(102).Y = 30

Points(103).X = 204
Points(103).Y = 31

Points(104).X = 206
Points(104).Y = 32

Points(105).X = 207
Points(105).Y = 33

Points(106).X = 208
Points(106).Y = 34

Points(107).X = 209
Points(107).Y = 35

Points(108).X = 210
Points(108).Y = 36

Points(109).X = 211
Points(109).Y = 37

Points(110).X = 212
Points(110).Y = 38

Points(111).X = 213
Points(111).Y = 39

Points(112).X = 214
Points(112).Y = 40

Points(113).X = 215
Points(113).Y = 41

Points(114).X = 216
Points(114).Y = 43

Points(115).X = 217
Points(115).Y = 44

Points(116).X = 218
Points(116).Y = 45

Points(117).X = 219
Points(117).Y = 46

Points(118).X = 220
Points(118).Y = 47

Points(119).X = 221
Points(119).Y = 49

Points(120).X = 222
Points(120).Y = 50

Points(121).X = 223
Points(121).Y = 51

Points(122).X = 224
Points(122).Y = 52

Points(123).X = 225
Points(123).Y = 53

Points(124).X = 226
Points(124).Y = 55

Points(125).X = 227
Points(125).Y = 56

Points(126).X = 228
Points(126).Y = 58

Points(127).X = 229
Points(127).Y = 60

Points(128).X = 230
Points(128).Y = 61

Points(129).X = 231
Points(129).Y = 63

Points(130).X = 232
Points(130).Y = 64

Points(131).X = 233
Points(131).Y = 66

Points(132).X = 234
Points(132).Y = 69

Points(133).X = 235
Points(133).Y = 71

Points(134).X = 236
Points(134).Y = 73

Points(135).X = 237
Points(135).Y = 75

Points(136).X = 238
Points(136).Y = 77

Points(137).X = 239
Points(137).Y = 80

Points(138).X = 240
Points(138).Y = 83

Points(139).X = 241
Points(139).Y = 85

Points(140).X = 242
Points(140).Y = 88

Points(141).X = 243
Points(141).Y = 92

Points(142).X = 244
Points(142).Y = 96

Points(143).X = 245
Points(143).Y = 101

Points(144).X = 246
Points(144).Y = 107

Points(145).X = 247
Points(145).Y = 119

Points(146).X = 247
Points(146).Y = 128

Points(147).X = 246
Points(147).Y = 140

Points(148).X = 245
Points(148).Y = 146

Points(149).X = 244
Points(149).Y = 151

Points(150).X = 243
Points(150).Y = 155

Points(151).X = 242
Points(151).Y = 159

Points(152).X = 241
Points(152).Y = 162

Points(153).X = 240
Points(153).Y = 164

Points(154).X = 239
Points(154).Y = 167

Points(155).X = 238
Points(155).Y = 170

Points(156).X = 237
Points(156).Y = 172

Points(157).X = 236
Points(157).Y = 174

Points(158).X = 235
Points(158).Y = 176

Points(159).X = 234
Points(159).Y = 178

Points(160).X = 233
Points(160).Y = 181

Points(161).X = 232
Points(161).Y = 183

Points(162).X = 231
Points(162).Y = 184

Points(163).X = 230
Points(163).Y = 186

Points(164).X = 229
Points(164).Y = 187

Points(165).X = 228
Points(165).Y = 189

Points(166).X = 227
Points(166).Y = 191

Points(167).X = 226
Points(167).Y = 192

Points(168).X = 225
Points(168).Y = 194

Points(169).X = 224
Points(169).Y = 195

Points(170).X = 223
Points(170).Y = 196

Points(171).X = 222
Points(171).Y = 197

Points(172).X = 221
Points(172).Y = 198

Points(173).X = 220
Points(173).Y = 200

Points(174).X = 219
Points(174).Y = 201

Points(175).X = 218
Points(175).Y = 202

Points(176).X = 217
Points(176).Y = 203

Points(177).X = 216
Points(177).Y = 204

Points(178).X = 215
Points(178).Y = 206

Points(179).X = 214
Points(179).Y = 207

Points(180).X = 213
Points(180).Y = 208

Points(181).X = 212
Points(181).Y = 209

Points(182).X = 211
Points(182).Y = 210

Points(183).X = 210
Points(183).Y = 211

Points(184).X = 209
Points(184).Y = 212

Points(185).X = 208
Points(185).Y = 213

Points(186).X = 207
Points(186).Y = 214

Points(187).X = 206
Points(187).Y = 215

Points(188).X = 204
Points(188).Y = 216

Points(189).X = 203
Points(189).Y = 217

Points(190).X = 202
Points(190).Y = 218

Points(191).X = 201
Points(191).Y = 219

Points(192).X = 200
Points(192).Y = 220

Points(193).X = 198
Points(193).Y = 221

Points(194).X = 197
Points(194).Y = 222

Points(195).X = 196
Points(195).Y = 223

Points(196).X = 195
Points(196).Y = 224

Points(197).X = 194
Points(197).Y = 225

Points(198).X = 192
Points(198).Y = 226

Points(199).X = 191
Points(199).Y = 227

Points(200).X = 189
Points(200).Y = 228

Points(201).X = 187
Points(201).Y = 229

Points(202).X = 186
Points(202).Y = 230

Points(203).X = 184
Points(203).Y = 231

Points(204).X = 183
Points(204).Y = 232

Points(205).X = 181
Points(205).Y = 233

Points(206).X = 178
Points(206).Y = 234

Points(207).X = 176
Points(207).Y = 235

Points(208).X = 175
Points(208).Y = 235

Points(209).X = 175
Points(209).Y = 367

Points(210).X = 174
Points(210).Y = 369

Points(211).X = 173
Points(211).Y = 371

Points(212).X = 172
Points(212).Y = 372

Points(213).X = 171
Points(213).Y = 374

Points(214).X = 170
Points(214).Y = 375

Points(215).X = 169
Points(215).Y = 376

Points(216).X = 168
Points(216).Y = 377

Points(217).X = 167
Points(217).Y = 378

Points(218).X = 166
Points(218).Y = 379

Points(219).X = 164
Points(219).Y = 380

Points(220).X = 163
Points(220).Y = 381

Points(221).X = 161
Points(221).Y = 382

Points(222).X = 160
Points(222).Y = 383

Points(223).X = 158
Points(223).Y = 384

Points(224).X = 155
Points(224).Y = 385

Points(225).X = 153
Points(225).Y = 386

Points(226).X = 150
Points(226).Y = 387

Points(227).X = 147
Points(227).Y = 388

Points(228).X = 143
Points(228).Y = 389

Points(229).X = 138
Points(229).Y = 390

Points(230).X = 128
Points(230).Y = 391

Points(231).X = 119
Points(231).Y = 391

Points(232).X = 109
Points(232).Y = 390

Points(233).X = 104
Points(233).Y = 389

Points(234).X = 100
Points(234).Y = 388

Points(235).X = 97
Points(235).Y = 387

Points(236).X = 94
Points(236).Y = 386

Points(237).X = 92
Points(237).Y = 385

Points(238).X = 89
Points(238).Y = 384

Points(239).X = 87
Points(239).Y = 383

Points(240).X = 86
Points(240).Y = 382

Points(241).X = 84
Points(241).Y = 381

Points(242).X = 83
Points(242).Y = 380

Points(243).X = 81
Points(243).Y = 379

Points(244).X = 80
Points(244).Y = 378

Points(245).X = 79
Points(245).Y = 377

Points(246).X = 78
Points(246).Y = 376

Points(247).X = 77
Points(247).Y = 375

Points(248).X = 76
Points(248).Y = 374

Points(249).X = 75
Points(249).Y = 372

Points(250).X = 74
Points(250).Y = 371

Points(251).X = 73
Points(251).Y = 369

Points(252).X = 72
Points(252).Y = 367

Points(253).X = 72
Points(253).Y = 235

Points(254).X = 71
Points(254).Y = 235

Points(255).X = 69
Points(255).Y = 234

Points(256).X = 66
Points(256).Y = 233

Points(257).X = 64
Points(257).Y = 232

Points(258).X = 63
Points(258).Y = 231

Points(259).X = 61
Points(259).Y = 230

Points(260).X = 60
Points(260).Y = 229

Points(261).X = 58
Points(261).Y = 228

Points(262).X = 56
Points(262).Y = 227

Points(263).X = 55
Points(263).Y = 226

Points(264).X = 53
Points(264).Y = 225

Points(265).X = 52
Points(265).Y = 224

Points(266).X = 51
Points(266).Y = 223

Points(267).X = 50
Points(267).Y = 222

Points(268).X = 49
Points(268).Y = 221

Points(269).X = 47
Points(269).Y = 220

Points(270).X = 46
Points(270).Y = 219

Points(271).X = 45
Points(271).Y = 218

Points(272).X = 44
Points(272).Y = 217

Points(273).X = 43
Points(273).Y = 216

Points(274).X = 41
Points(274).Y = 215

Points(275).X = 40
Points(275).Y = 214

Points(276).X = 39
Points(276).Y = 213

Points(277).X = 38
Points(277).Y = 212

Points(278).X = 36
Points(278).Y = 211

Points(279).X = 35
Points(279).Y = 209

Points(280).X = 34
Points(280).Y = 208

Points(281).X = 33
Points(281).Y = 207

Points(282).X = 32
Points(282).Y = 206

Points(283).X = 31
Points(283).Y = 204

Points(284).X = 30
Points(284).Y = 203

Points(285).X = 29
Points(285).Y = 202

Points(286).X = 28
Points(286).Y = 201

Points(287).X = 27
Points(287).Y = 200

Points(288).X = 26
Points(288).Y = 198

Points(289).X = 25
Points(289).Y = 197

Points(290).X = 24
Points(290).Y = 196

Points(291).X = 23
Points(291).Y = 195

Points(292).X = 22
Points(292).Y = 194

Points(293).X = 21
Points(293).Y = 192

Points(294).X = 20
Points(294).Y = 191

Points(295).X = 19
Points(295).Y = 189

Points(296).X = 18
Points(296).Y = 187

Points(297).X = 17
Points(297).Y = 186

Points(298).X = 16
Points(298).Y = 184

Points(299).X = 15
Points(299).Y = 183

Points(300).X = 14
Points(300).Y = 181

Points(301).X = 13
Points(301).Y = 178

Points(302).X = 12
Points(302).Y = 176

Points(303).X = 11
Points(303).Y = 174

Points(304).X = 10
Points(304).Y = 172

Points(305).X = 9
Points(305).Y = 170

Points(306).X = 8
Points(306).Y = 167

Points(307).X = 7
Points(307).Y = 164

Points(308).X = 6
Points(308).Y = 162

Points(309).X = 5
Points(309).Y = 159

Points(310).X = 4
Points(310).Y = 155

Points(311).X = 3
Points(311).Y = 151

Points(312).X = 2
Points(312).Y = 146

Points(313).X = 1
Points(313).Y = 140

Points(314).X = 0
Points(314).Y = 128

Points(315).X = 0
Points(315).Y = 119

hRgn = CreatePolygonRgn(Points(1), 315, 1)
SetWindowRgn hwnd, hRgn, True
End Sub

Private Sub Form_Load()
ReshapeFormToCustomShape
ModifyHands

Left = Screen.Width / 2 - Width / 2
Top = Screen.Height / 2 - Height / 2

b = LoadResData(101, "CUSTOM")

Ind = 30
Subt = True
Subt2 = True
Subt3 = True

GetSettings

Timer1.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Timer6.Enabled = True
Timer7.Enabled = True

If Mid$(Trim$(Command$), 1, 2) = "1" Then
    If Not bSysTray Then bSysTray = True
    frmMain.WindowState = 1
End If

DisableMenu = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Result As Long
Dim msg As Long

If Not DisableMenu Then
    msg = ScaleX(X, ScaleMode, vbPixels)
    Select Case msg
        Case WM_LBUTTONDBLCLK
            WindowState = 0
            Result = SetForegroundWindow(Me.hwnd)
            Show
        Case WM_RBUTTONUP
            Result = SetForegroundWindow(Me.hwnd)
            PopupMenu frmMenu.mSys, , , , frmMenu.mnuRestore
    End Select
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next

Unload frmMenu
Unload frmAtomicClock
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
SaveSettings
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Form_Resize()
If bSysTray Then
    If WindowState = 1 Then
        Hide
        With nid
            .cbSize = Len(nid)
            .hwnd = hwnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallBackMessage = WM_MOUSEMOVE
            .hIcon = Icon
            .szTip = "Clock" & vbNullChar
        End With
        Shell_NotifyIcon NIM_ADD, nid
    Else
        Timer8.Enabled = True
        DisableMenu = True
    End If
End If
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MD X, Y
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MM X, Y
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MU Button
End Sub

Private Sub Image3_Click()
On Error Resume Next

Unload frmAtomicClock
Load frmAtomicClock

frmAtomicClock.Show 1
End Sub

Private Sub Image4_Click()
On Error Resume Next

Unload frmSettings
Load frmSettings

frmSettings.Show 1
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MD X, Y
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MM X, Y
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MU Button
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MD X, Y
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MM X, Y
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MU Button
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MD X, Y
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MM X, Y
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MU Button
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MD X, Y
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MM X, Y
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MU Button
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MD X, Y
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MM X, Y
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MU Button
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MD X, Y
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MM X, Y
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MU Button
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MD X, Y
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MM X, Y
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MU Button
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MD X, Y
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MM X, Y
End Sub

Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MU Button
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MD X, Y
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MM X, Y
End Sub

Private Sub Label9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MU Button
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MD X, Y
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MM X, Y
End Sub

Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MU Button
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MD X, Y
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MM X, Y
End Sub

Private Sub Label11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MU Button
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MD X, Y
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MM X, Y
End Sub

Private Sub Label12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MU Button
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MD X, Y
End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MM X, Y
End Sub

Private Sub Label13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MU Button
End Sub

Private Sub Timer1_Timer()
Ind = Ind + IIf(Subt, -0.5, 0.5)

If Ind = 27.5 Then
    Subt = False
    Ind = Ind + 0.5
ElseIf Ind = 32.5 Then
    Subt = True
    Ind = Ind - 0.5
End If

If Ind <> 27.5 And Ind <> 32.5 Then
    Line4.X2 = 1585 * Cos(PI / 180 * (6 * Ind - 90)) + Line4.X1
    Line4.Y2 = 1585 * Sin(PI / 180 * (6 * Ind - 90)) + Line4.Y1
    
    Shape7.Left = Line4.X2 - Shape7.Width / 2
    Shape7.Top = Line4.Y2 - Shape7.Height / 2
End If
End Sub

Private Sub Timer2_Timer()
ModifyHands
End Sub

Private Sub Timer3_Timer()
If UseChime Then
    If Minute(Now) = 0 And Second(Now) = 0 Then
        If Not SoundPlayed Then
            sndPlaySound b(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
            SoundPlayed = True
        End If
    Else
        If SoundPlayed Then SoundPlayed = False
    End If
End If
End Sub

Private Sub Timer4_Timer()
If Subt2 Then
    If Ind2 <= 0 Then
        If Ind2 <> 0 Then Ind2 = 0
    Else
        Ind2 = Ind2 - 1
    End If
Else
    If Ind2 >= 10 Then
        If Ind2 <> 10 Then Ind2 = 10
    Else
        Ind2 = Ind2 + 1
    End If
End If

If Image3.Picture <> Image5(Ind2).Picture Then Image3.Picture = Image5(Ind2).Picture
End Sub

Private Sub Timer5_Timer()
If Subt3 Then
    If Ind3 <= 0 Then
        If Ind3 <> 0 Then Ind3 = 0
    Else
        Ind3 = Ind3 - 1
    End If
Else
    If Ind3 >= 10 Then
        If Ind3 <> 10 Then Ind3 = 10
    Else
        Ind3 = Ind3 + 1
    End If
End If

If Image4.Picture <> Image6(Ind3).Picture Then Image4.Picture = Image6(Ind3).Picture
End Sub

Private Sub Timer6_Timer()
GetCursorPos Pt
ScreenToClient hwnd, Pt

Select Case False
    Case XPos = Pt.X, YPos = Pt.Y
        If XPos <> Pt.X Then XPos = Pt.X
        If YPos <> Pt.Y Then YPos = Pt.Y
        GlowButtons XPos, YPos
End Select
End Sub

Private Sub Timer7_Timer()
If DateExists Then
    If Now = ACDate Then
        If Not DontSync Then
            SyncTime
            DontSync = True
        End If
    Else
        If DontSync Then DontSync = False
    End If
End If
End Sub

Private Sub Timer8_Timer()
Timer8.Enabled = False
Shell_NotifyIcon NIM_DELETE, nid
DisableMenu = False
End Sub

Private Function NewDate(sData As String) As Date
Dim X As Integer, tmpDate, tmpTime
ErrCode = 0
X = InStr(1, sData, ":")

tmpDate = Mid$(sData, X - 8, 5) & "-" & Mid$(sData, X - 11, 2)
tmpTime = Mid$(sData, X - 2, 9)

If IsDate(tmpDate) And IsDate(tmpTime) Then
    If Mid$(sData, X + 12, 1) <> "0" Then
        ErrCode = 2
    Else
        msADV = CInt(Mid$(sData, X + 14, 3) & Mid$(sData, X + 18, 1))
        NewDate = CDate(tmpDate & " " & tmpTime)
    End If
Else
    ErrCode = 1
End If
End Function

Private Sub WebSock1_DataArrival(ByVal SocketID As Long, sData As String)
On Error Resume Next

Dim MySys As SYSTEMTIME, AccurateDate As Date
AccurateDate = NewDate(sData)
If msADV <> 0 Then AccurateDate = DateAdd("s", -1, AccurateDate)
AccurateDate = DateAdd("s", Int((GetTickCount - LastTicks2) / 1000), AccurateDate)

If ErrCode = 1 Then
    ProcessMsg ERR_UNERR
Else
    If ErrCode = 2 Then
        ProcessMsg ERR_INACC
    Else
        With MySys
            .wYear = Year(AccurateDate)
            .wMonth = Month(AccurateDate)
            .wDay = Day(AccurateDate)
            .wHour = Hour(AccurateDate)
            .wMinute = Minute(AccurateDate)
            .wSecond = Second(AccurateDate)
            If msADV = 0 Then
                .wMilliseconds = 0
            Else
                .wMilliseconds = (10000 - msADV) / 10
            End If
        End With
        
        SetSystemTime MySys
        ProcessMsg "Clock successfully synchronized on " & Format(Now, "MMMM D, YYYY") & " at " & TimeValue(Now) & " (Delay: " & GetTimeOffSet & " sec)"
    End If
End If
ProcessDt IIf(DateExists, DateAdd("h", Offset, Now), "")
If bUsed Then bUsed = False
End Sub

Private Sub WebSock1_Error(ByVal SocketID As Long, ByVal number As Integer, Description As String)
ProcessMsg "Clock unsuccessfully synchronized"
If bUsed Then bUsed = False
End Sub
