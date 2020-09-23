VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   Caption         =   "MP3 Player"
   ClientHeight    =   360
   ClientLeft      =   5040
   ClientTop       =   2775
   ClientWidth     =   1560
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   24
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   104
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimeKeeper 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2160
      Top             =   0
   End
   Begin VB.CommandButton Command3o 
      Height          =   135
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   105
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2o 
      Caption         =   "Stop"
      Height          =   195
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton Command1o 
      Height          =   195
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Image Image2 
      Height          =   120
      Left            =   0
      Picture         =   "frmMain.frx":08CA
      Top             =   0
      Width           =   90
   End
   Begin VB.Image Image3 
      Height          =   120
      Left            =   90
      Picture         =   "frmMain.frx":09AC
      Top             =   0
      Width           =   90
   End
   Begin VB.Image Image1 
      Height          =   120
      Left            =   180
      Picture         =   "frmMain.frx":0A8E
      Top             =   0
      Width           =   90
   End
   Begin VB.Image Command2 
      Height          =   135
      Left            =   180
      Picture         =   "frmMain.frx":0B70
      Top             =   225
      Width           =   90
   End
   Begin VB.Image Command3 
      Height          =   135
      Left            =   90
      Picture         =   "frmMain.frx":0C66
      Top             =   225
      Width           =   90
   End
   Begin VB.Image Command1 
      Height          =   135
      Left            =   0
      Picture         =   "frmMain.frx":0D5C
      Top             =   225
      Width           =   90
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   150
      Left            =   0
      TabIndex        =   3
      Top             =   60
      Width           =   270
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
CommonDialog1.ShowOpen
mStop
Filename = GetShortPath(CommonDialog1.Filename)
'FileCopy CommonDialog1.Filename, "C:\Windows\Temp\1.mp3"
'Filename = CommonDialog1.Filename
'Filename = "C:\program files\napster\smusic\1.mp3"
'Filename = "C:\Windows\Temp\1.mp3"
mPlay
Length = mGetLength / 1000
TimeKeeper.Enabled = True
End Sub

Private Sub Command2_Click()
mStop
End Sub

Private Sub Command3_Click()
'If Command3.Caption = "Pause" Then
'Command3.Caption = "Unpause"
'Else
'Command3.Caption = "Pause"
'End If
'mPause
End Sub

Private Sub Form_Load()
Form1.Width = 18 * Screen.TwipsPerPixelX
Form1.Height = 24 * Screen.TwipsPerPixelY
mPause
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Form1
End Sub

Private Sub Form_Terminate()
Command2_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
Command2_Click
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'FormMove Form1
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Form1
End Sub

Private Sub Image3_Click()
'Me.WindowState = 1
Me.Visible = False
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'FormMove Form1
End Sub

Private Sub TimeKeeper_Timer()
Dim M As Double
Dim S As Double
Dim K
Position = mGetPosition
Position = Position / 1000
K = InStr(1, ".", Str(Position))
If K > 0 Then
Position = Mid(Str(Position), 1, K - 1)
End If
M = CInt(Position / 60)
Do
If M * 60 > Position Then
M = M - 1
End If
Loop Until M * 60 <= Position
S = CInt((Position - (M * 60)))
Do
If S + M * 60 > Position Then
S = S - 1
End If
Loop Until S + M * 60 <= Position
If S < 0 Then S = 60 - S
Label1.Caption = Format(M, "00") & ":" & Format(S, "00")
If Position = Length Then
mStop
End If
End Sub
