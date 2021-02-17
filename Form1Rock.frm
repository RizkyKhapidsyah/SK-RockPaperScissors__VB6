VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "PAPER  SCISSORS  STONE!!!!"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Form1Rock.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1Rock.frx":0442
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   8295
      Left            =   0
      ScaleHeight     =   8235
      ScaleWidth      =   11835
      TabIndex        =   17
      Top             =   480
      Width           =   11895
   End
   Begin VB.Timer TimerKeep 
      Interval        =   1000
      Left            =   4200
      Top             =   600
   End
   Begin VB.CommandButton Command2 
      Caption         =   " COMPUTER RANDOMLY CHOOSES"
      Height          =   1095
      Left            =   480
      TabIndex        =   13
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Timer TimerSC 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5160
      Top             =   600
   End
   Begin VB.CommandButton CmdInf 
      Caption         =   " "
      Height          =   855
      Left            =   600
      Picture         =   "Form1Rock.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   " "
      Height          =   855
      Left            =   0
      Picture         =   "Form1Rock.frx":0FC6
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   615
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   2160
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1680
   End
   Begin VB.OptionButton OptionS 
      Caption         =   " SCISSORS"
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.OptionButton OptionP 
      Caption         =   " PAPER"
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
   Begin VB.OptionButton OptionR 
      Caption         =   " ROCK"
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   1440
      Width           =   855
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1200
   End
   Begin VB.CommandButton CmdSciss 
      Caption         =   " "
      Height          =   1815
      Left            =   9720
      Picture         =   "Form1Rock.frx":1408
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton CmdPaper 
      Caption         =   " "
      Height          =   1815
      Left            =   7800
      Picture         =   "Form1Rock.frx":963A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton CmdRock 
      Caption         =   " "
      Height          =   1815
      Left            =   5880
      Picture         =   "Form1Rock.frx":1186C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   1800
      Top             =   480
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   1200
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   " START"
      Height          =   495
      Left            =   5520
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   9120
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   8640
      Top             =   480
   End
   Begin VB.Label lbld 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DRAW GAMES:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   2400
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblcs 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " COMPUTER SCORE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblys 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "YOUR SCORE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   1440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblcomp 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   7
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Image LHPAP 
      Height          =   2835
      Left            =   240
      Picture         =   "Form1Rock.frx":19A9E
      Top             =   4560
      Visible         =   0   'False
      Width           =   5985
   End
   Begin VB.Image RHPAP 
      Height          =   2835
      Left            =   6840
      Picture         =   "Form1Rock.frx":510D0
      Top             =   4560
      Visible         =   0   'False
      Width           =   5985
   End
   Begin VB.Image LHRO 
      Height          =   2835
      Left            =   960
      Picture         =   "Form1Rock.frx":88702
      Top             =   4680
      Visible         =   0   'False
      Width           =   4725
   End
   Begin VB.Image LHSC 
      Height          =   2880
      Left            =   -120
      Picture         =   "Form1Rock.frx":B4328
      Top             =   4680
      Visible         =   0   'False
      Width           =   6225
   End
   Begin VB.Image RHRO 
      Height          =   2835
      Left            =   7680
      Picture         =   "Form1Rock.frx":EEB6A
      Top             =   4560
      Visible         =   0   'False
      Width           =   4725
   End
   Begin VB.Image RHSC 
      Height          =   2880
      Left            =   6240
      Picture         =   "Form1Rock.frx":11A790
      Top             =   4680
      Visible         =   0   'False
      Width           =   6225
   End
   Begin VB.Label lblYou 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      TabIndex        =   6
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   " COMPUTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "PAPER SCISSORS STONE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   6015
   End
   Begin VB.Image LHUP 
      Height          =   4725
      Left            =   120
      Picture         =   "Form1Rock.frx":154FD2
      Top             =   2520
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Image RHUP 
      Height          =   4725
      Left            =   9000
      Picture         =   "Form1Rock.frx":180AFC
      Top             =   2520
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Image LHSTR 
      Height          =   2835
      Left            =   960
      Picture         =   "Form1Rock.frx":1AC626
      Top             =   4680
      Width           =   4725
   End
   Begin VB.Image RHSTR 
      Height          =   2835
      Left            =   6360
      Picture         =   "Form1Rock.frx":1D824C
      Top             =   4800
      Width           =   4725
   End
   Begin VB.Menu Fmenu 
      Caption         =   "File"
      Begin VB.Menu NewMenu 
         Caption         =   "New Game"
      End
      Begin VB.Menu SaveM 
         Caption         =   "Save Game"
      End
      Begin VB.Menu ConMenu 
         Caption         =   "Continue"
      End
      Begin VB.Menu Exitmenu 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Smenu 
      Caption         =   "Scores"
      Begin VB.Menu SeeMenu 
         Caption         =   "See"
      End
      Begin VB.Menu Clemenu 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim compscore As Integer
Dim YouScore As Integer
Dim draw As Integer
Dim FileName As String
Dim FileHandle As Integer
Dim Answer As String

Private Sub Clemenu_Click()
YouScore = 0
compscore = 0
draw = 0
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub CmdInf_Click()
lblys.Visible = True
lblcs.Visible = True
lbld.Visible = True
lblys.Caption = "your score: " + Str(YouScore)
lblcs.Caption = "computer score: " + Str(compscore)
lbld.Caption = "draw games: " + Str(draw)
End Sub

Private Sub CmdPaper_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
RHSTR.Visible = False
RHUP.Visible = False
RHSC.Visible = False
RHRO.Visible = False
RHPAP.Visible = True
CmdSciss.Enabled = False
CmdRock.Enabled = False
If OptionR.Value Then LHRO.Visible = True
If OptionP.Value Then LHPAP.Visible = True
If OptionS.Value Then LHSC.Visible = True
If LHRO.Visible = True Then LHPAP.Visible = False And LHSC.Visible = False
If LHPAP.Visible = True Then LHRO.Visible = False And LHSC.Visible = False
If LHSC.Visible = True Then LHPAP.Visible = False And LHRO.Visible = False
LHUP.Visible = False
TimerSC.Enabled = True
Call sndPlaySound(App.Path & "\Sounds\Explotion.wav", &H1)
End Sub

Private Sub CmdRock_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
RHSTR.Visible = False
RHUP.Visible = False
RHSC.Visible = False
RHRO.Visible = True
CmdSciss.Enabled = False
CmdPaper.Enabled = False
If OptionR.Value Then LHRO.Visible = True
If OptionP.Value Then LHPAP.Visible = True
If OptionS.Value Then LHSC.Visible = True
If LHRO.Visible = True Then LHPAP.Visible = False And LHSC.Visible = False
If LHPAP.Visible = True Then LHRO.Visible = False And LHSC.Visible = False
If LHSC.Visible = True Then LHPAP.Visible = False And LHRO.Visible = False
LHUP.Visible = False
TimerSC.Enabled = True
Call sndPlaySound(App.Path & "\Sounds\Explotion.wav", &H1)
End Sub

Private Sub CmdSciss_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
RHSTR.Visible = False
RHUP.Visible = False
RHSC.Visible = True
CmdRock.Enabled = False
CmdPaper.Enabled = False
If OptionR.Value Then LHRO.Visible = True
If OptionP.Value Then LHPAP.Visible = True
If OptionS.Value Then LHSC.Visible = True
If LHRO.Visible = True Then LHPAP.Visible = False And LHSC.Visible = False
If LHPAP.Visible = True Then LHRO.Visible = False And LHSC.Visible = False
If LHSC.Visible = True Then LHPAP.Visible = False And LHRO.Visible = False
LHUP.Visible = False
TimerSC.Enabled = True
Call sndPlaySound(App.Path & "\Sounds\Explotion.wav", &H1)
End Sub

Private Sub Command1_Click()
Timer5.Enabled = True
Timer1.Enabled = True
Timer3.Enabled = True
CmdRock.Enabled = True
CmdPaper.Enabled = True
CmdSciss.Enabled = True
RHSC.Visible = False
RHRO.Visible = False
RHPAP.Visible = False
LHSC.Visible = False
LHRO.Visible = False
LHPAP.Visible = False
lblcomp.Caption = ""
lblYou.Caption = ""
'lblys.Caption = "your score: " + Str(YouScore)
TimerKeep.Enabled = True
End Sub

Private Sub ConMenu_Click()
Dim FileName As String
Dim FileHandle As Integer
FileName = App.Path & "\scores.txt"
FileHandle = FreeFile()

On Error GoTo errhandler:

Open FileName For Input As #FileHandle
Input #FileHandle, YouScore, compscore, draw
Close #FileHandle
Exit Sub

errhandler:
MsgBox "Can't open a file"
End

End Sub

Private Sub Exitmenu_Click()

Dim Answer As String
Dim FileName As String
Dim FileHandle As Integer
Answer = MsgBox("Are you sure you want to stop", vbYesNo)
If Answer = vbYes Then

   MsgBox "  Final Score   " & vbCrLf & _
          "Player:   " & YouScore & vbCrLf & _
          "Computer: " & compscore
          
FileName = App.Path & "\scores.txt"
FileHandle = FreeFile()
Open FileName For Output As #FileHandle
Write #FileHandle, YouScore
Write #FileHandle, compscore
Write #FileHandle, draw
Close #FileHandle
End
End If


End Sub


Private Sub Form_Click()
lblys.Visible = False
lblcs.Visible = False
lbld.Visible = False
End Sub


Private Sub Form_Load()
Picture1.Visible = False
If Timer1.Enabled = False And Timer2.Enabled = False Then CmdRock.Enabled = False
If Timer1.Enabled = False And Timer2.Enabled = False Then CmdSciss.Enabled = False
If Timer1.Enabled = False And Timer2.Enabled = False Then CmdPaper.Enabled = False
YouScore = 0
End Sub

Private Sub SeeMenu_Click()
lblys.Visible = True
lblcs.Visible = True
lbld.Visible = True
lblys.Caption = "your score: " + Str(YouScore)
lblcs.Caption = "computer score: " + Str(compscore)
lbld.Caption = "draw games: " + Str(draw)
End Sub

Private Sub Timer1_Timer()
If Timer1.Enabled = True Then RHSTR.Visible = True
If Timer1.Interval Then Timer2.Enabled = True
                        RHUP.Visible = True
                        RHSTR.Visible = False
                        Timer1.Enabled = False
'If Timer2.Enabled = True Then RHUP.Visible = True
'If Timer2.Enabled = True Then Timer1.Enabled = False
'If Timer1.Enabled = False Then RHSTR.Visible = False
End Sub

Private Sub Timer2_Timer()
If Timer2.Interval Then Timer1.Enabled = True
If Timer1.Enabled = True Then Timer2.Enabled = False
If Timer2.Enabled = False Then RHUP.Visible = False
If RHUP.Visible = False Then RHSTR.Visible = True
End Sub

Private Sub Timer3_Timer()
If Timer3.Enabled = True Then LHSTR.Visible = True
If Timer3.Interval Then Timer4.Enabled = True
If Timer4.Enabled = True Then LHUP.Visible = True
If Timer4.Enabled = True Then Timer3.Enabled = False
If Timer3.Enabled = False Then LHSTR.Visible = False
End Sub

Private Sub Timer4_Timer()
If Timer4.Interval Then Timer3.Enabled = True
If Timer3.Enabled = True Then Timer4.Enabled = False
If Timer4.Enabled = False Then LHUP.Visible = False
If LHUP.Visible = False Then LHSTR.Visible = True
End Sub

Private Sub Timer5_Timer()
OptionR.Value = True
OptionP.Visible = False
OptionS.Visible = False
OptionR.Visible = False 'True
If Timer5.Interval Then Timer6.Enabled = True
If Timer6.Enabled = True Then Timer5.Enabled = False
End Sub

Private Sub Timer6_Timer()
OptionP.Visible = True
OptionP.Value = True
OptionR.Visible = False
OptionS.Visible = False
If Timer6.Interval Then Timer7.Enabled = True
If Timer7.Enabled = True Then Timer6.Enabled = False
End Sub

Private Sub Timer7_Timer()
OptionS.Visible = True
OptionS.Value = True
OptionR.Visible = False
OptionP.Visible = False
If Timer7.Interval Then Timer5.Enabled = True
If Timer5.Enabled = True Then Timer7.Enabled = False
End Sub
Private Sub TimerKeep_Timer()
If lblYou.Caption = "Winner" Then TimerKeep.Enabled = False
If lblcomp.Caption = "Winner" Then TimerKeep.Enabled = False
If lblcomp.Caption = "Draw" Then TimerKeep.Enabled = False

If lblYou.Caption = "Winner" Then YouScore = YouScore + 1
If lblcomp.Caption = "Winner" Then compscore = compscore + 1
If lblYou.Caption = "Draw" Then draw = draw + 1
End Sub

Private Sub TimerSC_Timer()
' this timer tells the difference between winner, looser, and draw
If LHRO.Visible = True And RHRO.Visible = True Then lblYou.Caption = "Draw"
If LHRO.Visible = True And RHRO.Visible = True Then lblcomp.Caption = "Draw"
If LHPAP.Visible = True And RHPAP.Visible = True Then lblYou.Caption = "Draw"
If LHPAP.Visible = True And RHPAP.Visible = True Then lblcomp.Caption = "Draw"
If LHSC.Visible = True And RHSC.Visible = True Then lblYou.Caption = "Draw"
If LHSC.Visible = True And RHSC.Visible = True Then lblcomp.Caption = "Draw"

If LHRO.Visible = True And RHPAP.Visible = True Then lblYou.Caption = "Winner"
If LHRO.Visible = True And RHSC.Visible = True Then lblcomp.Caption = "Winner"
If LHPAP.Visible = True And RHRO.Visible = True Then lblcomp.Caption = "Winner"
If LHPAP.Visible = True And RHSC.Visible = True Then lblYou.Caption = "Winner"
If LHSC.Visible = True And RHRO.Visible = True Then lblYou.Caption = "Winner"
If LHSC.Visible = True And RHPAP.Visible = True Then lblcomp.Caption = "Winner"
End Sub
