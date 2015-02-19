VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Эрудит-лото"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6825
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmTest.frx":0442
   ScaleHeight     =   584
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFix 
      BackColor       =   &H0000FFFF&
      Caption         =   "Зафиксировать ответ"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton cmdMenu 
      BackColor       =   &H0000FF00&
      Caption         =   "Вернуться в меню"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7920
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CDlg2 
      Left            =   120
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Выберите файл..."
   End
   Begin MSComDlg.CommonDialog CDlg1 
      Left            =   6240
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open media..."
      Filter          =   "Erudit-loto playlist(*.epl)|*.epl"
   End
   Begin MCI.MMControl MMControl1 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      _Version        =   393216
      PrevEnabled     =   -1  'True
      NextEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      PauseEnabled    =   -1  'True
      StopEnabled     =   -1  'True
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PlayVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.PictureBox P5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   6555
      TabIndex        =   1
      Top             =   6960
      Width           =   6615
   End
   Begin VB.PictureBox P1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2370
      Left            =   810
      Picture         =   "frmTest.frx":22DD6
      ScaleHeight     =   158
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   350
      TabIndex        =   0
      Top             =   120
      Width           =   5250
   End
   Begin VB.CommandButton cmdQuest 
      Caption         =   "Задать вопрос"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSForms.CommandButton cmdItog 
      Height          =   495
      Left            =   5280
      TabIndex        =   17
      Top             =   6360
      Width           =   1455
      ForeColor       =   12582912
      VariousPropertyBits=   19
      Caption         =   "Подвести итоги"
      Size            =   "2566;873"
      FontName        =   "Calibri"
      FontHeight      =   195
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdBegin 
      Height          =   495
      Left            =   5280
      TabIndex        =   16
      Top             =   4560
      Width           =   1455
      ForeColor       =   12582912
      VariousPropertyBits=   19
      Caption         =   "Начать тест!"
      Size            =   "2566;873"
      FontName        =   "Calibri"
      FontHeight      =   195
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.OptionButton Op4 
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   6360
      Width           =   5055
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   12582912
      DisplayStyle    =   5
      Size            =   "8916;873"
      Value           =   "0"
      FontName        =   "Comic Sans MS"
      FontHeight      =   240
      FontCharSet     =   204
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton Op3 
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   5760
      Width           =   5055
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   12582912
      DisplayStyle    =   5
      Size            =   "8916;873"
      Value           =   "0"
      FontName        =   "Comic Sans MS"
      FontHeight      =   240
      FontCharSet     =   204
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton Op2 
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   5160
      Width           =   5055
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   12582912
      DisplayStyle    =   5
      Size            =   "8916;873"
      Value           =   "0"
      FontName        =   "Comic Sans MS"
      FontHeight      =   240
      FontCharSet     =   204
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton Op1 
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   5055
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   12582912
      DisplayStyle    =   5
      Size            =   "8916;873"
      Value           =   "0"
      FontName        =   "Comic Sans MS"
      FontHeight      =   240
      FontCharSet     =   204
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton Command5 
      Height          =   735
      Left            =   3000
      TabIndex        =   10
      Top             =   7920
      Width           =   735
      ForeColor       =   65535
      VariousPropertyBits=   19
      Caption         =   "Next>"
      Size            =   "1296;1296"
      FontName        =   "Calibri"
      FontHeight      =   240
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox Text1 
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   6615
      VariousPropertyBits=   -1400879081
      ForeColor       =   12582912
      Size            =   "11668;2990"
      FontName        =   "Comic Sans MS"
      FontHeight      =   315
      FontCharSet     =   204
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton Command4 
      Height          =   735
      Left            =   2280
      TabIndex        =   8
      Top             =   7920
      Width           =   735
      ForeColor       =   65535
      VariousPropertyBits=   19
      Caption         =   "Stop"
      Size            =   "1296;1296"
      FontName        =   "Calibri"
      FontHeight      =   240
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Command3 
      Height          =   735
      Left            =   1560
      TabIndex        =   7
      Top             =   7920
      Width           =   735
      ForeColor       =   65535
      VariousPropertyBits=   19
      Caption         =   "Pause"
      Size            =   "1296;1296"
      FontName        =   "Calibri"
      FontHeight      =   240
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Command2 
      Height          =   735
      Left            =   840
      TabIndex        =   6
      Top             =   7920
      Width           =   735
      ForeColor       =   65535
      VariousPropertyBits=   17
      Caption         =   "Play"
      Size            =   "1296;1296"
      FontName        =   "Calibri"
      FontEffects     =   1073750016
      FontHeight      =   240
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Command1 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   7920
      Width           =   735
      ForeColor       =   65535
      VariousPropertyBits=   19
      Caption         =   "<Prev."
      Size            =   "1296;1296"
      FontName        =   "Calibri"
      FontHeight      =   240
      FontCharSet     =   204
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   2295
      Left            =   75
      TabIndex        =   15
      Top             =   4560
      Width           =   5055
      VariousPropertyBits=   746604565
      ForeColor       =   65535
      Size            =   "8916;4048"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   204
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer, a(1 To 10) As String, t(1 To 10) As String, k As Integer
Dim PlayPos As Byte, PlayKey As String


Private Sub cmdBegin_Click()
k = 0
P1.Cls

P1.DrawWidth = 1
P5.Cls


i = 1
cmdFix.Enabled = True
cmdBegin.Enabled = False
cmdQuest_Click
cmdFix.Enabled = True
cmdItog.Enabled = True
cmdMenu.Visible = False
End Sub

Private Sub cmdFix_Click()
If Op1.Value = True Then a(i - 1) = "А"
If Op2.Value = True Then a(i - 1) = "Б"
If Op3.Value = True Then a(i - 1) = "В"
If Op4.Value = True Then a(i - 1) = "Г"

Call zakraska(i)
cmdQuest.Enabled = True
cmdQuest_Click
End Sub

Private Sub cmdItog_Click()
Call prav
Dim ocenka As String, slovo As String
Select Case k
Case Is < 3
    If k = 1 Then slovo = " правильный ответ! " Else If k = 0 Then slovo = " правильных ответов :-(" Else slovo = " правильных ответa! "
    ocenka = "отстой"
Case Is = 3, Is = 4
    slovo = " правильных ответа! "
    ocenka = "ДУБровский"
Case Is = 5, Is = 6
    slovo = " правильных ответов! "
    ocenka = "лопух"
Case Is = 7, Is = 8
    slovo = " правильных ответов! "
    ocenka = "ботаник"
Case Is >= 9
    slovo = " правильных ответов! "
    ocenka = "ГЕНИЙ"
End Select
P5.Print "Поздравляю, "; nam; " , вы - "; ocenka; " !!!" + Chr(13); "У вас "; k; slovo
cmdItog.Enabled = False
cmdFix.Enabled = False
cmdBegin.Enabled = True
cmdMenu.Visible = True
End Sub

Private Sub cmdMenu_Click()
For i = 1 To 10
a(i) = ""
t(i) = ""
Next
Op1.Caption = ""
Op1.Enabled = False
Op2.Caption = ""
Op2.Enabled = False
Op3.Caption = ""
Op3.Enabled = False
Op4.Caption = ""
Op4.Enabled = False
Text1.Text = ""
P1.Cls
P5.Cls
Load FrmBegin
FrmBegin.Show (vbModal)
End Sub

Private Sub cmdQuest_Click()
i = i + 1
Call vopros(i)
Op1.Enabled = True
Op2.Enabled = True
Op3.Enabled = True
Op4.Enabled = True
cmdQuest.Enabled = False
'cmdFix.Enabled = True
End Sub

Sub vopros(i As Integer)
Dim q As AboutQ, j As Integer
j = i - 1
If j <> 11 Then
Get #1, i, q
Text1.Text = Trim(q.Quest)
Op1.Caption = Trim(q.AnswerA)
Op2.Caption = Trim(q.AnswerB)
Op3.Caption = Trim(q.AnswerC)
Op4.Caption = Trim(q.AnswerD)
t(j) = q.AnswerR
Else
cmdFix.Enabled = True
cmdItog_Click
End If
End Sub

Sub zakraska(i As Integer)
Select Case a(i - 1)
Case Is = "А"
P1.Line ((i - 1) * 32, 32)-((i - 1) * 32 + 30, 61), vbRed, BF
Case Is = "Б"
P1.Line ((i - 1) * 32, 64)-((i - 1) * 32 + 30, 93), vbRed, BF
Case Is = "В"
P1.Line ((i - 1) * 32, 96)-((i - 1) * 32 + 30, 125), vbRed, BF
Case Is = "Г"
P1.Line ((i - 1) * 32, 128)-((i - 1) * 32 + 30, 157), vbRed, BF
End Select
End Sub
Sub prav()
P1.DrawWidth = 5
For i = 1 To 10
If a(i) = t(i) And a(i) <> "" Then k = k + 1
Select Case t(i)
Case Is = "А"
P1.Line (i * 32, 32 + 15)-(i * 32 + 15, 62), vbGreen
P1.Line (i * 32 + 15, 32 + 30)-(i * 32 + 30, 32), vbGreen
Case Is = "Б"
P1.Line (i * 32, 64 + 15)-(i * 32 + 15, 94), vbGreen
P1.Line (i * 32 + 15, 64 + 30)-(i * 32 + 30, 64), vbGreen
Case Is = "В"
P1.Line (i * 32, 96 + 15)-(i * 32 + 15, 126), vbGreen
P1.Line (i * 32 + 15, 126)-(i * 32 + 30, 96), vbGreen
Case Is = "Г"
P1.Line (i * 32, 128 + 15)-(i * 32 + 15, 158), vbGreen
P1.Line (i * 32 + 15, 128 + 30)-(i * 32 + 30, 128), vbGreen
End Select
Next
End Sub

Private Sub Command1_Click()
PlayKey = "Stop"
MMControl1.Command = "close"
If MMControl1.FileName = App.path + "\Sounds\1.wav" Then
MMControl1.FileName = App.path + "\Sounds\10.wav"
PlayPos = 10
Else
PlayPos = PlayPos - 1
MMControl1.FileName = App.path + "\Sounds\" & PlayPos & ".wav"
End If
MMControl1.Command = "open"
MMControl1.Command = "play"
Command3.Enabled = True
Command2.Enabled = False
Command4.Enabled = True

End Sub

Private Sub Command2_Click()
MMControl1.Command = "play"
Command3.Enabled = True
Command2.Enabled = False
Command4.Enabled = True
End Sub

Private Sub Command3_Click()
PlayKey = "Stop"
MMControl1.Command = "pause"
Command3.Enabled = False
Command2.Enabled = True
Command4.Enabled = False
End Sub

Private Sub Command4_Click()
PlayKey = "Stop"
MMControl1.Command = "close"
'MMControl1.FileName = App.path + "\Sounds\" & PlayPos & ".wav"
MMControl1.Command = "open"
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Command5_Click()
PlayKey = "Stop"
MMControl1.Command = "close"
If MMControl1.FileName = App.path + "\Sounds\10.wav" Then
MMControl1.FileName = App.path + "\Sounds\1.wav"
PlayPos = 1
Else
PlayPos = PlayPos + 1
MMControl1.FileName = App.path + "\Sounds\" & PlayPos & ".wav"
End If
MMControl1.Command = "open"
MMControl1.Command = "play"
Command3.Enabled = True
Command2.Enabled = False
Command4.Enabled = True

End Sub

Private Sub Form_Load()
PlayPos = 1
MMControl1.Notify = False
MMControl1.Wait = True
MMControl1.Shareable = False
MMControl1.DeviceType = "WaveAudio"
MMControl1.FileName = App.path + "\Sounds\1.wav"
MMControl1.Command = "open"
MMControl1.Command = "play"

i = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)
If PlayKey <> "Stop" Then Command5_Click
PlayKey = ""
End Sub

Private Sub MMControl1_NextClick(Cancel As Integer)
'MMControl1.FileName = App.path + "\Sounds\2.wav"
End Sub

Private Sub MMControl1_PlayClick(Cancel As Integer)
MMControl1.Command = "play"
End Sub

Private Sub MMControl1_PlayCompleted(Errorcode As Long)

'MMControl1.Command = "close"
'If MMControl1.FileName = App.path + "\Sounds\10.wav" Then
'MMControl1.FileName = App.path + "\Sounds\1.wav"
'PlayPos = 1
'Else
'PlayPos = PlayPos + 1
'MMControl1.FileName = App.path + "\Sounds\" & PlayPos & ".wav"
'End If

'MMControl1.Command = "open"
'mMControl1.Command = "play"

End Sub

Private Sub MMControl1_StopClick(Cancel As Integer)
MMControl1.Command = "Stop"
End Sub

