VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmBegin 
   Caption         =   "�����������"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8700
   Icon            =   "FrmBegin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDlg1 
      Left            =   7800
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "������� ����..."
      FileName        =   "000.elt"
      Filter          =   "*.elt|*.elt"
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "������� ����!!!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "������� ����!!!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "������ ����������!"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3000
      TabIndex        =   3
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox TxtName 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Text            =   "���� �����"
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "�������������, ����������:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3000
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Image Image2 
      Height          =   2820
      Left            =   6000
      Picture         =   "FrmBegin.frx":0442
      Top             =   1320
      Width           =   2625
   End
   Begin VB.Image Image1 
      Height          =   2865
      Left            =   0
      Picture         =   "FrmBegin.frx":13AC
      Top             =   1200
      Width           =   2790
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  ������-���� �� ����    ""Excel � ������� �������"""
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   1080
      TabIndex        =   0
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "FrmBegin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub cat()

End Sub
Private Sub cmdChange_Click()
Dim pt As String, q As AboutQ, n As Integer, thms As String * 23
Close #1
n = Len(q)
CDlg1.ShowOpen
pt = CDlg1.FileName
Open pt For Random As #1 Len = n
Options SaveOption, LastPath, pt
Get #1, 1, thms
Label1.Caption = "������-���� �� ����" + Chr(13) + Chr(132) + Trim(Left(thms, 23)) + Chr(148)
End Sub

Private Sub cmdCreate_Click()

Dim f As Integer, n As Integer, FileNa As String, file As String, theme As String * 23, j As Integer
Dim q As AboutQ
n = Len(q)
Close #1
FileNa = InputBox("������� ��� �����", "��� �����")

file = App.path + "\Texts\" + FileNa + ".elt"

Open App.path + "\Texts\" + FileNa + ".elt" For Random As #1 Len = n
theme = InputBox("������� ��� ����. �������� - 23 �������(��������� �� ����� �������)", "��� ����")
Put #1, 1, theme
Label1.Caption = "������-���� �� ����" + Chr(13) + Chr(132) + Trim(Left(theme, 23)) + Chr(148)
For i = 2 To 11
j = i - 1
q.Quest = InputBox("������� ������. �������� �������� - 255, ���������� ��������� � 160.", "���� ������. ������ �" & j)
q.AnswerA = InputBox("������� ����� �.�������� - 50 ��������", "���� ������. ����� � � ������� �" & j)
q.AnswerB = InputBox("������� ����� �.�������� - 50 ��������", "���� ������. ����� � � ������� �" & j)
q.AnswerC = InputBox("������� ����� �.�������� - 50 ��������", "���� ������. ����� � � ������� �" & j)
q.AnswerD = InputBox("������� ����� �. ��������  - 50 ��������", "���� ������. ����� � � ������� �" & j)
q.AnswerR = UCase(InputBox("������� ����� ����������� ������ (�,�,�,�,�,�,�,�). � � � - �������!!! ", "���� ������. ���������� �����  � ������� �" & j))
Put #1, i, q
Next
'Set q = Nothing
Set i = Nothing
Options SaveOption, LastPath, file

End Sub


Private Sub cmdRead_Click()
nam = TxtName.Text
Options SaveOption, LastName, nam
FrmBegin.Hide
FrmInstr.Show
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim q As AboutQ
CDlg1.InitDir = App.path + "\Texts"
n = Len(q)
Dim Thm As String * 23, path As String
If Options(GetOption, LastPath) = vbNullString Then
Options SaveOption, LastPath, App.path + "\Texts\000.elt"
Options SaveOption, LastName, "���� �����"
End If
path = Options(GetOption, LastPath)
Open path For Random As #1 Len = n
Get #1, 1, Thm
nam = Options(GetOption, LastName)
TxtName.Text = nam
Label1.Caption = "������-���� �� ����" + Chr(13) + Chr(132) + Trim(Left(Thm, 23)) + Chr(148)
Set i = Nothing
'Set q = Nothing

End Sub
