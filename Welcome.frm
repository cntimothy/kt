VERSION 5.00
Begin VB.Form Welcome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   Icon            =   "Welcome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6975
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command4 
      Caption         =   "ע��"
      Height          =   615
      Left            =   1320
      TabIndex        =   5
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�˳�"
      Height          =   615
      Left            =   4080
      TabIndex        =   4
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��¼"
      Height          =   615
      Left            =   4080
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      Default         =   -1  'True
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��ѡ���û�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Top             =   1920
      Width           =   2205
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ӭʹ��KT�ִ�����ϵͳ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   5790
   End
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
liulan.Show 1
End Sub

Private Sub Command2_Click()
login.Show 1
End Sub

Private Sub Command3_Click()
Unload Welcome
End Sub

Private Sub Command4_Click()
zhuce.Show 1
End Sub

Private Sub Form_Load()
q = True
'���Ȩ��
limit = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
If q = True Then
    Dim str As Integer
    str = MsgBox("ȷ���˳�ϵͳ��", 1 + 32 + 256, "��ʾ")
    If str = 1 Then
        Cancel = 0
    Else
        Cancel = 1
    End If
Else:
    Cancel = 0
End If
End Sub
