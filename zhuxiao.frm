VERSION 5.00
Begin VB.Form zhuxiao1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ע��"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10245
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   10245
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   7800
      TabIndex        =   16
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   495
      Left            =   5400
      TabIndex        =   15
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   6600
      PasswordChar    =   "*"
      TabIndex        =   14
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   5640
      TabIndex        =   13
      Top             =   1800
      Width           =   360
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "�������Ϣȷ���������������������벢����ȷ����ť"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4800
      TabIndex        =   12
      Top             =   720
      Width           =   5250
   End
   Begin VB.Line Line1 
      X1              =   4560
      X2              =   4560
      Y1              =   240
      Y2              =   3960
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "��ϵ��ʽ"
      Height          =   180
      Left            =   480
      TabIndex        =   5
      Top             =   3360
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   600
      TabIndex        =   4
      Top             =   2760
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      Height          =   180
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�û���"
      Height          =   180
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   540
   End
End
Attribute VB_Name = "zhuxiao1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
q = False
'�������ݿ�
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
cn.Open "provider = microsoft.jet.oledb.4.0;data source=" & App.Path & "\db\kt.mdb;persist security info = false"
sql = "select * from ktuser where uid = '" & username & "'"
rs.Open sql, cn, 3, 2
'��֤�����Ƿ���д
If Text7.Text = "" Then
    MsgBox "���벻��Ϊ��", vbCritical, "��ʾ"
    Text7.SetFocus
Else
    '��֤�����Ƿ���ȷ
    If Text7.Text <> rs("upasswd") Then
        MsgBox "�������", vbCritical, "��ʾ"
        Text7.SetFocus
    Else
        Dim str As String
        str = MsgBox("ȷ��ע����", 1 + 32 + 256, "��ʾ")
        If str = 1 Then
            rs.Delete
            MsgBox "���û���ע���������µ�¼"
            '���µ�¼
            Unload zhuxiao1
            Unload guanli
            Welcome.Show 0
            login.Show 1
        End If
    End If
End If
End Sub

Private Sub Command2_Click()
Dim str As Integer
str = MsgBox("ȷ���˳���", 1 + 32 + 256, "��ʾ")
If str = 1 Then
    Unload zhuxiao1
End If
End Sub

Private Sub Form_Load()
'�������ݿ⣬��������û���Ϣ
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
cn.Open "provider = microsoft.jet.oledb.4.0;data source=" & App.Path & "\db\kt.mdb;persist security info = false"
sql = "select * from ktuser where uid = '" & username & "'"
rs.Open sql, cn, 3, 2
Text1.Text = rs("uid")
Text2.Text = rs("xingming")
Text3.Text = rs("nianling")
Text4.Text = rs("xingbie")
Text5.Text = rs("keshi")
Text6.Text = rs("lianxi")
End Sub
