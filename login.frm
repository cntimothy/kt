VERSION 5.00
Begin VB.Form login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��¼"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6030
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6030
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�û���"
      Height          =   180
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   540
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
q = False
'�������ݿ�
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\kt.mdb;Persist Security Info=False"
sql = "select * from ktuser where uid = '" & Text1.Text & "'"
rs.Open sql, cn, 1, 1
    '�û���δ�ҵ�����ʾ����
    If rs.EOF = True Then
    MsgBox "�û��������ڣ���ȷ��", vbCritical, "��ʾ"
    '�û����������ÿ�
    Text2.Text = ""
    Text1.Text = ""
    '���ý��㵽�û����ı���
    Text1.SetFocus
Else
    '�û����ҵ�����֤����
    If rs("upasswd") = Trim(Text2.Text) Then
        '���û���Ȩ��ֵ����ȫ�ֱ���limit���û�������ȫ�ֱ���username
        limit = rs("quanxian")
        username = rs("uid")
        '�������ϵͳ
        Unload Welcome
        Unload login
        guanli.Show
    Else:
        'δͨ��������֤�������ı����ÿ�
        MsgBox "�������", vbCritical, "��ʾ"
        Text2.Text = ""
    End If
End If
End Sub

Private Sub Command2_Click()
Dim str As Integer
str = MsgBox("ȷ���˳���¼��", 1 + 32 + 256, "��ʾ")
If str = 1 Then
Unload login
q = True
End If
End Sub

Private Sub Label1_Click()

End Sub

