VERSION 5.00
Begin VB.Form chushihua1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ʼ��"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   7230
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�����¿������볬������Ա���룬������ȷ��"
      Height          =   180
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   3600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���棡��ʼ�����������ݶ���ɾ�������ɻָ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   5040
   End
End
Attribute VB_Name = "chushihua1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer
i = MsgBox("ȷ����ʼ������ʼ�������ݿ⽫�޷��ָ�", 1 + 32 + 256, "����")
If i = 1 Then
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\kt.mdb;Persist Security Info=False"
    sql = "select * from ktuser where uid = 'root'"
    rs.Open sql, cn, 1, 1
    If Trim(Text1.Text) <> rs("upasswd") Then
        MsgBox "�������", vbCritical, "��ʾ"
        Text1.Text = ""
    Else
        rs.Close
        cn.Close
        cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\kt.mdb;Persist Security Info=False"
        sql = "select * from ktkucun "
        rs.Open , cn, 3, 2
        Do While rs.EOF <> True
            rs.Delete
            rs.Update
            rs.MoveNext
        Loop
        rs.Close
        cn.Close
        cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\kt.mdb;Persist Security Info=False"
        sql = "select * from ktchuru "
        rs.Open , cn, 3, 2
        Do While rs.EOF <> True
            rs.Delete
            rs.Update
            rs.MoveNext
        Loop
        rs.Close
        cn.Close
        cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\kt.mdb;Persist Security Info=False"
        sql = "select gaojing from ktpeizhi"
        rs.Open sql, cn, 3, 2
        rs("gaojing") = 10
        rs.Update
        rs.Close
        MsgBox "��ʼ���ɹ�"
    End If
End If
End Sub

Private Sub Command2_Click()
Unload chushihua1
End Sub
