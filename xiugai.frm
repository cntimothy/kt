VERSION 5.00
Begin VB.Form xiugai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�޸��û���Ϣ"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6135
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command3 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��������"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ɾ���û�"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ѡ���û���"
      Height          =   180
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   900
   End
End
Attribute VB_Name = "xiugai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'�������ݿ�
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
cn.Open "provider = microsoft.jet.oledb.4.0;data source=" & App.Path & "\db\kt.mdb;persist security info = false"
sql = "select * from ktuser where uid = '" & Combo1.Text & "'"
rs.Open sql, cn, 3, 2
Dim str As Integer
str = MsgBox("ȷ��ɾ����", 1 + 32 + 256, "��ʾ")
'ȷ��ɾ�����������ݿ�
If str = 1 Then
    rs.Delete
    MsgBox "���û���ɾ��"
    '�ر����ݿ�
    rs.Close
    cn.Close
    '���������û���Ϣ
    Unload xiugai
    xiugai.Show 1
    End If
End Sub

Private Sub Command2_Click()
Dim str As Integer
str = MsgBox("ȷ�����ã�", 1 + 32 + 256, "��ʾ")
If str = 1 Then
    '�������ݿ�
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    cn.Open "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\db\kt.mdb;persist security info = false"
    sql = "select upasswd from ktuser where uid = '" & Combo1.Text & "'"
    rs.Open sql, cn, 3, 2
    '��������Ϊ000000
    rs("upasswd") = "000000"
    '�������ݿ�
    rs.Update
    MsgBox "�����޸ĳɹ�"
End If
End Sub

Private Sub Command3_Click()
Dim str As Integer
str = MsgBox("ȷ���˳���", 1 + 32 + 256, "��ʾ")
If str = 1 Then
Unload xiugai
End If
End Sub

Private Sub Form_Load()
'�������ݿ�
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\kt.mdb;Persist Security Info=False"
'������ͨ����Ա��Ϣ
sql = "select uid from ktuser where quanxian = '1' "
rs.Open sql, cn, 1, 1
'ѭ�����б���������ͨ����Ա�û���
Do While rs.EOF <> True
    Combo1.AddItem rs("uid")
    rs.MoveNext
    Loop
'�ر����ݿ�
rs.Close
cn.Close
End Sub
