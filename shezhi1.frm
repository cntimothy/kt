VERSION 5.00
Begin VB.Form shezhi1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�澯����"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5355
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command3 
      Caption         =   "�ָ�Ĭ��ֵ"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "������澯ֵ"
      Height          =   180
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   1080
   End
End
Attribute VB_Name = "shezhi1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer
i = MsgBox("ȷ���޸ģ�", 1 + 32 + 256, "��ʾ")
If i = 1 Then
    '��֤������Ƿ�Ϊ����
    If (IsNumeric(Text1.Text) = False) Or (Text1.Text = "") Then
        MsgBox "����������", vbCritical, "��ʾ"
        Text1.Text = ""
    Else
        '�������ݿ⣬�޸����ñ��еĸ澯ֵ
        Set cn = New ADODB.Connection
        Set rs = New ADODB.Recordset
        cn.Open "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\db\kt.mdb;persist security info = false"
        sql = "select gaojing from ktpeizhi"
        rs.Open sql, cn, 3, 2
        rs("gaojing") = Text1.Text
        rs.Update
        '��ʵ����gֵ
        g = Text1.Text
        MsgBox "�޸ĳɹ�!���ȷ������"
        Unload shezhi1
    End If
End If
End Sub

Private Sub Command2_Click()
Unload shezhi1
End Sub

Private Sub Command3_Click()
Dim i As Integer
i = MsgBox("ȷ���ָ���ʼֵ��", 1 + 32 + 256, "��ʾ")
If i = 1 Then
    '�������ݿ⣬�������ñ��е�gֵΪ10
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    cn.Open "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\db\kt.mdb;persist security info = false"
    sql = "select gaojing from ktpeizhi"
    rs.Open sql, cn, 3, 2
    rs("gaojing") = 10
    rs.Update
    MsgBox "�޸ĳɹ�"
End If
End Sub

Private Sub Form_Load()
'���ı�������ʾ��ǰ�澯ֵ
Text1.Text = g
End Sub
