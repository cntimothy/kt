VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form chuku1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   8220
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command5 
      Caption         =   "��ѯ"
      Height          =   495
      Left            =   6480
      TabIndex        =   18
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "���"
      Height          =   495
      Left            =   2040
      TabIndex        =   17
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ɾ��"
      Height          =   495
      Left            =   480
      TabIndex        =   16
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1560
      TabIndex        =   14
      Top             =   240
      Width           =   3975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1560
      Top             =   1440
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   6720
      TabIndex        =   13
      Top             =   5520
      Width           =   1215
   End
   Begin VB.ListBox List3 
      Height          =   3120
      Left            =   4920
      TabIndex        =   12
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   495
      Left            =   5160
      TabIndex        =   11
      Top             =   5520
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   3120
      Left            =   3480
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   3120
      Left            =   600
      TabIndex        =   9
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��һ��"
      Default         =   -1  'True
      Height          =   495
      Left            =   6480
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   4920
      TabIndex        =   8
      Top             =   1680
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   3480
      TabIndex        =   7
      Top             =   1680
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   600
      TabIndex        =   6
      Top             =   1680
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ʣ������"
      Height          =   180
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "chuku1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'�����������Ƿ�С�ڿ��
If IsNumeric(Trim(Text1.Text)) = False Then
    MsgBox "��������ӦΪ����", vbCritical, "��ʾ"
    Else
    If CInt(Text1.Text) <= CInt(Text2.Text) And CInt(Text1.Text) > 0 Then
        '������Ϣ�����б���
        List1.AddItem (Combo1.Text)
        List2.AddItem (Text1.Text)
        List3.AddItem (Now())
        '���ź��������ÿ�
        Text1.Text = ""
        Text2.Text = ""
        Command1.Enabled = False
    Else:
        MsgBox "�����������ԣ�", vbCritical, "����"
        Text1.Text = ""
    End If
End If
End Sub

Private Sub Command2_Click()
Dim i As Integer
i = MsgBox("ȷ������?", 1 + 32 + 256, "����")
If i = 1 Then
    '�������ݿ�
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\kt.mdb;Persist Security Info=False"
    'ѭ��������д�����ݿ�
    For j = 0 To List1.ListCount - 1
        List1.ListIndex = j
        List2.ListIndex = j
        List3.ListIndex = j
        '������־
        sql = "select * from ktchuru"
        rs.Open sql, cn, 3, 2
        rs.AddNew
        rs("riqi") = List3.Text
        rs("churu") = "����"
        rs("huohao") = List1.Text
        rs("shuliang") = CInt(List2.Text)
        rs("jingban") = username
        rs.Update
        rs.Close
        '������
        sql = "select * from ktkucun where huohao = '" & List1.Text & "'"
        rs.Open sql, cn, 3, 2
        rs("kucun") = rs("kucun") - CInt(List2.Text)
        rs("riqi") = List3.Text
        rs.Update
        rs.Close
    Next
    MsgBox "������Ŀ�ѳ���", , "��ʾ"
    Combo1.SetFocus
End If
End Sub

Private Sub Command4_Click()
Dim i As Integer
i = MsgBox("ȷ���˳����ģʽ��", 1 + 32 + 256, "��ʾ")
If i = 1 Then
Unload chuku1
End If
End Sub

Private Sub Command5_Click()
'ʹ��һ�ť����
Command1.Enabled = True
'��֤��ѡ�����
If Combo1.Text <> "" Then
    '�������ݿ�
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\kt.mdb;Persist Security Info=False"
    sql = "select huohao, kucun from ktkucun where huohao = '" & Combo1.Text & "'"
    rs.Open sql, cn, 1, 1
    Text2.Text = rs("kucun")
    rs.Close
    cn.Close
Else
    MsgBox "���Ų���Ϊ��", vbCritical, "��ʾ"
End If
End Sub

Private Sub Command6_Click()
'ɾ���б���ѡ�е���Ŀ
Dim i As Integer
For i = List1.ListCount - 1 To 0 Step -1
'��֤��Щ��Ŀ��ѡ��
If List1.Selected(i) = True Or List2.Selected(i) = True Or List3.Selected(i) = True Then
    List1.RemoveItem (i)
    List2.RemoveItem (i)
    List3.RemoveItem (i)
End If
Next
End Sub

Private Sub Command7_Click()
'����б�
Dim i, j As Integer
i = MsgBox("ȷ�������", 1 + 32 + 256, "��ʾ")
If i = 1 Then
    For j = List1.ListCount - 1 To 0 Step -1
        List1.RemoveItem (j)
        List2.RemoveItem (j)
        List3.RemoveItem (j)
    Next
End If
End Sub

Private Sub Form_Load()
'���������Ϣ����������������
Command1.Enabled = False
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\kt.mdb;Persist Security Info=False"
sql = "select huohao from ktkucun order by huohao asc"
rs.Open sql, cn, 1, 1
Do While rs.EOF <> True
    Combo1.AddItem rs("huohao")
    rs.MoveNext
    Loop
rs.Close
cn.Close
End Sub
