VERSION 5.00
Begin VB.Form excel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   6420
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      Width           =   3735
   End
   Begin VB.DirListBox Dir1 
      Height          =   3030
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   3960
      TabIndex        =   0
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "���ѣ�������������̸�Ŀ¼����ȷ��Ŀ¼����ͬ���ļ���"
      Height          =   180
      Left            =   960
      TabIndex        =   6
      Top             =   360
      Width           =   4680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�����Ŀ¼"
      Height          =   180
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   900
   End
End
Attribute VB_Name = "excel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ktpath As String

Private Sub Command1_Click()
If ktpath = "" Then
    MsgBox "·������Ϊ��", vbCritical, "��ʾ"
Else
    Dim conn As New ADODB.Connection
    Dim rsqrecord As ADODB.Recordset
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\kt.mdb;Persist Security Info=False"
    response = MsgBox("ȷ��Ҫ����������", 1 + 32 + 256, "��ʾ")
    If response = vbOK Then
        Dim str As String
        '��ȡ��������
        str = CStr(Date)
'ѡ���������
        Select Case e
            Case 1
                conn.Execute "select huohao as ����, kucun as ���, riqi as ������������ into [excel 8.0;database=" & ktpath & "\" & str & "���.xls].sheet1 from ktkucun order by huohao"
                MsgBox "�����ɹ�,�ļ���Ϊ��������+���"
            Case 0
                conn.Execute "select riqi as ����, churu as �������, huohao as ����, shuliang as ����,jingban as ������,gongying as ��Ӧ�� into [excel 8.0;database=" & ktpath & "\" & str & "��־.xls].sheet1 from ktchuru order by riqi"
                MsgBox "�����ɹ�,�ļ���Ϊ��������+��־"
            Case 2
                conn.Execute "select huohao as ����, kucun as ���, riqi as ������������  into [excel 8.0;database=" & ktpath & "\" & str & "�澯.xls].sheet1 from ktkucun where kucun < " & g & " order by huohao"
                MsgBox "�����ɹ�,�ļ���Ϊ��������+�澯"
            Case 3
                conn.Execute "select uid as �û���, xingming as ����,nianling as ����, xingbie as �Ա�, keshi as ����,lianxi as ��ϵ��ʽ  into [excel 8.0;database=" & ktpath & "\" & str & "�û�.xls].sheet1 from ktuser"
                MsgBox "�����ɹ�,�ļ���Ϊ��������+�û�"
            End Select
            Unload excel
    End If
End If
End Sub
    
Private Sub Command2_Click()
Dim str As Integer
str = MsgBox("ȷ���˳���", 1 + 32 + 256, "��ʾ")
If str = 1 Then
Unload excel
End If
End Sub
'��ȡ���·��
Private Sub Dir1_Change()
ktpath = Dir1.Path
Text1.Text = ktpath
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
ktpath = Drive1.Drive
Text1.Text = ktpath
End Sub

