VERSION 5.00
Begin VB.Form gaimi1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�������"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   4680
   StartUpPosition =   1  '����������
   Begin VB.TextBox Text4 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      TabIndex        =   6
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "��6~16λ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   480
      TabIndex        =   10
      Top             =   2400
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "����ȷ��"
      Height          =   180
      Left            =   480
      TabIndex        =   3
      Top             =   2760
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ԭ����"
      Height          =   180
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�û���"
      Height          =   180
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   540
   End
End
Attribute VB_Name = "gaimi1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim str As Integer
str = MsgBox("ȷ���޸ģ�", 1 + 32 + 256, "��ʾ")
If str = 1 Then
    '��֤ԭ�������������ı����Ƿ�Ϊ��
    If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
        MsgBox "��Ŀ����Ϊ��", vbCritical, "��ʾ"
    Else
        '��֤�������������Ƿ���ͬ
        If Text3.Text <> Text4.Text Then
            MsgBox "�����������벻ͬ", vbCritical, "��ʾ"
            Text3.Text = ""
            Text4.Text = ""
        Else
            '��֤���볤���Ƿ�Ϊ6~16
            If Len(Text3.Text) < 6 Or Len(Text3.Text) > 16 Then
                MsgBox "���볤�ȴ���", vbCritical, "��ʾ"
                Text3.Text = ""
                Text4.Text = ""
            Else
                '�������ݿ�
                Set cn = New ADODB.Connection
                Set rs = New ADODB.Recordset
                cn.Open "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\db\kt.mdb;persist security info = false"
                sql = "select * from ktuser where uid = '" & Text1.Text & "'"
                rs.Open sql, cn, 3, 2
                '��֤ԭ�����Ƿ���ȷ
                If rs("upasswd") <> Trim(Text2.Text) Then
                    MsgBox "ԭ�����������������", vbCritical, "��ʾ"
                    'ԭ������󣬽�ԭ�����ı�����������ı����ÿ�
                    Text2.Text = ""
                    Text3.Text = ""
                    Text4.Text = ""
                Else
                    'ԭ����ͨ����֤���������ݿ⣬�޸�����
                    rs("upasswd") = Trim(Text3.Text)
                    rs.Update
                    MsgBox "�����޸ĳɹ�,�����µ�¼"
                    '���µ�¼
                    Unload gaimi1
                    q = False
                    Unload guanli
                    Welcome.Show 0
                    login.Show 1
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub Command2_Click()
Dim str As Integer
str = MsgBox("ȷ���˳���", 1 + 32 + 256, "��ʾ")
If str = 1 Then
    Unload gaimi1
End If
End Sub

Private Sub Form_Load()
'�Զ������û���
Text1.Text = username
End Sub
