VERSION 5.00
Begin VB.Form guanli 
   Caption         =   "����"
   ClientHeight    =   8445
   ClientLeft      =   3075
   ClientTop       =   2280
   ClientWidth     =   12345
   Icon            =   "guanli.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "guanli.frx":030A
   ScaleHeight     =   8445
   ScaleWidth      =   12345
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "�澯"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Menu xitong 
      Caption         =   "ϵͳ"
      Begin VB.Menu yonghu 
         Caption         =   "�û�����"
         Begin VB.Menu gaimi 
            Caption         =   "��������"
         End
         Begin VB.Menu zhuxiao 
            Caption         =   "ע��"
         End
      End
      Begin VB.Menu fenge 
         Caption         =   "-"
      End
      Begin VB.Menu genghuan 
         Caption         =   "�����û�"
      End
      Begin VB.Menu tuichu 
         Caption         =   "�˳�ϵͳ"
      End
   End
   Begin VB.Menu yewu 
      Caption         =   "ҵ��"
      Begin VB.Menu ruku 
         Caption         =   "���"
      End
      Begin VB.Menu chuku 
         Caption         =   "����"
      End
   End
   Begin VB.Menu zhangbiao 
      Caption         =   "�ʱ�"
      Begin VB.Menu kucun 
         Caption         =   "���"
      End
      Begin VB.Menu sousuo 
         Caption         =   "����"
      End
      Begin VB.Menu fenge3 
         Caption         =   "-"
      End
      Begin VB.Menu baobiao 
         Caption         =   "���ɱ���"
      End
   End
   Begin VB.Menu gaojing 
      Caption         =   "�澯"
      Begin VB.Menu chakan 
         Caption         =   "�鿴"
      End
      Begin VB.Menu shezhi 
         Caption         =   "����"
      End
   End
   Begin VB.Menu chaoji 
      Caption         =   "����Ȩ��"
      Begin VB.Menu rizhi 
         Caption         =   "�鿴��־"
      End
      Begin VB.Menu Ȩ�� 
         Caption         =   "Ȩ�޹���"
      End
   End
   Begin VB.Menu bangzhu 
      Caption         =   "����"
      Begin VB.Menu wenti 
         Caption         =   "��������"
      End
      Begin VB.Menu fenge2 
         Caption         =   "-"
      End
      Begin VB.Menu guanyu 
         Caption         =   "����"
      End
   End
End
Attribute VB_Name = "guanli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub baobiao_Click()
e = 1
excel.Show 1
End Sub

Private Sub chakan_Click()
gaojing1.Show 1
End Sub

Private Sub chuku_Click()
chuku1.Show 1
End Sub

Private Sub chushihua_Click()
chushihua1.Show 1
End Sub

Private Sub Command1_Click()
chuku1.Show 1
End Sub

Private Sub Command2_Click()
ruku1.Show 1
End Sub

Private Sub Command3_Click()
liulan.Show 1
End Sub

Private Sub Command4_Click()
gaojing1.Show 1
End Sub

Private Sub Form_Load()
q = True
'����澯ֵ
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
cn.Open "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\db\kt.mdb;persist security info = false"
sql = "select gaojing from ktpeizhi"
rs.Open sql, cn, 1, 1
g = rs("gaojing")
rs.Close
cn.Close
'�ڱ�������ʾ��ǰ�û���Ϣ
If limit = 1 Then
    chaoji.Enabled = False
    guanli.Caption = "����  " & "��ǰ��¼�û�Ϊ" & username & "  Ȩ��Ϊ��ͨ����Ա"
Else:
    guanli.Caption = "����  " & "��ǰ��¼�û�Ϊ" & username & "  Ȩ��Ϊ��������Ա"
    zhuxiao.Enabled = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If q = True Then
    Dim str As Integer
    str = MsgBox("ȷ���˳���ϵͳ��", 1 + 32 + 256, "��ʾ")
    If str = 1 Then
        Cancel = 0
    Else
        Cancel = 1
    
    End If
Else: Cancel = 0
End If
End Sub

Private Sub gaimi_Click()
gaimi1.Show 1
End Sub

Private Sub genghuan_Click()
q = False
Dim str As Integer
str = MsgBox("ȷ�������û���", 1 + 32 + 256, "�����û�")
If str = 1 Then
Unload guanli
Welcome.Show 0
login.Show 1
End If
End Sub

Private Sub guanyu_Click()
about.Show 1
End Sub

Private Sub kucun_Click()
liulan.Show 1
End Sub

Private Sub rizhi_Click()
rizhi1.Show 1
End Sub

Private Sub ruku_Click()
ruku1.Show 1
End Sub

Private Sub shezhi_Click()
shezhi1.Show 1
End Sub

Private Sub sousuo_Click()
sousuo1.Show 1
End Sub

Private Sub tuichu_Click()
Unload guanli
End Sub

Private Sub zhuxiao_Click()
zhuxiao1.Show 1
End Sub

Private Sub Ȩ��_Click()
yonghu1.Show 1
End Sub
