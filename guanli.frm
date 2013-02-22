VERSION 5.00
Begin VB.Form guanli 
   Caption         =   "管理"
   ClientHeight    =   8445
   ClientLeft      =   3075
   ClientTop       =   2280
   ClientWidth     =   12345
   Icon            =   "guanli.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "guanli.frx":030A
   ScaleHeight     =   8445
   ScaleWidth      =   12345
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "告警"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "库存"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "入库"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "出库"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Menu xitong 
      Caption         =   "系统"
      Begin VB.Menu yonghu 
         Caption         =   "用户管理"
         Begin VB.Menu gaimi 
            Caption         =   "更改密码"
         End
         Begin VB.Menu zhuxiao 
            Caption         =   "注销"
         End
      End
      Begin VB.Menu fenge 
         Caption         =   "-"
      End
      Begin VB.Menu genghuan 
         Caption         =   "更换用户"
      End
      Begin VB.Menu tuichu 
         Caption         =   "退出系统"
      End
   End
   Begin VB.Menu yewu 
      Caption         =   "业务"
      Begin VB.Menu ruku 
         Caption         =   "入库"
      End
      Begin VB.Menu chuku 
         Caption         =   "出库"
      End
   End
   Begin VB.Menu zhangbiao 
      Caption         =   "帐表"
      Begin VB.Menu kucun 
         Caption         =   "库存"
      End
      Begin VB.Menu sousuo 
         Caption         =   "搜索"
      End
      Begin VB.Menu fenge3 
         Caption         =   "-"
      End
      Begin VB.Menu baobiao 
         Caption         =   "生成报表"
      End
   End
   Begin VB.Menu gaojing 
      Caption         =   "告警"
      Begin VB.Menu chakan 
         Caption         =   "查看"
      End
      Begin VB.Menu shezhi 
         Caption         =   "设置"
      End
   End
   Begin VB.Menu chaoji 
      Caption         =   "超级权限"
      Begin VB.Menu rizhi 
         Caption         =   "查看日志"
      End
      Begin VB.Menu 权限 
         Caption         =   "权限管理"
      End
   End
   Begin VB.Menu bangzhu 
      Caption         =   "帮助"
      Begin VB.Menu wenti 
         Caption         =   "常见问题"
      End
      Begin VB.Menu fenge2 
         Caption         =   "-"
      End
      Begin VB.Menu guanyu 
         Caption         =   "关于"
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
'载入告警值
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
cn.Open "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\db\kt.mdb;persist security info = false"
sql = "select gaojing from ktpeizhi"
rs.Open sql, cn, 1, 1
g = rs("gaojing")
rs.Close
cn.Close
'在标题中显示当前用户信息
If limit = 1 Then
    chaoji.Enabled = False
    guanli.Caption = "管理  " & "当前登录用户为" & username & "  权限为普通管理员"
Else:
    guanli.Caption = "管理  " & "当前登录用户为" & username & "  权限为超级管理员"
    zhuxiao.Enabled = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If q = True Then
    Dim str As Integer
    str = MsgBox("确定退出该系统？", 1 + 32 + 256, "提示")
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
str = MsgBox("确定更换用户？", 1 + 32 + 256, "更换用户")
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

Private Sub 权限_Click()
yonghu1.Show 1
End Sub
