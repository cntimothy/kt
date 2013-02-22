VERSION 5.00
Begin VB.Form login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "登录"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6030
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6030
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
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
      Caption         =   "密码"
      Height          =   180
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "用户名"
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
'连接数据库
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\kt.mdb;Persist Security Info=False"
sql = "select * from ktuser where uid = '" & Text1.Text & "'"
rs.Open sql, cn, 1, 1
    '用户名未找到，提示错误
    If rs.EOF = True Then
    MsgBox "用户名不存在！请确认", vbCritical, "提示"
    '用户名及密码置空
    Text2.Text = ""
    Text1.Text = ""
    '设置焦点到用户名文本框
    Text1.SetFocus
Else
    '用户名找到，验证密码
    If rs("upasswd") = Trim(Text2.Text) Then
        '将用户的权限值赋给全局变量limit，用户名赋给全局变量username
        limit = rs("quanxian")
        username = rs("uid")
        '进入管理系统
        Unload Welcome
        Unload login
        guanli.Show
    Else:
        '未通过密码验证，密码文本框置空
        MsgBox "密码错误！", vbCritical, "提示"
        Text2.Text = ""
    End If
End If
End Sub

Private Sub Command2_Click()
Dim str As Integer
str = MsgBox("确定退出登录？", 1 + 32 + 256, "提示")
If str = 1 Then
Unload login
q = True
End If
End Sub

Private Sub Label1_Click()

End Sub

