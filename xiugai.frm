VERSION 5.00
Begin VB.Form xiugai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "修改用户信息"
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
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重置密码"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "删除用户"
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
      Caption         =   "选择用户名"
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
'连接数据库
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
cn.Open "provider = microsoft.jet.oledb.4.0;data source=" & App.Path & "\db\kt.mdb;persist security info = false"
sql = "select * from ktuser where uid = '" & Combo1.Text & "'"
rs.Open sql, cn, 3, 2
Dim str As Integer
str = MsgBox("确定删除？", 1 + 32 + 256, "提示")
'确认删除，更新数据库
If str = 1 Then
    rs.Delete
    MsgBox "该用户已删除"
    '关闭数据库
    rs.Close
    cn.Close
    '重新载入用户信息
    Unload xiugai
    xiugai.Show 1
    End If
End Sub

Private Sub Command2_Click()
Dim str As Integer
str = MsgBox("确定重置？", 1 + 32 + 256, "提示")
If str = 1 Then
    '连接数据库
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    cn.Open "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\db\kt.mdb;persist security info = false"
    sql = "select upasswd from ktuser where uid = '" & Combo1.Text & "'"
    rs.Open sql, cn, 3, 2
    '重置密码为000000
    rs("upasswd") = "000000"
    '更新数据库
    rs.Update
    MsgBox "密码修改成功"
End If
End Sub

Private Sub Command3_Click()
Dim str As Integer
str = MsgBox("确定退出？", 1 + 32 + 256, "提示")
If str = 1 Then
Unload xiugai
End If
End Sub

Private Sub Form_Load()
'连接数据库
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\kt.mdb;Persist Security Info=False"
'载入普通管理员信息
sql = "select uid from ktuser where quanxian = '1' "
rs.Open sql, cn, 1, 1
'循环向列表框中添加普通管理员用户名
Do While rs.EOF <> True
    Combo1.AddItem rs("uid")
    rs.MoveNext
    Loop
'关闭数据库
rs.Close
cn.Close
End Sub
