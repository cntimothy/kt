VERSION 5.00
Begin VB.Form shezhi1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "告警设置"
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
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command3 
      Caption         =   "恢复默认值"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
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
      Caption         =   "请输入告警值"
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
i = MsgBox("确定修改？", 1 + 32 + 256, "提示")
If i = 1 Then
    '验证输入的是否为数字
    If (IsNumeric(Text1.Text) = False) Or (Text1.Text = "") Then
        MsgBox "请输入数字", vbCritical, "提示"
        Text1.Text = ""
    Else
        '连接数据库，修改配置表中的告警值
        Set cn = New ADODB.Connection
        Set rs = New ADODB.Recordset
        cn.Open "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\db\kt.mdb;persist security info = false"
        sql = "select gaojing from ktpeizhi"
        rs.Open sql, cn, 3, 2
        rs("gaojing") = Text1.Text
        rs.Update
        '事实更新g值
        g = Text1.Text
        MsgBox "修改成功!点击确定返回"
        Unload shezhi1
    End If
End If
End Sub

Private Sub Command2_Click()
Unload shezhi1
End Sub

Private Sub Command3_Click()
Dim i As Integer
i = MsgBox("确定恢复初始值？", 1 + 32 + 256, "提示")
If i = 1 Then
    '连接数据库，更新配置表中的g值为10
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    cn.Open "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\db\kt.mdb;persist security info = false"
    sql = "select gaojing from ktpeizhi"
    rs.Open sql, cn, 3, 2
    rs("gaojing") = 10
    rs.Update
    MsgBox "修改成功"
End If
End Sub

Private Sub Form_Load()
'在文本框中显示当前告警值
Text1.Text = g
End Sub
