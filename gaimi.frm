VERSION 5.00
Begin VB.Form gaimi1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "密码管理"
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
   StartUpPosition =   1  '所有者中心
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
      Caption         =   "退出"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "（6~16位）"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "密码确认"
      Height          =   180
      Left            =   480
      TabIndex        =   3
      Top             =   2760
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "新密码"
      Height          =   180
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "原密码"
      Height          =   180
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "用户名"
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
str = MsgBox("确定修改？", 1 + 32 + 256, "提示")
If str = 1 Then
    '验证原密码与新密码文本框是否为空
    If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
        MsgBox "项目不能为空", vbCritical, "提示"
    Else
        '验证两次密码输入是否相同
        If Text3.Text <> Text4.Text Then
            MsgBox "两次密码输入不同", vbCritical, "提示"
            Text3.Text = ""
            Text4.Text = ""
        Else
            '验证密码长度是否为6~16
            If Len(Text3.Text) < 6 Or Len(Text3.Text) > 16 Then
                MsgBox "密码长度错误！", vbCritical, "提示"
                Text3.Text = ""
                Text4.Text = ""
            Else
                '连接数据库
                Set cn = New ADODB.Connection
                Set rs = New ADODB.Recordset
                cn.Open "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\db\kt.mdb;persist security info = false"
                sql = "select * from ktuser where uid = '" & Text1.Text & "'"
                rs.Open sql, cn, 3, 2
                '验证原密码是否正确
                If rs("upasswd") <> Trim(Text2.Text) Then
                    MsgBox "原密码错误，请重新输入", vbCritical, "提示"
                    '原密码错误，将原密码文本框和新密码文本框置空
                    Text2.Text = ""
                    Text3.Text = ""
                    Text4.Text = ""
                Else
                    '原密码通过验证，更新数据库，修改密码
                    rs("upasswd") = Trim(Text3.Text)
                    rs.Update
                    MsgBox "密码修改成功,请重新登录"
                    '重新登录
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
str = MsgBox("确定退出？", 1 + 32 + 256, "提示")
If str = 1 Then
    Unload gaimi1
End If
End Sub

Private Sub Form_Load()
'自动载入用户名
Text1.Text = username
End Sub
