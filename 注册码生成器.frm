VERSION 5.00
Begin VB.Form 注册码生成器 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "注册码生成器"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6525
   Icon            =   "注册码生成器.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6525
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3360
      TabIndex        =   5
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "生成注册码"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "请输入超级管理员密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   2400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请输入申请码(6位数字)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   2520
   End
End
Attribute VB_Name = "注册码生成器"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'验证输入是否有误
If Not IsNumeric(Text1.Text) Or Len(Text1.Text) <> 6 Then
    MsgBox "您输入有误，不是数字或位数不对，请核对后重新输入", vbCritical, "错误！"
Else
    Set cn = New ADODB.connection
    Set rs = New ADODB.Recordset
    Dim sql As String
    cn.Open "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\db\kt.mdb;persist security info = false"
    sql = "select upasswd from ktuser where uid = 'root'"
    rs.Open sql, cn, 3, 2
    '验证超级管理员密码
    If Text2.Text = rs("upasswd") Then
        '生成注册码
        Dim m As Long
        Dim str As String
        m = Trim(Text1.Text)
        str = Left(CStr(m * 2 - Sqr(m)), 6)
        Text3.Text = str
    End If
End If
End Sub

