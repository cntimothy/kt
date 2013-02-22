VERSION 5.00
Begin VB.Form zhuce 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "注册"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   13215
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command4 
      Caption         =   "退出注册"
      Height          =   375
      Left            =   8280
      TabIndex        =   26
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "重填"
      Height          =   375
      Left            =   5520
      TabIndex        =   25
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "完成注册"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   24
      Top             =   7800
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      Caption         =   "女"
      Height          =   495
      Left            =   11040
      TabIndex        =   23
      Top             =   4320
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "男"
      Height          =   495
      Left            =   9480
      TabIndex        =   22
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   8880
      TabIndex        =   20
      Top             =   6960
      Width           =   2655
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   18
      Top             =   6960
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "生成申请码"
      Height          =   375
      Left            =   840
      TabIndex        =   17
      Top             =   6960
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   7200
      TabIndex        =   16
      Top             =   5160
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   5160
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6600
      TabIndex        =   14
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      Height          =   180
      Left            =   8760
      TabIndex        =   21
      Top             =   4440
      Width           =   360
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "请输入获得的注册码"
      Height          =   180
      Left            =   6600
      TabIndex        =   19
      Top             =   7080
      Width           =   1620
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   12480
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "联系方式"
      Height          =   180
      Left            =   6000
      TabIndex        =   12
      Top             =   5280
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "所在科室"
      Height          =   180
      Left            =   1200
      TabIndex        =   11
      Top             =   5280
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      Height          =   180
      Left            =   5880
      TabIndex        =   10
      Top             =   4440
      Width           =   360
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Left            =   1440
      TabIndex        =   9
      Top             =   4440
      Width           =   360
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   12480
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "以下信息必填"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   8
      Top             =   3600
      Width           =   1710
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "(请输入6~16位数字或英文字符)"
      Height          =   180
      Left            =   6000
      TabIndex        =   7
      Top             =   1680
      Width           =   2520
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(请输入5~15位英文字符)"
      Height          =   180
      Left            =   6000
      TabIndex        =   6
      Top             =   840
      Width           =   1980
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "请再次输入密码"
      Height          =   180
      Left            =   1080
      TabIndex        =   2
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      Height          =   180
      Left            =   1560
      TabIndex        =   1
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "用户名"
      Height          =   180
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   540
   End
End
Attribute VB_Name = "zhuce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim m As Long
m = CLng(Rnd * (999999 - 100000 + 1) + 100000)
Text8.Text = m
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Then
    MsgBox "项目不能为空！", vbCritical, "提示"
Else
    If Len(Text1.Text) > 15 Or Len(Text1.Text) < 5 Then
        MsgBox "用户名长度错误!", vbCritical, "提示"
        '用户名和密码置空
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
    Else
        If Len(Text2.Text) > 16 Or Len(Text2.Text) < 6 Then
            MsgBox "密码长度错误!", vbCritical, "提示"
            '密码置空
            Text2.Text = ""
            Text3.Text = ""
        Else
            If Text2.Text <> Text3.Text Then
                MsgBox "两次密码不同", vbCritical, "提示"
                '密码置空
                Text2.Text = ""
                Text3.Text = ""
            Else
                Set cn = New ADODB.Connection
                Set rs = New ADODB.Recordset
                cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\kt.mdb;Persist Security Info=False"
                sql = "select * from ktuser where uid = '" & Text1.Text & "'"
                rs.Open sql, cn, 1, 1
                If rs.EOF <> True Then
                    MsgBox "该用户名已存在！请重新输入"
                    '用户名置空
                    Text1.Text = ""
                    Text1.SetFocus
                Else
                    rs.Close
                    Dim s1 As Double
                    Dim s2 As String
                    s1 = Text8.Text
                    s2 = Text9.Text
                    '注册码通过验证
                    If s2 = Left(CStr(s1 * 2 - Sqr(s1)), 6) Then
                        sql = "select * from ktuser"
                        '向数据库添加新用户
                        rs.Open sql, cn, 2, 3
                        rs.AddNew
                        rs("uid") = Trim(Text1.Text)
                        rs("upasswd") = Trim(Text2.Text)
                        rs("xingming") = Trim(Text4.Text)
                        rs("nianling") = Trim(Text5.Text)
                        If Option1.Value = True Then
                            rs("xingbie") = "男"
                        Else: rs("xingbie") = "女"
                        End If
                        rs("keshi") = Trim(Text6.Text)
                        rs("lianxi") = Trim(Text7.Text)
                        rs("quanxian") = 1
                        rs.Update
                        MsgBox "注册成功 请登录"
                        Unload zhuce
                    Else
                        MsgBox "注册码错误!请重新输入", vbCritical, "提示"
                        Text9.Text = ""
                    End If
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub Command3_Click()
Dim s As Integer
s = MsgBox("确定重填？重填后丢失所有数据！", 1 + 32 + 256, "重填")
'所有文本框置空
If s = 1 Then
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
End If
End Sub

Private Sub Command4_Click()
Dim str As Integer
str = MsgBox("确定退出注册？", 1 + 32 + 256, "提示")
If str = 1 Then
Unload zhuce
End If
End Sub

