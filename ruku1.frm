VERSION 5.00
Begin VB.Form ruku1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "入库"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11085
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command6 
      Caption         =   "清除所有项目"
      Height          =   495
      Left            =   2160
      TabIndex        =   18
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "删除"
      Height          =   495
      Left            =   480
      TabIndex        =   17
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      Height          =   495
      Left            =   9360
      TabIndex        =   16
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "入库"
      Height          =   495
      Left            =   7680
      TabIndex        =   15
      Top             =   7320
      Width           =   1215
   End
   Begin VB.ListBox List4 
      Height          =   4380
      Left            =   7920
      TabIndex        =   14
      Top             =   2400
      Width           =   2175
   End
   Begin VB.ListBox List3 
      Height          =   4380
      Left            =   5880
      TabIndex        =   13
      Top             =   2400
      Width           =   1935
   End
   Begin VB.ListBox List2 
      Height          =   4380
      Left            =   4320
      TabIndex        =   12
      Top             =   2400
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   4380
      Left            =   480
      TabIndex        =   11
      Top             =   2400
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "下一项"
      Default         =   -1  'True
      Height          =   495
      Left            =   8280
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "日期"
      Height          =   180
      Left            =   7920
      TabIndex        =   10
      Top             =   2040
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "供应商"
      Height          =   180
      Left            =   5880
      TabIndex        =   9
      Top             =   2040
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "数量"
      Height          =   180
      Left            =   4320
      TabIndex        =   8
      Top             =   2040
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "货号"
      Height          =   180
      Left            =   480
      TabIndex        =   7
      Top             =   2040
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "供应商"
      Height          =   180
      Left            =   7320
      TabIndex        =   2
      Top             =   480
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "数量"
      Height          =   180
      Left            =   4680
      TabIndex        =   1
      Top             =   480
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "货号"
      Height          =   180
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   360
   End
End
Attribute VB_Name = "ruku1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'验证货物信息是否填写完整
If Trim(Text1.Text) <> "" And Trim(Text2.Text) <> "" And Trim(Text3.Text) <> "" Then
    '验证数量文本框是否填写的是数字
    If IsNumeric(Text2.Text) = True Then
        '将货物信息填入列表框
        List1.AddItem (Trim(Text1.Text))
        List2.AddItem (Trim(Text2.Text))
        List3.AddItem (Trim(Text3.Text))
        List4.AddItem (Now())
        '货物信息加入列表框之后将文本框置空
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        '焦点回到货号文本框
        Text1.SetFocus
    Else
        MsgBox "数量应为数字！", vbCritical, "提示"
        Text2.SetFocus
    End If
Else: MsgBox "项目不能为空", vbCritical, "错误！"
End If
End Sub

Private Sub Command2_Click()
Dim i, j As Integer
i = MsgBox("确定入库？", 1 + 32 + 256, "入库")
If i = 1 Then
    '连接数据库
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\kt.mdb;Persist Security Info=False"
    '通过循环将列表框中的信息写入数据库
    For j = 0 To List1.ListCount - 1
        List1.ListIndex = j
        List2.ListIndex = j
        List3.ListIndex = j
        List4.ListIndex = j
        '写入管理日志
        sql = "select * from ktchuru"
        rs.Open sql, cn, 3, 2
        rs.AddNew
        rs("riqi") = List4.Text
        rs("churu") = "入库"
        rs("huohao") = List1.Text
        rs("shuliang") = CInt(List2.Text)
        rs("gongying") = List3.Text
        rs("jingban") = username
        rs.Update
        rs.Close
        '写入库存表
        sql = "select * from ktkucun where huohao = '" & List1.Text & "'"
        rs.Open sql, cn, 3, 2
        If rs.EOF <> True Then
        '若库存中已有同货号的物品，将更新数量
            rs.Close
            sql = "select * from ktkucun where huohao = '" & List1.Text & "'"
            rs.Open sql, cn, 3, 2
            rs("kucun") = rs("kucun") + List2.Text
            rs("riqi") = List4.Text
        Else
            rs.Close
        '若库存中无相同货号物品，将增加该物品
            sql = "select * from ktkucun"
            rs.Open sql, cn, 3, 2
            rs.AddNew
            rs("huohao") = List1.Text
            rs("kucun") = List2.Text
            rs("riqi") = List4.Text
        End If
    rs.Update
    rs.Close
    Next
    MsgBox "上述项目已入库", , "提示"
    Text1.SetFocus
End If
End Sub

Private Sub Command4_Click()
Dim i As Integer
i = MsgBox("确定退出入库模式？", 1 + 32 + 256, "提示")
If i = 1 Then
    Unload ruku1
End If
End Sub

Private Sub Command5_Click()
'删除选中项目
Dim i As Integer
For i = List1.ListCount - 1 To 0 Step -1
'验证哪些项目被选中
If List1.Selected(i) = True Or List2.Selected(i) = True Or List3.Selected(i) = True Or List4.Selected(i) = True Then
    List1.RemoveItem (i)
    List2.RemoveItem (i)
    List3.RemoveItem (i)
    List4.RemoveItem (i)
End If
Next
End Sub

Private Sub Command6_Click()
'清空列表
Dim i, j As Integer
i = MsgBox("确定清除？", 1 + 32 + 256, "提示")
If i = 1 Then
    '循环删除列表框中的所有信息
    For j = List1.ListCount - 1 To 0 Step -1
        List1.RemoveItem (j)
        List2.RemoveItem (j)
        List3.RemoveItem (j)
        List4.RemoveItem (j)
    Next
End If
End Sub

