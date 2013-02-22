VERSION 5.00
Begin VB.Form excel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "输出"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   6420
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      Width           =   3735
   End
   Begin VB.DirListBox Dir1 
      Height          =   3030
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   3960
      TabIndex        =   0
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "提醒：请勿输出至磁盘根目录，并确保目录下无同名文件！"
      Height          =   180
      Left            =   960
      TabIndex        =   6
      Top             =   360
      Width           =   4680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "输出至目录"
      Height          =   180
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   900
   End
End
Attribute VB_Name = "excel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ktpath As String

Private Sub Command1_Click()
If ktpath = "" Then
    MsgBox "路径不能为空", vbCritical, "提示"
Else
    Dim conn As New ADODB.Connection
    Dim rsqrecord As ADODB.Recordset
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\kt.mdb;Persist Security Info=False"
    response = MsgBox("确定要导出数据吗？", 1 + 32 + 256, "提示")
    If response = vbOK Then
        Dim str As String
        '获取当日日期
        str = CStr(Date)
'选择输出数据
        Select Case e
            Case 1
                conn.Execute "select huohao as 货号, kucun as 库存, riqi as 最后进出货日期 into [excel 8.0;database=" & ktpath & "\" & str & "库存.xls].sheet1 from ktkucun order by huohao"
                MsgBox "导出成功,文件名为今天日期+库存"
            Case 0
                conn.Execute "select riqi as 日期, churu as 出库入库, huohao as 货号, shuliang as 数量,jingban as 经办人,gongying as 供应商 into [excel 8.0;database=" & ktpath & "\" & str & "日志.xls].sheet1 from ktchuru order by riqi"
                MsgBox "导出成功,文件名为今天日期+日志"
            Case 2
                conn.Execute "select huohao as 货号, kucun as 库存, riqi as 最后进出货日期  into [excel 8.0;database=" & ktpath & "\" & str & "告警.xls].sheet1 from ktkucun where kucun < " & g & " order by huohao"
                MsgBox "导出成功,文件名为今天日期+告警"
            Case 3
                conn.Execute "select uid as 用户名, xingming as 姓名,nianling as 年龄, xingbie as 性别, keshi as 科室,lianxi as 联系方式  into [excel 8.0;database=" & ktpath & "\" & str & "用户.xls].sheet1 from ktuser"
                MsgBox "导出成功,文件名为今天日期+用户"
            End Select
            Unload excel
    End If
End If
End Sub
    
Private Sub Command2_Click()
Dim str As Integer
str = MsgBox("确定退出？", 1 + 32 + 256, "提示")
If str = 1 Then
Unload excel
End If
End Sub
'获取输出路径
Private Sub Dir1_Change()
ktpath = Dir1.Path
Text1.Text = ktpath
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
ktpath = Drive1.Drive
Text1.Text = ktpath
End Sub

