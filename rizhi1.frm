VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form rizhi1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "日志"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   14415
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command2 
      Caption         =   "生成Excel"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "退出"
      Height          =   495
      Left            =   8280
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   600
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Align           =   1  'Align Top
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   8070
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   16
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "rizhi1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim str As Integer
str = MsgBox("确定退出？", 1 + 32 + 256, "提示")
If str = 1 Then
Unload rizhi1
End If
End Sub

Private Sub Command2_Click()
'设置输出类型
e = 0
excel.Show 1
End Sub

Private Sub Form_Load()
'连接数据库
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
'清除90天以上的日志记录
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\kt.mdb;Persist Security Info=False"
Dim s, s1 As String
s = Format(Now, "yyyymmdd")
s = s - 90
sql = "select * from ktchuru "
rs.Open sql, cn, 3, 2
Do While rs.EOF <> True
    s1 = Format(rs("riqi"), "yyyymmdd")
    If s1 < s Then
        rs.Delete
    End If
    rs.MoveNext
Loop
rs.Close
'设置数据源
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\kt.mdb;Persist Security Info=False"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select riqi as 日期,churu as 出库入库, huohao as 货号,shuliang as 数量, jingban as 经办人,gongying as 供应商 from ktchuru order by riqi asc"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub
