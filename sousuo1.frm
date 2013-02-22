VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form sousuo1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "搜索"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   6300
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   7320
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2400
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
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
      Height          =   5700
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   10054
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
   Begin VB.CommandButton Command1 
      Caption         =   "搜索"
      Default         =   -1  'True
      Height          =   615
      Left            =   5280
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请输入货号"
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   900
   End
End
Attribute VB_Name = "sousuo1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'点击搜索按钮，将设置datagrid控件为可见
DataGrid1.Visible = True
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\kt.mdb;Persist Security Info=False"
'模糊搜索
sql = "select * from ktkucun where huohao like '%" & Text1.Text & "%'"
rs.Open sql, cn, 1, 1
'未找到符合关键词的货物
If rs.EOF = True Then
    MsgBox "没有该货物，请确认", vbCritical, "提示"
Else
    rs.Close
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db\kt.mdb;Persist Security Info=False"
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select huohao as 货号,kucun as 库存, riqi as 最后进货日期 from ktkucun where huohao like '%" & Text1.Text & "%' order by huohao asc"
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
    DataGrid1.Refresh
End If
End Sub

Private Sub Command2_Click()
Dim i As Integer
i = MsgBox("确定退出？", 1 + 32 + 256, "提示")
If i = 1 Then
    Unload sousuo1
End If
End Sub

Private Sub Form_Load()
'设置datagrid控件为不可见
DataGrid1.Visible = False
End Sub
