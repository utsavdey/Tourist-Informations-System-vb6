VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14700
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   8925
   ScaleWidth      =   14700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   8
      ToolTipText     =   "Click here to go to the home page"
      Top             =   6480
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   9000
      Top             =   6840
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
      CommandType     =   1
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
      Connect         =   $"Form3.frx":4268D
      OLEDBString     =   $"Form3.frx":42726
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from table3"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form3.frx":427BF
      Height          =   6375
      Left            =   3840
      TabIndex        =   3
      Top             =   1320
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11245
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "PA_ID"
         Caption         =   "PA_ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "PACKAGE DETAILS"
         Caption         =   "PACKAGE DETAILS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "package_cost"
         Caption         =   "package_cost"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "package_duration"
         Caption         =   "package_duration"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "type"
         Caption         =   "type"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Form3.frx":427D4
      Left            =   1200
      List            =   "Form3.frx":42850
      TabIndex        =   2
      Top             =   5280
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "PACKAGE DURATION"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      TabIndex        =   5
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      Caption         =   "PACKAGE TYPE"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      TabIndex        =   6
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "PACKAGE COST"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   720
      TabIndex        =   4
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PACKAGE AND TOUR OPTIONS"
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   2640
      TabIndex        =   7
      Top             =   360
      Width           =   9975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Combo1_Click()
If (Combo1 <> "") Then
Adodc1.RecordSource = "select * from table3 where package_cost='" & Combo1 & "'"
Adodc1.Refresh
End If
End Sub
Private Sub Combo2_Click()
If (Combo2 <> "") Then
Adodc1.RecordSource = "select * from table3 where package_duration='" & Combo2 & "'"
Adodc1.Refresh
End If
End Sub
Private Sub Combo3_Click()
If (Combo3 <> "") Then
Adodc1.RecordSource = "select * from table3 where type='" & Combo3 & "'"
Adodc1.Refresh
End If
End Sub



Private Sub Command1_Click()
Form3.Hide
Form1.Show
End Sub

Private Sub Form_Load()
cn.ConnectionString = Form1.str
cn.Open
rs.ActiveConnection = cn
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
rs.Source = "table3"
rs.Open
Adodc1.ConnectionString = Form1.str
Combo1.Clear
Combo2.Clear
'Combo3.Clear

Combo1.AddItem "3500/-"
Combo1.AddItem "4000/-"
Combo1.AddItem "5000/-"
Combo1.AddItem "6000/-"
Combo1.AddItem "7500/-"
Combo1.AddItem "8000/-"
Combo1.AddItem "8500/-"
Combo1.AddItem "9000/-"
Combo1.AddItem "9500/-"
Combo1.AddItem "10000/-"
Combo1.AddItem "10500/-"
Combo1.AddItem "11000/-"
Combo1.AddItem "15000/-"
Combo1.AddItem "17000/-"
Combo2.AddItem "5 Days"
Combo2.AddItem "6 Days"
Combo2.AddItem "7 Days"
Combo2.AddItem "8 Days"
Combo2.AddItem "9 Days"
Combo2.AddItem "10 Days"
Combo2.AddItem "11 Days"
Combo2.AddItem "12 Days"
Combo2.AddItem "15 Days"

'Adodc1.RecordSource = "select distinct PACKAGE_COST,PACKAGE_DURATION,TYPE from table3"
'Adodc1.Refresh
'    Do Until .EOF
'        Combo1.AddItem ![package_cost]
'        Combo2.AddItem ![package_duration]
'        Combo3.AddItem ![Type]
 '   .MoveNext
'    Loop
'End With
End Sub



