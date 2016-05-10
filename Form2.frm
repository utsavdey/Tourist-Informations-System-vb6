VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SEARCH YOUR TRAVEL DESTINATION"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12930
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   5715
   ScaleWidth      =   12930
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      TabIndex        =   12
      ToolTipText     =   "Click here to go back to home screen"
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   10
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium Cond"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   9
      Text            =   "Enter name of place to search"
      Top             =   960
      Width           =   4695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   5040
      Top             =   5880
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
      Connect         =   $"Form2.frx":34DC2
      OLEDBString     =   $"Form2.frx":34E5B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from table1"
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
      Bindings        =   "Form2.frx":34EF4
      Height          =   3135
      Left            =   3720
      TabIndex        =   8
      Top             =   1800
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5530
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Search Result(s)"
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "Form2.frx":34F09
      Left            =   960
      List            =   "Form2.frx":34F10
      Style           =   2  'Dropdown List
      TabIndex        =   7
      ToolTipText     =   "Enter the type of your travel destination"
      Top             =   4680
      Width           =   1815
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Form2.frx":34F20
      Left            =   960
      List            =   "Form2.frx":34F27
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   "Enter the city to which you wish to travel"
      Top             =   3360
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form2.frx":34F34
      Left            =   960
      List            =   "Form2.frx":34F3B
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Enter the name of the state in which you wish to travel"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form2.frx":34F4C
      Left            =   960
      List            =   "Form2.frx":34F5C
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Enter the name of the region in which you wish to travel"
      Top             =   720
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0080C0FF&
      Caption         =   "SELECT      TYPE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   855
      Left            =   720
      TabIndex        =   3
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      Caption         =   "SELECT        CITY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   855
      Left            =   720
      TabIndex        =   2
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "SELECT      STATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   855
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "SELECT    REGION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "SEARCH YOUR TRAVEL DESTINATION"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      Top             =   360
      Width           =   8055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00004080&
      Height          =   3375
      Left            =   3600
      Top             =   1680
      Width           =   7935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Combo1_Click()
If (Combo1 <> "") Then
Adodc1.RecordSource = "select * from table1 where region_name='" & Combo1 & "'"
Adodc1.Refresh
Combo3.Clear
If (Combo1 = "Eastern") Then
Combo2.Clear
Combo2.AddItem "Odisha"
Combo2.AddItem "West Bengal"
Combo2.AddItem "Bihar"
Combo2.AddItem "Jharkhand"
ElseIf (Combo1 = "North") Then
Combo2.Clear
Combo2.AddItem "Jammu & Kashmir"
Combo2.AddItem "Delhi"
Combo2.AddItem "Uttar Pradesh"
Combo2.AddItem "Himachal Pradesh"
Combo2.AddItem "Uttarakhand"
Combo2.AddItem "Punjab"
Combo2.AddItem "Haryana"
ElseIf (Combo1 = "North-East") Then
Combo2.Clear
Combo2.AddItem "Assam"
Combo2.AddItem "Sikkim"
Combo2.AddItem "Arunachal Pradesh"
Combo2.AddItem "Meghalay"
Combo2.AddItem "Nagaland"
Combo2.AddItem "Manipur"
Combo2.AddItem "Mizoram"
ElseIf (Combo1 = "West") Then
Combo2.Clear
Combo2.AddItem "Rajasthan"
Combo2.AddItem "Gujarat"
Combo2.AddItem "Maharashtra"
Combo2.AddItem "Goa"
Combo2.AddItem "Daman and Diu"
ElseIf (Combo1 = "Middle") Then
Combo2.Clear
Combo2.AddItem "Chattisgarh"
Combo2.AddItem "Madhya Pradesh"
ElseIf (Combo1 = "South") Then
Combo2.Clear
Combo2.AddItem "Karnataka"
Combo2.AddItem "Tamilnadu"
Combo2.AddItem "Kerala"
Combo2.AddItem "Lakshadweep"
Combo2.AddItem "Dadra and Nagar Haveli"
Combo2.AddItem "Puducherry"
End If
End If
End Sub

Private Sub Combo2_Click()
If (Combo2 <> "") Then
Adodc1.RecordSource = "select * from table1 where state_name='" & Combo2 & "'"
Adodc1.Refresh
'=============================EAST============================
If (Combo2 = "Odisha") Then
Combo3.Clear
Combo3.AddItem "Puri"
Combo3.AddItem "Bhubaneswar"
Combo3.AddItem "Chandipur"
Combo3.AddItem "Simlipal"
ElseIf (Combo2 = "West Bengal") Then
Combo3.Clear
Combo3.AddItem "Kolkata"
Combo3.AddItem "Darjiling"
Combo3.AddItem "Sundarban"
Combo3.AddItem "Dooars"
Combo3.AddItem "Midnapore"
Combo3.AddItem "Murshidabad"
Combo3.AddItem "Siliguri"
Combo3.AddItem "Hoogly"
Combo3.AddItem "Kalimpong"
Combo3.AddItem "Kamarpukur"
Combo3.AddItem "Bishnupur"
Combo3.AddItem "Jalpaiguri"
Combo3.AddItem "Malda"
Combo3.AddItem "Cooch Behar"
Combo3.AddItem "Birbhum"
Combo3.AddItem "Barrackpore"
ElseIf (Combo2 = "Bihar") Then
Combo3.Clear
Combo3.AddItem ""
Combo3.AddItem ""
ElseIf (Combo2 = "Jharkhand") Then
Combo3.Clear
Combo3.AddItem "Ranchi"
'==================================NORTH-EAST
ElseIf (Combo2 = "Assam") Then
Combo3.Clear
Combo3.AddItem "Guwahati"
ElseIf (Combo2 = "Sikkim") Then
Combo3.Clear
Combo3.AddItem "Silk Route"
Combo3.AddItem "Gangtok"
Combo3.AddItem "Peling"
ElseIf (Combo2 = "Nagaland") Then
Combo3.Clear
Combo3.AddItem "Kohima"
Combo3.AddItem "Dimapur"
Combo3.AddItem "Mokokchung"
Combo3.AddItem "Wokha"
Combo3.AddItem "Mon"
Combo3.AddItem "Phek"
Combo3.AddItem "Kiphire"
ElseIf (Combo2 = "Manipur") Then
Combo3.Clear
Combo3.AddItem "Thoupal"
Combo3.AddItem "Chandel"
Combo3.AddItem "Senapati"
Combo3.AddItem "Tamenglong"
Combo3.AddItem "Churachandpur"
ElseIf (Combo2 = "Mizoram") Then
Combo3.Clear
Combo3.AddItem "Champhai"
Combo3.AddItem "Lunglei"
Combo3.AddItem "Serchhip"
Combo3.AddItem "Lawngtlai"
ElseIf (Combo2 = "Meghalay") Then
Combo3.Clear
Combo3.AddItem "Shilong"
ElseIf (Combo2 = "Arunachal Pradesh") Then
Combo3.Clear
Combo3.AddItem "Tawang"
'====================================NORTH
ElseIf (Combo2 = "Punjab") Then
Combo3.Clear
Combo3.AddItem "Amritswar"
ElseIf (Combo2 = "Jammu & Kashmir") Then
Combo3.Clear
Combo3.AddItem "Srinagar"
Combo3.AddItem "Leh"
Combo3.AddItem "Amarnath"
Combo3.AddItem "Jammu"
ElseIf (Combo2 = "Uttar Pradesh") Then
Combo3.Clear
Combo3.AddItem "Lucknow"
Combo3.AddItem "Nainital"
Combo3.AddItem "Haridwar"
Combo3.AddItem "Varanasi"
Combo3.AddItem "Agra"
ElseIf (Combo2 = "West Bengal") Then
Combo3.Clear
Combo3.AddItem "Chandigarh"
Combo3.AddItem "Gurgaon"
Combo3.AddItem "Kurukshetra"
Combo3.AddItem "Panipat"
Combo3.AddItem "Ambala"
Combo3.AddItem "Hisar"
Combo3.AddItem "Faridabad"
Combo3.AddItem "Karnal"
Combo3.AddItem "Pinjore"
Combo3.AddItem "Panchkula"
Combo3.AddItem "Morni Hills"
Combo3.AddItem " Sonipat"
ElseIf (Combo2 = "Himachal Pradesh") Then
Combo3.Clear
Combo3.AddItem "Simla"
Combo3.AddItem "Kinnor"
Combo3.AddItem "Dalhousi"
ElseIf (Combo2 = "Uttarakhand") Then
Combo3.Clear
Combo3.AddItem "Dehradun"
Combo3.AddItem "Munsiary"
Combo3.AddItem "Nainital"
ElseIf (Combo2 = "Delhi") Then
Combo3.Clear
Combo3.AddItem "Delhi"
'===============================WEST
ElseIf (Combo2 = "Rajasthan") Then
Combo3.Clear
Combo3.AddItem "Jodhpur"
Combo3.AddItem "Jaipur"
ElseIf (Combo2 = "Daman and Diu") Then
Combo3.Clear
Combo3.AddItem "Daman"
Combo3.AddItem "Diu"
ElseIf (Combo2 = "Maharashtra") Then
Combo3.Clear
Combo3.AddItem "Mumbai"
Combo3.AddItem "Pune"
Combo3.AddItem "Aurangabad"
ElseIf (Combo2 = "Gujarat") Then
Combo3.Clear
Combo3.AddItem "Ahmedabad"
ElseIf (Combo2 = "Goa") Then
Combo3.Clear
Combo3.AddItem "Panaji"
'=============================Middle
ElseIf (Combo2 = "Madhya Pradesh") Then
Combo3.Clear
Combo3.AddItem "Jabbalpur"
Combo3.AddItem "Khajuraho"
Combo3.AddItem "Indore"
ElseIf (Combo2 = "Chhattisgarh") Then
Combo3.Clear
Combo3.AddItem "Raipur"
'=============================SOUTH
ElseIf (Combo2 = "Karnataka") Then
Combo3.Clear
Combo3.AddItem "Bangaluru"
Combo3.AddItem "Mysore"
Combo3.AddItem "Hampi"
ElseIf (Combo2 = "Andhra Pradesh") Then
Combo3.Clear
Combo3.AddItem "Visakhapatnam"
Combo3.AddItem "Hydrabad"
Combo3.AddItem "Tirupati"
Combo3.AddItem "Vijayawada"
Combo3.AddItem "Nellore"
Combo3.AddItem "Chittoor"
Combo3.AddItem "Anantapur"
ElseIf (Combo2 = "Tamilnadu") Then
Combo3.Clear
Combo3.AddItem "Chennai"
Combo3.AddItem "Ooty"
ElseIf (Combo2 = "Lakshadweep") Then
Combo3.Clear
Combo3.AddItem "Kavaratti Island"
ElseIf (Combo2 = "Dadra and Nagar Haveli") Then
Combo3.Clear
Combo3.AddItem "Dadra"
Combo3.AddItem "Silvassa"
ElseIf (Combo2 = "Puducherry") Then
Combo3.Clear
Combo3.AddItem "Pondicherry"
Combo3.AddItem "Auroville"
ElseIf (Combo2 = "Andaman & Nicobar") Then
Combo3.Clear
Combo3.AddItem "Portblair"
ElseIf (Combo2 = "Kerala") Then
Combo3.Clear
Combo3.AddItem "Ernakulam"
End If
End If
End Sub

Private Sub Combo3_Click()
If (Combo3 <> "") Then
Adodc1.RecordSource = "select * from table1 where place_name='" & Combo3 & "'"
Adodc1.Refresh
End If
End Sub

Private Sub Combo4_Click()
If (Combo4 <> "") Then
Adodc1.RecordSource = "select * from table1 where TYPE_NAME='" & Combo4 & "'"
Adodc1.Refresh
End If
End Sub

Private Sub Command1_Click()
If (Text1.Text <> "" Or Text1.Text <> "Enter name of place to search") Then
Adodc1.RecordSource = "select * from table1 where place_name like '" & Text1 & "'"
Adodc1.Refresh
End If
End Sub

Private Sub Command2_Click()
Form1.Show
Form2.Hide

End Sub

Private Sub Form_Load()
cn.ConnectionString = Form1.str
cn.Open
rs.ActiveConnection = cn
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
rs.Source = "table1"
rs.Open
Adodc1.ConnectionString = Form1.str
Combo1.Clear
Combo2.Clear
Combo3.Clear
Combo4.Clear
'Adodc1.RecordSource = "select distinct region_name from table1"
'Adodc1.Refresh
Combo1.AddItem "North"
Combo1.AddItem "Eastern"
Combo1.AddItem "North-East"
Combo1.AddItem "Middle"
Combo1.AddItem "West"
Combo1.AddItem "South"
Combo4.AddItem "Pilgrimage"
Combo4.AddItem "Capital"
Combo4.AddItem "Beach"
Combo4.AddItem "Forest"
Combo4.AddItem "Hill Station"
Combo4.AddItem "Nature"
Combo4.AddItem "Historical"
Combo4.AddItem "Island"
'Adodc1.RecordSource = "select distinct region_name,state_name,place_name,TYPE_NAME from table1"
'Adodc1.Refresh
'With Adodc1.Recordset
'    Do Until .EOF
'        Combo1.AddItem ![region_name]
'        Combo2.AddItem ![state_name]
'        Combo3.AddItem ![place_name]
'        Combo4.AddItem ![TYPE_NAME]
'    .MoveNext
'    Loop
'End With
End Sub
Private Sub Text1_Click()
Text1.Text = ""
End Sub
