VERSION 5.00
Begin VB.Form SignIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login Page"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6645
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "SignIn.frx":0000
   ScaleHeight     =   4200
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000B&
      Caption         =   "Here"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Click here for going back to the home page"
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000B&
      Caption         =   "Here"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Click here if you are not registered"
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Enter your password"
      Top             =   1745
      Width           =   2200
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Log In"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Press when finshed"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      ToolTipText     =   "Enter your login id"
      Top             =   1285
      Width           =   2200
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "To Continue As Guest User Click "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   7
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Did not Register yet? Click   "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   6
      Top             =   3720
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "To Register"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Index           =   2
      Left            =   4800
      TabIndex        =   5
      Top             =   3720
      Width           =   1815
   End
End
Attribute VB_Name = "SignIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strsql1 As String
Dim pid As Integer



Private Sub Command1_Click()
cn.ConnectionString = Form1.str
cn.Open
rs.ActiveConnection = cn
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
rs.Source = "Table2"
strsql1 = "select * from Table2 where UNAME LIKE '%" & Text1.Text & "%'"
rs.Open strsql1
With rs
    On Error GoTo eh1246:
    If !Password = Text2.Text Then
        Form1.username = Text1.Text
        
        Form1.login_status = 1
        Form1.admin_status = rs.Fields(3)
        rs.Close
        cn.Close
        Form1.d.Enabled = False
        Form1.G.Enabled = True
        SignIn.Hide
        Form1.Show
    End If
End With
If 1 = 2 Then
eh1246: Text1.Text = "Failed"
End If
End Sub


Private Sub Command2_Click()
Form1.Show
SignIn.Hide
End Sub

Private Sub Command3_Click()
AddUser.Show
SignIn.Hide

End Sub
