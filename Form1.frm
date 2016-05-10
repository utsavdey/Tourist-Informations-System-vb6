VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tourist Information System"
   ClientHeight    =   3885
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3885
   ScaleWidth      =   10335
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   998
      Left            =   3240
      Top             =   3120
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "INDIA"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   975
      Left            =   6240
      TabIndex        =   4
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   4920
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1095
      Left            =   1680
      TabIndex        =   2
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   855
      Left            =   7200
      TabIndex        =   1
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Menu R 
      Caption         =   "&Registration"
      Begin VB.Menu b 
         Caption         =   "&Add User"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu d 
         Caption         =   "Log &In"
      End
      Begin VB.Menu g 
         Caption         =   "Log &Out"
      End
   End
   Begin VB.Menu C 
      Caption         =   "Admin &Control"
   End
   Begin VB.Menu U 
      Caption         =   "Tour G&uide"
      Begin VB.Menu A 
         Caption         =   "&Search"
      End
      Begin VB.Menu poll 
         Caption         =   "&Gallery"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu packages 
         Caption         =   "&Packages"
      End
   End
   Begin VB.Menu X 
      Caption         =   "E&xtras"
      Begin VB.Menu opinion 
         Caption         =   "P&oll"
      End
      Begin VB.Menu map 
         Caption         =   "Map"
      End
      Begin VB.Menu Dash4 
         Caption         =   "-"
      End
      Begin VB.Menu EX 
         Caption         =   "Exi&t"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Hel&p"
      Begin VB.Menu About1 
         Caption         =   "Abo&ut"
      End
      Begin VB.Menu dash5 
         Caption         =   "-"
      End
      Begin VB.Menu CR 
         Caption         =   "Cred&its"
      End
      Begin VB.Menu Us 
         Caption         =   "Contact Us"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public username, str As String
Public admin_status As Integer
Public login_status As Integer
Private Sub a_Click()
Form2.Show
Form1.Hide

End Sub


Private Sub About1_Click()
About.Show
Form1.Hide
End Sub

Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub b_Click()
AddUser.Show
Form1.Hide
End Sub

Private Sub c_Click()

If Form1.admin_status = 0 Then
MsgBox "You Are Not An Admin!", vbCritical, "Admin Status"
Else
Adminp.Show
Form1.Hide
 End If

End Sub



Private Sub CR_Click()
Credits.Timer1.Enabled = True
Credits.Show
Form1.Hide
End Sub

Private Sub d_Click()

SignIn.Show
opinion.Enabled = True
C.Enabled = True
Form1.Hide
End Sub

Private Sub EX_Click()
End
End Sub

Private Sub Form_Load()
Label1.Caption = Time
Label2.Caption = Date
C.Enabled = False
g.Enabled = False
opinion.Enabled = False
login_status = 0
username = ""
d.Enabled = True
str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=I:\Tourist-Information-System Actual\Database4_3_2.mdb;Persist Security Info=False"
    
End Sub

Private Sub g_Click()
    
    MsgBox "Successfully Logged Out!", vbInformation, "Log Out"
    login_status = 0
    username = ""
    d.Enabled = True
    g.Enabled = False
    C.Enabled = False
    
End Sub






Private Sub map_Click()
Form6.Show
Form1.Hide

End Sub

Private Sub opinion_Click()
Form5.Show
Form1.Hide

End Sub

Private Sub P_Click()
Help.Show
Form1.Hide
End Sub

Private Sub packages_Click()
Form3.Show
Form1.Hide
End Sub

Private Sub Poll_Click()
Form4.Show
Form1.Hide
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Time
Label2.Caption = Date
End Sub

Private Sub Us_Click()
contact.Show
Form1.Hide

End Sub
