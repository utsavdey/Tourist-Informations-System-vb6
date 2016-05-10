VERSION 5.00
Begin VB.Form adminhelp2 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remove An User"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7170
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "adimnhelp2.frx":0000
   ScaleHeight     =   3390
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Text            =   "Enter Username Here..."
      Top             =   960
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   1920
      Width           =   2175
   End
End
Attribute VB_Name = "adminhelp2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p_id As Integer
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click(Index As Integer)

If Text1.Text = "" Then
MsgBox "Enter Username!", vbCritical, "Username"
Else
rs.Find ("UNAME='" & Text1.Text & "'")
If rs.EOF = True Then
MsgBox "Username Not Found!", vbCritical, "Username"
rs.MoveFirst
Else
rs.Delete
MsgBox "Deletiton Performed Perfectly", vbInformation, "Delete User"
rs.Update
End If
End If


End Sub

Private Sub Command2_Click()

adminhelp2.Hide
Adminp.Show

End Sub

Private Sub Form_Load()
cn.ConnectionString = Form1.str
cn.Open
rs.ActiveConnection = cn
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
rs.Source = "Table2"
rs.Open

End Sub

Private Sub Text1_Click()

Text1.Text = ""

End Sub

Private Sub Text1_LostFocus()
If Text1.Text = "" Then
Text1.Text = "Enter Username Here..."
End If
End Sub



