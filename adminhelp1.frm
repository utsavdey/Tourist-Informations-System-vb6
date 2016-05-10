VERSION 5.00
Begin VB.Form adminhelp1 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Remove Admin"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9780
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "adminhelp1.frx":0000
   ScaleHeight     =   6525
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "BACK"
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
      Left            =   3840
      TabIndex        =   4
      Top             =   4800
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove Admin"
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
      Index           =   1
      Left            =   5760
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make Admin"
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
      Left            =   2160
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Text            =   "Enter Username Here..."
      Top             =   1320
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   1
      Top             =   3480
      Width           =   6015
   End
End
Attribute VB_Name = "adminhelp1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p_id As Integer
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click(Index As Integer)

Label1.Caption = "Declarations"
If Text1.Text = "" Then
MsgBox "Enter Username!", vbCritical, "Username"
Else
rs.Find ("UNAME='" & Text1.Text & "'")
If rs.EOF = True Then
Label1.Caption = "Username Not Found"
rs.MoveFirst
Else
    With rs
        ck = !admin
    End With
rs.Update
If ck = 1 And Index = 0 Then
Label1.Caption = "User Already Admin"
ElseIf ck = 1 And Index = 1 Then
rs.Fields(3) = 0
Label1.Caption = "Admin Successfully Removed"
rs.Update
ElseIf ck = 0 And Index = 0 Then
rs.Fields(3) = 1
Label1.Caption = "Admin Successfully Made"
rs.Update
Else
Label1.Caption = "User Not An Admin"
End If
End If
End If


End Sub

Private Sub Command2_Click()

adminhelp1.Hide
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


