VERSION 5.00
Begin VB.Form AddUser 
   BackColor       =   &H00008000&
   Caption         =   "Add User"
   ClientHeight    =   8445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7590
   LinkTopic       =   "Form3"
   ScaleHeight     =   8445
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
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
      Left            =   3960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Click here if you are already registered"
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
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
      Left            =   5040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Click here for going back to the home page"
      Top             =   7320
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   600
      Picture         =   "AddUser.frx":0000
      ScaleHeight     =   1905
      ScaleWidth      =   6345
      TabIndex        =   5
      Top             =   240
      Width           =   6375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Eras Demi ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   3
      ToolTipText     =   "Click to register"
      Top             =   6480
      Width           =   5535
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Text            =   "Re enter Password"
      Top             =   5400
      Width           =   5295
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   1080
      TabIndex        =   1
      Text            =   "Password"
      Top             =   3720
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Text            =   "User Name"
      Top             =   2520
      Width           =   5295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "To Log In"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   5160
      TabIndex        =   10
      Top             =   7800
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Already A User? Click   "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   9
      Top             =   7800
      Width           =   3735
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   8
      Top             =   7320
      Width           =   3735
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   2880
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Password Strength:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      ToolTipText     =   "Your password should be >=4 characters with at least 1 digiit and 1 alphabet"
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   3855
      Left            =   960
      Top             =   2400
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2175
      Left            =   480
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   15
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   15
   End
End
Attribute VB_Name = "AddUser"
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
    If Text2.Text = Text3.Text Then
    If .EOF Then
     GoTo eh1246
    ElseIf Text1.Text = !UNAME Then
        MsgBox ("Username Taken")
            Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
   
   End If
   
    If 1 = 2 Then
    
eh1246: rs.Close
        cn.Close
        cn.Open
        rs.ActiveConnection = cn
        rs.CursorType = adOpenDynamic
        rs.LockType = adLockOptimistic
        rs.Source = "Table2"
        rs.Open
        
             rs.AddNew
             rs.Fields("UNAME").Value = Text1.Text
             rs.Fields("PASSWORD").Value = Text2.Text
            rs.Fields("ADMIN").Value = 0
            rs.Update
             MsgBox ("Success")
             AddUser.Hide
             SignIn.Show
             
         Else
            MsgBox ("Please Reenter the Password")
            Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
        End If
    End If
    End With
    rs.Close
    cn.Close
End Sub

Private Sub Command2_Click()
Form1.Show
AddUser.Hide

End Sub

Private Sub Command3_Click()
SignIn.Show
AddUser.Hide

End Sub

Private Sub Form_Load()
Shape3.Visible = False

End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub
Private Sub Text2_Change()
Text2.PasswordChar = "*"
Dim pass, ckpass As String
pass = Text2.Text
Dim plen, i, dig, letr, ck As Integer
i = 1
dig = 0
letr = 0
plen = Len(pass)
While i <= plen
ckpass = Mid(pass, i, 1)
ck = Asc(ckpass)
If (ck > 47 And ck < 58) Then
dig = 1
ElseIf (ck > 64 And ck < 91) Or (ck > 96 And ck < 123) Then
letr = 1
End If
Shape3.Visible = True

If (plen < 4 And (dig = 0 Or letr = 0)) Then
Shape3.Width = 135
Shape3.FillColor = &HC0&
Label3.Caption = "Very Weak"
Label3.ForeColor = &HC0&
Command1.Enabled = False

ElseIf (plen < 4 Or dig = 0 Or letr = 0) Then
Shape3.Width = 375
Shape3.FillColor = &H8080FF
Label3.Caption = "Weak"
Command1.Enabled = False

ElseIf (plen > 4 And plen < 7 And dig = 1 And letr = 1) Then
Shape3.Width = 1415
Shape3.FillColor = &H80FF80
Label3.Caption = "Storng"
Label3.ForeColor = &H80FF80
Command1.Enabled = True

ElseIf (plen = 4 And dig = 1 And letr = 1) Then
Shape3.Width = 900
Shape3.FillColor = &HFFFF&
Label3.Caption = "Normal"
Label3.ForeColor = &HFFFF&
Command1.Enabled = True

Else
Shape3.Width = 1935
Shape3.FillColor = &H8000&
Label3.Caption = "Very Strong"
Label3.ForeColor = &H8000&
Command1.Enabled = True
End If
i = i + 1
Wend





End Sub


Private Sub Text3_Click()
Text3.Text = ""
End Sub
Private Sub Text3_Change()
Text3.PasswordChar = "*"
End Sub
Private Sub Text1_LostFocus()
If Text1.Text = "" Then
    Text1.Text = "Username"
End If
End Sub
Private Sub Text2_LostFocus()
If Text2.Text = "" Then
    Text2.PasswordChar = ""
    Text2.Text = "Password"
End If
End Sub
Private Sub Text3_LostFocus()
If Text3.Text = "" Then
    Text3.PasswordChar = ""
    Text3.Text = "Re enter Password"
End If
End Sub

