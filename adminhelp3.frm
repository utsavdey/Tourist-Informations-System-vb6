VERSION 5.00
Begin VB.Form adminhelp3 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Places"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11415
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "adminhelp3.frx":0000
   ScaleHeight     =   7950
   ScaleWidth      =   11415
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text16 
      Height          =   495
      Left            =   5880
      TabIndex        =   24
      Text            =   "Food Cost..."
      Top             =   5640
      Width           =   3615
   End
   Begin VB.TextBox Text15 
      Height          =   495
      Left            =   1560
      TabIndex        =   23
      Text            =   "Hotel Cost..."
      Top             =   5640
      Width           =   3495
   End
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   5880
      TabIndex        =   22
      Text            =   "Tourist Attractions..."
      Top             =   4920
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1560
      TabIndex        =   21
      Text            =   "Best Time To Visit..."
      Top             =   4920
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5880
      TabIndex        =   20
      Text            =   "Resturants..."
      Top             =   4200
      Width           =   3615
   End
   Begin VB.TextBox Text14 
      Height          =   495
      Left            =   7320
      TabIndex        =   14
      Text            =   "Temperature in Summer..."
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   4440
      TabIndex        =   13
      Text            =   "Temperature in Monsoon..."
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Text            =   "Temperature in Winter..."
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   7320
      TabIndex        =   11
      Text            =   "Airport Nearby..."
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   4440
      TabIndex        =   10
      Text            =   "Bus Terminus Nearby..."
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   1560
      TabIndex        =   9
      Text            =   "Place Name..."
      Top             =   600
      Width           =   7935
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   5880
      TabIndex        =   8
      Text            =   "Region Name..."
      Top             =   1320
      Width           =   3615
   End
   Begin VB.TextBox Text13 
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Text            =   "Hotels..."
      Top             =   4200
      Width           =   3495
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Text            =   "State Name..."
      Top             =   1320
      Width           =   3615
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Text            =   "Type..."
      Top             =   2040
      Width           =   7935
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Text            =   "Railway Station Nearby..."
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remove"
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
      Left            =   5520
      TabIndex        =   3
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update"
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
      Left            =   3000
      TabIndex        =   2
      Top             =   6480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
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
      Left            =   480
      TabIndex        =   1
      Top             =   6480
      Width           =   2295
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
      Left            =   8040
      TabIndex        =   0
      Top             =   6480
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5040
      TabIndex        =   26
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9480
      TabIndex        =   25
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Star (*) Marked Are Required Fields While You ADD A Record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2160
      TabIndex        =   19
      Top             =   7320
      Width           =   7455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9480
      TabIndex        =   18
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9480
      TabIndex        =   17
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5160
      TabIndex        =   16
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9480
      TabIndex        =   15
      Top             =   1200
      Width           =   255
   End
End
Attribute VB_Name = "adminhelp3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p_id As Integer
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()

If Text7.Text = "Region Name..." Or Text7.Text = "" Then
MsgBox "Region Name is a Required Field", vbCritical, "Region Name"
ElseIf Text6.Text = "State Name..." Or Text6.Text = "" Then
MsgBox "State Name is a Required Field", vbCritical, "State Name"
ElseIf Text8.Text = "Place Name..." Or Text8.Text = "" Then
MsgBox "Place Name is a Required Field", vbCritical, "Place Name"
ElseIf Text5.Text = "Type..." Or Text5.Text = "" Then
MsgBox "Type is a Required Field", vbCritical, "Type"
ElseIf Text15.Text = "Hotel Cost..." Or Text15.Text = "" Then
MsgBox "Hotel Cost is a Required Field", vbCritical, "Hotel Cost"
ElseIf Text16.Text = "Food Cost..." Or Text16.Text = "" Then
MsgBox "Food Cost is a Required Field", vbCritical, "Food Cost"
Else
rs.Find ("P_NAME= '" & Text8.Text & "'")
If rs.EOF Then
rs.AddNew
rs.Fields(1) = Text7.Text
rs.Fields(2) = Text6.Text
rs.Fields(3) = Text8.Text
rs.Fields(4) = Text5.Text
rs.Fields(15) = Text15.Text
rs.Fields(16) = Text16.Text
If Text4.Text = "Railway Station Nearby..." Or Text4.Text = "" Then
rs.Fields(5) = ""
Else
rs.Fields(5) = Text4.Text
End If
If Text10.Text = "Bus Terminus Nearby..." Or Text10.Text = "" Then
rs.Fields(6) = ""
Else
rs.Fields(6) = Text10.Text
End If
If Text9.Text = "Airport Nearby..." Or Text9.Text = "" Then
rs.Fields(7) = ""
Else
rs.Fields(7) = Text9.Text
End If
If Text3.Text = "Temperature in Winter..." Or Text3.Text = "" Then
rs.Fields(8) = ""
Else
rs.Fields(8) = Text3.Text
End If
If Text11.Text = "Temperature in Monsoon..." Or Text11.Text = "" Then
rs.Fields(9) = ""
Else
rs.Fields(9) = Text11.Text
End If
If Text14.Text = "Temperature in Summer..." Or Text14.Text = "" Then
rs.Fields(10) = ""
Else
rs.Fields(10) = Text14.Text
End If
If Text13.Text = "Hotels..." Or Text13.Text = "" Then
rs.Fields(11) = ""
Else
rs.Fields(11) = Text13.Text
End If
If Text1.Text = "Resturants..." Or Text1.Text = "" Then
rs.Fields(12) = ""
Else
rs.Fields(12) = Text1.Text
End If
If Text2.Text = "Best Time To Visit..." Or Text2.Text = "" Then
rs.Fields(13) = ""
Else
rs.Fields(13) = Text2.Text
End If
If Text12.Text = "Tourist Attractions..." Or Text12.Text = "" Then
rs.Fields(14) = ""
Else
rs.Fields(14) = Text12.Text
End If
rs.Update
MsgBox "Successfully Added!", vbInformation, "Success"
Else
MsgBox "Enter Different Place Name", vbCritical, "Place Name Already Exists"
rs.MoveFirst
End If
End If

End Sub

Private Sub Command2_Click()

adminhelp3.Hide
Adminp.Show

End Sub

Private Sub Command3_Click()

If Text8.Text = "" Or Text8.Text = "Place Name..." Then
MsgBox "Place Name is a Required Field", vbCritical, "Place Name"
Else
rs.Find ("P_NAME= '" & Text8.Text & "'")
If rs.EOF Then
MsgBox "Place Name Not Found!", vbCritical, "Not Found"
rs.MoveFirst
Else
rs.Update
If Text7.Text = "Region Name..." Or Text7.Text = "" Then
Else
rs.Fields(1) = Text7.Text
End If
If Text6.Text = "State Name..." Or Text6.Text = "" Then
Else
rs.Fields(2) = Text6.Text
End If
If Text5.Text = "Type..." Or Text5.Text = "" Then
Else
rs.Fields(4) = Text5.Text
End If
If Text4.Text = "Railway Station Nearby..." Or Text4.Text = "" Then
Else
rs.Fields(5) = Text4.Text
End If
If Text10.Text = "Bus Terminus Nearby..." Or Text10.Text = "" Then
Else
rs.Fields(6) = Text10.Text
End If
If Text9.Text = "Airport Nearby..." Or Text9.Text = "" Then
Else
rs.Fields(7) = Text9.Text
End If
If Text3.Text = "Temperature in Winter..." Or Text3.Text = "" Then
Else
rs.Fields(8) = Text3.Text
End If
If Text11.Text = "Temperature in Monsoon..." Or Text11.Text = "" Then
Else
rs.Fields(9) = Text11.Text
End If
If Text14.Text = "Temperature in Summer..." Or Text14.Text = "" Then
Else
rs.Fields(10) = Text14.Text
End If
If Text13.Text = "Hotels..." Or Text13.Text = "" Then
Else
rs.Fields(11) = Text13.Text
End If
If Text1.Text = "Resturants..." Or Text1.Text = "" Then
Else
rs.Fields(12) = Text1.Text
End If
If Text2.Text = "Best Time To Visit..." Or Text2.Text = "" Then
Else
rs.Fields(13) = Text2.Text
End If
If Text12.Text = "Tourist Attractions..." Or Text12.Text = "" Then
Else
rs.Fields(14) = Text12.Text
End If
If Text12.Text = "Tourist Attractions..." Or Text12.Text = "" Then
Else
If Text12.Text = "Tourist Attractions..." Or Text12.Text = "" Then
Else
If Text15.Text = "Hotel Cost..." Or Text15.Text = "" Then
Else
rs.Fields(15) = Text15.Text
End If
If Text16.Text = "Food Cost..." Or Text16.Text = "" Then
Else
rs.Fields(16) = Text16.Text
End If
rs.Update
MsgBox "Updated Successfully!", vbInformation, " Success"
End If
rs.MoveFirst
End If
End If

End Sub

Private Sub Command4_Click()

If Text8.Text = "Place Name..." Or Text8.Text = "" Then
MsgBox "Place Name is a Required Field", vbCritical, "Place ID"
Else
rs.Find ("P_NAME= '" & Text8.Text & "'")
If rs.EOF Then
MsgBox "Place Name Not Found!", vbCritical, "Not Found"
rs.MoveFirst
Else
rs.Delete
MsgBox "Deletion Done Successfully", vbInformation, "Remove"
rs.MoveFirst
rs.Update
End If
End If

End Sub

Private Sub Form_Load()
cn.ConnectionString = Form1.str
cn.Open
rs.ActiveConnection = cn
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
rs.Source = "Table1"
rs.Open

End Sub




Private Sub Text14_Click()

Text14.Text = ""

End Sub

Private Sub Text14_LostFocus()
If Text14.Text = "" Then
Text14.Text = "Temperature in Summer..."
End If
End Sub

Private Sub Text13_Click()

Text13.Text = ""

End Sub

Private Sub Text13_LostFocus()
If Text13.Text = "" Then
Text13.Text = "Hotels..."
End If
End Sub
Private Sub Text1_Click()

Text1.Text = ""

End Sub

Private Sub Text1_LostFocus()
If Text13.Text = "" Then
Text13.Text = "Resturants..."
End If
End Sub



Private Sub Text2_Click()

Text2.Text = ""

End Sub

Private Sub Text2_LostFocus()
If Text2.Text = "" Then
Text2.Text = "Best Time To Visit..."
End If
End Sub
Private Sub Text16_Click()

Text16.Text = ""

End Sub

Private Sub Text16_LostFocus()
If Text16.Text = "" Then
Text16.Text = "Food Cost..."
End If
End Sub
Private Sub Text15_Click()

Text15.Text = ""

End Sub

Private Sub Text15_LostFocus()
If Text15.Text = "" Then
Text15.Text = "Hotel Cost..."
End If
End Sub

Private Sub Text12_Click()

Text12.Text = ""

End Sub

Private Sub Text12_LostFocus()
If Text12.Text = "" Then
Text12.Text = "Tourist Attractions..."
End If
End Sub

Private Sub Text3_Click()

Text3.Text = ""

End Sub

Private Sub Text3_LostFocus()
If Text3.Text = "" Then
Text3.Text = "Temperature in Winter..."
End If
End Sub

Private Sub Text11_Click()

Text11.Text = ""

End Sub

Private Sub Text11_LostFocus()
If Text11.Text = "" Then
Text11.Text = "Temperature in Monsoon..."
End If
End Sub
Private Sub Text10_Click()

Text10.Text = ""

End Sub

Private Sub Text10_LostFocus()
If Text10.Text = "" Then
Text10.Text = "Bus Terminus Nearby..."
End If
End Sub

Private Sub Text9_Click()

Text9.Text = ""

End Sub

Private Sub Text9_LostFocus()
If Text9.Text = "" Then
Text9.Text = "Airport Nearby..."
End If
End Sub
Private Sub Text5_Click()

Text5.Text = ""

End Sub

Private Sub Text5_LostFocus()
If Text5.Text = "" Then
Text5.Text = "Type..."
End If
End Sub

Private Sub Text4_Click()

Text4.Text = ""

End Sub

Private Sub Text4_LostFocus()
If Text4.Text = "" Then
Text4.Text = "Railway Station Nearby..."
End If
End Sub

Private Sub Text6_Click()

Text6.Text = ""

End Sub

Private Sub Text6_LostFocus()
If Text6.Text = "" Then
Text6.Text = "State Name..."
End If
End Sub

Private Sub Text8_Click()

Text8.Text = ""

End Sub

Private Sub Text8_LostFocus()
If Text8.Text = "" Then
Text8.Text = "Place Name..."
End If
End Sub


Private Sub Text7_Click()

Text7.Text = ""

End Sub

Private Sub Text7_LostFocus()
If Text7.Text = "" Then
Text7.Text = "Region Name..."
End If
End Sub



