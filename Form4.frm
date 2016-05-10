VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15345
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   8625
   ScaleWidth      =   15345
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form4.frx":851AE
      Left            =   9360
      List            =   "Form4.frx":852FC
      TabIndex        =   4
      ToolTipText     =   "Select Place Name"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   10320
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      TabIndex        =   1
      ToolTipText     =   "Click To Go Back To The Home Screen"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DISPLAY"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11760
      TabIndex        =   0
      ToolTipText     =   "Click To See Image"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "GALLERY"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   4080
      TabIndex        =   5
      Top             =   240
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   600
      ToolTipText     =   "Image"
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "SELECT ANY PLACE TO VIEW IMAGE"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   7215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Text1.Text = Combo1.Text
Text1.Text = Text1.Text + ".jpg"
End Sub

Private Sub Command1_Click()
Dim picname As String
On Error GoTo errmsg
picname = getimage(Text1.Text)
Image1.Visible = True
Image1.Picture = LoadPicture(picname)
Text1.Text = ""
'Text1.SetFocus
Exit Sub
errmsg:
Image1.Visible = False
MsgBox "PICTURE NOT FOUND"
Text1.Text = ""
'Text1.SetFocus
End Sub

Private Sub Command2_Click()
Form1.Show
Form4.Hide
End Sub

Private Function getimage(pname As String) As String
getimage = App.Path + "\pics\" + pname
End Function

