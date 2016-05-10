VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form5"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12510
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   7170
   ScaleWidth      =   12510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   30
      Top             =   5520
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form5.frx":2A73D
      Left            =   5160
      List            =   "Form5.frx":2A88B
      TabIndex        =   28
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text_Bad 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7560
      TabIndex        =   27
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text_Sat 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7560
      TabIndex        =   26
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text_G 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7560
      TabIndex        =   25
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text_VG 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7560
      TabIndex        =   24
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text_Excel 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7560
      TabIndex        =   23
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "VOTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   18
      Top             =   5520
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
      Height          =   4095
      Left            =   2160
      ScaleHeight     =   4035
      ScaleWidth      =   4755
      TabIndex        =   5
      Top             =   1080
      Width           =   4815
      Begin VB.PictureBox Picture2 
         Height          =   135
         Left            =   960
         ScaleHeight     =   135
         ScaleWidth      =   15
         TabIndex        =   6
         Top             =   360
         Width           =   15
      End
   End
   Begin VB.OptionButton Bad 
      Caption         =   "Bad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   4680
      Width           =   1455
   End
   Begin VB.OptionButton Sat 
      Caption         =   "Satisfactory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3960
      Width           =   1455
   End
   Begin VB.OptionButton G 
      Caption         =   "Good"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin VB.OptionButton VG 
      Caption         =   "Very Good"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.OptionButton Excel 
      Caption         =   "Excellent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   29
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "         How   Was"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   22
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   9840
      TabIndex        =   21
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   7560
      TabIndex        =   20
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   2520
      TabIndex        =   19
      Top             =   5520
      Width           =   4215
   End
   Begin VB.Label lbl_Total 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   7560
      TabIndex        =   17
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label l_b 
      Height          =   495
      Left            =   9840
      TabIndex        =   16
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label l_s 
      Height          =   495
      Left            =   9840
      TabIndex        =   15
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label l_g 
      Height          =   495
      Left            =   9840
      TabIndex        =   14
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label l_vg 
      Height          =   495
      Left            =   9840
      TabIndex        =   13
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label l_e 
      Height          =   495
      Left            =   9840
      TabIndex        =   12
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lbl_Bad 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   7560
      TabIndex        =   11
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label lbl_Sat 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   7560
      TabIndex        =   10
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lbl_G 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   7560
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lbl_VG 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   7560
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lbl_Excel 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   7560
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim total, Excel_total, VG_total, G_total, Sat_total, Bad_total As Integer
Dim Excel_percent, VG_percent, G_percent, Sat_percent, Bad_percent As Single
Dim done As Boolean
Dim str As String

Private Sub Combo1_Click()

Label5.Caption = Combo1 + " ?"
rs.Find ("PLACE_NAME= '" & Combo1 & "'")

Text_Excel.Text = rs.Fields(2).Value
Text_VG.Text = rs.Fields(3).Value
Text_G.Text = rs.Fields(4).Value
Text_Sat.Text = rs.Fields(5).Value
Text_Bad.Text = rs.Fields(6).Value

Excel_total = CDbl(Text_Excel.Text)
VG_total = CDbl(Text_VG.Text)
G_total = CDbl(Text_G.Text)
Sat_total = CDbl(Text_Sat.Text)
Bad_total = CDbl(Text_Bad.Text)
total = Excel_total + VG_total + G_total + Sat_total + Bad_total
Excel_percent = Excel_total / total
VG_percent = VG_total / total
G_percent = G_total / total
Sat_percent = Sat_total / total
Bad_percent = Bad_total / total
lbl_Total.Caption = total
l_e.Caption = Format(Excel_percent, "Percent")
l_vg.Caption = Format(VG_percent, "Percent")
l_g.Caption = Format(G_percent, "Percent")
l_s.Caption = Format(Sat_percent, "Percent")
l_b.Caption = Format(Bad_percent, "Percent")
Picture1.Cls
Picture1.Line (100, 750)-(3800 * Excel_percent, 950), vbGreen, BF
Picture1.Line (100, 1450)-(3800 * VG_percent, 1650), vbBlue, BF
Picture1.Line (100, 2150)-(3800 * G_percent, 2350), vbYellow, BF
Picture1.Line (100, 2850)-(3800 * Sat_percent, 3050), vbOrange, BF
Picture1.Line (100, 3550)-(3800 * Bad_percent, 3750), vbRed, BF
rs.MoveFirst
End Sub

Private Sub Command1_Click()
Picture1.Cls
If Excel.Value = True Then
Excel_total = Excel_total + 1
Text_Excel.Text = Excel_total
ElseIf VG.Value = True Then
VG_total = VG_total + 1
Text_VG.Text = VG_total
ElseIf G.Value = True Then
G_total = G_total + 1
Text_G.Text = G_total
ElseIf Sat.Value = True Then
Sat_total = Sat_total + 1
Text_Sat.Text = Sat_total
ElseIf Bad.Value = True Then
Bad_total = Bad_total + 1
Text_Bad.Text = Bad_total
End If
total = Excel_total + VG_total + G_total + Sat_total + Bad_total
Excel_percent = Excel_total / total
VG_percent = VG_total / total
G_percent = G_total / total
Sat_percent = Sat_total / total
Bad_percent = Bad_total / total
lbl_Total.Caption = total
l_e.Caption = Format(Excel_percent, "Percent")
l_vg.Caption = Format(VG_percent, "Percent")
l_g.Caption = Format(G_percent, "Percent")
l_s.Caption = Format(Sat_percent, "Percent")
l_b.Caption = Format(Bad_percent, "Percent")
Picture1.Line (100, 750)-(3800 * Excel_percent, 950), vbGreen, BF
Picture1.Line (100, 1450)-(3800 * VG_percent, 1650), vbBlue, BF
Picture1.Line (100, 2150)-(3800 * G_percent, 2350), vbYellow, BF
Picture1.Line (100, 2850)-(3800 * Sat_percent, 3050), vbOrange, BF
Picture1.Line (100, 3550)-(3800 * Bad_percent, 3750), vbRed, BF
rs.Find ("PLACE_NAME= '" & Combo1 & "'")

rs.Fields(2).Value = Text_Excel.Text
rs.Fields(3).Value = Text_VG.Text
rs.Fields(4).Value = Text_G.Text
rs.Fields(5).Value = Text_Sat.Text
rs.Fields(6).Value = Text_Bad.Text
rs.MoveFirst
End Sub

Private Sub Command2_Click()
Form1.Show
Form5.Hide
End Sub

Private Sub Form_Load()
cn.ConnectionString = Form1.str
cn.Open
rs.ActiveConnection = cn
rs.CursorType = adOpenDynamic
rs.LockType = adLockOptimistic
rs.Source = "table4"
rs.Open
End Sub

