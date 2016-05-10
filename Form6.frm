VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8280
   ForeColor       =   &H00404000&
   LinkTopic       =   "Form6"
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   5760
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
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
      Left            =   6360
      TabIndex        =   37
      Top             =   4560
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1200
      ItemData        =   "Form6.frx":F772
      Left            =   3240
      List            =   "Form6.frx":F788
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   5880
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form6.frx":F7B9
      Left            =   5760
      List            =   "Form6.frx":F7BB
      TabIndex        =   0
      Text            =   "SELECT STATE"
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Dadra and Nagar Haveli"
      Height          =   495
      Left            =   0
      TabIndex        =   36
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "Delhi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   35
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label32 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Lakshadweep"
      Height          =   255
      Left            =   -120
      TabIndex        =   34
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Line Line12 
      Visible         =   0   'False
      X1              =   2040
      X2              =   2520
      Y1              =   4560
      Y2              =   4440
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Puducherry"
      Height          =   255
      Left            =   2520
      TabIndex        =   33
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Line Line11 
      Visible         =   0   'False
      X1              =   1080
      X2              =   480
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Goa"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Daman && Diu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   31
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Line Line10 
      X1              =   3360
      X2              =   3600
      Y1              =   2640
      Y2              =   3000
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   2760
      Top             =   0
      Width           =   1695
   End
   Begin VB.Line Line9 
      Visible         =   0   'False
      X1              =   4920
      X2              =   3960
      Y1              =   2880
      Y2              =   2520
   End
   Begin VB.Line Line8 
      Visible         =   0   'False
      X1              =   5040
      X2              =   4200
      Y1              =   2640
      Y2              =   2520
   End
   Begin VB.Line Line7 
      Visible         =   0   'False
      X1              =   4920
      X2              =   4320
      Y1              =   2400
      Y2              =   2280
   End
   Begin VB.Line Line6 
      Visible         =   0   'False
      X1              =   5040
      X2              =   3840
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line5 
      Visible         =   0   'False
      X1              =   4200
      X2              =   4920
      Y1              =   2040
      Y2              =   1920
   End
   Begin VB.Line Line4 
      Visible         =   0   'False
      X1              =   4920
      X2              =   4350
      Y1              =   1680
      Y2              =   2040
   End
   Begin VB.Line Line3 
      DrawMode        =   6  'Mask Pen Not
      Visible         =   0   'False
      X1              =   4320
      X2              =   4920
      Y1              =   1680
      Y2              =   1200
   End
   Begin VB.Line Line2 
      DrawMode        =   1  'Blackness
      Visible         =   0   'False
      X1              =   3480
      X2              =   3400
      Y1              =   1560
      Y2              =   1870
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Assam"
      Height          =   255
      Left            =   5040
      TabIndex        =   30
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Himachal pradesh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1200
      TabIndex        =   29
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Andaman && Nicobar islands"
      Height          =   495
      Left            =   3720
      TabIndex        =   28
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Tripura"
      Height          =   255
      Left            =   5040
      TabIndex        =   27
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Meghalaya"
      Height          =   255
      Left            =   5040
      TabIndex        =   26
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Jharkhand"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2520
      TabIndex        =   25
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   1560
      X2              =   840
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Nagaland"
      Height          =   255
      Left            =   5040
      TabIndex        =   24
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Arunachal Pradesh"
      Height          =   375
      Left            =   5040
      TabIndex        =   23
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Manipur"
      Height          =   255
      Left            =   5040
      TabIndex        =   22
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Mizoram"
      Height          =   255
      Left            =   5040
      TabIndex        =   21
      Top             =   2520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Tamil Nadu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   1620
      TabIndex        =   20
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Karnataka"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Kerela"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Andhra Pradesh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1680
      TabIndex        =   17
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Sikkim"
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Odisha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Chattisgarh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Bihar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "West Bengal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Gujarat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Maharashtra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Uttar Pradesh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Madhya Pradesh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Rajasthan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Uttarakhand"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Punjab"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Haryana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Jammu && Kashmir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
Label18.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label33.Visible = False
Label34.Visible = False
Line1.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
Line5.Visible = False
Line6.Visible = False
Line7.Visible = False
Line8.Visible = False
Line9.Visible = False
Line10.Visible = False
Line11.Visible = False
Line12.Visible = False
Select Case Combo1.Text
Case "Jammu & Kashmir"
Label1.Visible = True
Case "Odisha"
Label13.Visible = True
Case "West Bengal"
Label10.Visible = True
Line10.Visible = True
Case "Assam"
Label28.Visible = True
Line5.Visible = True
Case "Sikkim"
Label14.Visible = True
Line2.Visible = True
Case "Tripura"
Label25.Visible = True
Line9.Visible = True
Case "Bihar"
Label11.Visible = True
Case "Meghalay"
Label24.Visible = True
Line6.Visible = True
Case "Jharkhand"
Label23.Visible = True
Case "Delhi"
Label33.Visible = True
Case "Uttar Pradesh"
Label7.Visible = True
Case "UttaraKhand"
Label4.Visible = True
Case "Himachal Pradesh"
Label27.Visible = True
Case "Punjab"
Label3.Visible = True
Case "Rajsthan"
Label5.Visible = True
Case "Gujrat"
Label9.Visible = True
Case "Maharashtra"
Label8.Visible = True
Case "Madhyapradesh"
Label6.Visible = True
Case "Chhattisgarh"
Label12.Visible = True
Case "Andhrapradesh"
Label15.Visible = True
Case "Karnatak"
Label17.Visible = True
Case "Tamilnadu"
Label18.Visible = True
Case "Kerala"
Label16.Visible = True
Line1.Visible = True
Case "Andaman & Nicobar"
Label26.Visible = True
Case "Nagaland"
Label22.Visible = True
Line4.Visible = True
Case "Manipur"
Label20.Visible = True
Line7.Visible = True
Case "Mizoram"
Label19.Visible = True
Line8.Visible = True
Case "Haryana"
Label2.Visible = True
Case "Arunachal Pradesh"
Label21.Visible = True
Line3.Visible = True
Case "Goa"
Label30.Visible = True
Line11.Visible = True
Case "Puducherry"
Label31.Visible = True
Line12.Visible = True
Case "Dadra and Nagar Haveli"
Label34.Visible = True
Case "Daman and Diu"
Label29.Visible = True
Case "Lakshadweep"
Label32.Visible = True
End Select
End Sub

Private Sub Command1_Click()
Form1.Show
Form6.Hide
End Sub

Private Sub Form_Load()
Combo1.Clear
Combo1.AddItem "Odisha"
Combo1.AddItem "West Bengal"
Combo1.AddItem "Assam"
Combo1.AddItem "Sikkim"
Combo1.AddItem "Tripura"
Combo1.AddItem "Bihar"
Combo1.AddItem "Arunachal Pradesh"
Combo1.AddItem "Meghalay"
Combo1.AddItem "Jharkhand"
Combo1.AddItem "Jammu & Kashmir"
Combo1.AddItem "Delhi"
Combo1.AddItem "Uttar Pradesh"
Combo1.AddItem "UttaraKhand"
Combo1.AddItem "Himachal Pradesh"
Combo1.AddItem "Punjab"
Combo1.AddItem "Rajsthan"
Combo1.AddItem "Gujrat"
Combo1.AddItem "Maharashtra"
Combo1.AddItem "Goa"
Combo1.AddItem "Madhyapradesh"
Combo1.AddItem "Chhattisgarh"
Combo1.AddItem "Andhrapradesh"
Combo1.AddItem "Karnatak"
Combo1.AddItem "Tamilnadu"
Combo1.AddItem "Kerala"
Combo1.AddItem "Andaman & Nicobar"
Combo1.AddItem "Nagaland"
Combo1.AddItem "Manipur"
Combo1.AddItem "Mizoram"
Combo1.AddItem "Haryana"
Combo1.AddItem "Lakshadweep"
Combo1.AddItem "Dadra and Nagar Haveli"
Combo1.AddItem "Daman and Diu"
Combo1.AddItem "Puducherry"
Combo1.AddItem "Arunachal Pradesh"
End Sub

