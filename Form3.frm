VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Place Your Bets "
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6600
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4380
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      TabIndex        =   20
      Top             =   3600
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   2760
      TabIndex        =   16
      Text            =   "0"
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2760
      TabIndex        =   15
      Text            =   "0"
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2760
      TabIndex        =   14
      Text            =   "0"
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2760
      TabIndex        =   13
      Text            =   "0"
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2760
      TabIndex        =   12
      Text            =   "0"
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2760
      TabIndex        =   11
      Text            =   "0"
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2760
      TabIndex        =   10
      Text            =   "0"
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2760
      TabIndex        =   9
      Text            =   "0"
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label19 
      Caption         =   "3:1"
      Height          =   495
      Left            =   5520
      TabIndex        =   28
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label18 
      Caption         =   "7:1"
      Height          =   495
      Left            =   5520
      TabIndex        =   27
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label17 
      Caption         =   "2:1"
      Height          =   495
      Left            =   5520
      TabIndex        =   26
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "8:1"
      Height          =   495
      Left            =   5520
      TabIndex        =   25
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "50:1"
      Height          =   495
      Left            =   5520
      TabIndex        =   24
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "4:1"
      Height          =   495
      Left            =   5520
      TabIndex        =   23
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "13:1"
      Height          =   495
      Left            =   5520
      TabIndex        =   22
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "5:1"
      Height          =   495
      Left            =   5520
      TabIndex        =   21
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Here are your total Winnings:"
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   3360
      Width           =   5415
   End
   Begin VB.Label Label10 
      Caption         =   "Please place your bet to the pound, please: (no dots or spaces"
      Height          =   495
      Left            =   2760
      TabIndex        =   17
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Powerful Pete"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Knotting Ken"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Zapping Zora"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Innovative Ian"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Thunder Thorn"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pumping Paula"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Strong Seth"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Riot Robin"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Horse Name:"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Label9 = Text1
Form1.Label10 = Text2
Form1.Label11 = Text3
Form1.Label12 = Text4
Form1.Label13 = Text5
Form1.Label30 = Text6
Form1.Label31 = Text7
Form1.Label32 = Text8
Unload Form3
End Sub

Private Sub Form_Load()
Text1 = Form1.Label9.Caption
Text2 = Form1.Label10.Caption
Text3 = Form1.Label11.Caption
Text4 = Form1.Label12.Caption
Text5 = Form1.Label13.Caption
Text6 = Form1.Label30.Caption
Text7 = Form1.Label31.Caption
Text8 = Form1.Label32.Caption
Text9 = Form1.Label33.Caption
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
Form1.Label33.Caption = Text9.Text
End Sub
