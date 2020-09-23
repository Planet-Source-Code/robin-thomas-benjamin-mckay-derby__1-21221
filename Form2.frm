VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Horse Statistics"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6615
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7035
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   2520
      TabIndex        =   16
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Horse Victories:"
      Height          =   495
      Left            =   4560
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "0"
      Height          =   735
      Left            =   4560
      TabIndex        =   15
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   870
      Index           =   7
      Left            =   3720
      Picture         =   "Form2.frx":030A
      Top             =   5640
      Width           =   555
   End
   Begin VB.Label Label15 
      Caption         =   "Powerful Pete"
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   5760
      Width           =   3615
   End
   Begin VB.Label Label14 
      Caption         =   "0"
      Height          =   735
      Left            =   4560
      TabIndex        =   13
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   870
      Index           =   6
      Left            =   3720
      Picture         =   "Form2.frx":1CAC
      Top             =   4920
      Width           =   555
   End
   Begin VB.Label Label13 
      Caption         =   "Knotting Ken"
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   5040
      Width           =   3615
   End
   Begin VB.Label Label12 
      Caption         =   "0"
      Height          =   735
      Left            =   4560
      TabIndex        =   11
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   870
      Index           =   5
      Left            =   3720
      Picture         =   "Form2.frx":364E
      Top             =   4200
      Width           =   555
   End
   Begin VB.Label Label11 
      Caption         =   "Zapping Zara"
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   4320
      Width           =   3615
   End
   Begin VB.Label Label10 
      Caption         =   "0"
      Height          =   735
      Left            =   4560
      TabIndex        =   9
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   870
      Index           =   4
      Left            =   3720
      Picture         =   "Form2.frx":4FF0
      Top             =   3480
      Width           =   555
   End
   Begin VB.Label Label9 
      Caption         =   "Innovative Ian"
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   3600
      Width           =   3495
   End
   Begin VB.Label Label8 
      Caption         =   "0"
      Height          =   735
      Left            =   4560
      TabIndex        =   7
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   870
      Index           =   3
      Left            =   3720
      Picture         =   "Form2.frx":6992
      Top             =   2760
      Width           =   555
   End
   Begin VB.Label Label7 
      Caption         =   "Thunder Thorn"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   735
      Left            =   4560
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   870
      Index           =   2
      Left            =   3720
      Picture         =   "Form2.frx":8334
      Top             =   2040
      Width           =   555
   End
   Begin VB.Label Label5 
      Caption         =   "Pumping Paula"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   735
      Left            =   4560
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   870
      Index           =   0
      Left            =   3720
      Picture         =   "Form2.frx":9CD6
      Top             =   1320
      Width           =   555
   End
   Begin VB.Label Label3 
      Caption         =   "Strong Seth"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   870
      Index           =   1
      Left            =   3720
      Picture         =   "Form2.frx":B678
      Top             =   600
      Width           =   555
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   735
      Left            =   4560
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Riot Robin"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   3375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form2
Set Form2 = Nothing
End Sub

Private Sub Form_Load()
Label2.Caption = Form1.Label22.Caption
Label4.Caption = Form1.Label23.Caption
Label6.Caption = Form1.Label24.Caption
Label8.Caption = Form1.Label25.Caption
Label10.Caption = Form1.Label26.Caption
Label12.Caption = Form1.Label27.Caption
Label14.Caption = Form1.Label28.Caption
Label16.Caption = Form1.Label29.Caption
End Sub

