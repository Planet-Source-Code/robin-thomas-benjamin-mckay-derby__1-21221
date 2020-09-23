VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Rename Horses"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   4395
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   3840
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Rename Horses"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   3480
      Width           =   4695
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   4455
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   4455
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   4455
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   4455
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rename Horses"
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Label14.Caption = "Lane 1: " + Text1.Text
Form1.Label15.Caption = "Lane 2: " + Text2.Text
Form1.Label16.Caption = "Lane 3: " + Text3.Text
Form1.Label17.Caption = "Lane 4: " + Text4.Text
Form1.Label18.Caption = "Lane 5: " + Text5.Text
Form1.Label19.Caption = "Lane 6: " + Text6.Text
Form1.Label20.Caption = "Lane 7: " + Text7.Text
Form1.Label21.Caption = "Lane 8: " + Text8.Text
Form2.Label1.Caption = Text1
Form2.Label3.Caption = Text2
Form2.Label5.Caption = Text3
Form2.Label7.Caption = Text4
Form2.Label9.Caption = Text5
Form2.Label11.Caption = Text6
Form2.Label13.Caption = Text7
Form2.Label15.Caption = Text8
Unload Form4
End Sub

Private Sub Command2_Click()
Unload Form4
End Sub
Private Sub form_unload(cancel As Integer)
Form1.Text1.Text = Text1.Text
Form1.Text2.Text = Text2.Text
Form1.Text3.Text = Text3.Text
Form1.Text4.Text = Text4.Text
Form1.Text5.Text = Text5.Text
Form1.Text6.Text = Text6.Text
Form1.Text7.Text = Text7.Text
Form1.Text8.Text = Text8.Text
End Sub
Private Sub Form_Load()
Text1.Text = Form1.Text1.Text
Text2.Text = Form1.Text2.Text
Text3.Text = Form1.Text3.Text
Text4.Text = Form1.Text4.Text
Text5.Text = Form1.Text5.Text
Text6.Text = Form1.Text6.Text
Text7.Text = Form1.Text7.Text
Text8.Text = Form1.Text8.Text
End Sub
