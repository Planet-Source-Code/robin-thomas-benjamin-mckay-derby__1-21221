VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "The Derby- Written By Robin McKay"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8160
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   4080
      TabIndex        =   40
      Text            =   "Text8"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   3720
      TabIndex        =   39
      Text            =   "Text7"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   2640
      TabIndex        =   38
      Text            =   "Text6"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   2520
      TabIndex        =   37
      Text            =   "Text5"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3600
      TabIndex        =   36
      Text            =   "Text4"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2040
      TabIndex        =   35
      Text            =   "Text3"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5040
      TabIndex        =   34
      Text            =   "Text2"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1320
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   720
      Top             =   360
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   5160
      Top             =   3000
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Label Label33 
      Caption         =   "0"
      Height          =   495
      Left            =   3600
      TabIndex        =   32
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label32 
      Caption         =   "0"
      Height          =   495
      Left            =   2280
      TabIndex        =   31
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label31 
      Caption         =   "0"
      Height          =   495
      Left            =   4560
      TabIndex        =   30
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label30 
      Caption         =   "0"
      Height          =   495
      Left            =   1560
      TabIndex        =   29
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label29 
      Caption         =   "0"
      Height          =   495
      Left            =   1440
      TabIndex        =   28
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label28 
      Caption         =   "0"
      Height          =   495
      Left            =   240
      TabIndex        =   27
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label27 
      Caption         =   "0"
      Height          =   495
      Left            =   4440
      TabIndex        =   26
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label26 
      Caption         =   "0"
      Height          =   495
      Left            =   2160
      TabIndex        =   25
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label25 
      Caption         =   "0"
      Height          =   495
      Left            =   3840
      TabIndex        =   24
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label24 
      Caption         =   "0"
      Height          =   495
      Left            =   720
      TabIndex        =   23
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label23 
      Caption         =   "0"
      Height          =   495
      Left            =   5280
      TabIndex        =   22
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label22 
      Caption         =   "0"
      Height          =   495
      Left            =   3000
      TabIndex        =   21
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Lane 8: Powerful Pete"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3240
      TabIndex        =   20
      Top             =   6000
      Width           =   2775
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Lane 7: Knotting Ken"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3240
      TabIndex        =   19
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Lane 6: Zapping Zara"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3240
      TabIndex        =   18
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Lane 5: Innovative Ian"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3240
      TabIndex        =   17
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Lane 4: Thunder Thorn"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3240
      TabIndex        =   16
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Lane 3: Pumping Paula"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3240
      TabIndex        =   15
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Lane 2: Strong Seth"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3240
      TabIndex        =   14
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Lane 1: Riot Robin"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3240
      TabIndex        =   13
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label13 
      Caption         =   "0"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "0"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "0"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "0"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "0"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "0"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "0"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FF80FF&
      X1              =   360
      X2              =   7320
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FF8080&
      X1              =   360
      X2              =   7320
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFF80&
      X1              =   360
      X2              =   7320
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0080FF80&
      X1              =   360
      X2              =   7320
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FFFF&
      X1              =   360
      X2              =   7320
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080C0FF&
      X1              =   360
      X2              =   7320
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H008080FF&
      X1              =   360
      X2              =   7320
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   360
      X2              =   7320
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image8 
      Height          =   870
      Left            =   7320
      Picture         =   "Form1.frx":030A
      Top             =   5880
      Width           =   555
   End
   Begin VB.Image Image7 
      Height          =   870
      Left            =   7320
      Picture         =   "Form1.frx":1CAC
      Top             =   5040
      Width           =   555
   End
   Begin VB.Image Image6 
      Height          =   870
      Left            =   7320
      Picture         =   "Form1.frx":364E
      Top             =   4200
      Width           =   555
   End
   Begin VB.Image Image5 
      Height          =   870
      Left            =   7320
      Picture         =   "Form1.frx":4FF0
      Top             =   3360
      Width           =   555
   End
   Begin VB.Image Image4 
      Height          =   870
      Left            =   7320
      Picture         =   "Form1.frx":6992
      Top             =   2520
      Width           =   555
   End
   Begin VB.Image Image3 
      Height          =   870
      Left            =   7320
      Picture         =   "Form1.frx":8334
      Top             =   1680
      Width           =   555
   End
   Begin VB.Image Image2 
      Height          =   870
      Left            =   7320
      Picture         =   "Form1.frx":9CD6
      Top             =   840
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   7320
      Picture         =   "Form1.frx":B678
      Top             =   0
      Width           =   555
   End
   Begin VB.Menu Game 
      Caption         =   "&Game"
      Begin VB.Menu BRACE 
         Caption         =   "&Begin Race..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu HSTATS 
         Caption         =   "&Horse Statistics"
         Shortcut        =   {F2}
      End
      Begin VB.Menu HBET 
         Caption         =   "&Horse Bet"
         Shortcut        =   {F3}
      End
      Begin VB.Menu border1 
         Caption         =   "-"
      End
      Begin VB.Menu ADERBY 
         Caption         =   "&About Derby"
         Shortcut        =   {F4}
      End
      Begin VB.Menu BORDER2 
         Caption         =   "-"
      End
      Begin VB.Menu CMVEM 
         Caption         =   "&Contact Me Via E-Mail"
         Shortcut        =   {F5}
      End
      Begin VB.Menu BORDER3 
         Caption         =   "-"
      End
      Begin VB.Menu RENAMEHORSES 
         Caption         =   "&Rename Horses..."
         Shortcut        =   {F6}
      End
      Begin VB.Menu rahb 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
         Shortcut        =   {F7}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Welcome to the Derby 1.0

' The aim of the game is to see which horse wins
Private Sub time()
Timer1.Enabled = False ' Disables the Timer
End Sub

Private Sub ADERBY_Click()
MsgBox "Thanks for using Derby 1.0. It means a lot to me. This is just a fun game and an innovation on Racing Command Buttons 1.0. No further plans have been made to make a sequel to Racing Command Buttons. However, if you would like to see a silly sequel, please let me know." + vbCrLf + vbCrLf + "Thanks for using The Derby. This is a more advanced program than my more recent submission. Comments are most welcome. Please click Contact Author to open up your email program. I appreciate that the Internet IS expensive and not everyone has access. Here is my home address:" + vbCrLf + vbCrLf + "ROBIN MCKAY" + vbCrLf + "20 CONISTON ROAD" + vbCrLf + "KEMPSHOTT" + vbCrLf + "BASINGSTOKE" + vbCrLf + "HAMPSHIRE" + vbCrLf + "RG22 5HS" + vbCrLf + "UNITED KINGDOM" + vbCrLf + vbCrLf + "Happy Coding, everyone!" + vbCrLf + vbCrLf + "ROBIN MCKAY - THE DERBY AUTHOR", vbInformation ' Displays a complicated message box
End Sub

Private Sub BRACE_Click()
Timer1.Enabled = True ' Enables the timer and starts the race (See Timer1 for more details)
End Sub

Private Sub CMVEM_Click()
Shell ("start mailto:ian@imckay.fsnet.co.uk"), vbHide ' Opens up the default email program and inserts my email address
End Sub

Private Sub Exit_Click()
Unload Me ' These two lines stop the program
End
End Sub

Private Sub form_unload(cancel As Integer)
On Error Resume Next ' When the form is unloaded, all other forms are unloaded and freed from memory
Unload Form1
Unload Form2
Unload Form3
Set Form1 = Nothing ' This means form1 will not be now taking up memory
Set Form2 = Nothing
Set Form3 = Nothing
End Sub
Private Sub Form_Load()
' What happens when the program starts for the first time
Agent1.Characters.Load "Hanz", "Hanz.acs"
Set hanz = Agent1.Characters.Character("Hanz")
hanz.Show
Timer1.Enabled = False
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False
Label22.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Text1.Text = "Riot Robin"
Text2.Text = "Strong Seth"
Text3.Text = "Pumping Paula"
Text4.Text = "Thunder Thorn"
Text5.Text = "Innovative Ian"
Text6.Text = "Zapping Zara"
Text7.Text = "Knotting Ken"
Text8.Text = "Powerful Pete"
Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
Text5.Visible = False
Text6.Visible = False
Text7.Visible = False
Text8.Visible = False
End Sub

Private Sub HBET_Click()
' Opens up form3 so one can bet on the horses.
' However, if the race has started, this is not possible.
If Timer1.Enabled = True Then
    MsgBox "A bet cannot be made whilst the horses are racing.", vbInformation
Else
    Form3.Show
End If
End Sub

Private Sub HSTATS_Click()
' Show the statistics of each horse.
Form2.Show
End Sub

Private Sub RENAMEHORSES_Click()
Form4.Show
End Sub

Private Sub Timer1_Timer()
' This is long and complicated. Basically (if that's the right word), you
' need to know that Randomize will create a random number between 0 and 5.
' Then, when the race starts, the number is called at random for each horse
' and it will move at that speed of which it is designated. Once a horse
' moves across the 360 barrier at the far left, that horse has won. Oh, you
' can't rig races, either. Unless you know how.
On Error Resume Next
Set hanz = Agent1.Characters.Character("Hanz")
Randomize
Label1 = Int(Rnd * 5)
Label2 = Int(Rnd * 5)
Label3 = Int(Rnd * 5)
Label4 = Int(Rnd * 5)
Label5 = Int(Rnd * 5)
Label6 = Int(Rnd * 5)
Label7 = Int(Rnd * 5)
Label8 = Int(Rnd * 5)
        If Label1 = 0 Then
            Image1.Left = Image1.Left - 10 ' The next chunk of code controls the speed at which the horses race
            time
        End If
        If Label1 = 1 Then
            Image1.Left = Image1.Left - 10
            time
        End If
        If Label1 = 2 Then
            Image1.Left = Image1.Left - 20
            time
        End If
        If Label1 = 3 Then
            Image1.Left = Image1.Left - 30
            time
        End If
        If Label1 = 4 Then
            Image1.Left = Image1.Left - 40
            time
        End If
        If Label1 = 5 Then
            Image1.Left = Image1.Left - 50
            time
        End If
        If Label2 = 0 Then
            Image2.Left = Image2.Left - 10
            time
        End If
        If Label2 = 1 Then
            Image2.Left = Image2.Left - 10
            time
        End If
        If Label2 = 2 Then
            Image2.Left = Image2.Left - 20
            time
        End If
        If Label2 = 3 Then
            Image2.Left = Image2.Left - 30
            time
        End If
        If Label2 = 4 Then
            Image2.Left = Image2.Left - 40
            time
        End If
        If Label2 = 5 Then
            Image2.Left = Image2.Left - 50
            time
        End If
        If Label3 = 0 Then
            Image3.Left = Image3.Left - 10
            time
        End If
        If Label3 = 1 Then
            Image3.Left = Image3.Left - 10
            time
        End If
        If Label3 = 2 Then
            Image3.Left = Image3.Left - 20
            time
        End If
        If Label3 = 3 Then
            Image3.Left = Image3.Left - 30
            time
        End If
        If Label3 = 4 Then
            Image3.Left = Image3.Left - 40
            time
        End If
        If Label3 = 5 Then
            Image3.Left = Image3.Left - 50
            time
        End If
         If Label4 = 0 Then
            Image4.Left = Image4.Left - 10
            time
        End If
        If Label4 = 1 Then
            Image4.Left = Image4.Left - 10
            time
        End If
        If Label4 = 2 Then
            Image4.Left = Image4.Left - 20
            time
        End If
        If Label4 = 3 Then
            Image4.Left = Image4.Left - 30
            time
        End If
        If Label4 = 4 Then
            Image4.Left = Image4.Left - 40
            time
        End If
        If Label4 = 5 Then
            Image4.Left = Image4.Left - 50
            time
        End If
        If Label5 = 0 Then
            Image5.Left = Image5.Left - 10
            time
        End If
        If Label5 = 1 Then
            Image5.Left = Image5.Left - 10
            time
        End If
        If Label5 = 2 Then
            Image5.Left = Image5.Left - 20
            time
        End If
        If Label5 = 3 Then
            Image5.Left = Image5.Left - 30
            time
        End If
        If Label5 = 4 Then
            Image5.Left = Image5.Left - 40
            time
        End If
        If Label5 = 5 Then
            Image5.Left = Image5.Left - 50
            time
        End If
        If Label6 = 0 Then
            Image6.Left = Image6.Left - 10
            time
        End If
        If Label6 = 1 Then
            Image6.Left = Image6.Left - 10
            time
        End If
        If Label6 = 2 Then
            Image6.Left = Image6.Left - 20
            time
        End If
        If Label6 = 3 Then
            Image6.Left = Image6.Left - 30
            time
        End If
        If Label6 = 4 Then
            Image6.Left = Image6.Left - 40
            time
        End If
        If Label6 = 5 Then
            Image6.Left = Image6.Left - 50
            time
        End If
        If Label7 = 0 Then
            Image7.Left = Image7.Left - 10
            time
        End If
        If Label7 = 1 Then
            Image7.Left = Image7.Left - 10
            time
        End If
        If Label7 = 2 Then
            Image7.Left = Image7.Left - 20
            time
        End If
        If Label7 = 3 Then
            Image7.Left = Image7.Left - 30
            time
        End If
        If Label7 = 4 Then
            Image7.Left = Image7.Left - 40
            time
        End If
        If Label7 = 5 Then
            Image7.Left = Image7.Left - 50
            time
        End If
        If Label8 = 0 Then
            Image8.Left = Image8.Left - 10
            time
        End If
        If Label8 = 1 Then
            Image8.Left = Image8.Left - 10
            time
        End If
        If Label8 = 2 Then
            Image8.Left = Image8.Left - 20
            time
        End If
        If Label8 = 3 Then
            Image8.Left = Image8.Left - 30
            time
        End If
        If Label8 = 4 Then
            Image8.Left = Image8.Left - 40
            time
        End If
        If Label8 = 5 Then
            Image8.Left = Image8.Left - 50
            time
        End If
        Dim i As Integer
            Randomize
            i = Int(Rnd * 3)
                If (Image1.Left < Image2.Left) And (Image1.Left < Image3.Left) And (Image1.Left < Image4.Left) And (Image1.Left < Image5.Left) And (Image1.Left < Image6.Left) And (Image1.Left < Image7.Left) And (Image1.Left < Image8.Left) And (i = 0) Then
                    hanz.Speak "Riot Robin Fuelling"
                ElseIf (Image1.Left < Image2.Left) And (Image1.Left < Image3.Left) And (Image1.Left < Image4.Left) And (Image1.Left < Image5.Left) And (Image1.Left < Image6.Left) And (Image1.Left < Image7.Left) And (Image1.Left < Image8.Left) And (i = 1) Then
                    hanz.Speak "Robin Rocking"
                ElseIf (Image1.Left < Image2.Left) And (Image1.Left < Image3.Left) And (Image1.Left < Image4.Left) And (Image1.Left < Image5.Left) And (Image1.Left < Image6.Left) And (Image1.Left < Image7.Left) And (Image1.Left < Image8.Left) And (i = 2) Then
                    hanz.Speak "Rob On A Roll"
                ElseIf (Image1.Left < Image2.Left) And (Image1.Left < Image3.Left) And (Image1.Left < Image4.Left) And (Image1.Left < Image5.Left) And (Image1.Left < Image6.Left) And (Image1.Left < Image7.Left) And (Image1.Left < Image8.Left) And (i = 3) Then
                    hanz.Speak "Robin Reliant"
                ElseIf (Image2.Left < Image1.Left) And (Image2.Left < Image3.Left) And (Image2.Left < Image4.Left) And (Image2.Left < Image5.Left) And (Image2.Left < Image6.Left) And (Image2.Left < Image7.Left) And (Image2.Left < Image8.Left) And (i = 0) Then
                    hanz.Speak "Strong Seth Roadworthy"
                ElseIf (Image2.Left < Image1.Left) And (Image2.Left < Image3.Left) And (Image2.Left < Image4.Left) And (Image2.Left < Image5.Left) And (Image2.Left < Image6.Left) And (Image2.Left < Image7.Left) And (Image2.Left < Image8.Left) And (i = 1) Then
                    hanz.Speak "Seth's Stomping Away!"
                ElseIf (Image2.Left < Image1.Left) And (Image2.Left < Image3.Left) And (Image2.Left < Image4.Left) And (Image2.Left < Image5.Left) And (Image2.Left < Image6.Left) And (Image2.Left < Image7.Left) And (Image2.Left < Image8.Left) And (i = 2) Then
                    hanz.Speak "Don't spit at Seth!"
                ElseIf (Image2.Left < Image1.Left) And (Image2.Left < Image3.Left) And (Image2.Left < Image4.Left) And (Image2.Left < Image5.Left) And (Image2.Left < Image6.Left) And (Image2.Left < Image7.Left) And (Image2.Left < Image8.Left) And (i = 3) Then
                    hanz.Speak "Seth Super!"
                ElseIf (Image3.Left < Image1.Left) And (Image3.Left < Image2.Left) And (Image3.Left < Image4.Left) And (Image3.Left < Image5.Left) And (Image3.Left < Image6.Left) And (Image3.Left < Image7.Left) And (Image3.Left < Image8.Left) And (i = 0) Then
                    hanz.Speak "Paula Performing!"
                ElseIf (Image3.Left < Image1.Left) And (Image3.Left < Image2.Left) And (Image3.Left < Image4.Left) And (Image3.Left < Image5.Left) And (Image3.Left < Image6.Left) And (Image3.Left < Image7.Left) And (Image3.Left < Image8.Left) And (i = 1) Then
                    hanz.Speak "Paula Staggering!"
                ElseIf (Image3.Left < Image1.Left) And (Image3.Left < Image2.Left) And (Image3.Left < Image4.Left) And (Image3.Left < Image5.Left) And (Image3.Left < Image6.Left) And (Image3.Left < Image7.Left) And (Image3.Left < Image8.Left) And (i = 2) Then
                    hanz.Speak "Pretty Paula Pressing!"
                ElseIf (Image3.Left < Image1.Left) And (Image3.Left < Image2.Left) And (Image3.Left < Image4.Left) And (Image3.Left < Image5.Left) And (Image3.Left < Image6.Left) And (Image3.Left < Image7.Left) And (Image3.Left < Image8.Left) And (i = 3) Then
                    hanz.Speak "Paula Outclassing!"
                ElseIf (Image4.Left < Image1.Left) And (Image4.Left < Image2.Left) And (Image4.Left < Image3.Left) And (Image4.Left < Image5.Left) And (Image4.Left < Image6.Left) And (Image4.Left < Image7.Left) And (Image4.Left < Image8.Left) And (i = 0) Then
                    hanz.Speak "Thunder Thorn Thorny!"
                ElseIf (Image4.Left < Image1.Left) And (Image4.Left < Image2.Left) And (Image4.Left < Image3.Left) And (Image4.Left < Image5.Left) And (Image4.Left < Image6.Left) And (Image4.Left < Image7.Left) And (Image4.Left < Image8.Left) And (i = 1) Then
                    hanz.Speak "Thunder Thorn Thornier Than The Rest!"
                ElseIf (Image4.Left < Image1.Left) And (Image4.Left < Image2.Left) And (Image4.Left < Image3.Left) And (Image4.Left < Image5.Left) And (Image4.Left < Image6.Left) And (Image4.Left < Image7.Left) And (Image4.Left < Image8.Left) And (i = 2) Then
                    hanz.Speak "Thorn Under Way!"
                ElseIf (Image4.Left < Image1.Left) And (Image4.Left < Image2.Left) And (Image4.Left < Image3.Left) And (Image4.Left < Image5.Left) And (Image4.Left < Image6.Left) And (Image4.Left < Image7.Left) And (Image4.Left < Image8.Left) And (i = 3) Then
                    hanz.Speak "Thorn Lucky!"
                ElseIf (Image5.Left < Image1.Left) And (Image5.Left < Image2.Left) And (Image5.Left < Image3.Left) And (Image5.Left < Image4.Left) And (Image5.Left < Image6.Left) And (Image6.Left < Image7.Left) And (Image7.Left < Image8.Left) And (i = 0) Then
                    hanz.Speak "Innovation Occurring!"
                ElseIf (Image5.Left < Image1.Left) And (Image5.Left < Image2.Left) And (Image5.Left < Image3.Left) And (Image5.Left < Image4.Left) And (Image5.Left < Image6.Left) And (Image6.Left < Image7.Left) And (Image7.Left < Image8.Left) And (i = 1) Then
                    hanz.Speak "Ian Is Winning!"
                ElseIf (Image5.Left < Image1.Left) And (Image5.Left < Image2.Left) And (Image5.Left < Image3.Left) And (Image5.Left < Image4.Left) And (Image5.Left < Image6.Left) And (Image6.Left < Image7.Left) And (Image7.Left < Image8.Left) And (i = 2) Then
                    hanz.Speak "Ian Tantalizing!"
                ElseIf (Image5.Left < Image1.Left) And (Image5.Left < Image2.Left) And (Image5.Left < Image3.Left) And (Image5.Left < Image4.Left) And (Image5.Left < Image6.Left) And (Image6.Left < Image7.Left) And (Image7.Left < Image8.Left) And (i = 3) Then
                    hanz.Speak "Ian Good!"
                ElseIf (Image6.Left < Image1.Left) And (Image6.Left < Image2.Left) And (Image6.Left < Image3.Left) And (Image6.Left < Image4.Left) And (Image6.Left < Image5.Left) And (Image6.Left < Image7.Left) And (Image6.Left < Image8.Left) And (i = 0) Then
                    hanz.Speak "Zara Cool!"
                ElseIf (Image6.Left < Image1.Left) And (Image6.Left < Image2.Left) And (Image6.Left < Image3.Left) And (Image6.Left < Image4.Left) And (Image6.Left < Image5.Left) And (Image6.Left < Image7.Left) And (Image6.Left < Image8.Left) And (i = 1) Then
                    hanz.Speak "Flaming BK Zara!"
                ElseIf (Image6.Left < Image1.Left) And (Image6.Left < Image2.Left) And (Image6.Left < Image3.Left) And (Image6.Left < Image4.Left) And (Image6.Left < Image5.Left) And (Image6.Left < Image7.Left) And (Image6.Left < Image8.Left) And (i = 2) Then
                    hanz.Speak "Is Zara Zummeting"
                ElseIf (Image6.Left < Image1.Left) And (Image6.Left < Image2.Left) And (Image6.Left < Image3.Left) And (Image6.Left < Image4.Left) And (Image6.Left < Image5.Left) And (Image6.Left < Image7.Left) And (Image6.Left < Image8.Left) And (i = 3) Then
                    hanz.Speak "Zara Coming Up Top!"
                ElseIf (Image7.Left < Image1.Left) And (Image7.Left < Image2.Left) And (Image7.Left < Image3.Left) And (Image7.Left < Image4.Left) And (Image7.Left < Image5.Left) And (Image7.Left < Image6.Left) And (Image7.Left < Image8.Left) And (i = 0) Then
                    hanz.Speak "Ken De Old Speed Demon!"
                ElseIf (Image7.Left < Image1.Left) And (Image7.Left < Image2.Left) And (Image7.Left < Image3.Left) And (Image7.Left < Image4.Left) And (Image7.Left < Image5.Left) And (Image7.Left < Image6.Left) And (Image7.Left < Image8.Left) And (i = 1) Then
                    hanz.Speak "Krafty Ken!"
                ElseIf (Image7.Left < Image1.Left) And (Image7.Left < Image2.Left) And (Image7.Left < Image3.Left) And (Image7.Left < Image4.Left) And (Image7.Left < Image5.Left) And (Image7.Left < Image6.Left) And (Image7.Left < Image8.Left) And (i = 2) Then
                    hanz.Speak "Is Ken Old And Good?"
                ElseIf (Image7.Left < Image1.Left) And (Image7.Left < Image2.Left) And (Image7.Left < Image3.Left) And (Image7.Left < Image4.Left) And (Image7.Left < Image5.Left) And (Image7.Left < Image6.Left) And (Image7.Left < Image8.Left) And (i = 3) Then
                    hanz.Speak "Ken Knitting His Victories!"
                ElseIf (Image8.Left < Image1.Left) And (Image8.Left < Image2.Left) And (Image8.Left < Image3.Left) And (Image8.Left < Image4.Left) And (Image8.Left < Image5.Left) And (Image8.Left < Image6.Left) And (Image8.Left < Image7.Left) And (i = 0) Then
                    hanz.Speak "Let's not forget Pete!"
                ElseIf (Image8.Left < Image1.Left) And (Image8.Left < Image2.Left) And (Image8.Left < Image3.Left) And (Image8.Left < Image4.Left) And (Image8.Left < Image5.Left) And (Image8.Left < Image6.Left) And (Image8.Left < Image7.Left) And (i = 1) Then
                    hanz.Speak "Pete Purgeful!"
                ElseIf (Image8.Left < Image1.Left) And (Image8.Left < Image2.Left) And (Image8.Left < Image3.Left) And (Image8.Left < Image4.Left) And (Image8.Left < Image5.Left) And (Image8.Left < Image6.Left) And (Image8.Left < Image7.Left) And (i = 2) Then
                    hanz.Speak "Pete Packing His Power!"
                ElseIf (Image8.Left < Image1.Left) And (Image8.Left < Image2.Left) And (Image8.Left < Image3.Left) And (Image8.Left < Image4.Left) And (Image8.Left < Image5.Left) And (Image8.Left < Image6.Left) And (Image8.Left < Image7.Left) And (i = 3) Then
                    hanz.Speak "Pete Is Plotting His Winnings!"
                End If
        Timer1.Enabled = True
        If (Timer1.Enabled = True) And (Image1.Left < 360) Then
            Timer1.Enabled = False
            hanz.Stop
            MsgBox "Congratulations " + Form4.Text1.Text, vbInformation
            If Form3.Text1.Text = 0 Then Form3.Text1.Text = Form3.Text1.Text * 0
            If Form3.Text1.Text > 0 Then Form3.Text1.Text = Form3.Text1.Text * 5
            Form3.Text9.Text = 0 + Form3.Text1.Text
            Form2.Label2.Caption = Form2.Label2.Caption + 1
            Label22.Caption = Label22.Caption + 1
            Form3.Text1 = ""
            Form3.Text2 = ""
            Form3.Text3 = ""
            Form3.Text4 = ""
            Form3.Text5 = ""
            Form3.Text6 = ""
            Form3.Text7 = ""
            Form3.Text8 = ""
            Form1.Label6 = ""
            Form1.Label7 = ""
            Form1.Label8 = ""
            Form1.Label9 = ""
            Form1.Label10 = ""
            Form1.Label11 = ""
            Form1.Label12 = ""
            Form1.Label13 = ""
            Image1.Left = 7320
            Image2.Left = 7320
            Image3.Left = 7320
            Image4.Left = 7320
            Image5.Left = 7320
            Image6.Left = 7320
            Image7.Left = 7320
            Image8.Left = 7320
        Else
        If (Timer1.Enabled = True) And (Image2.Left < 360) Then
            Timer1.Enabled = False
            hanz.Stop
            MsgBox "Congratulations " + Form4.Text2.Text, vbInformation
            If Form3.Text2.Text = 0 Then Form3.Text2.Text = Form3.Text2.Text * 0
            If Form3.Text2.Text > 0 Then Form3.Text2.Text = Form3.Text2.Text * 13
            Form3.Text9.Text = 0 + Form3.Text2.Text
            Form2.Label4.Caption = Form2.Label4.Caption + 1
            Label23.Caption = Label23.Caption + 1
            Form3.Text1 = ""
            Form3.Text2 = ""
            Form3.Text3 = ""
            Form3.Text4 = ""
            Form3.Text5 = ""
            Form3.Text6 = ""
            Form3.Text7 = ""
            Form3.Text8 = ""
            Form1.Label6 = ""
            Form1.Label7 = ""
            Form1.Label8 = ""
            Form1.Label9 = ""
            Form1.Label10 = ""
            Form1.Label11 = ""
            Form1.Label12 = ""
            Form1.Label13 = ""
            Image1.Left = 7320
            Image2.Left = 7320
            Image3.Left = 7320
            Image4.Left = 7320
            Image5.Left = 7320
            Image6.Left = 7320
            Image7.Left = 7320
            Image8.Left = 7320
        Else
        If (Timer1.Enabled = True) And (Image3.Left < 360) Then
            Timer1.Enabled = False
            hanz.Stop
            MsgBox "Congratulations " + Form4.Text3.Text, vbInformation
            If Form3.Text3.Text = 0 Then Form3.Text3.Text = Form3.Text3.Text * 0
            If Form3.Text3.Text > 0 Then Form3.Text3.Text = Form3.Text3.Text * 4
            Form3.Text9.Text = 0 + Form3.Text3.Text
            Form2.Label6.Caption = Form2.Label6.Caption + 1
            Label24.Caption = Label24.Caption + 1
            Form3.Text1 = ""
            Form3.Text2 = ""
            Form3.Text3 = ""
            Form3.Text4 = ""
            Form3.Text5 = ""
            Form3.Text6 = ""
            Form3.Text7 = ""
            Form3.Text8 = ""
            Form1.Label6 = ""
            Form1.Label7 = ""
            Form1.Label8 = ""
            Form1.Label9 = ""
            Form1.Label10 = ""
            Form1.Label11 = ""
            Form1.Label12 = ""
            Form1.Label13 = ""
            Image1.Left = 7320
            Image2.Left = 7320
            Image3.Left = 7320
            Image4.Left = 7320
            Image5.Left = 7320
            Image6.Left = 7320
            Image7.Left = 7320
            Image8.Left = 7320
        Else
        If (Timer1.Enabled = True) And (Image4.Left < 360) Then
            hanz.Stop
            Timer1.Enabled = False
            MsgBox "Congratulations " + Form4.Text4.Text, vbInformation
            If Form3.Text4.Text = 0 Then Form3.Text4.Text = Form3.Text4.Text * 0
            If Form3.Text4.Text > 0 Then Form3.Text4.Text = Form3.Text4.Text * 5
            Form3.Text9.Text = 0 + Form3.Text4.Text
            Form2.Label8.Caption = Form2.Label8.Caption + 1
            Label25.Caption = Label25.Caption + 1
            Form3.Text1 = ""
            Form3.Text2 = ""
            Form3.Text3 = ""
            Form3.Text4 = ""
            Form3.Text5 = ""
            Form3.Text6 = ""
            Form3.Text7 = ""
            Form3.Text8 = ""
            Form1.Label6 = ""
            Form1.Label7 = ""
            Form1.Label8 = ""
            Form1.Label9 = ""
            Form1.Label10 = ""
            Form1.Label11 = ""
            Form1.Label12 = ""
            Form1.Label13 = ""
            Image1.Left = 7320
            Image2.Left = 7320
            Image3.Left = 7320
            Image4.Left = 7320
            Image5.Left = 7320
            Image6.Left = 7320
            Image7.Left = 7320
            Image8.Left = 7320
        Else
        If (Timer1.Enabled = True) And (Image5.Left < 360) Then
            Timer1.Enabled = False
            hanz.Stop
            MsgBox "Congratulations " + Form4.Text5.Text, vbInformation
            If Form3.Text5.Text = 0 Then Form3.Text5.Text = Form3.Text5.Text * 0
            If Form3.Text5.Text > 0 Then Form3.Text5.Text = Form3.Text5.Text * 8
            Form3.Text9.Text = 0 + Form3.Text5.Text
            Form2.Label10.Caption = Form2.Label10.Caption + 1
            Label26.Caption = Label26.Caption + 1
            Form3.Text1 = ""
            Form3.Text2 = ""
            Form3.Text3 = ""
            Form3.Text4 = ""
            Form3.Text5 = ""
            Form3.Text6 = ""
            Form3.Text7 = ""
            Form3.Text8 = ""
            Form1.Label6 = ""
            Form1.Label7 = ""
            Form1.Label8 = ""
            Form1.Label9 = ""
            Form1.Label10 = ""
            Form1.Label11 = ""
            Form1.Label12 = ""
            Form1.Label13 = ""
            Image1.Left = 7320
            Image2.Left = 7320
            Image3.Left = 7320
            Image4.Left = 7320
            Image5.Left = 7320
            Image6.Left = 7320
            Image7.Left = 7320
            Image8.Left = 7320
        Else
        If (Timer1.Enabled = True) And (Image6.Left < 360) Then
            Timer2.Enabled = False
            hanz.Stop
            MsgBox "Congratulations " + Form4.Text6.Text, vbInformation
            If Form3.Text6.Text = 0 Then Form3.Text6.Text = Form3.Text6.Text * 0
            If Form3.Text6.Text > 0 Then Form3.Text6.Text = Form3.Text6.Text * 2
            Form3.Text9.Text = 0 + Form3.Text6.Text
            Form2.Label12.Caption = Form2.Label12.Caption + 1
            Label27.Caption = Label27.Caption + 1
            Form3.Text1 = ""
            Form3.Text2 = ""
            Form3.Text3 = ""
            Form3.Text4 = ""
            Form3.Text5 = ""
            Form3.Text6 = ""
            Form3.Text7 = ""
            Form3.Text8 = ""
            Form1.Label6 = ""
            Form1.Label7 = ""
            Form1.Label8 = ""
            Form1.Label9 = ""
            Form1.Label10 = ""
            Form1.Label11 = ""
            Form1.Label12 = ""
            Form1.Label13 = ""
            Image1.Left = 7320
            Image2.Left = 7320
            Image3.Left = 7320
            Image4.Left = 7320
            Image5.Left = 7320
            Image6.Left = 7320
            Image7.Left = 7320
            Image8.Left = 7320
        Else
        If (Timer1.Enabled = True) And (Image7.Left < 360) Then
            Timer1.Enabled = False
            hanz.Stop
            MsgBox "Congratulations " + Form4.Text7.Text, vbInformation
            If Form3.Text7.Text = 0 Then Form3.Text7.Text = Form3.Text7.Text * 0
            If Form3.Text7.Text > 0 Then Form3.Text7.Text = Form3.Text7.Text * 7
            Form3.Text9.Text = 0 + Form3.Text7.Text
            Form2.Label14.Caption = Form2.Label14.Caption + 1
            Label28.Caption = Label28.Caption + 1
            Form3.Text1 = ""
            Form3.Text2 = ""
            Form3.Text3 = ""
            Form3.Text4 = ""
            Form3.Text5 = ""
            Form3.Text6 = ""
            Form3.Text7 = ""
            Form3.Text8 = ""
            Form1.Label6 = ""
            Form1.Label7 = ""
            Form1.Label8 = ""
            Form1.Label9 = ""
            Form1.Label10 = ""
            Form1.Label11 = ""
            Form1.Label12 = ""
            Form1.Label13 = ""
            Image1.Left = 7320
            Image2.Left = 7320
            Image3.Left = 7320
            Image4.Left = 7320
            Image5.Left = 7320
            Image6.Left = 7320
            Image7.Left = 7320
            Image8.Left = 7320
        Else
        If (Timer1.Enabled = True) And (Image8.Left < 360) Then
            Timer1.Enabled = False
            hanz.Stop
            MsgBox "Congratulations " + Form4.Text8.Text, vbInformation
            If Form3.Text8.Text = 0 Then Form3.Text8.Text = Form3.Text8.Text * 0
            If Form3.Text8.Text > 0 Then Form3.Text8.Text = Form3.Text8.Text * 3
            Form3.Text9.Text = 0 + Form3.Text8.Text
            Form2.Label16.Caption = Form2.Label16.Caption + 1
            Label29.Caption = Label29.Caption + 1
            Form3.Text1 = ""
            Form3.Text2 = ""
            Form3.Text3 = ""
            Form3.Text4 = ""
            Form3.Text5 = ""
            Form3.Text6 = ""
            Form3.Text7 = ""
            Form3.Text8 = ""
            Form1.Label6 = ""
            Form1.Label7 = ""
            Form1.Label8 = ""
            Form1.Label9 = ""
            Form1.Label10 = ""
            Form1.Label11 = ""
            Form1.Label12 = ""
            Form1.Label13 = ""
            Image1.Left = 7320
            Image2.Left = 7320
            Image3.Left = 7320
            Image4.Left = 7320
            Image5.Left = 7320
            Image6.Left = 7320
            Image7.Left = 7320
            Image8.Left = 7320
        End If ' Eight end if's are needed to support the else statements
        End If
        End If
        End If
        End If
        End If
        End If
        End If
End Sub

