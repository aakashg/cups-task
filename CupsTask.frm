VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   11655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   11655
   ScaleWidth      =   18330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   120
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3495
      Left            =   720
      Picture         =   "CupsTask.frx":0000
      ScaleHeight     =   3495
      ScaleWidth      =   8940
      TabIndex        =   2
      Top             =   3480
      Width           =   8940
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3495
      Left            =   9720
      Picture         =   "CupsTask.frx":3294
      ScaleHeight     =   3495
      ScaleWidth      =   8940
      TabIndex        =   1
      Top             =   3480
      Width           =   8940
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   0
      Top             =   7320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FF00&
      Caption         =   " Running Total ($):"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13800
      TabIndex        =   7
      Top             =   10320
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   16320
      TabIndex        =   6
      Top             =   10320
      Width           =   1695
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Press 's' to start when you are ready."
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   5880
      TabIndex        =   5
      Top             =   7680
      Width           =   8775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   3720
      TabIndex        =   4
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   12960
      TabIndex        =   3
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   4215
      Left            =   8040
      Top             =   3000
      Visible         =   0   'False
      Width           =   3975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Private Declare Function GetCurrentTime Lib "kernel32" Alias "GetTickCount" () As Long
Dim path, j, SafeSide, running_total, subID
Dim domain(90), category(90), CupN(90), pay(90), order(90), order0(90)
Dim trial, Choice, payoff, s(4), jishi, suiji, trialN, position

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()

path = App.path & "\material"
subID = Form1.subID

'###### Domain with half gain and half lose(0-gain, 1-lose)
For i = 1 To 45: domain(i) = 0: Next i
For i = 46 To 90: domain(i) = 1: Next i

'###### category-1: equal; category-2: Risk Advantageous; category-3: Risk Disadvantageous
For i = 1 To 90
    j = (i - Int((i - 1) / 45) * 45)
    If j < 16 Then category(i) = 1
    If j > 15 And j < 31 Then category(i) = 2
    If j > 30 Then category(i) = 3
Next i

'###### CupN
For i = 1 To 45
    j = Int((i - 1) / 5) + 1
    If j = 1 Then CupN(i) = 2: pay(i) = "2"
    If j = 2 Then CupN(i) = 3: pay(i) = "3"
    If j = 3 Then CupN(i) = 5: pay(i) = "5"
    If j = 4 Then CupN(i) = 2: pay(i) = "3"
    If j = 5 Then CupN(i) = 2: pay(i) = "5"
    If j = 6 Then CupN(i) = 3: pay(i) = "5"
    If j = 7 Then CupN(i) = 3: pay(i) = "2"
    If j = 8 Then CupN(i) = 5: pay(i) = "2"
    If j = 9 Then CupN(i) = 5: pay(i) = "3"
Next i

For i = 46 To 90
    j = Int((i - 46) / 5) + 1
    If j = 1 Then CupN(i) = 2: pay(i) = "-2"
    If j = 2 Then CupN(i) = 3: pay(i) = "-3"
    If j = 3 Then CupN(i) = 5: pay(i) = "-5"
    If j = 4 Then CupN(i) = 3: pay(i) = "-2"
    If j = 5 Then CupN(i) = 5: pay(i) = "-2"
    If j = 6 Then CupN(i) = 5: pay(i) = "-3"
    If j = 7 Then CupN(i) = 2: pay(i) = "-3"
    If j = 8 Then CupN(i) = 2: pay(i) = "-5"
    If j = 9 Then CupN(i) = 3: pay(i) = "-5"
Next i


For i = 1 To 90: order0(i) = i: Next i
Randomize
For i = 1 To 90
    j = Int(Rnd * (91 - i)) + 1: order(i) = order0(j)
    For k = j To (90 - i): order0(k) = order0(k + 1): Next k
Next i


trialN = 90: trial = 1
running_total = 0: Label4.Caption = running_total


Randomize
SafeSide = Int(Rnd * 2)

j = order(trial)
If SafeSide = 0 Then
    If domain(j) = 0 And CupN(j) = 5 Then Picture1 = LoadPicture(path & "\1-cup.jpg"): Label1.Caption = "1": Picture2 = LoadPicture(path & "\5-cup.jpg"): Label2.Caption = pay(j)
    If domain(j) = 0 And CupN(j) = 3 Then Picture1 = LoadPicture(path & "\1-cup.jpg"): Label1.Caption = "1": Picture2 = LoadPicture(path & "\3-cup.jpg"): Label2.Caption = pay(j)
    If domain(j) = 0 And CupN(j) = 2 Then Picture1 = LoadPicture(path & "\1-cup.jpg"): Label1.Caption = "1": Picture2 = LoadPicture(path & "\2-cup.jpg"): Label2.Caption = pay(j)
    
    If domain(j) = 1 And CupN(j) = 5 Then Picture1 = LoadPicture(path & "\1-cup-lose.jpg"): Label1.Caption = "-1": Picture2 = LoadPicture(path & "\5-cup-lose.jpg"): Label2.Caption = pay(j)
    If domain(j) = 1 And CupN(j) = 3 Then Picture1 = LoadPicture(path & "\1-cup-lose.jpg"): Label1.Caption = "-1": Picture2 = LoadPicture(path & "\3-cup-lose.jpg"): Label2.Caption = pay(j)
    If domain(j) = 1 And CupN(j) = 2 Then Picture1 = LoadPicture(path & "\1-cup-lose.jpg"): Label1.Caption = "-1": Picture2 = LoadPicture(path & "\2-cup-lose.jpg"): Label2.Caption = pay(j)
End If

If SafeSide = 1 Then
    If domain(j) = 0 And CupN(j) = 5 Then Picture2 = LoadPicture(path & "\1-cup.jpg"): Label2.Caption = "1": Picture1 = LoadPicture(path & "\5-cup.jpg"): Label1.Caption = pay(j)
    If domain(j) = 0 And CupN(j) = 3 Then Picture2 = LoadPicture(path & "\1-cup.jpg"): Label2.Caption = "1": Picture1 = LoadPicture(path & "\3-cup.jpg"): Label1.Caption = pay(j)
    If domain(j) = 0 And CupN(j) = 2 Then Picture2 = LoadPicture(path & "\1-cup.jpg"): Label2.Caption = "1": Picture1 = LoadPicture(path & "\2-cup.jpg"): Label1.Caption = pay(j)
    
    If domain(j) = 1 And CupN(j) = 5 Then Picture2 = LoadPicture(path & "\1-cup-lose.jpg"): Label2.Caption = "-1": Picture1 = LoadPicture(path & "\5-cup-lose.jpg"): Label1.Caption = pay(j)
    If domain(j) = 1 And CupN(j) = 3 Then Picture2 = LoadPicture(path & "\1-cup-lose.jpg"): Label2.Caption = "-1": Picture1 = LoadPicture(path & "\3-cup-lose.jpg"): Label1.Caption = pay(j)
    If domain(j) = 1 And CupN(j) = 2 Then Picture2 = LoadPicture(path & "\1-cup-lose.jpg"): Label2.Caption = "-1": Picture1 = LoadPicture(path & "\2-cup-lose.jpg"): Label1.Caption = pay(j)
End If



Label3.Left = 7500: Label3.Top = 7600: Label3.Height = 1400: Label3.Width = 9000: Label3.FontSize = 42: Label3.Visible = True: Label3.Caption = "Pick up a cup.": t1 = GetCurrentTime
Picture1.Enabled = True: Picture2.Enabled = True
End Sub

Private Sub Picture1_Click()

j = order(trial)

't2 = GetCurrentTime
Picture1.Visible = False: Picture2.Visible = False: Label1.Visible = False: Label2.Visible = False

Choice = SafeSide

position = "L"

If SafeSide = 0 Then
    If domain(j) = 0 Then payoff = "1":     running_total = running_total + Val(payoff): Label4.Caption = running_total
    If domain(j) = 1 Then payoff = "-1":    running_total = running_total + Val(payoff): Label4.Caption = running_total

End If

If SafeSide = 1 Then
    payoff = pay(j)
    
    Randomize
    suiji = Int(Rnd * CupN(j)) + 1
    If suiji <> CupN(j) Then payoff = "0"
    running_total = running_total + Val(payoff): Label4.Caption = running_total
    
End If

'RT = (t2 - t1) / 1000: Prof = Prof + win + loss
Open "P:\TaskOutput\CupsTaskOutput_" & subID & ".txt" For Append As #1
Print #1, trial, domain(j), category(j), CupN(j), pay(j), Choice, position, payoff
Close #1

If domain(j) = 0 Then Label3.Caption = "    Win($)   " & payoff: Image1.Picture = LoadPicture(path & "\WinFace.gif"): Image1.Visible = True
If domain(j) = 1 Then Label3.Caption = "   Lose($)   " & payoff: Image1.Picture = LoadPicture(path & "\LoseFace.gif"): Image1.Visible = True

trial = trial + 1
If trial < trialN + 1 Then jishi = 0: Timer1.Enabled = True
If trial = trialN + 1 Then Image1.Visible = False: Command1.Visible = True: Label3.Left = 4000: Label3.Top = 5000: Label3.Width = 24000: Label3.FontSize = 20: Label3.Caption = "This is the end of this task. Please click 'EXIT' button to exit the task."

Randomize
SafeSide = Int(Rnd * 2)

End Sub
Private Sub Picture2_Click()

j = order(trial)

Picture1.Visible = False: Picture2.Visible = False: Label1.Visible = False: Label2.Visible = False
Choice = 1 - SafeSide
position = "R"

If SafeSide = 1 Then
    If domain(j) = 0 Then payoff = "1":     running_total = running_total + Val(payoff): Label4.Caption = running_total
    If domain(j) = 1 Then payoff = "-1":    running_total = running_total + Val(payoff): Label4.Caption = running_total
End If

If SafeSide = 0 Then
    
    payoff = pay(j)
    
    Randomize
    suiji = Int(Rnd * CupN(j)) + 1
    If suiji <> CupN(j) Then payoff = "0"
     running_total = running_total + Val(payoff): Label4.Caption = running_total
End If
   
'RT = (t2 - t1) / 1000: Prof = Prof + win + loss
Open "P:\TaskOutput\CupsTaskOutput_" & subID & ".txt" For Append As #1
Print #1, trial, domain(j), category(j), CupN(j), pay(j), Choice, position, payoff
Close #1

If domain(j) = 0 Then Label3.Caption = "    Win($)   " & payoff: Image1.Picture = LoadPicture(path & "\WinFace.gif"): Image1.Visible = True
If domain(j) = 1 Then Label3.Caption = "   Lose($)   " & payoff: Image1.Picture = LoadPicture(path & "\LoseFace.gif"): Image1.Visible = True

trial = trial + 1
If trial < trialN + 1 Then jishi = 0: Timer1.Enabled = True
If trial = trialN + 1 Then Image1.Visible = False: Command1.Visible = True: Label3.Left = 4000: Label3.Top = 5000: Label3.Width = 24000: Label3.FontSize = 20: Label3.Caption = "This is the end of this task. Please click 'EXIT' button to exit the task."

Randomize
SafeSide = Int(Rnd * 2)

End Sub

Private Sub Timer1_Timer()

jishi = jishi + 1
If jishi = 2 Then
 
j = order(trial)

If SafeSide = 0 Then
    If domain(j) = 0 And CupN(j) = 5 Then Picture1 = LoadPicture(path & "\1-cup.jpg"): Label1.Caption = "1": Picture2 = LoadPicture(path & "\5-cup.jpg"): Label2.Caption = pay(j)
    If domain(j) = 0 And CupN(j) = 3 Then Picture1 = LoadPicture(path & "\1-cup.jpg"): Label1.Caption = "1": Picture2 = LoadPicture(path & "\3-cup.jpg"): Label2.Caption = pay(j)
    If domain(j) = 0 And CupN(j) = 2 Then Picture1 = LoadPicture(path & "\1-cup.jpg"): Label1.Caption = "1": Picture2 = LoadPicture(path & "\2-cup.jpg"): Label2.Caption = pay(j)
    
    If domain(j) = 1 And CupN(j) = 5 Then Picture1 = LoadPicture(path & "\1-cup-lose.jpg"): Label1.Caption = "-1": Picture2 = LoadPicture(path & "\5-cup-lose.jpg"): Label2.Caption = pay(j)
    If domain(j) = 1 And CupN(j) = 3 Then Picture1 = LoadPicture(path & "\1-cup-lose.jpg"): Label1.Caption = "-1": Picture2 = LoadPicture(path & "\3-cup-lose.jpg"): Label2.Caption = pay(j)
    If domain(j) = 1 And CupN(j) = 2 Then Picture1 = LoadPicture(path & "\1-cup-lose.jpg"): Label1.Caption = "-1": Picture2 = LoadPicture(path & "\2-cup-lose.jpg"): Label2.Caption = pay(j)
End If

If SafeSide = 1 Then
    If domain(j) = 0 And CupN(j) = 5 Then Picture2 = LoadPicture(path & "\1-cup.jpg"): Label2.Caption = "1": Picture1 = LoadPicture(path & "\5-cup.jpg"): Label1.Caption = pay(j)
    If domain(j) = 0 And CupN(j) = 3 Then Picture2 = LoadPicture(path & "\1-cup.jpg"): Label2.Caption = "1": Picture1 = LoadPicture(path & "\3-cup.jpg"): Label1.Caption = pay(j)
    If domain(j) = 0 And CupN(j) = 2 Then Picture2 = LoadPicture(path & "\1-cup.jpg"): Label2.Caption = "1": Picture1 = LoadPicture(path & "\2-cup.jpg"): Label1.Caption = pay(j)
    
    If domain(j) = 1 And CupN(j) = 5 Then Picture2 = LoadPicture(path & "\1-cup-lose.jpg"): Label2.Caption = "-1": Picture1 = LoadPicture(path & "\5-cup-lose.jpg"): Label1.Caption = pay(j)
    If domain(j) = 1 And CupN(j) = 3 Then Picture2 = LoadPicture(path & "\1-cup-lose.jpg"): Label2.Caption = "-1": Picture1 = LoadPicture(path & "\3-cup-lose.jpg"): Label1.Caption = pay(j)
    If domain(j) = 1 And CupN(j) = 2 Then Picture2 = LoadPicture(path & "\1-cup-lose.jpg"): Label2.Caption = "-1": Picture1 = LoadPicture(path & "\2-cup-lose.jpg"): Label1.Caption = pay(j)
End If
 
Timer1.Enabled = False: Timer2.Enabled = True: jishi = 0
        
End If

End Sub
Private Sub Timer2_Timer()
jishi = jishi + 1

If jishi = 1 Then
    Image1.Visible = False
    Label3.Caption = "Pick up a cup."
    Picture1.Visible = True: Picture2.Visible = True: Label1.Visible = True: Label2.Visible = True
    Timer2.Enabled = False
End If

End Sub
