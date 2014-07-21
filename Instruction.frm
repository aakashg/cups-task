VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Instruction"
   ClientHeight    =   10710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11244.09
   ScaleMode       =   0  'User
   ScaleWidth      =   18000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   5
      Top             =   9960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3495
      Left            =   9120
      Picture         =   "Instruction.frx":0000
      ScaleHeight     =   3495
      ScaleWidth      =   8940
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   8940
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3495
      Left            =   240
      Picture         =   "Instruction.frx":1091
      ScaleHeight     =   3495
      ScaleWidth      =   8940
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   8940
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8000
      TabIndex        =   0
      Top             =   5040
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continue"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   2
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   13080
      TabIndex        =   7
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4680
      TabIndex        =   6
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Please put in your Participation ID:"
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
      Height          =   4425
      Left            =   7680
      TabIndex        =   1
      Top             =   3360
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num, path
Public subID As String

Private Sub Command2_Click()
Unload Me
Form2.Visible = True
End Sub

Private Sub Form_Load()
path = App.path & "\material"
End Sub

Private Sub Text1_Change()
If Val(Text1.Text) > 0 Then Command1.Enabled = True
End Sub

Private Sub Command1_Click()

KeyPreview = True
Text1.Visible = False: Command1.Visible = False: Form1.BackColor = &H808000
Picture1.Visible = True: Picture2.Visible = True: Picture1 = LoadPicture(path & "\1-cup-demo.jpg"): Picture2 = LoadPicture(path & "\3-cup-demo.jpg")

Label1.Left = 2200: Label1.Top = 6800: Label1.ForeColor = &HFFFFFF: Label1.Width = 15400: Label1.FontSize = 24
Label1.Caption = "          Read through the instructions carefully before you start the game." & vbNewLine & vbNewLine & "                                Spacebar: Proceed to next page" & vbNewLine & "                                'b' key: Go back to previous page"
num = 0

subID = Text1.Text
Open "P:\TaskOutput\CupsTaskOutput_" & subID & ".txt" For Append As #1
Print #1, "Subject ID:"; subID; "    "; "Time: "; Date; "   "; Time; "    "
Print #1, "Trial", "Domain", "Category", "CupsNumber", "AmountShown", "Choice(1-rsk)", "Position", "Payoff"
Close #1


End Sub

Private Sub Form_DblClick()
End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 32 Then
    num = num + 1
    If num = 1 Then Label1.Caption = vbNewLine & "           On each side of the screen, you will see certain number of cups." & vbNewLine & vbNewLine & "There will be 1 cup on one side and multiple cups (either 2, 3, or 5) on the other side."
    If num = 2 Then Label1.Caption = vbNewLine & vbNewLine & "For each trial, you will be given the option of choosing a cup from either side by clicking on your choice. "
    If num = 3 Then Label2.Caption = "$1": Label3.Caption = "$3": Picture1 = LoadPicture(path & "\1-cup.jpg"): Picture2 = LoadPicture(path & "\3-cup.jpg"): Label1.Caption = "When the cups are shaded BLUE:" & vbNewLine & vbNewLine & "a) Choosing from the side with one cup will result in winning $1." & vbNewLine & vbNewLine & "b) Choosing a cup from the side with multiple cups, only one cup contains the amount shown above the cups (either $2, $3, or $5). However, if you choose the wrong cup, you will win $0."
    If num = 4 Then Picture1 = LoadPicture(path & "\1-cup-lose.jpg"): Picture2 = LoadPicture(path & "\3-cup-lose.jpg"):  Label1.Caption = "When the cups are shaded RED:" & vbNewLine & vbNewLine & "a) Choosing from the side with one cup will result in losing $1." & vbNewLine & vbNewLine & "b) Choosing a cup from the side with multiple cups, only one cup causes you to lose the amount shown above the cups (either $2, $3, or $5). However, if you do not choose that cup, you will lose $0."
    If num = 5 Then Picture1.Visible = False: Picture2.Visible = False: Label2.Caption = "": Label3.Caption = "": Label1.Caption = "                       Please read the payouts for each trial carefully." & vbNewLine & vbNewLine & "                              The sure payout side switches randomly."
    If num = 6 Then Label1.Caption = "The goal of the game is to win as much money as possible and avoid losing as much money as possible." & vbNewLine & vbNewLine & "         Your running total amount will be displayed at the bottom-right corner." & vbNewLine & "               Please note that it is possible to get negative amount of money."
    If num = 7 Then Command2.Visible = True: Label1.Caption = "                     Click 'START' button to start the task if you are ready." & vbNewLine & vbNewLine & "                        Press 'b' if you need to go back to previous pages."
    If num > 7 Then num = 7
End If

If KeyAscii = 98 Or KeyAscii = 66 Then
    num = num - 1
    If num = 0 Then Label1.Caption = "          Read through the instructions carefully before you start the game." & vbNewLine & vbNewLine & "                                Spacebar: Proceed to next page" & vbNewLine & "                                'b' key: Go back to previous page"
    If num = 1 Then Label1.Caption = vbNewLine & "           On each side of the screen, you will see certain number of cups." & vbNewLine & vbNewLine & "There will be 1 cup on one side and multiple cups (either 2, 3, or 5) on the other side."
    If num = 2 Then Label2.Caption = "": Label3.Caption = "": Picture1 = LoadPicture(path & "\1-cup-demo.jpg"): Picture2 = LoadPicture(path & "\3-cup-demo.jpg"): Label1.Caption = vbNewLine & vbNewLine & "For each trial, you will be given the option of choosing a cup from either side by clicking on your choice. "
    If num = 3 Then Picture1 = LoadPicture(path & "\1-cup.jpg"): Picture2 = LoadPicture(path & "\3-cup.jpg"): Label1.Caption = "When the cups are shaded BLUE:" & vbNewLine & vbNewLine & "a) Choosing from one side will result in a 100% chance of winning money ($1)." & vbNewLine & vbNewLine & "b) Under 1 cup on the other side, you have a chance to win a larger prize (either       $2, $3, or $5). However, if you choose the wrong cup, you will win $0."
    If num = 4 Then Picture1.Visible = True: Picture2.Visible = True: Label2.Caption = "$1": Label3.Caption = "$3": Picture1 = LoadPicture(path & "\1-cup-lose.jpg"): Picture2 = LoadPicture(path & "\3-cup-lose.jpg"): Label1.Caption = "When the cups are shaded RED:" & vbNewLine & "a) You will be given a certain amount of money for each trial." & vbNewLine & "b) One side will take away $1 from your pile for certain." & vbNewLine & "c) The other side offers a chance that you will not lose any money. However, one      cup on this side is an incorrect choice. If you happen to choose this cup, you          will lose all of the money that were given to you for that trial." & vbNewLine & "d) Any money that you don't lose will be added to your total amount earned at the     end of the task."
    If num = 5 Then Label1.Caption = "                       Please read the payouts for each trial carefully." & vbNewLine & vbNewLine & "                              The sure payout side switches randomly."
    If num = 6 Then Command2.Visible = False: Label1.Caption = "The goal of the game is to win as much money as possible and avoid losing as much money as possible." & vbNewLine & vbNewLine & "               Your total amount will be displayed at the end of the game."
    If num < 0 Then num = 0
End If

End Sub
