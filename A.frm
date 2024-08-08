VERSION 5.00
Begin VB.Form A 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "简易计算器"
   ClientHeight    =   6255
   ClientLeft      =   2445
   ClientTop       =   1575
   ClientWidth     =   4605
   BeginProperty Font 
      Name            =   "楷体"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6096.491
   ScaleMode       =   0  'User
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "00"
      Height          =   718
      Left            =   480
      TabIndex        =   20
      Top             =   5040
      Width           =   700
   End
   Begin VB.CommandButton Command1 
      Caption         =   "."
      Height          =   718
      Left            =   2400
      TabIndex        =   19
      Top             =   5040
      Width           =   700
   End
   Begin VB.CommandButton button0 
      Caption         =   "0"
      Height          =   718
      Left            =   1440
      TabIndex        =   18
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton buttonmul 
      Caption         =   "*"
      Height          =   462
      Left            =   3360
      TabIndex        =   17
      Top             =   1320
      Width           =   700
   End
   Begin VB.CommandButton buttonsub 
      Caption         =   "-"
      Height          =   462
      Left            =   3360
      TabIndex        =   16
      Top             =   2760
      Width           =   700
   End
   Begin VB.CommandButton buttonadd 
      Caption         =   "+"
      Height          =   462
      Left            =   3360
      TabIndex        =   15
      Top             =   3480
      Width           =   700
   End
   Begin VB.CommandButton buttondivision 
      Caption         =   "/"
      Height          =   462
      Left            =   3360
      TabIndex        =   14
      Top             =   2040
      Width           =   700
   End
   Begin VB.CommandButton buttondelete 
      Caption         =   "<"
      Height          =   462
      Left            =   2400
      TabIndex        =   13
      Top             =   1320
      Width           =   700
   End
   Begin VB.CommandButton buttonsign 
      Caption         =   "%"
      Height          =   462
      Left            =   1440
      TabIndex        =   12
      Top             =   1320
      Width           =   700
   End
   Begin VB.CommandButton cls 
      Caption         =   "C"
      Height          =   462
      Left            =   480
      TabIndex        =   11
      Top             =   1320
      Width           =   700
   End
   Begin VB.CommandButton buttonequ 
      Caption         =   "="
      Height          =   1545
      Left            =   3360
      TabIndex        =   10
      Top             =   4200
      Width           =   700
   End
   Begin VB.CommandButton button9 
      Caption         =   "9"
      Height          =   718
      Left            =   2400
      TabIndex        =   9
      Top             =   2160
      Width           =   700
   End
   Begin VB.CommandButton button8 
      Caption         =   "8"
      Height          =   718
      Left            =   1440
      TabIndex        =   8
      Top             =   2160
      Width           =   700
   End
   Begin VB.CommandButton button7 
      Caption         =   "7"
      Height          =   718
      Left            =   480
      TabIndex        =   7
      Top             =   2160
      Width           =   700
   End
   Begin VB.CommandButton button6 
      Caption         =   "6"
      Height          =   718
      Left            =   2400
      TabIndex        =   6
      Top             =   3120
      Width           =   700
   End
   Begin VB.CommandButton button5 
      Caption         =   "5"
      Height          =   718
      Left            =   1440
      TabIndex        =   5
      Top             =   3120
      Width           =   700
   End
   Begin VB.CommandButton button4 
      Caption         =   "4"
      Height          =   718
      Left            =   480
      TabIndex        =   4
      Top             =   3120
      Width           =   700
   End
   Begin VB.CommandButton button3 
      Caption         =   "3"
      Height          =   718
      Left            =   2400
      TabIndex        =   3
      Top             =   4080
      Width           =   700
   End
   Begin VB.CommandButton button2 
      Caption         =   "2"
      Height          =   718
      Left            =   1440
      TabIndex        =   2
      Top             =   4080
      Width           =   700
   End
   Begin VB.CommandButton button1 
      Caption         =   "1"
      Height          =   718
      Left            =   480
      TabIndex        =   1
      Top             =   4080
      Width           =   700
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label ABOUT 
      BackStyle       =   0  'Transparent
      Caption         =   "about"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   7.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3720
      TabIndex        =   21
      Top             =   5880
      Width           =   495
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   4440
      Y1              =   5964.912
      Y2              =   5964.912
   End
   Begin VB.Line Line3 
      X1              =   4440
      X2              =   4440
      Y1              =   116.959
      Y2              =   5964.912
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4440
      Y1              =   116.959
      Y2              =   116.959
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   120
      Y1              =   116.959
      Y2              =   5964.912
   End
End
Attribute VB_Name = "A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim firstNumber As Double
Dim secondNumber As Double
Dim operation As String

Private Sub ABOUT_Click()
B.Show
End Sub

Private Sub button1_Click()
    AppendDigit "1"
End Sub

Private Sub button2_Click()
    AppendDigit "2"
End Sub

Private Sub button3_Click()
    AppendDigit "3"
End Sub

Private Sub button4_Click()
    AppendDigit "4"
End Sub

Private Sub button5_Click()
    AppendDigit "5"
End Sub

Private Sub button6_Click()
    AppendDigit "6"
End Sub

Private Sub button7_Click()
    AppendDigit "7"
End Sub

Private Sub button8_Click()
    AppendDigit "8"
End Sub

Private Sub button9_Click()
    AppendDigit "9"
End Sub

Private Sub button0_Click()
    AppendDigit "0"
End Sub

Private Sub buttonadd_Click()
    PerformOperation "+"
End Sub

Private Sub buttondelete_Click()
    If Len(Text1.Text) > 0 Then
        Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
    End If
End Sub


Private Sub buttonsub_Click()
    PerformOperation "-"
End Sub

Private Sub buttonmul_Click()
    PerformOperation "*"
End Sub

Private Sub buttondivision_Click()
    PerformOperation "/"
End Sub

Private Sub buttonequ_Click()
    CalculateResult
End Sub

Private Sub buttonsign_Click()
    Dim currentValue As Double
    currentValue = CDbl(Text1.Text)
    currentValue = currentValue / 100
    Text1.Text = Format(currentValue, "0.00")
End Sub

Private Sub cls_Click()
    ClearAll
End Sub

Private Sub Command1_Click()
    AppendDecimalPoint
End Sub

Private Sub Command2_Click()
    AppendTwoZeros
End Sub

Private Sub AppendDecimalPoint()
    If InStr(Text1.Text, ".") = 0 Then
        Text1.Text = Text1.Text & "."
    End If
End Sub

Private Sub AppendTwoZeros()
    Text1.Text = Text1.Text & "00"
End Sub


Private Sub Command4_Click()

End Sub

Private Sub Form_Load()
    ClearAll
End Sub

Private Sub AppendDigit(ByVal digit As String)
    If Text1.Text = "0" Then
        Text1.Text = digit
    Else
        Text1.Text = Text1.Text & digit
    End If
End Sub

Private Sub PerformOperation(ByVal op As String)
    firstNumber = CDbl(Text1.Text)
    operation = op
    Text1.Text = "0"
End Sub

Private Sub CalculateResult()
    secondNumber = CDbl(Text1.Text)
    Dim result As Double

    Select Case operation
        Case "+"
            result = firstNumber + secondNumber
        Case "-"
            result = firstNumber - secondNumber
        Case "*"
            result = firstNumber * secondNumber
        Case "/"
            If secondNumber <> 0 Then
                result = firstNumber / secondNumber
            Else
                MsgBox "除数不能为0！", vbExclamation
                Exit Sub
            End If
    End Select

    Text1.Text = result
End Sub

Private Sub ClearAll()
    Text1.Text = "0"
    firstNumber = 0
    secondNumber = 0
    operation = ""
End Sub
