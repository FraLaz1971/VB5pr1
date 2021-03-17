VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   63
   ClientTop       =   342
   ClientWidth     =   8910
   BeginProperty Font 
   EndProperty
   Font            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Quit 
      Caption         =   "Quit"
      BeginProperty Font 
      EndProperty
      Font            =   "Form1.frx":0018
      Height          =   360
      Left            =   6669
      TabIndex        =   4
      Top             =   117
      Width           =   1062
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LoadImage"
      BeginProperty Font 
      EndProperty
      Font            =   "Form1.frx":0030
      Height          =   360
      Left            =   5148
      TabIndex        =   3
      Top             =   117
      Width           =   1062
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
      EndProperty
      Font            =   "Form1.frx":0048
      Height          =   1530
      Left            =   3861
      ScaleHeight     =   1494
      ScaleWidth      =   4536
      TabIndex        =   2
      Top             =   1755
      Width           =   4572
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SQRT"
      BeginProperty Font 
      EndProperty
      Font            =   "Form1.frx":0060
      Height          =   360
      Left            =   4095
      TabIndex        =   1
      Top             =   117
      Width           =   828
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
      EndProperty
      Font            =   "Form1.frx":0078
      Height          =   1179
      Left            =   3861
      TabIndex        =   0
      Text            =   "Hello textbox"
      Top             =   585
      Width           =   4572
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
100 Rem THIS PROGRAM EXTRACTS SQUARE ROOTS
105 CLS
110 Print "*************************"
120 Print "*** THIS PROGRAM"
130 Print "*** EXTRACTS"
135 Print "*** SQUARE ROOTS"
140 Print "*************************"
150 Dim A As Integer
160 Text1.Text = " "
220 A = InputBox("Can you input a number?", "Expecting a number", "4")
230 If A = -1 Then GoTo 990
990 If A < -1 Then MsgBox "negative numbers not handled ", vbExclamation: GoTo 220
1000 Text1.Text = "the square root of " & A & " is " & SQR(A)
End Sub

Private Sub Command2_Click()
 'Picture1.AutoSize = True
 Picture1.Visible = True
 Picture1.Picture = LoadPicture("JupiterMoonToEarth.jpg")
 Dim s As Double
 Let s = 18
    Picture1.ScaleMode = 3
    Picture1.AutoRedraw = True
    Picture1.PaintPicture Picture1.Picture, _
    0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, _
    0, 0, Picture1.Picture.Width / s, _
    Picture1.Picture.Height / s
    Picture1.Picture = Picture1.Image
    Picture1.Refresh

End Sub

Private Sub Quit_Click()
    Unload Form1
    End ' Ends application.
End Sub
