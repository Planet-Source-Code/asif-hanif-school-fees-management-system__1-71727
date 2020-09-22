VERSION 5.00
Begin VB.Form frmClickme 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CLICK ME"
   ClientHeight    =   6585
   ClientLeft      =   2625
   ClientTop       =   2730
   ClientWidth     =   9615
   FillColor       =   &H0000C000&
   Icon            =   "Clickme.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Clickme.frx":0152
   MousePointer    =   99  'Custom
   ScaleHeight     =   6585
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   2880
      Top             =   5760
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H008080FF&
      Caption         =   "    CLICK ME      WHER ARE YOU GOING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6960
      MaskColor       =   &H00FFFF00&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "    CLICK ME      WHER ARE YOU GOING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      MaskColor       =   &H00FFFF00&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "    CLICK ME      WHER ARE YOU GOING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      MaskColor       =   &H00FFFF00&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0FF&
      Caption         =   "    CLICK ME      WHER ARE YOU GOING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6960
      MaskColor       =   &H00FFFF00&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "    CLICK ME      WHER ARE YOU GOING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      MaskColor       =   &H00FFFF00&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "    CLICK ME      WHER ARE YOU GOING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      MaskColor       =   &H00FFFF00&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "    CLICK ME      WHER ARE YOU GOING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      MaskColor       =   &H00FFFF00&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "    CLICK ME      WHER ARE YOU GOING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      MaskColor       =   &H00FFFF00&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "    CLICK ME      WHER ARE YOU GOING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      MaskColor       =   &H00FFFF00&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "    CLICK ME      WHER ARE YOU GOING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      MaskColor       =   &H00FFFF00&
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
End
Attribute VB_Name = "frmClickme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim ans As Integer




Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Visible = False
Command2.Visible = True
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command9.Visible = False
Command10.Visible = False
End Sub


Private Sub Command10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Visible = True
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command9.Visible = False
Command10.Visible = False
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Visible = False
Command2.Visible = False
Command3.Visible = True
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command9.Visible = False
Command10.Visible = False
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = True
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command9.Visible = False
Command10.Visible = False
End Sub


Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = True
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command9.Visible = False
Command10.Visible = False
End Sub


Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = True
Command7.Visible = False
Command8.Visible = False
Command9.Visible = False
Command10.Visible = False
End Sub


Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = True
Command8.Visible = False
Command9.Visible = False
Command10.Visible = False
End Sub


Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = True
Command9.Visible = False
Command10.Visible = False
End Sub


Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command9.Visible = True
Command10.Visible = False
End Sub


Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command9.Visible = False
Command10.Visible = True
End Sub

Private Sub Form_Load()
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command9.Visible = False
Command10.Visible = True
Timer1.Interval = 1000
Timer1.Enabled = True
i = 0


End Sub

Private Sub Timer1_Timer()

i = i + 1
Me.Caption = "Start Time    " & i
If i = 10 Then
ans = MsgBox("Game is Over,Do You Want Play Agan", vbYesNo + vbDefaultButton2, "Play Again")
If ans = vbYes Then
Timer1.Enabled = True
i = 0
Else
Timer1.Enabled = False
i = 0
Me.Hide
End If
End If
End Sub
