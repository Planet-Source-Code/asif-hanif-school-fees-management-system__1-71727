VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Programmer"
   ClientHeight    =   4455
   ClientLeft      =   5715
   ClientTop       =   4155
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3074.921
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   1260
   End
   Begin VB.Label MessageList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email :future_trick1@hotmail.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   3450
      WordWrap        =   -1  'True
   End
   Begin VB.Label MessageList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email :future_trick1@yahoo.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   6
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   3450
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1065
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   5295
   End
   Begin VB.Label MessageList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contect No 0333 -  3506186"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   2850
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Develped By :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1485
   End
   Begin VB.Label MessageList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FUTURE TRICK SOLUTIONS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   3210
      WordWrap        =   -1  'True
   End
   Begin VB.Label MessageList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "School Management System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   3555
      WordWrap        =   -1  'True
   End
   Begin VB.Label MessageList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2008"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   1530
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1695
      Left            =   3960
      Top             =   600
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1680
      Left            =   3960
      Picture         =   "frmAbout.frx":0125
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   197.201
      X2              =   5422.084
      Y1              =   1987.828
      Y2              =   1987.828
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   211.287
      X2              =   5422.084
      Y1              =   1987.828
      Y2              =   1987.828
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version  1.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1005
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bhide As Boolean

Private Sub cmdOK_Click()
Timer1.Interval = 30
bhide = True
'Unload Me
End Sub

Private Sub Form_Activate()
Me.Height = 0
bhide = False
Timer1.Interval = 30
End Sub

Private Sub Timer1_Timer()
If bhide = False Then
    If Me.Height >= 4965 Then
        Me.Width = 5850
        Me.Height = 4965
        Timer1.Interval = 0
        Else
        Me.Height = Me.Height + 300
    End If
Else
 If Me.Height <= 600 Then
        Me.Width = 0
        Me.Height = 0
        Timer1.Interval = 0
        Unload Me
        Else
        Me.Height = Me.Height - 300
        DoEvents
    End If
End If

 Me.Top = (Screen.Height / 2) - (Me.Height / 2)
End Sub

