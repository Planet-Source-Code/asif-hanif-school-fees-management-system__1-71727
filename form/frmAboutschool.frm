VERSION 5.00
Begin VB.Form frmAboutschool 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About School"
   ClientHeight    =   4350
   ClientLeft      =   5775
   ClientTop       =   4815
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3002.448
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFC0C0&
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
   Begin VB.Image imgLogo 
      Height          =   2145
      Left            =   120
      Picture         =   "frmAboutschool.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lbladdress2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "fdsfdsfdsfsfd"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   2280
      Width           =   3570
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAboutschool.frx":A27B
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
      TabIndex        =   6
      Top             =   3000
      Width           =   5295
   End
   Begin VB.Label lbladdress1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dsfsdfsdfsadfas"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   1800
      Width           =   3570
      WordWrap        =   -1  'True
   End
   Begin VB.Label MessageList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   1440
      Width           =   930
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "School Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   1395
   End
   Begin VB.Label lblschoolname 
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
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
      Width           =   3570
      WordWrap        =   -1  'True
   End
   Begin VB.Label MessageList 
      Alignment       =   2  'Center
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
      TabIndex        =   1
      Top             =   120
      Width           =   5355
      WordWrap        =   -1  'True
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
End
Attribute VB_Name = "frmAboutschool"
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

Private Sub Form_Load()
lblschoolname.Caption = schoolname
lbladdress1.Caption = address1
lbladdress2.Caption = address2
End Sub
