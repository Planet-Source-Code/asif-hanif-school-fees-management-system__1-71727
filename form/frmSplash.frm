VERSION 5.00
Object = "{CED4D45D-8DB4-42E5-A708-BCA2899963C8}#2.0#0"; "ProgressBar.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4035
   ClientLeft      =   3990
   ClientTop       =   3375
   ClientWidth     =   7095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   4050
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7080
      Begin progressbarocx.ProgressBar ProgressBar1 
         Height          =   450
         Left            =   1440
         TabIndex        =   6
         Top             =   2760
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   794
      End
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   5400
         Top             =   480
      End
      Begin VB.Line Line2 
         BorderStyle     =   2  'Dash
         BorderWidth     =   2
         X1              =   2400
         X2              =   6240
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         BorderWidth     =   2
         X1              =   1920
         X2              =   6720
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "Please Wait..... Application is Loading"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2235
         TabIndex        =   5
         Top             =   2400
         Width           =   3405
      End
      Begin VB.Image imgLogo 
         Height          =   1665
         Left            =   240
         Picture         =   "frmSplash.frx":0000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H80000013&
         Caption         =   "Copyright :  Asif Hanif"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   3420
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H80000013&
         Caption         =   "Company :  future trick solutions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   3630
         Width           =   2415
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "Version :  1.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   5610
         TabIndex        =   3
         Top             =   3600
         Width           =   1275
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         Caption         =   "JOHAR PUBLIC SCHOOL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   465
         Left            =   1920
         TabIndex        =   4
         Top             =   1080
         Width           =   4920
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim RetValue
RetValue = ChangeRes(1024, 768, 32)
End Sub



Private Sub Timer1_Timer()
With ProgressBar1
    .Percent = .Percent + 2
    If .Percent = 100 Then
        Unload Me
        frmmain.Show
    End If
End With
End Sub
