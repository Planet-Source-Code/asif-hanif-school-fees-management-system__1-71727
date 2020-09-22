VERSION 5.00
Begin VB.Form frmplaceholder 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Picture Chooser"
   ClientHeight    =   5220
   ClientLeft      =   5055
   ClientTop       =   3390
   ClientWidth     =   5490
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2520
      TabIndex        =   8
      Top             =   4080
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   4680
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   120
      Pattern         =   "*.JPG;*.JPEG;*.BMP"
      TabIndex        =   2
      Top             =   3480
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   1890
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   300
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Selection Path :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   195
      Left            =   2520
      TabIndex        =   9
      Top             =   3840
      Width           =   1380
   End
   Begin VB.Image pic 
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   2558
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Preview Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   300
      Left            =   2880
      TabIndex        =   7
      Top             =   840
      Width           =   2130
   End
   Begin VB.Label Label3 
      Caption         =   "File Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Folder Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Disk Drive"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmplaceholder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub Command1_Click()
frmstudent.picc.Picture = LoadPicture(Text1)
'frmstudent.txtpict.Text = Dir1.Path + "\" + File1.FileName
'frmstudent.txtpict.Text = "\picture\" + File1.FileName
frmstudent.txtpict.Text = File1.FileName
Unload Me
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Dim st As String
st = Dir1.Path + "\" + File1.FileName
Text1 = st
pic.Picture = LoadPicture(Text1)
End Sub

Private Sub Form_Load()
Me.Top = Val(Screen.Height - Me.Height) / 2
Me.Left = Val(Screen.Width - Me.Width) / 2
End Sub
