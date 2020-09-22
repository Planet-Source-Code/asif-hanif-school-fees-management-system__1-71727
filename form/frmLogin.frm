VERSION 5.00
Begin VB.Form frmlogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3060
   ClientLeft      =   4755
   ClientTop       =   4110
   ClientWidth     =   4485
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4485
   Begin VB.Timer Timer3 
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox ctrlLiner2 
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   6615
      TabIndex        =   9
      Top             =   960
      Width           =   6615
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   240
      ScaleHeight     =   30
      ScaleWidth      =   4095
      TabIndex        =   6
      Top             =   2280
      Width           =   4095
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00ECF4F4&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Picture         =   "frmLogin.frx":08CA
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00ECF4F4&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      Picture         =   "frmLogin.frx":1264
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtpassword 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox txtusername 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      TabIndex        =   0
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Login !"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   240
      Picture         =   "frmLogin.frx":1BFE
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type Username and Password to login."
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   600
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim rsuser As New ADODB.Recordset
Dim txtuserid As String
Dim bhide As Boolean



Private Sub cmdCancel_Click()
    Timer3.Interval = 30
    bhide = True
'    frmmain.Show
'    Me.Hide
End Sub

Private Sub cmdOK_Click()
Dim CurrentRec As Variant
 With rsuser
 If .State = closed Then
 .Open
 End If
 If txtusername.Text = "" Then
 MsgBox "Write User Name", vbInformation, "Empty User"
 txtusername.SetFocus
 Exit Sub
 End If
 If txtpassword.Text = "" Then
 MsgBox "Write Password", vbInformation, "Empty Password"
 txtpassword.SetFocus
 Exit Sub
 End If
 .MoveFirst
 .Find "USERNAME='" & txtusername & "'"
 If .EOF = True Then
 MsgBox "Invalid User Name", vbInformation, "Invalid User Name"
 txtusername.Text = ""
 txtusername.SetFocus
 Exit Sub
 End If
If txtpassword.Text <> .Fields("password") Then
    MsgBox "Wrong Password", vbCritical, "Wrong Password"
    txtpassword.SetFocus
    txtpassword.Text = ""
    Exit Sub
End If
txtuserid = .Fields("USERID")
End With
Call nameopen
schoolname = rsname.Fields("schoolname")
address1 = rsname.Fields("address1")
address2 = rsname.Fields("address2")
pubusername = UCase(txtusername.Text)
pubpassword = txtpassword.Text
pubuserid = txtuserid
Call userlogopen
With rsuserlog
.AddNew
.Fields("USERID") = txtuserid
.Fields("USERNAME") = UCase(txtusername.Text)
.Fields("PASSWORD") = txtpassword.Text
.Fields("LOGINTIME") = Now
.Update
.Close
End With
Timer2.Interval = 30
bhide = True

'frmmain.Show
'frmmain.menu.Visible = True

'Me.Hide
'frmmain.cmdmarksheet.Enabled = True
'frmmain.cmdnew.Enabled = True
'frmmain.cmdsubject.Enabled = True
'frmmain.cmdviewall.Enabled = True
'frmmain.cmdfeepayment.Enabled = True
'frmmain.cmduser.Enabled = True
'frmmain.cmdLogin.Enabled = False

'frmmain.cmdLogin.Visible = False
End Sub

Private Sub Form_Activate()
Me.Height = 0
bhide = False
Timer1.Interval = 30
End Sub

Private Sub Form_Load()
Call conopen
With rsuser
If .State = closed Then
Call useropen
End If
End With
txtusername.Text = ""
txtpassword.Text = ""
End Sub



Private Sub Timer2_Timer()
If bhide = False Then
    If Me.Height >= 3570 Then
        Me.Width = 4605
        Me.Height = 3570
        Timer1.Interval = 0
        Else
        Me.Height = Me.Height + 300
    End If
Else
 If Me.Height <= 600 Then
        Me.Width = 0
        Me.Height = 0
        Timer1.Interval = 0
        If Me.Height <= 510 And Me.Width <= 1005 Then
            frmmain.menu.Visible = True
            frmmain.cmdLogin.Visible = False
            frmmain.Show
            Unload Me
        End If
        Else
        Me.Height = Me.Height - 300
        DoEvents
    End If
End If

 Me.Top = (Screen.Height / 2) - (Me.Height / 2)

End Sub

Private Sub Timer3_Timer()
If bhide = False Then
    If Me.Height >= 3570 Then
        Me.Width = 4605
        Me.Height = 3570
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
If Me.Height <= 0 And Me.Width <= 0 Then
frmmain.Show
End If
End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdok.SetFocus
End If
End Sub

Private Sub txtusername_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtpassword.SetFocus
End If
End Sub
Private Sub Timer1_Timer()
If bhide = False Then
    If Me.Height >= 3570 Then
        Me.Width = 4605
        Me.Height = 3570
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


