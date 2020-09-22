VERSION 5.00
Begin VB.Form frmfindlevel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find Level "
   ClientHeight    =   1965
   ClientLeft      =   2910
   ClientTop       =   7395
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   0
      Picture         =   "frmfindlevel.frx":0000
      ScaleHeight     =   1995
      ScaleWidth      =   8295
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.TextBox txtfind1 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5040
         TabIndex        =   2
         Top             =   240
         Width           =   2640
      End
      Begin VB.TextBox txtfind 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         TabIndex        =   1
         Top             =   240
         Width           =   1680
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Level Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   7
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Level Code"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label cmdback 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   5
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label cmdfind 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Label cmdfind1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmfindlevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
Me.Hide
End Sub

Private Sub cmdfind_Click()
Dim CurrentRec As Variant
If rssubjects.State = closed Then
Call subjectsopen
End If
If Trim(txtfind.Text) = "" Then
    MsgBox "Please Write Level Code Not Empty Box", vbInformation, "Find"
    txtfind.SetFocus
    Exit Sub
Else
rssubjects.MoveFirst
rssubjects.Find "LEVELCODE='" & txtfind & "'"
CurrentRec = rssubjects.AbsolutePosition
If rssubjects.EOF = True Then
   rssubjects.MoveFirst
   MsgBox "Record is not found", vbCritical, "Find"
   txtfind.Text = ""
   txtfind.SetFocus
Else
    rssubjects.AbsolutePosition = CurrentRec
    Me.Hide
    Call showdata
    frmsubjects.txtlevelcode.SetFocus
End If
End If
End Sub


Private Sub Form_Load()
txtfind = ""
txtfind1 = ""
End Sub

Private Sub cmdFind1_Click()
Dim CurrentRec As Variant
If rssubjects.State = closed Then
Call subjectsopen
End If
If txtfind1.Text = "" Then
    MsgBox "Please Write Level Name.", vbInformation, "Find"
    txtfind1.SetFocus
    Exit Sub
Else
    With rssubjects
    .MoveFirst
    .Find "DESCR='" & txtfind1 & "'"
     CurrentRec = .AbsolutePosition
        If .EOF Then
            MsgBox "Record is Not Found.", vbCritical, "Find"
            txtfind1.Text = ""
            txtfind1.SetFocus
            .AbsolutePosition = 1
        Else
            Me.Hide
            Call showdata
            frmsubjects.txtlevelcode.SetFocus
        End If
        End With
End If
End Sub
Public Sub showdata()
With rssubjects
If .RecordCount > 0 Then
frmsubjects.txtlevelcode.Text = .Fields("LEVELCODE")
frmsubjects.txtdescr.Text = .Fields("DESCR")

' subjects

 frmsubjects.txtsubject(0).Text = .Fields("SUB01")
 frmsubjects.txtsubject(1).Text = .Fields("SUB02")
 frmsubjects.txtsubject(2).Text = .Fields("SUB03")
 frmsubjects.txtsubject(3).Text = .Fields("SUB04")
 frmsubjects.txtsubject(4).Text = .Fields("SUB05")
 frmsubjects.txtsubject(5).Text = .Fields("SUB06")
 frmsubjects.txtsubject(6).Text = .Fields("SUB07")
 frmsubjects.txtsubject(7).Text = .Fields("SUB08")
 frmsubjects.txtsubject(8).Text = .Fields("SUB09")
 frmsubjects.txtsubject(9).Text = .Fields("SUB10")
 frmsubjects.txtsubject(10).Text = .Fields("SUB11")
 frmsubjects.txtsubject(11).Text = .Fields("SUB12")

' maximum marks

frmsubjects.txtmax(0).Text = .Fields("MAX01")
frmsubjects.txtmax(1).Text = .Fields("MAX02")
frmsubjects.txtmax(2).Text = .Fields("MAX03")
frmsubjects.txtmax(3).Text = .Fields("MAX04")
frmsubjects.txtmax(4).Text = .Fields("MAX05")
frmsubjects.txtmax(5).Text = .Fields("MAX06")
frmsubjects.txtmax(6).Text = .Fields("MAX07")
frmsubjects.txtmax(7).Text = .Fields("MAX08")
frmsubjects.txtmax(8).Text = .Fields("MAX09")
frmsubjects.txtmax(9).Text = .Fields("MAX10")
frmsubjects.txtmax(10).Text = .Fields("MAX11")
frmsubjects.txtmax(11).Text = .Fields("MAX12")

'minimum marks

frmsubjects.txtmin(0).Text = .Fields("MIN01")
frmsubjects.txtmin(1).Text = .Fields("MIN02")
frmsubjects.txtmin(2).Text = .Fields("MIN03")
frmsubjects.txtmin(3).Text = .Fields("MIN04")
frmsubjects.txtmin(4).Text = .Fields("MIN05")
frmsubjects.txtmin(5).Text = .Fields("MIN06")
frmsubjects.txtmin(6).Text = .Fields("MIN07")
frmsubjects.txtmin(7).Text = .Fields("MIN08")
frmsubjects.txtmin(8).Text = .Fields("MIN09")
frmsubjects.txtmin(9).Text = .Fields("MIN10")
frmsubjects.txtmin(10).Text = .Fields("MIN11")
frmsubjects.txtmin(11).Text = .Fields("MIN12")

' charges fee

frmsubjects.txtaddmission.Text = Format(.Fields("addmissionfee"), "###,###0.00")
frmsubjects.txtmonthly.Text = Format(.Fields("monthlyfee"), "###,###0.00")
frmsubjects.txttution.Text = Format(.Fields("tutionfee"), "###,###0.00")
frmsubjects.txtexamination.Text = Format(.Fields("examinationfee"), "###,###0.00")
frmsubjects.txtcomputer.Text = Format(.Fields("computerfee"), "###,###0.00")
frmsubjects.txtlab.Text = Format(.Fields("labfee"), "###,###0.00")
frmsubjects.txtother.Text = Format(.Fields("otherfee"), "###,###0.00")
frmsubjects.txttotal.Text = Format(.Fields("totalfee"), "###,###0.00")
End If
End With
End Sub


