VERSION 5.00
Begin VB.Form frmfindstudent 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find Student"
   ClientHeight    =   1965
   ClientLeft      =   2835
   ClientTop       =   7755
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   2115
      Left            =   0
      Picture         =   "frmfindstudent.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   8295
      TabIndex        =   0
      Top             =   0
      Width           =   8295
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
         Left            =   2040
         TabIndex        =   2
         Top             =   120
         Width           =   1680
      End
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
         Left            =   4680
         TabIndex        =   1
         Top             =   120
         Width           =   2640
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Student GR No"
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
         Left            =   2040
         TabIndex        =   7
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Name"
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
         Left            =   4680
         TabIndex        =   6
         Top             =   480
         Width           =   1695
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
         Left            =   3960
         TabIndex        =   5
         Top             =   120
         Width           =   615
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
         Left            =   1320
         TabIndex        =   4
         Top             =   120
         Width           =   615
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
         Left            =   6600
         TabIndex        =   3
         Top             =   1320
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmfindstudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
Me.Hide
End Sub

Private Sub cmdfind_Click()
Dim CurrentRec As Variant
If rsSTUDENT.State = closed Then
Call studentopen
End If
If Trim(txtfind.Text) = "" Then
    MsgBox "Please Write GrNo Not Empty Box", vbInformation, "Find"
    txtfind.SetFocus
    Exit Sub
ElseIf Not IsNumeric(txtfind.Text) Then
    MsgBox "Please Write Numeric Value", vbInformation, "Find"
    txtfind.Text = ""
    txtfind.SetFocus
    Exit Sub
Else
rsSTUDENT.MoveFirst
rsSTUDENT.Find "GRNO='" & txtfind & "'"
CurrentRec = rsSTUDENT.AbsolutePosition
If rsSTUDENT.EOF = True Then
   rsSTUDENT.MoveFirst
   MsgBox "Record is not found", vbCritical, "Find"
   txtfind.Text = ""
   txtfind.SetFocus
Else
    rsSTUDENT.AbsolutePosition = CurrentRec
    frmstudent.dg.Bookmark = CurrentRec
    Call showdata
    Me.Hide
End If
End If
End Sub

Private Sub Form_Load()
txtfind = ""
txtfind1 = ""
End Sub

Private Sub cmdFind1_Click()
Dim CurrentRec As Variant
If rsSTUDENT.State = closed Then
Call studentopen
End If
If txtfind1.Text = "" Then
    MsgBox "Please Write Student Name.", vbInformation, "Find"
    txtfind1.SetFocus
    Exit Sub
Else
    With rsSTUDENT
    .MoveFirst
    .Find "NAME='" & txtfind1 & "'"
     CurrentRec = .AbsolutePosition
        If .EOF Then
            MsgBox "Record is Not Found.", vbCritical, "Find"
            txtfind1.Text = ""
            txtfind1.SetFocus
            .AbsolutePosition = 1
        Else
            frmstudent.dg.Bookmark = CurrentRec
            Call showdata
            Me.Hide
        End If
        End With
End If
End Sub

Public Sub showdata()
With rsSTUDENT
If .RecordCount > 0 Then
frmstudent.txtgrno.Text = .Fields(0)
frmstudent.txtname.Text = .Fields(1)
frmstudent.txtfname.Text = .Fields(2)
frmstudent.txtaddress.Text = .Fields(3)
frmstudent.txtdob.Text = .Fields(4)
frmstudent.txttel.Text = .Fields(5)
frmstudent.txtmobile.Text = .Fields(6)
frmstudent.cmbage.Text = .Fields(7)
frmstudent.cmbsex.Text = .Fields(8)
frmstudent.cmbreligion.Text = .Fields(9)
frmstudent.cmbpqualification.Text = .Fields(10)
frmstudent.txtpinstitution.Text = .Fields(11)
frmstudent.cmbclass.Text = .Fields(12)
frmstudent.cmbshift.Text = .Fields(13)
frmstudent.cmbclasstime.Text = .Fields(14)
frmstudent.txtpict.Text = .Fields(15)
If frmstudent.txtpict.Text = "" Then
On Error Resume Next
frmstudent.picc.Picture = Nothing
Else
frmstudent.picc.Picture = LoadPicture(App.Path & "\picture\" & frmstudent.txtpict.Text)
End If
frmstudent.cmbstatus.Text = .Fields(16)
End If
End With
End Sub

