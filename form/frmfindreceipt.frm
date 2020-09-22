VERSION 5.00
Begin VB.Form frmfindreceipt 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find Receipt"
   ClientHeight    =   1965
   ClientLeft      =   3015
   ClientTop       =   7125
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   0
      Picture         =   "frmfindreceipt.frx":0000
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
         Left            =   3000
         TabIndex        =   3
         Top             =   720
         Width           =   4320
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
         Left            =   5640
         TabIndex        =   2
         Top             =   120
         Width           =   1680
      End
      Begin VB.TextBox txtfind2 
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
         Left            =   3000
         TabIndex        =   1
         Top             =   120
         Width           =   1680
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
         Left            =   7080
         TabIndex        =   10
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
         Left            =   4920
         TabIndex        =   9
         Top             =   120
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
         Left            =   2280
         TabIndex        =   8
         Top             =   720
         Width           =   615
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
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
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
         Left            =   5640
         TabIndex        =   6
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt  No"
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
         Left            =   3000
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label cmdfind2 
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
         Left            =   2280
         TabIndex        =   4
         Top             =   120
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmfindreceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback_Click()
Me.Hide
End Sub

Private Sub cmdfind_Click()
Dim CurrentRec As Variant
If rsfeepayment.State = closed Then
Call feespaymentopen
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
rsfeepayment.MoveFirst
rsfeepayment.Find "GRNO='" & txtfind & "'"
CurrentRec = rsfeepayment.AbsolutePosition
If rsfeepayment.EOF = True Then
   rsfeepayment.MoveFirst
   MsgBox "Record is not found", vbCritical, "Find"
   txtfind.Text = ""
   txtfind.SetFocus
Else
    rsfeepayment.AbsolutePosition = CurrentRec
    frmfeepayment.dg.Bookmark = CurrentRec
    Me.Hide
    Call showdata
End If
End If

End Sub

Private Sub cmdfind2_Click()
Dim CurrentRec As Variant
If rsfeepayment.State = closed Then
Call feespaymentopen
End If
If Trim(txtfind2.Text) = "" Then
    MsgBox "Please Write Receipt No Not Empty Box", vbInformation, "Find"
    txtfind2.SetFocus
    Exit Sub
ElseIf Not IsNumeric(txtfind2.Text) Then
    MsgBox "Please Write Numeric Value", vbInformation, "Find"
    txtfind2.Text = ""
    txtfind2.SetFocus
    Exit Sub
Else
rsfeepayment.MoveFirst
rsfeepayment.Find "RECEIPTNO='" & txtfind2 & "'"
CurrentRec = rsfeepayment.AbsolutePosition
If rsfeepayment.EOF = True Then
   rsfeepayment.MoveFirst
   MsgBox "Record is not found", vbCritical, "Find"
   txtfind2.Text = ""
   txtfind2.SetFocus
Else
    rsfeepayment.AbsolutePosition = CurrentRec
    frmfeepayment.dg.Bookmark = CurrentRec
    Me.Hide
    Call showdata
End If
End If

End Sub

Private Sub Form_Load()
If rsfeepayment.State = adStateOpen Then rsfeepayment.Close
rsfeepayment.Open "select * from feespayment", cn, adOpenStatic, adLockOptimistic

txtfind = ""
txtfind1 = ""
txtfind2 = ""
End Sub

Private Sub cmdFind1_Click()
Dim CurrentRec As Variant
If rsfeepayment.State = closed Then
Call feespaymentopen
End If
If txtfind1.Text = "" Then
    MsgBox "Please Write Student Name.", vbInformation, "Find"
    txtfind1.SetFocus
    Exit Sub
Else
    With rsfeepayment
    .MoveFirst
    .Find "NAME='" & txtfind1 & "'"
     CurrentRec = .AbsolutePosition
        If .EOF Then
            MsgBox "Record is Not Found.", vbCritical, "Find"
            txtfind1.Text = ""
            txtfind1.SetFocus
            .AbsolutePosition = 1
        Else
            frmfeepayment.dg.Bookmark = CurrentRec
            Me.Hide
            Call showdata
        End If
        End With
End If
End Sub
Public Sub feespaymentopen()
rsfeepayment.Open "select * from feespayment", cn, adOpenStatic, adLockOptimistic
End Sub

Public Sub showdata()
With rsfeepayment
If .RecordCount > 0 Then
frmfeepayment.txtreceiptno.Text = .Fields("RECEIPTNO")
frmfeepayment.txtdate.Text = .Fields("DATE")
frmfeepayment.cmbmonth.Text = .Fields("FEEMONTH")
frmfeepayment.cmbgrno.Text = .Fields("GRNO")
frmfeepayment.txtlevelcode.Text = .Fields("LEVELCODE")
frmfeepayment.txtadmissionfee.Text = Format(.Fields("ADMISSIONFEE"), "###,###,##0.00")
frmfeepayment.txtmonthfee.Text = Format(.Fields("MONTHLYFEE"), "###,###,##0.00")
frmfeepayment.lbltotalamount.Caption = Format(.Fields("AMOUNT"), "###,###,##0.00")
frmfeepayment.txtcontectno.Text = .Fields("CONTECTNO")
frmfeepayment.txtremarks.Text = .Fields("REMARKS")
frmfeepayment.txtname.Text = .Fields("NAME")
frmfeepayment.txtdescr.Text = .Fields("DESCR")
frmfeepayment.dg.Refresh
End If
End With
CurrentRec = 0
End Sub


