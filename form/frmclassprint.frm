VERSION 5.00
Begin VB.Form frmclassprint 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Class & Subjects Report"
   ClientHeight    =   4785
   ClientLeft      =   3750
   ClientTop       =   3570
   ClientWidth     =   7530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmclassprint.frx":0000
   ScaleHeight     =   4785
   ScaleWidth      =   7530
   Begin VB.Frame frmselectlevel 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   480
      TabIndex        =   5
      Top             =   2160
      Width           =   6615
      Begin VB.ComboBox cmblevelfrom 
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
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmblevelto 
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
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtlevelfrom 
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
         Left            =   3360
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   7
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtlevelto 
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
         Left            =   3360
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   6
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Level Code  From :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Level Code  To     :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   780
         Width           =   1695
      End
   End
   Begin VB.OptionButton optselectlevel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Selection By Level Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Print"
      Height          =   855
      Left            =   2160
      Picture         =   "frmclassprint.frx":0467
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Back"
      Height          =   855
      Left            =   3720
      Picture         =   "frmclassprint.frx":0D84
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.OptionButton optall 
      BackColor       =   &H00C0C0C0&
      Caption         =   "All Level  && Subjects"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1320
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PRINT CLASS && SUBJECTS  REPORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   -840
      Picture         =   "frmclassprint.frx":1832
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15855
   End
   Begin VB.Shape Shape1 
      Height          =   3495
      Left            =   240
      Top             =   1200
      Width           =   7095
   End
End
Attribute VB_Name = "frmclassprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstemp As New ADODB.Recordset
Dim rsprint As New ADODB.Recordset

Private Sub cmblevelfrom_Click()
With rstemp
If .State = 1 Then
.Close
End If
rstemp.Open "Select * from level_details where levelcode='" & cmblevelfrom & "'", cn, adOpenDynamic, adLockOptimistic
If .RecordCount <= 0 Then
txtlevelfrom.Text = ""
Else
txtlevelfrom.Text = UCase(.Fields("DESCR"))
.Close
End If
End With
End Sub


Private Sub cmblevelto_Click()
With rstemp
If .State = 1 Then
.Close
End If
rstemp.Open "Select * from level_details where levelcode='" & cmblevelto & "'", cn, adOpenDynamic, adLockOptimistic
If .RecordCount <= 0 Then
txtlevelto.Text = ""
Else
txtlevelto.Text = UCase(.Fields("DESCR"))
.Close
End If
End With

End Sub

Private Sub cmdback_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
If optall.Value = True Then
With rsprint
    If .State = adStateOpen Then .Close
    .Open "SELECT * from subjects order by levelcode ", cn, adOpenForwardOnly, adLockReadOnly
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        Exit Sub
    End If
    Set rptclass.DataSource = rsprint
    rptclass.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptclass.Sections("section4").Controls("lbladdress1").Caption = address1
    rptclass.Sections("section4").Controls("lbladdress2").Caption = address2
    rptclass.Sections("section4").Controls("lblreportname").Caption = "REPORT FOR ALL CLASS & SUBJECTS"
    rptclass.Sections("section4").Controls("lblusername").Caption = pubusername
    rptclass.Sections("section4").Controls("lbldate").Caption = Now
    rptclass.Show vbModal
End With
End If
If optselectlevel.Value = True Then
If Trim(cmblevelfrom.Text) = "" Then
MsgBox "Write Level Code From", vbInformation, "Empty Level Code"
cmblevelfrom.SetFocus
Exit Sub
ElseIf Trim(cmblevelto.Text) = "" Then
MsgBox "Write Level Code To", vbInformation, "Empty Level Code"
cmblevelto.SetFocus
Exit Sub
End If
With rsprint
    If .State = adStateOpen Then .Close
  '  .Open "SELECT * from student_details where GRNO='" & cmbgrnofrom & "'", cn, adOpenForwardOnly, adLockReadOnly
 .Open "SELECT * from subjects where levelcode between """ & cmblevelfrom & """ And """ & cmblevelto & """", cn, adOpenForwardOnly, adLockReadOnly
    
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        cmblevelfrom.SetFocus
        Exit Sub
    End If
    Set rptclass.DataSource = rsprint
    rptclass.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptclass.Sections("section4").Controls("lbladdress1").Caption = address1
    rptclass.Sections("section4").Controls("lbladdress2").Caption = address2
    rptclass.Sections("section4").Controls("lblreportname").Caption = "REPORT FOR SELECTED LEVEL CODE"
    rptclass.Sections("section4").Controls("lblusername").Caption = pubusername
    rptclass.Sections("section4").Controls("lbldate").Caption = Now
    rptclass.Show vbModal
End With
cmblevelfrom.SetFocus
End If
End Sub

Private Sub Form_Load()
Call conopen
If rssubjects.State = 0 Then
rssubjects.Open "select * from subjects order by levelcode", cn, adOpenDynamic, adLockOptimistic
End If


With rssubjects
If .State = closed Then
rssubjects.Open "select * from subjects order by levelcode", cn, adOpenDynamic, adLockOptimistic
End If
.AbsolutePosition = 1
While .EOF = False
cmblevelfrom.AddItem .Fields("LEVELCODE")
cmblevelto.AddItem .Fields("LEVELCODE")
.MoveNext
Wend
End With

frmselectlevel.Visible = False

optall.Value = True


End Sub

Private Sub optall_Click()
If optall.Value = True Then
frmselectlevel.Visible = False
cmdprint.SetFocus
End If
End Sub

Private Sub optselectlevel_Click()
If optselectlevel.Value = True Then
frmselectlevel.Visible = True
cmblevelfrom.SetFocus
End If
End Sub

