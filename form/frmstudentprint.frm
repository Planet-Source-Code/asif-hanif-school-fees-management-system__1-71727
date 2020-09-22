VERSION 5.00
Begin VB.Form frmstudentprint 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Report"
   ClientHeight    =   6225
   ClientLeft      =   3735
   ClientTop       =   1965
   ClientWidth     =   7530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmstudentprint.frx":0000
   ScaleHeight     =   6225
   ScaleWidth      =   7530
   Begin VB.Frame frmselectgendar 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   480
      TabIndex        =   25
      Top             =   3240
      Width           =   2655
      Begin VB.ComboBox cmbgendar 
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
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Gendar :"
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
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Frame frmselectlevel 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   480
      TabIndex        =   18
      Top             =   3240
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   780
         Width           =   1695
      End
   End
   Begin VB.OptionButton optselectgendar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Selection By Gendar"
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
      TabIndex        =   17
      Top             =   2760
      Width           =   2175
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
      TabIndex        =   16
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Frame frmselectname 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   480
      TabIndex        =   13
      Top             =   3240
      Width           =   6615
      Begin VB.TextBox txtselectname 
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
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   14
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Name :"
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
         Index           =   7
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.OptionButton optselectname 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Selection By Student  Name"
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
      TabIndex        =   12
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Print"
      Height          =   855
      Left            =   2160
      Picture         =   "frmstudentprint.frx":0467
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Back"
      Height          =   855
      Left            =   3720
      Picture         =   "frmstudentprint.frx":0D84
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Frame frmselectgr 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   480
      TabIndex        =   2
      Top             =   3240
      Width           =   6615
      Begin VB.TextBox txtnameto 
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
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   9
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txtnamefrom 
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
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   8
         Top             =   240
         Width           =   3375
      End
      Begin VB.ComboBox cmbgrnoto 
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
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox cmbgrnofrom 
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
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "GR No To     :"
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
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "GR No From :"
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
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.OptionButton optselectgr 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Selection By Student GR No"
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
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
   End
   Begin VB.OptionButton optall 
      BackColor       =   &H00C0C0C0&
      Caption         =   "All Students"
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
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PRINT REPORT"
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
      TabIndex        =   7
      Top             =   120
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   -840
      Picture         =   "frmstudentprint.frx":1832
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15855
   End
   Begin VB.Shape Shape1 
      Height          =   4815
      Left            =   240
      Top             =   1200
      Width           =   7095
   End
End
Attribute VB_Name = "frmstudentprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstemp As New ADODB.Recordset
Dim rsprint As New ADODB.Recordset


Private Sub cmbgrnofrom_Click()
With rstemp
If .State = 1 Then
.Close
End If
rstemp.Open "Select * from student_details where GRNO='" & cmbgrnofrom.Text & "'", cn, adOpenForwardOnly, adLockOptimistic
If .RecordCount <= 0 Then
txtnamefrom.Text = ""
Else
txtnamefrom.Text = UCase(.Fields("NAME"))
.Close
End If
End With
End Sub


Private Sub cmbgrnoto_Click()
With rstemp
If .State = 1 Then
.Close
End If
rstemp.Open "Select * from student_details where GRNO='" & cmbgrnoto.Text & "'", cn, adOpenForwardOnly, adLockOptimistic
If .RecordCount <= 0 Then
txtnameto.Text = ""
Else
txtnameto.Text = UCase(.Fields("NAME"))
.Close
End If
End With
End Sub


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
    .Open "SELECT * from student_details ", cn, adOpenForwardOnly, adLockReadOnly
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        Exit Sub
    End If
    Set rptstudent.DataSource = rsprint
    rptstudent.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptstudent.Sections("section4").Controls("lbladdress1").Caption = address1
    rptstudent.Sections("section4").Controls("lbladdress2").Caption = address2
    rptstudent.Sections("section4").Controls("lblreportname").Caption = "REPORT FOR ALL STUDENT"
    rptstudent.Sections("section4").Controls("lblusername").Caption = pubusername
    rptstudent.Sections("section4").Controls("lbldate").Caption = Now
    rptstudent.Show vbModal
End With
End If
If optselectgr.Value = True Then
If Trim(cmbgrnofrom.Text) = "" Then
MsgBox "Write GR No From", vbInformation, "Empty GR No"
cmbgrnofrom.SetFocus
Exit Sub
ElseIf Trim(cmbgrnoto.Text) = "" Then
MsgBox "Write GR No To", vbInformation, "Empty GR No"
cmbgrnoto.SetFocus
Exit Sub
End If
With rsprint
    If .State = adStateOpen Then .Close
  '  .Open "SELECT * from student_details where GRNO='" & cmbgrnofrom & "'", cn, adOpenForwardOnly, adLockReadOnly
 .Open "SELECT * from student_details where grno between """ & cmbgrnofrom & """ And """ & cmbgrnoto & """", cn, adOpenForwardOnly, adLockReadOnly
    
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        cmbgrnofrom.SetFocus
        Exit Sub
    End If
    Set rptstudent.DataSource = rsprint
    rptstudent.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptstudent.Sections("section4").Controls("lbladdress1").Caption = address1
    rptstudent.Sections("section4").Controls("lbladdress2").Caption = address2
    rptstudent.Sections("section4").Controls("lblreportname").Caption = "REPORT FOR SELECTED STUDENT GR NO"
    rptstudent.Sections("section4").Controls("lblusername").Caption = pubusername
    rptstudent.Sections("section4").Controls("lbldate").Caption = Now
    rptstudent.Show vbModal
End With
cmbgrnofrom.SetFocus
End If
If optselectname.Value = True Then
If Trim(txtselectname.Text) = "" Then
MsgBox "Write Student Name", vbInformation, "Empty Student Name"
txtselectname.SetFocus
Exit Sub
End If
With rsprint
    If .State = adStateOpen Then .Close
  '  .Open "SELECT * from student_details where GRNO='" & cmbgrnofrom & "'", cn, adOpenForwardOnly, adLockReadOnly
 .Open "SELECT * from student_details where name like '" & txtselectname & "%" & "'", cn, adOpenForwardOnly, adLockReadOnly
    
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        txtselectname.Text = ""
        txtselectname.SetFocus
        Exit Sub
    End If
    Set rptstudent.DataSource = rsprint
    rptstudent.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptstudent.Sections("section4").Controls("lbladdress1").Caption = address1
    rptstudent.Sections("section4").Controls("lbladdress2").Caption = address2
    rptstudent.Sections("section4").Controls("lblreportname").Caption = "REPORT FOR SELECTED STUDENT NAME"
    rptstudent.Sections("section4").Controls("lblusername").Caption = pubusername
    rptstudent.Sections("section4").Controls("lbldate").Caption = Now
    rptstudent.Show vbModal
End With
txtselectname.Text = ""
txtselectname.SetFocus
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
 .Open "SELECT * from student_details where levelcode between """ & cmblevelfrom & """ And """ & cmblevelto & """", cn, adOpenForwardOnly, adLockReadOnly
    
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        cmblevelfrom.SetFocus
        Exit Sub
    End If
    Set rptstudent.DataSource = rsprint
    rptstudent.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptstudent.Sections("section4").Controls("lbladdress1").Caption = address1
    rptstudent.Sections("section4").Controls("lbladdress2").Caption = address2
    rptstudent.Sections("section4").Controls("lblreportname").Caption = "REPORT FOR SELECTED LEVEL CODE"
    rptstudent.Sections("section4").Controls("lblusername").Caption = pubusername
    rptstudent.Sections("section4").Controls("lbldate").Caption = Now
    rptstudent.Show vbModal
End With
cmblevelfrom.SetFocus
End If
If optselectgendar.Value = True Then
If Trim(cmbgendar.Text) = "" Then
MsgBox "Write Student Gendar", vbInformation, "Empty Student Gendar"
cmbgendar.SetFocus
Exit Sub
End If
With rsprint
    If .State = adStateOpen Then .Close
    .Open "SELECT * from student_details where sex='" & cmbgendar & "'", cn, adOpenForwardOnly, adLockReadOnly
     
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        cmbgendar.SetFocus
        Exit Sub
    End If
    Set rptstudent.DataSource = rsprint
    rptstudent.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptstudent.Sections("section4").Controls("lbladdress1").Caption = address1
    rptstudent.Sections("section4").Controls("lbladdress2").Caption = address2
    rptstudent.Sections("section4").Controls("lblreportname").Caption = "REPORT FOR SELECTED STUDENT GENDAR"
    rptstudent.Sections("section4").Controls("lblusername").Caption = pubusername
    rptstudent.Sections("section4").Controls("lbldate").Caption = Now
    rptstudent.Show vbModal
End With
cmbgendar.SetFocus
End If
End Sub

Private Sub Form_Load()
Call conopen
If rsSTUDENT.State = 0 Then
rsSTUDENT.Open "select * from student order by grno", cn, adOpenDynamic, adLockOptimistic
End If

With rsSTUDENT
If .State = closed Then
.Open
End If
.AbsolutePosition = 1
While .EOF = False
cmbgrnofrom.AddItem .Fields("GRNO")
cmbgrnoto.AddItem .Fields("GRNO")
.MoveNext
Wend
End With

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

cmbgendar.AddItem "Male"
cmbgendar.AddItem "Female"
cmbgendar.Text = "Male"

frmselectgr.Visible = False
frmselectname.Visible = False
frmselectlevel.Visible = False
frmselectgendar.Visible = False

optall.Value = True


End Sub

Private Sub optall_Click()
If optall.Value = True Then
frmselectgr.Visible = False
frmselectname.Visible = False
frmselectlevel.Visible = False
frmselectgendar.Visible = False
cmdprint.SetFocus
End If
End Sub

Private Sub optselectgendar_Click()
If optselectgendar.Value = True Then
frmselectname.Visible = False
frmselectgr.Visible = False
frmselectlevel.Visible = False
frmselectgendar.Visible = True
cmbgendar.SetFocus
End If
End Sub

Private Sub optselectgr_Click()
If optselectgr.Value = True Then
frmselectgr.Visible = True
frmselectname.Visible = False
frmselectlevel.Visible = False
frmselectgendar.Visible = False
cmbgrnofrom.SetFocus
End If
End Sub

Private Sub optselectlevel_Click()
If optselectlevel.Value = True Then
frmselectname.Visible = False
frmselectgr.Visible = False
frmselectlevel.Visible = True
frmselectgendar.Visible = False
cmblevelfrom.SetFocus
End If
End Sub

Private Sub optselectname_Click()
If optselectname.Value = True Then
frmselectname.Visible = True
frmselectgr.Visible = False
frmselectlevel.Visible = False
frmselectgendar.Visible = False
txtselectname.SetFocus
End If
End Sub
