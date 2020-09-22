VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form frmreceiptprint 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receipt  Report"
   ClientHeight    =   6150
   ClientLeft      =   2880
   ClientTop       =   1740
   ClientWidth     =   7530
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmreceiptprint.frx":0000
   ScaleHeight     =   6150
   ScaleWidth      =   7530
   Begin VB.OptionButton optselectnotfee 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Not Paid Fee Student Details  ( Only Receipt  Details )"
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
      TabIndex        =   29
      Top             =   3120
      Width           =   5055
   End
   Begin VB.CommandButton cmddetailprint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Print Receipt Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4440
      Picture         =   "frmreceiptprint.frx":0467
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Frame frmselectno 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   480
      TabIndex        =   19
      Top             =   3600
      Width           =   6135
      Begin VB.ComboBox cmbnofrom 
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
         Top             =   247
         Width           =   1215
      End
      Begin VB.ComboBox cmbnoto 
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
         Left            =   4800
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   247
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt  No From :"
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
         Index           =   6
         Left            =   120
         TabIndex        =   23
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt  No To :"
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
         Index           =   1
         Left            =   3240
         TabIndex        =   22
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.Frame frmselectmonth 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   480
      TabIndex        =   16
      Top             =   3600
      Width           =   5295
      Begin VB.ComboBox cmbyear 
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
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmbmonth 
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
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Year :"
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
         Left            =   2760
         TabIndex        =   27
         Top             =   300
         Width           =   615
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Month :"
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
         TabIndex        =   18
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.Frame frmselectdate 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   480
      TabIndex        =   15
      Top             =   3480
      Width           =   3855
      Begin TDBDate6Ctl.TDBDate txtdatefrom 
         Bindings        =   "frmreceiptprint.frx":0D84
         Height          =   360
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   3570
         _Version        =   65536
         _ExtentX        =   6297
         _ExtentY        =   635
         Calendar        =   "frmreceiptprint.frx":0D8F
         Caption         =   "frmreceiptprint.frx":0EA7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmreceiptprint.frx":0F2A
         Keys            =   "frmreceiptprint.frx":0F48
         Spin            =   "frmreceiptprint.frx":0FA6
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "mmm dd, yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   1
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   " "
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "01/08/2008"
         ValidateMode    =   0
         ValueVT         =   2118189063
         Value           =   39661
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate txtdateto 
         Bindings        =   "frmreceiptprint.frx":0FCE
         Height          =   360
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   3570
         _Version        =   65536
         _ExtentX        =   6297
         _ExtentY        =   635
         Calendar        =   "frmreceiptprint.frx":0FD9
         Caption         =   "frmreceiptprint.frx":10F1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmreceiptprint.frx":1170
         Keys            =   "frmreceiptprint.frx":118E
         Spin            =   "frmreceiptprint.frx":11EC
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "mmm dd, yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   1
         ForeColor       =   -2147483640
         Format          =   "dd/mm/yyyy"
         HighlightText   =   0
         IMEMode         =   3
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxDate         =   2958465
         MinDate         =   -657434
         MousePointer    =   0
         MoveOnLRKey     =   0
         OLEDragMode     =   0
         OLEDropMode     =   0
         PromptChar      =   " "
         ReadOnly        =   0
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "01/08/2008"
         ValidateMode    =   0
         ValueVT         =   2118189063
         Value           =   39661
         CenturyMode     =   0
      End
   End
   Begin VB.OptionButton optselectmonth 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Selection By Receipt  Month && Year Wise"
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
      TabIndex        =   14
      Top             =   2760
      Width           =   3975
   End
   Begin VB.OptionButton optselectdate 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Selection By Receipt Date Wise"
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
      TabIndex        =   13
      Top             =   2400
      Width           =   3135
   End
   Begin VB.OptionButton optselectgr 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Selection By Receipt Student GR No Wise"
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
      Width           =   3975
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Print Receipt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1440
      Picture         =   "frmreceiptprint.frx":1214
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2880
      Picture         =   "frmreceiptprint.frx":1B31
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
      Top             =   3480
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
   Begin VB.OptionButton optselectno 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Selection By Receipt No Wise"
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
      Width           =   3015
   End
   Begin VB.OptionButton optall 
      BackColor       =   &H00C0C0C0&
      Caption         =   "All Receipt "
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
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PRINT RECEIPT REPORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   -840
      Picture         =   "frmreceiptprint.frx":25DF
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
Attribute VB_Name = "frmreceiptprint"
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


Private Sub cmdback_Click()
Unload Me
End Sub

Private Sub cmddetailprint_Click()
If optall.Value = True Then
With rsprint
    If .State = adStateOpen Then .Close
    .Open "SELECT * from feespayment", cn, adOpenForwardOnly, adLockReadOnly
    
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        Exit Sub
    End If
    Set rptfeesreceiptdetail.DataSource = rsprint
    rptfeesreceiptdetail.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptfeesreceiptdetail.Sections("section4").Controls("lbladdress1").Caption = address1
    rptfeesreceiptdetail.Sections("section4").Controls("lbladdress2").Caption = address2
    rptfeesreceiptdetail.Sections("section4").Controls("lblusername").Caption = pubusername
    rptfeesreceiptdetail.Sections("section4").Controls("lblreportname").Caption = "REPORT ALL RECEIPT"
    rptfeesreceiptdetail.Sections("section4").Controls("lbldate").Caption = Now
    rptfeesreceiptdetail.Show vbModal
End With
End If
If optselectno.Value = True Then
    If Trim(cmbnofrom.Text) = "" Then
        MsgBox "Write Receipt No From", vbInformation, "Empty Receipt No"
        cmbnofrom.SetFocus
        Exit Sub
    ElseIf Trim(cmbnoto.Text) = "" Then
        MsgBox "Write Receipt No To", vbInformation, "Empty Receipt No"
        cmbnoto.SetFocus
        Exit Sub
    End If
    
    With rsprint
    
    If .State = adStateOpen Then .Close
    .Open "SELECT * from feespayment where receiptno between   """ & cmbnofrom & """  And """ & cmbnoto & """", cn, adOpenForwardOnly, adLockReadOnly
 
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        Exit Sub
    End If
    Set rptfeesreceiptdetail.DataSource = rsprint
    rptfeesreceiptdetail.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptfeesreceiptdetail.Sections("section4").Controls("lbladdress1").Caption = address1
    rptfeesreceiptdetail.Sections("section4").Controls("lbladdress2").Caption = address2
    rptfeesreceiptdetail.Sections("section4").Controls("lblusername").Caption = pubusername
    rptfeesreceiptdetail.Sections("section4").Controls("lblreportname").Caption = "RECEIPT NO WISE REPORT"
    rptfeesreceiptdetail.Sections("section4").Controls("lbldate").Caption = Now
    rptfeesreceiptdetail.Show vbModal
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
    .Open "SELECT * from feespayment where grno between """ & cmbgrnofrom & """ And """ & cmbgrnoto & """", cn, adOpenForwardOnly, adLockReadOnly
    
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        Exit Sub
    End If
    Set rptfeesreceiptdetail.DataSource = rsprint
    rptfeesreceiptdetail.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptfeesreceiptdetail.Sections("section4").Controls("lbladdress1").Caption = address1
    rptfeesreceiptdetail.Sections("section4").Controls("lbladdress2").Caption = address2
    rptfeesreceiptdetail.Sections("section4").Controls("lblusername").Caption = pubusername
    rptfeesreceiptdetail.Sections("section4").Controls("lblreportname").Caption = "GR NO WISE RECEIPT REPORT"
    rptfeesreceiptdetail.Sections("section4").Controls("lbldate").Caption = Now
    rptfeesreceiptdetail.Show vbModal
End With
End If

If optselectdate.Value = True Then
    If Trim(txtdatefrom.Value) = "" Then
        MsgBox "Write Receipt Date From", vbInformation, "Empty Receipt Date"
        txtdatefrom.SetFocus
        Exit Sub
    ElseIf Trim(txtdateto.Value) = "" Then
        MsgBox "Write Receipt Date To", vbInformation, "Empty Receipt Date"
        txtdateto.SetFocus
        Exit Sub
    End If
    
    With rsprint
    
    If .State = adStateOpen Then .Close
    .Open "SELECT * from feespayment where date between   """ & txtdatefrom & """  And """ & txtdateto & """", cn, adOpenForwardOnly, adLockReadOnly
 
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        Exit Sub
    End If
    Set rptfeesreceiptdetail.DataSource = rsprint
    rptfeesreceiptdetail.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptfeesreceiptdetail.Sections("section4").Controls("lbladdress1").Caption = address1
    rptfeesreceiptdetail.Sections("section4").Controls("lbladdress2").Caption = address2
    rptfeesreceiptdetail.Sections("section4").Controls("lblusername").Caption = pubusername
    rptfeesreceiptdetail.Sections("section4").Controls("lblreportname").Caption = "DATE WISE RECEIPT REPORT"
    rptfeesreceiptdetail.Sections("section4").Controls("lbldate").Caption = Now
    rptfeesreceiptdetail.Show vbModal
End With
End If

If optselectmonth.Value = True Then
    If Trim(cmbmonth.Text) = "" Then
        MsgBox "Write Month", vbInformation, "Empty Month"
        cmbmonth.SetFocus
        Exit Sub
    ElseIf Trim(cmbyear.Text) = "" Then
        MsgBox "Write Year", vbInformation, "Empty Year"
        cmbyear.SetFocus
        Exit Sub
    End If
    
    With rsprint
    
    If .State = adStateOpen Then .Close
    .Open "SELECT * from feespayment where feemonth in ('" & cmbmonth & "')  And year = '" & cmbyear & "'", cn, adOpenForwardOnly, adLockReadOnly
    
'    .Open "SELECT * from feespayment where feemonth like '%" & cmbmonth & "%'  And year = '" & cmbyear & "'", cn, adOpenForwardOnly, adLockReadOnly
'rs_search.Open "Select * from qry_StudentDetails_GS where studid like '%" & Text1.Text & "%' and sy = '" & "2007-2008" & "'", cn, adOpenStatic, adLockOptimistic
 
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        Exit Sub
    End If
    Set rptfeesreceiptdetail.DataSource = rsprint
    rptfeesreceiptdetail.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptfeesreceiptdetail.Sections("section4").Controls("lbladdress1").Caption = address1
    rptfeesreceiptdetail.Sections("section4").Controls("lbladdress2").Caption = address2
    rptfeesreceiptdetail.Sections("section4").Controls("lblusername").Caption = pubusername
    rptfeesreceiptdetail.Sections("section4").Controls("lblreportname").Caption = "MONTH WISE RECEIPT REPORT"
    rptfeesreceiptdetail.Sections("section4").Controls("lbldate").Caption = Now
    rptfeesreceiptdetail.Show vbModal
End With
End If
If optselectnotfee.Value = True Then
    If Trim(cmbmonth.Text) = "" Then
        MsgBox "Write Month", vbInformation, "Empty Month"
        cmbmonth.SetFocus
        Exit Sub
    ElseIf Trim(cmbyear.Text) = "" Then
        MsgBox "Write Year", vbInformation, "Empty Year"
        cmbyear.SetFocus
        Exit Sub
    End If
    
    With rsprint
    
    If .State = adStateOpen Then .Close
    .Open "SELECT * from feespayment where feemonth not in ('" & cmbmonth & "')  And year = '" & cmbyear & "'", cn, adOpenForwardOnly, adLockReadOnly

'rs.Open "select * from tbl_subjectonly WHERE studid= '" & fMain.curr_student & "' And subject = '" & Label5.Caption & "' and sy = '" & fMain.curr_sy & "'", cn, adOpenStatic, adLockOptimistic

'rs_search.Open "Select * from qry_StudentDetails_GS where studid like '%" & Text1.Text & "%' and sy = '" & "2007-2008" & "'", cn, adOpenStatic, adLockOptimistic
 
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        Exit Sub
    End If
    Set rptfeesdefalter.DataSource = rsprint
    rptfeesdefalter.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptfeesdefalter.Sections("section4").Controls("lbladdress1").Caption = address1
    rptfeesdefalter.Sections("section4").Controls("lbladdress2").Caption = address2
    rptfeesdefalter.Sections("section4").Controls("lblusername").Caption = pubusername
    rptfeesdefalter.Sections("section4").Controls("lblreportname").Caption = "NOT PAID FEE MONTH WISE REPORT"
    rptfeesdefalter.Sections("section4").Controls("lbldate").Caption = Now
    rptfeesdefalter.Show vbModal
End With
End If

End Sub

Private Sub cmdprint_Click()
If optall.Value = True Then
With rsprint
    If .State = adStateOpen Then .Close
    .Open "SELECT * from feespayment", cn, adOpenForwardOnly, adLockReadOnly
    
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        Exit Sub
    End If
    Set rptfeesreceipt.DataSource = rsprint
    rptfeesreceipt.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptfeesreceipt.Sections("section4").Controls("lbladdress1").Caption = address1
    rptfeesreceipt.Sections("section4").Controls("lbladdress2").Caption = address2
    rptfeesreceipt.Sections("section4").Controls("lblusername").Caption = pubusername
    rptfeesreceipt.Sections("section4").Controls("lbldate").Caption = Now
    rptfeesreceipt.Sections("section1").Controls("lblschoolname1").Caption = schoolname
    rptfeesreceipt.Show vbModal
End With
End If
If optselectno.Value = True Then
    If Trim(cmbnofrom.Text) = "" Then
        MsgBox "Write Receipt No From", vbInformation, "Empty Receipt No"
        cmbnofrom.SetFocus
        Exit Sub
    ElseIf Trim(cmbnoto.Text) = "" Then
        MsgBox "Write Receipt No To", vbInformation, "Empty Receipt No"
        cmbnoto.SetFocus
        Exit Sub
    End If
    
    With rsprint
    
    If .State = adStateOpen Then .Close
    .Open "SELECT * from feespayment where receiptno between   """ & cmbnofrom & """  And """ & cmbnoto & """", cn, adOpenForwardOnly, adLockReadOnly
 
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        Exit Sub
    End If
    Set rptfeesreceipt.DataSource = rsprint
    rptfeesreceipt.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptfeesreceipt.Sections("section4").Controls("lbladdress1").Caption = address1
    rptfeesreceipt.Sections("section4").Controls("lbladdress2").Caption = address2
    rptfeesreceipt.Sections("section4").Controls("lblusername").Caption = pubusername
    rptfeesreceipt.Sections("section4").Controls("lbldate").Caption = Now
    rptfeesreceipt.Sections("section1").Controls("lblschoolname1").Caption = schoolname
    rptfeesreceipt.Show vbModal
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
    .Open "SELECT * from feespayment where grno between """ & cmbgrnofrom & """ And """ & cmbgrnoto & """", cn, adOpenForwardOnly, adLockReadOnly
    
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        Exit Sub
    End If
    Set rptfeesreceipt.DataSource = rsprint
    rptfeesreceipt.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptfeesreceipt.Sections("section4").Controls("lbladdress1").Caption = address1
    rptfeesreceipt.Sections("section4").Controls("lbladdress2").Caption = address2
    rptfeesreceipt.Sections("section4").Controls("lblusername").Caption = pubusername
    rptfeesreceipt.Sections("section4").Controls("lbldate").Caption = Now
    rptfeesreceipt.Sections("section1").Controls("lblschoolname1").Caption = schoolname
    rptfeesreceipt.Show vbModal
End With
End If

If optselectdate.Value = True Then
    If Trim(txtdatefrom.Value) = "" Then
        MsgBox "Write Receipt Date From", vbInformation, "Empty Receipt Date"
        txtdatefrom.SetFocus
        Exit Sub
    ElseIf Trim(txtdateto.Value) = "" Then
        MsgBox "Write Receipt Date To", vbInformation, "Empty Receipt Date"
        txtdateto.SetFocus
        Exit Sub
    End If
    
    With rsprint
    
    If .State = adStateOpen Then .Close
    .Open "SELECT * from feespayment where date between   """ & txtdatefrom & """  And """ & txtdateto & """", cn, adOpenForwardOnly, adLockReadOnly
 
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        Exit Sub
    End If
    Set rptfeesreceipt.DataSource = rsprint
    rptfeesreceipt.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptfeesreceipt.Sections("section4").Controls("lbladdress1").Caption = address1
    rptfeesreceipt.Sections("section4").Controls("lbladdress2").Caption = address2
    rptfeesreceipt.Sections("section4").Controls("lblusername").Caption = pubusername
    rptfeesreceipt.Sections("section4").Controls("lbldate").Caption = Now
    rptfeesreceipt.Sections("section1").Controls("lblschoolname1").Caption = schoolname
    rptfeesreceipt.Show vbModal
End With
End If

If optselectmonth.Value = True Then
    If Trim(cmbmonth.Text) = "" Then
        MsgBox "Write Month", vbInformation, "Empty Month"
        cmbmonth.SetFocus
        Exit Sub
    ElseIf Trim(cmbyear.Text) = "" Then
        MsgBox "Write Year", vbInformation, "Empty Year"
        cmbyear.SetFocus
        Exit Sub
    End If
    
    With rsprint
    
    If .State = adStateOpen Then .Close
    
    .Open "SELECT * from feespayment where feemonth in ('" & cmbmonth & "')  And year = '" & cmbyear & "'", cn, adOpenForwardOnly, adLockReadOnly
    
'    .Open "SELECT * from feespayment where feemonth like '%" & cmbmonth & "%'  And year = '" & cmbyear & "'", cn, adOpenForwardOnly, adLockReadOnly
'rs_search.Open "Select * from qry_StudentDetails_GS where studid like '%" & Text1.Text & "%' and sy = '" & "2007-2008" & "'", cn, adOpenStatic, adLockOptimistic
 
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        Exit Sub
    End If
    Set rptfeesreceipt.DataSource = rsprint
    rptfeesreceipt.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptfeesreceipt.Sections("section4").Controls("lbladdress1").Caption = address1
    rptfeesreceipt.Sections("section4").Controls("lbladdress2").Caption = address2
    rptfeesreceipt.Sections("section4").Controls("lblusername").Caption = pubusername
    rptfeesreceipt.Sections("section4").Controls("lbldate").Caption = Now
    rptfeesreceipt.Sections("section1").Controls("lblschoolname1").Caption = schoolname
    rptfeesreceipt.Show vbModal
End With
End If

End Sub

Private Sub Form_Load()
Call conopen
With rsfeepayment
If .State = adStateOpen Then .Close
.Open "SELECT * from feespayment ", cn, adOpenForwardOnly, adLockReadOnly

.AbsolutePosition = 1
While .EOF = False
cmbnofrom.AddItem .Fields("RECEIPTNO")
cmbnoto.AddItem .Fields("RECEIPTNO")
.MoveNext
Wend
End With

With rsSTUDENT
If .State = adStateOpen Then .Close
rsSTUDENT.Open "select * from student_details", cn, adOpenForwardOnly, adLockOptimistic
.AbsolutePosition = 1
While .EOF = False
cmbgrnofrom.AddItem .Fields("GRNO")
cmbgrnoto.AddItem .Fields("GRNO")
.MoveNext
Wend
End With

cmbmonth.AddItem "January"
cmbmonth.AddItem "February"
cmbmonth.AddItem "March"
cmbmonth.AddItem "April"
cmbmonth.AddItem "May"
cmbmonth.AddItem "June"
cmbmonth.AddItem "July"
cmbmonth.AddItem "August"
cmbmonth.AddItem "September"
cmbmonth.AddItem "October"
cmbmonth.AddItem "November"
cmbmonth.AddItem "December"
cmbmonth.Text = "January"

cmbyear.AddItem "2005"
cmbyear.AddItem "2006"
cmbyear.AddItem "2007"
cmbyear.AddItem "2008"
cmbyear.AddItem "2009"
cmbyear.AddItem "2010"
cmbyear.AddItem "2011"
cmbyear.AddItem "2012"
cmbyear.AddItem "2013"
cmbyear.AddItem "2014"
cmbyear.AddItem "2015"
cmbyear.AddItem "2016"
cmbyear.AddItem "2017"
cmbyear.AddItem "2018"
cmbyear.AddItem "2019"
cmbyear.AddItem "2020"
cmbyear.Text = "2008"




txtdatefrom.Value = Date
txtdateto.Value = Date

frmselectno.Visible = False
frmselectgr.Visible = False
frmselectdate.Visible = False
frmselectmonth.Visible = False

optall.Value = True


End Sub


Private Sub optall_Click()
If optall.Value = True Then
frmselectno.Visible = False
frmselectgr.Visible = False
frmselectdate.Visible = False
frmselectmonth.Visible = False
cmdprint.SetFocus
End If

End Sub

Private Sub optselectdate_Click()
If optselectdate.Value = True Then
frmselectno.Visible = False
frmselectgr.Visible = False
frmselectdate.Visible = True
frmselectmonth.Visible = False
txtdatefrom.SetFocus
End If

End Sub

Private Sub optselectgr_Click()
If optselectgr.Value = True Then
frmselectno.Visible = False
frmselectgr.Visible = True
frmselectdate.Visible = False
frmselectmonth.Visible = False
cmbgrnofrom.SetFocus
End If

End Sub

Private Sub optselectmonth_Click()
If optselectmonth.Value = True Then
frmselectno.Visible = False
frmselectgr.Visible = False
frmselectdate.Visible = False
frmselectmonth.Visible = True
cmbmonth.SetFocus
End If

End Sub

Private Sub optselectno_Click()
If optselectno.Value = True Then
frmselectno.Visible = True
frmselectgr.Visible = False
frmselectdate.Visible = False
frmselectmonth.Visible = False
cmbnofrom.SetFocus
End If

End Sub

Private Sub optselectnotfee_Click()
If optselectnotfee.Value = True Then
frmselectno.Visible = False
frmselectgr.Visible = False
frmselectdate.Visible = False
frmselectmonth.Visible = True
cmbmonth.SetFocus
End If
End Sub
