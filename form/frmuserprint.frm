VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form frmuserprint 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User Login Report"
   ClientHeight    =   5565
   ClientLeft      =   5580
   ClientTop       =   3045
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   5535
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4335
      Begin VB.Frame frmselectuser 
         BackColor       =   &H00E0E0E0&
         Height          =   735
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   3855
         Begin VB.ComboBox cmbuser 
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
            Left            =   1320
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "User Name :"
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
            TabIndex        =   14
            Top             =   300
            Width           =   1095
         End
      End
      Begin VB.Frame frmselectdate 
         BackColor       =   &H00E0E0E0&
         Height          =   1215
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   3855
         Begin TDBDate6Ctl.TDBDate txtdatefrom 
            Bindings        =   "frmuserprint.frx":0000
            Height          =   360
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   3570
            _Version        =   65536
            _ExtentX        =   6297
            _ExtentY        =   635
            Calendar        =   "frmuserprint.frx":000B
            Caption         =   "frmuserprint.frx":0123
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmuserprint.frx":01A4
            Keys            =   "frmuserprint.frx":01C2
            Spin            =   "frmuserprint.frx":0220
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
            Bindings        =   "frmuserprint.frx":0248
            Height          =   360
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   3570
            _Version        =   65536
            _ExtentX        =   6297
            _ExtentY        =   635
            Calendar        =   "frmuserprint.frx":0253
            Caption         =   "frmuserprint.frx":036B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmuserprint.frx":03E8
            Keys            =   "frmuserprint.frx":0406
            Spin            =   "frmuserprint.frx":0464
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
      Begin VB.CommandButton cmdback 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Back"
         Height          =   855
         Left            =   2280
         Picture         =   "frmuserprint.frx":048C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox txtusername 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         TabIndex        =   0
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton cmdprint 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Print"
         Height          =   855
         Left            =   960
         Picture         =   "frmuserprint.frx":0C67
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox txtpassword 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "User Login Report"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   17
         Top             =   120
         Width           =   4335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   2760
         TabIndex        =   16
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Label10"
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   855
         Width           =   1215
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1455
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmuserprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim txtuserid As String
Dim rsprint As New ADODB.Recordset
Dim reccount As Integer

Private Sub cmdback_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
If frmselectuser.Visible = True Then
 If Trim(txtusername.Text) = "" Then
 MsgBox "Write User Name", vbInformation, "Empty User"
 txtusername.SetFocus
 Exit Sub
 End If
 If Trim(txtpassword.Text) = "" Then
 MsgBox "Write Password", vbInformation, "Empty Password"
 txtpassword.SetFocus
 Exit Sub
 End If
 If Trim(cmbuser.Text) = "" Then
 MsgBox "Write User Name", vbInformation, "Empty User"
 cmbuser.SetFocus
 Exit Sub
 End If
With rsprint
    If .State = adStateOpen Then .Close
 
    .Open "SELECT * from userlog_details where username= """ & cmbuser & """ and logtime between """ & txtdatefrom & """ and """ & txtdateto & """", cn, adOpenForwardOnly, adLockReadOnly
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        Exit Sub
    End If
    reccount = .RecordCount
    Set rptuserlogin.DataSource = rsprint
    rptuserlogin.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptuserlogin.Sections("section4").Controls("lbladdress1").Caption = address1
    rptuserlogin.Sections("section4").Controls("lbladdress2").Caption = address2
    rptuserlogin.Sections("section4").Controls("lblreportname").Caption = "USER LOG REPORT"
    rptuserlogin.Sections("section4").Controls("lblusername").Caption = pubusername
    rptuserlogin.Sections("section4").Controls("lbldate").Caption = Now
    rptuserlogin.Sections("section5").Controls("lblnoofrecords").Caption = reccount
    rptuserlogin.Show vbModal
End With
txtusername.Text = ""
txtpassword.Text = ""
frmselectuser.Visible = False
End If
If frmselectuser.Visible = False Then
 If Trim(txtusername.Text) = "" Then
 MsgBox "Write User Name", vbInformation, "Empty User"
 txtusername.SetFocus
 Exit Sub
 End If
 If Trim(txtpassword.Text) = "" Then
 MsgBox "Write Password", vbInformation, "Empty Password"
 txtpassword.SetFocus
 Exit Sub
 End If
With rsprint
    If .State = adStateOpen Then .Close
 
    .Open "SELECT * from userlog_details where username= """ & txtusername & """ and logtime between """ & txtdatefrom & """ and """ & txtdateto & """", cn, adOpenForwardOnly, adLockReadOnly
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        Exit Sub
    End If
    reccount = .RecordCount
    Set rptuserlogin.DataSource = rsprint
    rptuserlogin.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptuserlogin.Sections("section4").Controls("lbladdress1").Caption = address1
    rptuserlogin.Sections("section4").Controls("lbladdress2").Caption = address2
    rptuserlogin.Sections("section4").Controls("lblreportname").Caption = "USER LOG REPORT"
    rptuserlogin.Sections("section4").Controls("lblusername").Caption = pubusername
    rptuserlogin.Sections("section4").Controls("lbldate").Caption = Now
    rptuserlogin.Sections("section5").Controls("lblnoofrecords").Caption = reccount
    rptuserlogin.Show vbModal
End With
txtusername.Text = ""
txtpassword.Text = ""
frmselectuser.Visible = False
End If
End Sub

Private Sub Form_Activate()
txtusername.SetFocus
End Sub

Private Sub Form_Load()
Call conopen
If rsuser.State = 0 Then
rsuser.Open "select * from user_details", cn, adOpenDynamic, adLockOptimistic
End If
With rsuser
If .State = closed Then
.Open
End If
.AbsolutePosition = 1
While .EOF = False
cmbuser.AddItem UCase(.Fields("USERNAME"))
.MoveNext
Wend
End With

txtusername.Text = ""
txtpassword.Text = ""
frmselectuser.Visible = False


End Sub


Private Sub txtpassword_GotFocus()
 With rsuser
 If .State = closed Then
 .Open
 End If
 .MoveFirst
 .Find "USERNAME ='" & txtusername & "'"
 If .EOF Then
      MsgBox "User Not Found", vbInformation, "Invelid User Name"
      txtusername.SetFocus
      Exit Sub
  Else
    txtpassword.SetFocus
    Exit Sub
  End If
  End With

End Sub

Private Sub txtpassword_LostFocus()
 With rsuser
 If .State = closed Then
 .Open
 End If
 If Trim(txtusername.Text) = "" Then
 MsgBox "Write User Name", vbInformation, "Empty User"
 txtusername.SetFocus
 Exit Sub
 End If
 If Trim(txtpassword.Text) = "" Then
 MsgBox "Write Password", vbInformation, "Empty Password"
 txtpassword.SetFocus
 Exit Sub
 End If
 .MoveFirst
 .Find "USERNAME='" & txtusername & "'"
If txtpassword.Text <> .Fields("password") Then
    MsgBox "Wrong Password", vbCritical, "Wrong Password"
    txtpassword.Text = ""
    txtpassword.SetFocus
    Exit Sub
End If
txtuserid = .Fields("USERID")
End With
If txtuserid <> "Admin" Then
frmselectuser.Visible = False
Else
frmselectuser.Visible = True
cmbuser.SetFocus
End If
End Sub
