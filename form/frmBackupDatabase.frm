VERSION 5.00
Object = "{CED4D45D-8DB4-42E5-A708-BCA2899963C8}#2.0#0"; "ProgressBar.ocx"
Begin VB.Form frmBackupDatabase 
   Caption         =   "Backup Database"
   ClientHeight    =   5130
   ClientLeft      =   135
   ClientTop       =   2145
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   11295
   Begin progressbarocx.ProgressBar prgbar 
      Height          =   450
      Left            =   3240
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   794
      Size            =   5
   End
   Begin VB.DirListBox Dir 
      Appearance      =   0  'Flat
      Height          =   1665
      Left            =   3240
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
   End
   Begin VB.DriveListBox Drive 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3240
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5400
      TabIndex        =   2
      Top             =   3240
      Width           =   3975
   End
   Begin VB.CommandButton cmdClose 
      Height          =   855
      Left            =   9960
      Picture         =   "frmBackupDatabase.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdBackup 
      Height          =   855
      Left            =   9960
      Picture         =   "frmBackupDatabase.frx":2D44
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   975
   End
   Begin VB.Timer timPrgBar 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10680
      Top             =   4200
   End
   Begin VB.Image Image2 
      Height          =   4830
      Left            =   120
      Picture         =   "frmBackupDatabase.frx":5A88
      Top             =   240
      Width           =   2490
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000001&
      X1              =   9480
      X2              =   5400
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lbl_fra_BackUp 
      BackStyle       =   0  'Transparent
      Caption         =   "Back Up Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   840
      Width           =   2295
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000001&
      X1              =   2880
      X2              =   3120
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000001&
      X1              =   9480
      X2              =   9480
      Y1              =   3960
      Y2              =   960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   2880
      X2              =   9480
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   2880
      X2              =   2880
      Y1              =   960
      Y2              =   3960
   End
   Begin VB.Label lblBackUpPath 
      BackStyle       =   0  'Transparent
      Caption         =   "Back Up Path :"
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
      Left            =   5400
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblFileName 
      BackStyle       =   0  'Transparent
      Caption         =   "Back Up Filename :"
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
      Left            =   5400
      TabIndex        =   6
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000001&
      X1              =   9720
      X2              =   9720
      Y1              =   960
      Y2              =   3960
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000001&
      X1              =   11160
      X2              =   11160
      Y1              =   960
      Y2              =   3960
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000001&
      X1              =   9720
      X2              =   11160
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000001&
      X1              =   9720
      X2              =   11160
      Y1              =   3960
      Y2              =   3960
   End
End
Attribute VB_Name = "frmBackupDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'--------------------------------------------------------------------------------
'Hospital Management System - Extended Edition
'Form Name: Backup Database Interface
'Programmer: Imran Sheriff
'Quality Assurance Engineer (Testing):  Isham Sally
'Start Date: 23/04/08
'Date Of Last Modification: 23/04/08
'The Name Of The Database Being Accessed: sdp
'The Name/s Of The Database Table/s Being Accessed:
'--------------------------------------------------------------------------------

Option Explicit
Dim FileSystemObject As Object
Dim strfilename As String


Private Sub cmdBackup_Click()
On Error GoTo e
    'set the copying functionality
    strfilename = "" + txtPath.Text + "\" + txtFile.Text + ".mdb"
    'Set the object contractions
    Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
    'Copy the file according the path settings
    FileSystemObject.CopyFile App.Path & "\database\school.mdb", strfilename
    prgbar.Visible = True
    timPrgBar.Enabled = True
Exit Sub
e:
MsgBox "Invalid Path Setting, Please Try Again", vbCritical, "Invalid Path Setting!"
End Sub


Private Sub Dir_Click()
    txtPath.Text = "" & Dir.Path
End Sub

Private Sub Drive_Change()
    Dim d, fs As Object
    
    'Set the constrctions to created objectes
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(Drive.Drive))
    
    'Set the contents of the selected drives
    If d.IsReady Then
        Dir.Path = Drive.Drive
        Dir.SetFocus
    Else
        MsgBox "The Drive Is Not Ready!", vbExclamation, "Drive Not Ready!"
    End If
End Sub

Private Sub Form_Load()
    'Display Today 's date
    txtFile.Text = FormatDateTime(Now, vbLongDate)
End Sub


Private Sub timPrgBar_Timer()
    Static iCnt As Integer
    'Run the timer and check the condition
    If iCnt <= 100 Then
        'prgbar.Value = iCnt
        prgbar.Percent = iCnt
        iCnt = iCnt + 1
    Else
        MsgBox "The Backup Procedure Has Been Successfully Completed!", vbInformation, "Successful Backup Procedure!"
        Drive.SetFocus
        prgbar.Visible = False
        timPrgBar.Enabled = False
    End If
End Sub



Private Sub cmdClose_Click()    'On click of the Close Button
    
    'Obtaining confirmation from the user
'    If MsgBox(" Are You Sure You Wish To Close This Interface?", vbYesNo + vbQuestion, "Close Interface?") = vbYes Then
        Unload Me
 '   End If
    
End Sub
