VERSION 5.00
Object = "{CED4D45D-8DB4-42E5-A708-BCA2899963C8}#2.0#0"; "ProgressBar.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRestore 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Restore Facility"
   ClientHeight    =   4275
   ClientLeft      =   1950
   ClientTop       =   3660
   ClientWidth     =   10200
   Icon            =   "frmRestore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Option's"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   2535
      Left            =   8400
      TabIndex        =   11
      Top             =   1080
      Width           =   1695
      Begin VB.CommandButton cmdRestore 
         Caption         =   "Restore &Now"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "Cl&ear All"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Back to Main"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8400
      Top             =   3720
   End
   Begin VB.Frame Frame2 
      Caption         =   "Location where to Save and New Filename:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   1920
      TabIndex        =   5
      Top             =   2400
      Width           =   6375
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   4335
      End
      Begin VB.CommandButton cmdBrowse2 
         Caption         =   "B&rowse"
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
         Left            =   2640
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdClear2 
         Caption         =   "C&lear"
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
         Left            =   4080
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Database Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   400
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select the Path &&  Database to Restore:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1335
      Left            =   1920
      TabIndex        =   0
      Top             =   960
      Width           =   6375
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   9
         Top             =   360
         Width           =   4335
      End
      Begin VB.CommandButton cmdBrowse1 
         Caption         =   "&Browse"
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
         Left            =   2640
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdClear1 
         Caption         =   "&Clear"
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
         Left            =   4080
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Database Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   400
         Width           =   1815
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   8880
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin progressbarocx.ProgressBar pb 
      Height          =   450
      Left            =   2040
      TabIndex        =   10
      Top             =   3720
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   794
      Size            =   4
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   1920
      Picture         =   "frmRestore.frx":030A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   8175
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   3465
      Left            =   120
      Picture         =   "frmRestore.frx":3C33
      Top             =   120
      Width           =   1755
   End
End
Attribute VB_Name = "frmRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    ans = MsgBox("Exit Restore Facility?", vbCritical + vbYesNo, "Confirm")
    If ans = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdRestore_Click()
    If Text1 = "" Or Text2 = "" Then
        MsgBox "Filenames and Locations are required!"
        Text1.SetFocus
    Else
        Dim backupfile As New FileSystemObject
        backupfile.GetFileName (Text2)
        Kill Text2
        FileCopy Text1, Text2
        pb.Visible = True
        Timer1.Enabled = True
    End If
End Sub

Private Sub cmdBrowse1_Click()
    cd.ShowOpen
    Text1 = cd.FileName
End Sub

Private Sub cmdBrowse2_Click()
    cd.ShowSave
    Text2 = cd.FileName
End Sub

Private Sub cmdClear1_Click()
    Text1 = ""
End Sub

Private Sub cmdClear2_Click()
    Text2 = ""
End Sub

Private Sub cmdClearAll_Click()
    cmdClear1_Click
    cmdClear2_Click
End Sub

Private Sub Timer1_Timer()
    pb.Percent = pb.Percent + 1
    If pb.Percent = 100 Then
        MsgBox "Restore Database Completed!"
        Timer1.Enabled = False
        pb.Visible = False
    End If
End Sub
