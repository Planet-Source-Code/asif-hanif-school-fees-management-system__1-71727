VERSION 5.00
Begin VB.Form frmallreports 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All Reports"
   ClientHeight    =   5760
   ClientLeft      =   3045
   ClientTop       =   2385
   ClientWidth     =   3600
   Icon            =   "frmallreports.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   3600
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3255
      Begin VB.CommandButton cmdreceipt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fee Receipt Details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   840
         Width           =   2775
      End
      Begin VB.CommandButton cmdsubjects 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Subjects Details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1440
         Width           =   2775
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2040
         Width           =   2775
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2640
         Width           =   2775
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3240
         Width           =   2775
      End
      Begin VB.CommandButton cmdclose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Close"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3840
         Width           =   2775
      End
      Begin VB.CommandButton cmdstudent 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Student Details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "View All Reports"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "frmallreports.frx":0F22
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15855
   End
End
Attribute VB_Name = "frmallreports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
frmmain.Show
Me.Hide
End Sub

Private Sub cmdreceipt_Click()
frmreceiptprint.Show
Me.Hide
End Sub

Private Sub cmdstudent_Click()
frmstudentprint.Show
Me.Hide
End Sub
