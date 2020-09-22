VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmuser 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Account"
   ClientHeight    =   6240
   ClientLeft      =   3765
   ClientTop       =   3060
   ClientWidth     =   6960
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   6960
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   240
      TabIndex        =   21
      Top             =   1200
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "User List"
      TabPicture(0)   =   "frmuser.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label21"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label20"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmddelete"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdadd"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdchange"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdback"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "List1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "List2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "New Record"
      TabPicture(1)   =   "frmuser.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Change Password"
      TabPicture(2)   =   "frmuser.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Delete Record"
      TabPicture(3)   =   "frmuser.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2(2)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.ListBox List2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3060
         ItemData        =   "frmuser.frx":0070
         Left            =   2280
         List            =   "frmuser.frx":0072
         TabIndex        =   49
         Top             =   960
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   3855
         Index           =   2
         Left            =   -74640
         TabIndex        =   44
         Top             =   600
         Width           =   5895
         Begin VB.TextBox txtpassword3 
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
            TabIndex        =   18
            Top             =   1440
            Width           =   2295
         End
         Begin VB.CommandButton cmddelete3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Delete"
            Height          =   855
            Left            =   1560
            Picture         =   "frmuser.frx":0074
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   2640
            Width           =   1095
         End
         Begin VB.TextBox txtusername3 
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
            TabIndex        =   17
            Top             =   840
            Width           =   2295
         End
         Begin VB.CommandButton cmdback3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Back"
            Height          =   855
            Left            =   2760
            Picture         =   "frmuser.frx":0A5E
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   2640
            Width           =   1095
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
            TabIndex        =   48
            Top             =   1455
            Width           =   1215
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
            TabIndex        =   47
            Top             =   855
            Width           =   1215
         End
         Begin VB.Label Label24 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Label10"
            Height          =   375
            Left            =   1920
            TabIndex        =   46
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label23 
            BackColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   1920
            TabIndex        =   45
            Top             =   1680
            Width           =   2175
         End
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3060
         ItemData        =   "frmuser.frx":1239
         Left            =   240
         List            =   "frmuser.frx":123B
         TabIndex        =   43
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdback 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Back"
         Height          =   855
         Left            =   4560
         Picture         =   "frmuser.frx":123D
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3405
         Width           =   1575
      End
      Begin VB.CommandButton cmdchange 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Change Password"
         Height          =   855
         Left            =   4560
         MaskColor       =   &H008080FF&
         Picture         =   "frmuser.frx":1A18
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2550
         Width           =   1575
      End
      Begin VB.CommandButton cmdadd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Add User"
         Height          =   855
         Left            =   4560
         Picture         =   "frmuser.frx":2402
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Delete User"
         Height          =   855
         Left            =   4560
         Picture         =   "frmuser.frx":2B6C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1695
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   3855
         Index           =   1
         Left            =   -74760
         TabIndex        =   27
         Top             =   480
         Width           =   6015
         Begin VB.CommandButton cmdback2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Back"
            Height          =   855
            Left            =   4560
            Picture         =   "frmuser.frx":3556
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   2040
            Width           =   1215
         End
         Begin VB.CommandButton cmdsave1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Save"
            Height          =   855
            Left            =   4560
            Picture         =   "frmuser.frx":3D31
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtverypass 
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
            Left            =   1920
            PasswordChar    =   "*"
            TabIndex        =   14
            Top             =   2880
            Width           =   2415
         End
         Begin VB.TextBox txtnewpass 
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
            Left            =   1920
            PasswordChar    =   "*"
            TabIndex        =   13
            Top             =   2160
            Width           =   2415
         End
         Begin VB.TextBox txtcurpass 
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
            Left            =   1920
            PasswordChar    =   "*"
            TabIndex        =   12
            Top             =   1440
            Width           =   2415
         End
         Begin VB.TextBox txtusername1 
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
            Left            =   1920
            TabIndex        =   11
            Top             =   720
            Width           =   2415
         End
         Begin VB.Label Label18 
            BackColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   2280
            TabIndex        =   41
            Top             =   2400
            Width           =   2175
         End
         Begin VB.Label Label17 
            BackColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   2160
            TabIndex        =   40
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label Label16 
            BackColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   2160
            TabIndex        =   39
            Top             =   1680
            Width           =   2295
         End
         Begin VB.Label Label15 
            BackColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   2280
            TabIndex        =   38
            Top             =   3120
            Width           =   2175
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Verify Password:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   2955
            Width           =   1575
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "New Password:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   2235
            Width           =   1455
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Current Password:"
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
            Left            =   120
            TabIndex        =   29
            Top             =   1455
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   " Username:"
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
            Left            =   120
            TabIndex        =   28
            Top             =   735
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   3855
         Index           =   0
         Left            =   -74640
         TabIndex        =   22
         Top             =   600
         Width           =   5895
         Begin VB.CommandButton cmdback1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Back"
            Height          =   855
            Left            =   4440
            Picture         =   "frmuser.frx":471B
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1800
            Width           =   1215
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
            Left            =   1800
            TabIndex        =   4
            Top             =   600
            Width           =   2295
         End
         Begin VB.CommandButton cmdsave 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Save"
            Height          =   855
            Left            =   4440
            Picture         =   "frmuser.frx":4EF6
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtname 
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
            Left            =   1800
            TabIndex        =   7
            Top             =   2400
            Width           =   2295
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
            Left            =   1800
            PasswordChar    =   "*"
            TabIndex        =   5
            Top             =   1200
            Width           =   2295
         End
         Begin VB.TextBox txtconfirmpass 
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
            Left            =   1800
            PasswordChar    =   "*"
            TabIndex        =   6
            Top             =   1800
            Width           =   2295
         End
         Begin VB.ComboBox cmblevel 
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
            Left            =   1800
            TabIndex        =   8
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label14 
            BackColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   2520
            TabIndex        =   37
            Top             =   3240
            Width           =   975
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   2160
            TabIndex        =   36
            Top             =   2040
            Width           =   2175
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   2160
            TabIndex        =   35
            Top             =   2640
            Width           =   2175
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   2160
            TabIndex        =   34
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label Label10 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Label10"
            Height          =   375
            Left            =   2160
            TabIndex        =   33
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   32
            Top             =   2475
            Width           =   735
         End
         Begin VB.Label Label1 
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
            Left            =   480
            TabIndex        =   26
            Top             =   615
            Width           =   1215
         End
         Begin VB.Label Label2 
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
            Left            =   480
            TabIndex        =   25
            Top             =   1215
            Width           =   1215
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm Password:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   24
            Top             =   1875
            Width           =   855
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "User Level:"
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
            Left            =   480
            TabIndex        =   23
            Top             =   3090
            Width           =   1335
         End
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2970
         TabIndex        =   51
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   660
         TabIndex        =   50
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmuser.frx":58E0
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "User Account"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   42
      Top             =   240
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      Height          =   5055
      Left            =   120
      Top             =   1080
      Width           =   6735
   End
End
Attribute VB_Name = "frmuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsuser As New ADODB.Recordset

Private Sub cmdAdd_Click()
SSTab1.Tab = 1
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = True
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
txtusername.Text = ""
txtpassword.Text = ""
txtconfirmpass.Text = ""
txtname.Text = ""
cmblevel.Text = ""
txtusername.SetFocus
End Sub

Private Sub cmdback_Click()
Unload Me
End Sub

Private Sub cmdback1_Click()
Dim ans As Integer
Dim ans1 As Integer
ans1 = MsgBox("Exit Now ?", vbYesNo + vbDefaultButton2 + vbQuestion, "Confirm")
If ans1 = vbYes Then
ans = MsgBox("Save This Record", vbYesNo + vbDefaultButton2 + vbInformation, "Confirm")
If ans = vbYes Then
cmdsave.SetFocus
Exit Sub
End If
'frmmain.Show
Me.Hide
Else
txtusername.SetFocus
End If
End Sub

Private Sub cmdback2_Click()
Dim ans As Integer
Dim ans1 As Integer
ans1 = MsgBox("Exit Now ?", vbYesNo + vbDefaultButton2 + vbQuestion, "Confirm")
If ans1 = vbYes Then
ans = MsgBox("Save This Record", vbYesNo + vbDefaultButton2 + vbInformation, "Confirm")
If ans = vbYes Then
cmdsave1.SetFocus
Exit Sub
End If
'frmmain.Show
Me.Hide
Else
txtusername1.SetFocus
End If
End Sub

Private Sub cmdback3_Click()
'frmmain.Show
Me.Hide
End Sub

Private Sub cmdchange_Click()
SSTab1.Tab = 2
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = True
SSTab1.TabEnabled(3) = False
txtcurpass.Text = ""
txtnewpass.Text = ""
txtverypass.Text = ""
txtusername1.SetFocus
End Sub

Private Sub cmdDelete_Click()
SSTab1.Tab = 3
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = True
txtpassword3.Text = ""
txtusername3.SetFocus
End Sub

Private Sub cmddelete3_Click()
Dim ans As Integer
Dim CurrentRec As Variant
 With rsuser
 If .State = closed Then
 .Open
 End If
 If txtusername3.Text = "" Then
 MsgBox "Write User Name", vbInformation, "Empty User"
 txtusername3.SetFocus
 Exit Sub
 End If
 If txtpassword3.Text = "" Then
 MsgBox "Write Password", vbInformation, "Empty Password"
 txtpassword3.SetFocus
 Exit Sub
 End If
 .MoveFirst
 .Find "USERNAME='" & txtusername3 & "'"
If txtpassword3.Text <> .Fields("password") Then
    MsgBox "Wrong Password", vbCritical, "Wrong Password"
    txtpassword3.SetFocus
    Exit Sub
End If
ans = MsgBox("Are You Sure Delete This Record", vbYesNo + vbDefaultButton2 + vbInformation, "Confirm")
If ans = vbYes Then
.Delete
MsgBox "Record Has Been Deleted.", vbInformation, "Delete Record"
Call fillusername
txtusername3.Text = ""
txtpassword3.Text = ""
txtusername3.SetFocus
Else
SSTab1.Tab = 0
SSTab1.TabEnabled(0) = True
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
txtusername3.Text = ""
txtpassword3.Text = ""
End If
End With

End Sub

Private Sub cmdsave_Click()
With rsuser
If .State = closed Then
.Open
End If
If Trim(txtusername.Text) = "" Then
MsgBox "User Name Is Empty", vbInformation, "Save"
txtusername.SetFocus
Exit Sub
ElseIf Trim(txtpassword.Text) = "" Then
MsgBox "Password Is Empty", vbInformation, "Save"
txtpassword.SetFocus
Exit Sub
ElseIf Trim(txtname.Text) = "" Then
MsgBox "Name Is Empty", vbInformation, "Save"
txtname.SetFocus
Exit Sub
ElseIf Trim(cmblevel.Text) = "" Then
MsgBox "User Level Is Empty", vbInformation, "Save"
cmblevel.SetFocus
Exit Sub
Else
.AddNew
.Fields("USERID") = cmblevel.Text
.Fields("username") = UCase(txtusername.Text)
.Fields("password") = txtpassword.Text
.Fields("name") = UCase(txtname.Text)
.Update
MsgBox "Saved Record", vbInformation, "Save"
Call fillusername
txtusername.Text = ""
txtpassword.Text = ""
txtconfirmpass.Text = ""
txtname.Text = ""
cmblevel.Text = ""
txtusername.SetFocus
End If
End With

End Sub

Private Sub cmdsave1_Click()
Dim CurrentRec As Variant
 With rsuser
 If .State = closed Then
 .Open
 End If
 .MoveFirst
 .Find "USERNAME ='" & txtusername1 & "'"
 CurrentRec = .AbsolutePosition
 .AbsolutePosition = CurrentRec
If txtcurpass.Text <> .Fields("password") Then
    MsgBox "Wrong Password", vbCritical, "Wrong Password"
End If
If txtusername1.Text = "" Then
MsgBox "User Name Is Empty", vbInformation, "Save"
txtusername1.SetFocus
Exit Sub
ElseIf txtcurpass.Text = "" Then
MsgBox "Current Password Is Empty", vbInformation, "Save"
txtcurpass.SetFocus
Exit Sub
ElseIf txtnewpass.Text = "" Then
MsgBox "New Password Is Empty", vbInformation, "Save"
txtnewpass.SetFocus
Exit Sub
ElseIf txtverypass.Text = "" Then
MsgBox "Verify PassWord Is Empty", vbInformation, "Save"
txtverypass.SetFocus
Exit Sub
Else
.Fields("Password") = txtnewpass.Text
.Update
MsgBox "Password successfully Change", vbInformation, "Changes Saved"
Call fillusername
txtusername1.Text = ""
txtcurpass.Text = ""
txtnewpass.Text = ""
txtverypass.Text = ""
txtusername1.SetFocus
End If
End With
 
End Sub

Private Sub Form_Activate()
SSTab1.Tab = 0
SSTab1.TabEnabled(0) = True
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
End Sub

Private Sub Form_Load()
Call conopen
If rsuser.State = 0 Then
rsuser.Open "select * from user_details", cn, adOpenDynamic, adLockOptimistic
End If
SSTab1.Tab = 0
SSTab1.TabEnabled(0) = True
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False

cmblevel.AddItem "Admin"
cmblevel.AddItem "User"
cmblevel.Text = "Admin"

Call fillusername

End Sub




Private Sub List1_Click()
With rsuser
Dim X As Integer
X = List1.ListIndex + 1
If .BOF = True Then
X = 0
End If
.AbsolutePosition = X
Me.txtusername1.Text = .Fields("username")
Me.txtusername3.Text = .Fields("username")
End With

End Sub

Private Sub txtconfirmpass_LostFocus()
If txtpassword.Text <> txtconfirmpass Then
MsgBox "Password Is Not Match Reconfirm Password", vbCritical, "Reconfirm Password"
txtconfirmpass.SetFocus
txtconfirmpass.Text = ""
Exit Sub
End If
End Sub





Private Sub txtcurpass_GotFocus()
 Dim CurrentRec As Variant
 With rsuser
 If .State = closed Then
 .Open
 End If
 .MoveFirst
 CurrentRec = .Bookmark
 .Find "USERNAME ='" & txtusername1 & "'"
 If .EOF Then
      MsgBox "User Not Found", vbInformation, "Invelid User Name"
      txtusername1.SetFocus
      Exit Sub
  Else
    txtcurpass.SetFocus
    Exit Sub
  End If
  End With

End Sub

Private Sub txtcurpass_LostFocus()
 Dim CurrentRec As Variant
 With rsuser
 If .State = closed Then
 .Open
 End If
 .MoveFirst
 .Find "USERNAME ='" & txtusername1 & "'"
 CurrentRec = .AbsolutePosition
 If .EOF Then
      MsgBox "User Not Found", vbInformation, "Invelid User Name"
      txtusername1.SetFocus
      Exit Sub
  Else
    .AbsolutePosition = CurrentRec
  End If
  If txtcurpass.Text <> .Fields("password") Then
    MsgBox "Wrong Password", vbCritical, "Wrong Password"
    txtcurpass.SetFocus
    txtcurpass.Text = ""
    Exit Sub
  End If
  End With
  txtnewpass.SetFocus
  Exit Sub
End Sub


Private Sub txtpassword_GotFocus()
 Dim CurrentRec As Variant
 With rsuser
 If .State = closed Then
 .Open
 End If
 .MoveFirst
 .Find "USERNAME ='" & txtusername & "'"
  If .EOF Then
    txtpassword.SetFocus
    Exit Sub
  Else
      MsgBox "User Is Alredy Exits", vbInformation, "Duplicate User Name"
      txtusername.SetFocus
      Exit Sub
  End If
  End With
End Sub




Private Sub txtpassword3_GotFocus()
 Dim CurrentRec As Variant
 With rsuser
 If .State = closed Then
 .Open
 End If
 .MoveFirst
 .Find "USERNAME ='" & txtusername3 & "'"
 If .EOF Then
      MsgBox "User Not Found", vbInformation, "Invelid User Name"
      txtusername3.SetFocus
      Exit Sub
  Else
    txtpassword3.SetFocus
    Exit Sub
  End If
  End With

End Sub



Private Sub txtverypass_LostFocus()
If txtnewpass.Text <> txtverypass.Text Then
MsgBox "Password Not Match Verify Password Again", vbInformation, "Verify Password"
txtverypass.SetFocus
txtverypass.Text = ""
Exit Sub
End If

End Sub
Public Sub fillusername()
List1.Clear
List2.Clear
With rsuser
If .State = closed Then
.Open
End If
.MoveFirst
While .EOF = False
List1.AddItem UCase(LTrim(.Fields("username")))
List2.AddItem UCase(LTrim(.Fields("name")))
.MoveNext
Wend
End With
End Sub
