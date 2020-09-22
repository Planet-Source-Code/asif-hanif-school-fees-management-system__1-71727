VERSION 5.00
Object = "{C5DE211B-7681-446C-89FF-487BF2A2BA98}#1.0#0"; "menubar.ocx"
Begin VB.Form frmmain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Main Menu"
   ClientHeight    =   6510
   ClientLeft      =   2595
   ClientTop       =   2760
   ClientWidth     =   9600
   FillColor       =   &H80000001&
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmmain.frx":0ECA
   ScaleHeight     =   6510
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H80000003&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdLogout 
      BackColor       =   &H80000003&
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin MenuBarOcx.MenuBar menu 
      Height          =   6375
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   11245
      BorderStyle     =   2
      ButtonBackColor =   12582912
      ButtonForeColor =   16777152
      ButtonGradientColor=   14737632
      ButtonGradientType=   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBoldButtonCaption=   0   'False
      ItemAlignment   =   0
      ItemBackColor   =   16773862
      ItemForeColor   =   12582912
      ItemIconSize    =   0
      MaxMenus        =   9
      ButtonCaption1  =   "School Details"
      ButtonIcon1     =   "frmmain.frx":B145
      ItemCaption1_1  =   "About School"
      ItemIcon1_1     =   "frmmain.frx":B161
      ButtonCaption2  =   "Registration"
      ButtonIcon2     =   "frmmain.frx":C03B
      ItemCaption2_1  =   "New Student"
      ItemIcon2_1     =   "frmmain.frx":CF15
      ButtonCaption3  =   "Class & Subject"
      ButtonIcon3     =   "frmmain.frx":D90F
      ItemCaption3_1  =   "Class Registration"
      ItemIcon3_1     =   "frmmain.frx":E309
      ButtonCaption4  =   "Fee Payment"
      ButtonIcon4     =   "frmmain.frx":ED03
      ItemCaption4_1  =   "Receipt Fee"
      ItemIcon4_1     =   "frmmain.frx":F6FD
      ButtonCaption5  =   "Utility Tools"
      MaxItems5       =   2
      ItemCaption5_1  =   "Backup Database"
      ItemIcon5_1     =   "frmmain.frx":100F7
      ItemCaption5_2  =   "Restore Database"
      ItemIcon5_2     =   "frmmain.frx":10491
      ButtonCaption6  =   "Reports Menu"
      ButtonIcon6     =   "frmmain.frx":107AB
      MaxItems6       =   4
      ItemCaption6_1  =   "Student Details"
      ItemIcon6_1     =   "frmmain.frx":116DD
      ItemCaption6_2  =   "Fee Receipt Details"
      ItemIcon6_2     =   "frmmain.frx":1260F
      ItemCaption6_3  =   "Subject Details"
      ItemIcon6_3     =   "frmmain.frx":13541
      ItemCaption6_4  =   "All Student Records"
      ItemIcon6_4     =   "frmmain.frx":14473
      ButtonCaption7  =   "User"
      ButtonIcon7     =   "frmmain.frx":1534D
      MaxItems7       =   2
      MenuPassword7   =   -1  'True
      ItemCaption7_1  =   "User"
      ItemIcon7_1     =   "frmmain.frx":16227
      ItemCaption7_2  =   "User Report"
      ItemIcon7_2     =   "frmmain.frx":169A1
      ButtonCaption8  =   "Entertenment"
      MaxItems8       =   2
      ItemCaption8_1  =   "Game"
      ItemIcon8_1     =   "frmmain.frx":1787B
      ItemCaption8_2  =   "Game"
      ItemIcon8_2     =   "frmmain.frx":17B95
      ButtonCaption9  =   "About Programmar"
      ButtonIcon9     =   "frmmain.frx":17EAF
      ItemCaption9_1  =   "Asif Hanif"
      ItemIcon9_1     =   "frmmain.frx":17ECB
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Student Registration Management System"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2880
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdLogin_Click()
frmlogin.Show
frmlogin.txtusername = ""
frmlogin.txtpassword = ""
frmlogin.txtusername.SetFocus
End Sub

Private Sub cmdLogout_Click()
Dim ans As Integer
ans = MsgBox("Do You Want To Exit This Programm ?", vbYesNo + vbDefaultButton2 + vbQuestion, "Exit")
If ans = vbYes Then
    'Call DoExitWindows
    Dim RetValue
    RetValue = ExitApplication(Val(OldWidth1), Val(OldHeight1), Val(OldBPP1))
    If cn.State = 1 Then cn.Close
    End
Else
    Exit Sub
End If
End Sub








Private Sub menu_MenuItemClick(MenuIndex As Integer, ItemIndex As Integer, ItemKey As String, ItemType As MenuBarOcx.ItemTypes, ItemValue As Boolean, PreviousOption As Integer)
If (MenuIndex = 1) And (ItemIndex = 1) Then frmAboutschool.Show vbModal
If (MenuIndex = 2) And (ItemIndex = 1) Then frmstudent.Show
If (MenuIndex = 3) And (ItemIndex = 1) Then frmsubjects.Show
If (MenuIndex = 4) And (ItemIndex = 1) Then frmfeepayment.Show
If (MenuIndex = 5) And (ItemIndex = 1) Then frmBackupDatabase.Show vbModal
If (MenuIndex = 5) And (ItemIndex = 2) Then frmRestore.Show vbModal
'If (MenuIndex = 5) And (ItemIndex = 1) Then frmmarksheet.Show vbModal
If (MenuIndex = 6) And (ItemIndex = 1) Then frmstudentprint.Show vbModal
If (MenuIndex = 6) And (ItemIndex = 2) Then frmreceiptprint.Show vbModal
If (MenuIndex = 6) And (ItemIndex = 3) Then frmclassprint.Show vbModal
If (MenuIndex = 6) And (ItemIndex = 4) Then frmviewall.Show vbModal
If (MenuIndex = 7) And (ItemIndex = 1) Then frmuser.Show vbModal
If (MenuIndex = 7) And (ItemIndex = 2) Then frmuserprint.Show vbModal
If (MenuIndex = 8) And (ItemIndex = 1) Then
Dim ans1 As Integer
ans1 = MsgBox("Do You Want To Play Game 'Click Me' ?", vbYesNo + vbDefaultButton2 + vbInformation, "Confermation")
If ans1 = vbYes Then
     frmClickme.Show vbModal
End If
End If
If (MenuIndex = 8) And (ItemIndex = 2) Then
Dim ans2 As Integer
ans2 = MsgBox("Do You Want To Play Game '3-D-Maze' ?", vbYesNo + vbDefaultButton2 + vbInformation, "Confermation")
If ans2 = vbYes Then
     frm3DMaze.Show vbModal
End If
End If
If (MenuIndex = 9) And (ItemIndex = 1) Then frmAbout.Show vbModal
'Dim ans As Integer
'ans = MsgBox("Do You Want To Exit", vbYesNo + vbDefaultButton2 + vbInformation, "Exit")
'If ans = vbYes Then
'    If cn.State = 1 Then cn.Close
'    End
'Else
'    Exit Sub
'End If
'End If

End Sub
