VERSION 5.00
Begin VB.Form frmmarksheet 
   Caption         =   "Marksheet"
   ClientHeight    =   5475
   ClientLeft      =   2745
   ClientTop       =   2085
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm1 
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtmarks 
         Height          =   375
         Index           =   0
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   12
         Top             =   885
         Width           =   735
      End
      Begin VB.TextBox txtmarks 
         Height          =   375
         Index           =   1
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   11
         Top             =   1365
         Width           =   735
      End
      Begin VB.TextBox txtmarks 
         Height          =   375
         Index           =   2
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   10
         Top             =   1845
         Width           =   735
      End
      Begin VB.TextBox txtmarks 
         Height          =   375
         Index           =   3
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   9
         Top             =   2325
         Width           =   735
      End
      Begin VB.TextBox txtmarks 
         Height          =   375
         Index           =   4
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   8
         Top             =   2805
         Width           =   735
      End
      Begin VB.TextBox txtmarks 
         Height          =   375
         Index           =   5
         Left            =   4545
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   885
         Width           =   735
      End
      Begin VB.TextBox txtmarks 
         Height          =   375
         Index           =   6
         Left            =   4545
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1365
         Width           =   735
      End
      Begin VB.TextBox txtmarks 
         Height          =   375
         Index           =   7
         Left            =   4545
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1845
         Width           =   735
      End
      Begin VB.TextBox txtmarks 
         Height          =   375
         Index           =   8
         Left            =   3112
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2820
         Width           =   2400
      End
      Begin VB.CommandButton Command1 
         Caption         =   "MARKSHEET"
         Height          =   480
         Left            =   2482
         TabIndex        =   3
         Top             =   4245
         Width           =   1410
      End
      Begin VB.CommandButton Command2 
         Caption         =   "NEW"
         Height          =   480
         Left            =   4102
         TabIndex        =   2
         Top             =   4245
         Width           =   1410
      End
      Begin VB.CommandButton Command3 
         Caption         =   "EXIT"
         Height          =   480
         Left            =   877
         TabIndex        =   1
         Top             =   4245
         Width           =   1410
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Marksheet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   60
         TabIndex        =   22
         Top             =   0
         Width           =   6225
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "English"
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   945
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Urdu"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   20
         Top             =   1425
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Math"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   19
         Top             =   1905
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Science"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   18
         Top             =   2385
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Computer"
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   17
         Top             =   2865
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   3120
         Left            =   255
         Top             =   570
         Width           =   2550
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Total"
         Height          =   240
         Left            =   3480
         TabIndex        =   16
         Top             =   945
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Percentage"
         Height          =   240
         Left            =   3480
         TabIndex        =   15
         Top             =   1425
         Width           =   945
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Grade"
         Height          =   240
         Left            =   3480
         TabIndex        =   14
         Top             =   1905
         Width           =   945
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks"
         Height          =   240
         Left            =   3840
         TabIndex        =   13
         Top             =   2505
         Width           =   945
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   3
         Height          =   3120
         Left            =   2790
         Top             =   570
         Width           =   2985
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H008080FF&
         BorderWidth     =   2
         Height          =   645
         Index           =   0
         Left            =   765
         Shape           =   4  'Rounded Rectangle
         Top             =   4170
         Width           =   1635
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H008080FF&
         BorderWidth     =   2
         Height          =   645
         Index           =   1
         Left            =   2370
         Shape           =   4  'Rounded Rectangle
         Top             =   4170
         Width           =   1635
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H008080FF&
         BorderWidth     =   2
         Height          =   645
         Index           =   2
         Left            =   3990
         Shape           =   4  'Rounded Rectangle
         Top             =   4170
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmmarksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TOTAL, per, i As Integer


Private Sub Command1_Click()

' Validations
' Empty checking
For i = 0 To 4
If Trim(txtmarks(i).Text) = "" Then
MsgBox "Please Type any numeric value"
txtmarks(i).Text = ""
txtmarks(i).SetFocus
Exit Sub
End If
Next i
'numeric value

For i = 0 To 4
If IsNumeric(txtmarks(i).Text) = False Then
MsgBox "Please Type any numeric value"
txtmarks(i).Text = ""
txtmarks(i).SetFocus
Exit Sub
End If
Next i

For i = 0 To 4
If txtmarks(i).Text <= 0 Then
MsgBox "Do not Type less then value"
txtmarks(i).Text = ""
txtmarks(i).SetFocus
ElseIf txtmarks(i).Text >= 100 Then
MsgBox "Please Type Tow Digit value"
txtmarks(i).Text = ""
txtmarks(i).SetFocus
Exit Sub
End If
Next i









TOTAL = Val(txtmarks(0).Text) + Val(txtmarks(1).Text) + Val(txtmarks(2).Text) + Val(txtmarks(3).Text) + Val(txtmarks(4).Text)
txtmarks(5).Text = TOTAL
per = TOTAL / 500 * 100
txtmarks(6).Text = per & " %"
If per >= 80 Then

txtmarks(7).Text = "A-One"
txtmarks(8).Text = "Excellent"
ElseIf per >= 70 Then
txtmarks(7).Text = "A"
txtmarks(8).Text = "Very Good"
ElseIf per >= 60 Then
txtmarks(7).Text = "B"
txtmarks(8).Text = "Good"
ElseIf per >= 50 Then
txtmarks(7).Text = "C"
txtmarks(8).Text = "Fair"
ElseIf per >= 40 Then
txtmarks(7).Text = "D"
txtmarks(8).Text = "Do Hardwork"
Else

txtmarks(7).Text = "Failed"
txtmarks(8).Text = "Very Poor"
End If

End Sub

Private Sub Command2_Click()
For i = 0 To 8
txtmarks(i).Text = ""
Next i
txtmarks(0).SetFocus
End Sub

Private Sub Command3_Click()
Me.Hide
End Sub

