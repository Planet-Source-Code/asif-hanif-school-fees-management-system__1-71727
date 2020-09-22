VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmviewall 
   Caption         =   "View All Records"
   ClientHeight    =   6915
   ClientLeft      =   3450
   ClientTop       =   1755
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   8175
   Begin VB.CommandButton cmdlast 
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   4620
      Picture         =   "frmviewall.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   3540
      Picture         =   "frmviewall.frx":0C43
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdpre 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   2460
      Picture         =   "frmviewall.frx":1886
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdfirst 
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   1380
      Picture         =   "frmviewall.frx":24C9
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdmain 
      Caption         =   "Main"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   5700
      Picture         =   "frmviewall.frx":310C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dg 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   765
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblnoofstudent 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "No Of Student :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   1095
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   6015
   End
   Begin VB.Label Label1 
      Caption         =   "View All Students Records"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmviewall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdfirst_Click()
rsSTUDENT.MoveFirst
End Sub

Private Sub cmdlast_Click()
rsSTUDENT.MoveLast
With rsSTUDENT
If .EOF Then
MsgBox "This is the Last Recorods"
.MoveLast
End If
End With
End Sub

Private Sub cmdmain_Click()
Unload Me
End Sub

Private Sub cmdnext_Click()
rsSTUDENT.MoveNext
With rsSTUDENT
If .EOF Then
MsgBox "This is the Last Recorods"
.MoveLast
End If
End With
End Sub

Private Sub cmdpre_Click()
rsSTUDENT.MovePrevious
With rsSTUDENT
If .BOF Then
MsgBox "This is the First Recorods"
.MoveFirst
End If
End With
End Sub

Private Sub Form_Load()
Call conopen
If rsSTUDENT.State = closed Then
Call studentopen
End If
Set dg.DataSource = rsSTUDENT
With rsSTUDENT
lblnoofstudent.Caption = .RecordCount
End With
End Sub
