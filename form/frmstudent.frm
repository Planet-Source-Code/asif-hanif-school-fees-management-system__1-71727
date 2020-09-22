VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Begin VB.Form frmstudent 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Student Registration"
   ClientHeight    =   9075
   ClientLeft      =   1530
   ClientTop       =   1425
   ClientWidth     =   10710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   10710
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
      Height          =   360
      Left            =   2760
      MaxLength       =   30
      TabIndex        =   1
      Top             =   3165
      Width           =   3375
   End
   Begin VB.TextBox txtdescr 
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
      Left            =   6240
      MaxLength       =   30
      TabIndex        =   51
      Top             =   7320
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   3480
      TabIndex        =   46
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton cmdprint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   840
         Picture         =   "frmstudent.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdfind 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmstudent.frx":0C43
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1560
         Picture         =   "frmstudent.frx":1885
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.ComboBox cmbstatus 
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
      Left            =   6240
      TabIndex        =   15
      Top             =   8280
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DG 
      CausesValidation=   0   'False
      Height          =   1215
      Left            =   240
      TabIndex        =   38
      Top             =   1320
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   2143
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   240
      TabIndex        =   39
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton cmdfirst 
         BackColor       =   &H80000009&
         Caption         =   "First"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmstudent.frx":2060
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdpre 
         BackColor       =   &H80000009&
         Caption         =   "Previous"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   840
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmstudent.frx":2CA3
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdnext 
         BackColor       =   &H80000009&
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmstudent.frx":38E6
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdlast 
         BackColor       =   &H80000009&
         Caption         =   "Last"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2280
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmstudent.frx":4529
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.TextBox txtmobile 
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
      Left            =   6240
      MaxLength       =   12
      TabIndex        =   6
      Top             =   4920
      Width           =   1815
   End
   Begin VB.ComboBox cmbclasstime 
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
      Left            =   2760
      TabIndex        =   14
      Top             =   8280
      Width           =   1575
   End
   Begin VB.ComboBox cmbshift 
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
      Left            =   2760
      TabIndex        =   13
      Top             =   7800
      Width           =   1575
   End
   Begin VB.ComboBox cmbreligion 
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
      Left            =   2760
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox txtpinstitution 
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
      Left            =   2760
      MaxLength       =   30
      TabIndex        =   11
      Top             =   6840
      Width           =   4575
   End
   Begin VB.ComboBox cmbsex 
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
      Left            =   6240
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   5400
      Width           =   1815
   End
   Begin VB.ComboBox cmbpqualification 
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
      Left            =   2760
      TabIndex        =   10
      Top             =   6360
      Width           =   1575
   End
   Begin VB.ComboBox cmbclass 
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
      Left            =   2760
      TabIndex        =   12
      Top             =   7320
      Width           =   1575
   End
   Begin VB.TextBox txtfname 
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
      Left            =   2760
      MaxLength       =   30
      TabIndex        =   2
      Top             =   3585
      Width           =   3375
   End
   Begin VB.TextBox txtgrno 
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
      Left            =   2760
      LinkTimeout     =   70
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2715
      Width           =   1455
   End
   Begin VB.ComboBox cmbage 
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
      Left            =   2760
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   5415
      Width           =   1575
   End
   Begin VB.TextBox txttel 
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
      Left            =   2760
      MaxLength       =   11
      TabIndex        =   5
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox txtaddress 
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
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   3
      Top             =   3990
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   6000
      TabIndex        =   23
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3720
         Picture         =   "frmstudent.frx":516C
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3000
         Picture         =   "frmstudent.frx":5A36
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2280
         Picture         =   "frmstudent.frx":6420
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1560
         Picture         =   "frmstudent.frx":6E0A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   840
         Picture         =   "frmstudent.frx":738C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   120
         Picture         =   "frmstudent.frx":7D76
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.TextBox txtpict 
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
      Left            =   9360
      MaxLength       =   70
      TabIndex        =   40
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin TDBDate6Ctl.TDBDate txtdob 
      Bindings        =   "frmstudent.frx":8760
      Height          =   315
      Left            =   1530
      TabIndex        =   4
      Top             =   4440
      Width           =   2970
      _Version        =   65536
      _ExtentX        =   5239
      _ExtentY        =   556
      Calendar        =   "frmstudent.frx":876B
      Caption         =   "frmstudent.frx":8883
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmstudent.frx":88DF
      Keys            =   "frmstudent.frx":88FD
      Spin            =   "frmstudent.frx":895B
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   16777215
      BorderStyle     =   1
      BtnPositioning  =   1
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   1
      DisplayFormat   =   "mmm dd, yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
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
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Description :"
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
      Index           =   17
      Left            =   4680
      TabIndex        =   52
      Top             =   7380
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Picture "
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
      Index           =   16
      Left            =   9300
      TabIndex        =   50
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Status :"
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
      Index           =   15
      Left            =   4680
      TabIndex        =   45
      Top             =   8340
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      X1              =   240
      X2              =   10560
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Image picc 
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Left            =   8760
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click Open Picture Browser"
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No :"
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
      Index           =   14
      Left            =   4680
      TabIndex        =   37
      Top             =   4980
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Class Timing :"
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
      Index           =   13
      Left            =   240
      TabIndex        =   36
      Top             =   8340
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Shift :"
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
      Index           =   12
      Left            =   240
      TabIndex        =   35
      Top             =   7860
      Width           =   735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Religion :"
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
      Index           =   11
      Left            =   240
      TabIndex        =   34
      Top             =   5940
      Width           =   975
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Institution Name:"
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
      Index           =   9
      Left            =   240
      TabIndex        =   33
      Top             =   6900
      Width           =   2415
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender :"
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
      Index           =   10
      Left            =   4680
      TabIndex        =   32
      Top             =   5460
      Width           =   735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Previous Qualification :"
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
      Left            =   240
      TabIndex        =   31
      Top             =   6420
      Width           =   2295
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Addmission in Class :"
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
      Left            =   240
      TabIndex        =   30
      Top             =   7380
      Width           =   2295
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Father Name :"
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
      Left            =   240
      TabIndex        =   29
      Top             =   3645
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
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
      Left            =   240
      TabIndex        =   28
      Top             =   3225
      Width           =   735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Address :"
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
      Left            =   240
      TabIndex        =   27
      Top             =   4050
      Width           =   855
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "D.O.B :"
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
      Left            =   240
      TabIndex        =   26
      Top             =   4470
      Width           =   735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Age :"
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
      Left            =   240
      TabIndex        =   25
      Top             =   5475
      Width           =   735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel :"
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
      Index           =   8
      Left            =   240
      TabIndex        =   24
      Top             =   4980
      Width           =   735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "GR No :"
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
      Left            =   240
      TabIndex        =   22
      Top             =   2775
      Width           =   735
   End
End
Attribute VB_Name = "frmstudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSTUDENT As New ADODB.Recordset
Dim rstemp As New ADODB.Recordset
Dim rsprint As New ADODB.Recordset


Private Sub cmbclass_Change()
If rstemp.State = 1 Then
rstemp.Close
End If
rstemp.Open "Select * from level_details where levelcode='" & cmbclass.Text & "'", cn, adOpenDynamic, adLockOptimistic
With rstemp
If .RecordCount <= 0 Then
txtdescr.Text = "Not Found"
Else
CurrentRec = .RecordCount
txtdescr.Text = UCase(.Fields("DESCR"))
End If
.Close
End With
End Sub

Private Sub cmbclass_Click()
If rstemp.State = 1 Then
rstemp.Close
End If
rstemp.Open "Select * from level_details where levelcode='" & cmbclass.Text & "'", cn, adOpenDynamic, adLockOptimistic
With rstemp
If .RecordCount <= 0 Then
txtdescr.Text = "Not Found"
Else
CurrentRec = .RecordCount
txtdescr.Text = UCase(.Fields("DESCR"))
End If
.Close
End With
End Sub



Private Sub cmdAdd_Click()

cmdadd.Enabled = False
cmdUpdate.Enabled = True
cmddelete.Enabled = False
cmdCancel.Enabled = True
cmdedit.Enabled = False
cmdprint.Enabled = False
cmdfind.Enabled = False
cmdfirst.Enabled = False
cmdpre.Enabled = False
cmdnext.Enabled = False
cmdlast.Enabled = False

' all fields blank for new entery

txtgrno.Text = ""
txtname.Text = ""
txtfname.Text = ""
txtaddress.Text = ""
txtdob.Value = Date
txttel.Text = ""
txtmobile.Text = ""
cmbage.Text = ""
cmbsex.Text = ""
cmbreligion.Text = ""
cmbpqualification.Text = ""
txtpinstitution.Text = ""
cmbclass.Text = ""
txtdescr.Text = ""
cmbshift.Text = ""
cmbclasstime.Text = ""
cmbstatus.Text = ""
txtpict.Text = ""
picc.Picture = LoadPicture("")

dg.Refresh
txtgrno.SetFocus
End Sub

Private Sub cmdCancel_Click()
cmdadd.Enabled = True
cmdUpdate.Enabled = False
cmddelete.Enabled = True
cmdCancel.Enabled = False
cmdedit.Enabled = True
cmdprint.Enabled = True
cmdfind.Enabled = True
cmdfirst.Enabled = True
cmdpre.Enabled = True
cmdnext.Enabled = True
cmdlast.Enabled = True

txtgrno.Text = ""
txtname.Text = ""
txtfname.Text = ""
txtaddress.Text = ""
txtdob.Value = ""
txttel.Text = ""
txtmobile.Text = ""
cmbage.Text = ""
cmbsex.Text = ""
cmbreligion.Text = ""
cmbpqualification.Text = ""
txtpinstitution.Text = ""
cmbclass.Text = ""
cmbshift.Text = ""
cmbclasstime.Text = ""
cmbstatus.Text = ""
txtpict.Text = ""
picc.Picture = LoadPicture("")
MsgBox "Record has been Cancel Updation"
dg.Refresh
rsSTUDENT.AbsolutePosition = 1
Call showdata
txtgrno.SetFocus
End Sub

Private Sub cmdDelete_Click()
Dim CurrentRec As Variant
With rsSTUDENT
If .RecordCount > 0 Then
CurrentRec = .AbsolutePosition
If MsgBox("Are you sure you want to delete this recorde", vbYesNo + vbCritical, "Warning") = vbYes Then
rsSTUDENT.Delete
MsgBox "Record has been deleted"
'.AbsolutePosition = CurrentRec
End If
'rsSTUDENT.AbsolutePosition = CurrentRec
dg.Refresh
Call showdata
End If
End With
End Sub

Private Sub cmdEdit_Click()
cmdadd.Enabled = True
cmdUpdate.Enabled = False
cmddelete.Enabled = True
cmdCancel.Enabled = False
cmdedit.Enabled = True
cmdprint.Enabled = True
cmdfind.Enabled = True
cmdfirst.Enabled = True
cmdpre.Enabled = True
cmdnext.Enabled = True
cmdlast.Enabled = True

With rsSTUDENT
.Fields(0) = txtgrno.Text
.Fields(1) = UCase(txtname.Text)
.Fields(2) = UCase(txtfname.Text)
.Fields(3) = UCase(txtaddress.Text)
.Fields(4) = txtdob.Value
.Fields(5) = txttel.Text
.Fields(6) = txtmobile.Text
.Fields(7) = cmbage.Text
.Fields(8) = cmbsex.Text
.Fields(9) = cmbreligion.Text
.Fields(10) = UCase(cmbpqualification.Text)
.Fields(11) = UCase(txtpinstitution.Text)
.Fields(12) = cmbclass.Text
.Fields(13) = cmbshift.Text
.Fields(14) = cmbclasstime.Text
.Fields(15) = txtpict.Text
.Fields(16) = cmbstatus.Text
.Update
MsgBox "Record has been Edit successfully"
dg.Refresh
Call showdata
End With
End Sub

Private Sub cmdexit_Click()
Me.Hide
End Sub
Private Sub cmdfind_Click()
frmfindstudent.Show
frmfindstudent.txtfind = ""
frmfindstudent.txtfind1 = ""
frmfindstudent.txtfind.SetFocus
End Sub

Private Sub cmdprint_Click()
With rsprint
    If .State = adStateOpen Then .Close
    .Open "SELECT * from student_details where GRNO='" & txtgrno & "'", cn, adOpenForwardOnly, adLockReadOnly
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        Exit Sub
    End If
    Set rptstudent.DataSource = rsprint
    rptstudent.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptstudent.Sections("section4").Controls("lbladdress1").Caption = address1
    rptstudent.Sections("section4").Controls("lbladdress2").Caption = address2
    rptstudent.Sections("section4").Controls("lblusername").Caption = pubusername
    rptstudent.Sections("section4").Controls("lbldate").Caption = Now
    rptstudent.Show vbModal
End With
End Sub

Private Sub cmdRefresh_Click()
dg.Refresh
End Sub

Private Sub cmdUpdate_Click()
Call chkdata(chk)
If chk = "True" Then
Exit Sub
End If
rsSTUDENT.AddNew
With rsSTUDENT
.Fields(0) = txtgrno.Text
.Fields(1) = UCase(txtname.Text)
.Fields(2) = UCase(txtfname.Text)
.Fields(3) = UCase(txtaddress.Text)
.Fields(4) = txtdob.Value
.Fields(5) = txttel.Text
.Fields(6) = txtmobile.Text
.Fields(7) = cmbage.Text
.Fields(8) = cmbsex.Text
.Fields(9) = cmbreligion.Text
.Fields(10) = UCase(cmbpqualification.Text)
.Fields(11) = UCase(txtpinstitution.Text)
.Fields(12) = cmbclass.Text
.Fields(13) = cmbshift.Text
.Fields(14) = cmbclasstime.Text
.Fields(15) = txtpict.Text
.Fields(16) = cmbstatus.Text
.Update
MsgBox "Record has been Saved successfully"
dg.Refresh
End With
cmdadd.Enabled = True
cmdUpdate.Enabled = False
cmddelete.Enabled = True
cmdCancel.Enabled = False
cmdedit.Enabled = True
cmdprint.Enabled = True
cmdfind.Enabled = True
cmdfirst.Enabled = True
cmdpre.Enabled = True
cmdnext.Enabled = True
cmdlast.Enabled = True
rsSTUDENT.MoveLast
txtgrno.SetFocus
End Sub


Private Sub DG_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'If rsSTUDENT.BOF = True Then
'rsSTUDENT.MoveFirst
'End If
'If rsSTUDENT.EOF = True Then
'rsSTUDENT.MoveLast
'End If
'Call showdata
End Sub

Private Sub Form_Load()
cmdadd.Enabled = True
cmdUpdate.Enabled = False
cmddelete.Enabled = True
cmdCancel.Enabled = False
cmdedit.Enabled = True
cmdprint.Enabled = False
cmdfind.Enabled = False
cmdfirst.Enabled = True
cmdpre.Enabled = True
cmdnext.Enabled = True
cmdlast.Enabled = True

Call conopen
If rsSTUDENT.State = 0 Then
rsSTUDENT.Open "select * from student order by grno", cn, adOpenDynamic, adLockOptimistic
End If
Set dg.DataSource = rsSTUDENT
With rsSTUDENT
If .State = closed Then
.Open
End If
If .RecordCount > 0 Then
cmdprint.Enabled = True
cmdfind.Enabled = True
txtgrno.Text = .Fields(0)
txtname.Text = .Fields(1)
txtfname.Text = .Fields(2)
txtaddress.Text = .Fields(3)
txtdob.Value = .Fields(4)
txttel.Text = .Fields(5)
txtmobile.Text = .Fields(6)
cmbage.Text = .Fields(7)
cmbsex.Text = .Fields(8)
cmbreligion.Text = .Fields(9)
cmbpqualification.Text = .Fields(10)
txtpinstitution.Text = .Fields(11)
cmbclass.Text = .Fields(12)
cmbshift.Text = .Fields(13)
cmbclasstime.Text = .Fields(14)
txtpict.Text = .Fields(15)
If txtpict.Text = "" Then
On Error Resume Next
picc.Picture = Nothing
Else
picc.Picture = LoadPicture(App.Path & "\picture\" + txtpict.Text)
End If
cmbstatus.Text = .Fields(16)
End If
End With

' fill combo box age
i = 15
For i = 15 To 99
cmbage.AddItem i
Next i
cmbage.Text = 15

' fill combo box sex
cmbsex.AddItem "Male"
cmbsex.AddItem "Female"
cmbsex.Text = "Male"

'fill combo box religion
cmbreligion.AddItem "Islam"
cmbreligion.AddItem "Non Muslim"
cmbreligion.Text = "Islam"

' fill combo box previous qualification
cmbpqualification.AddItem "None"
cmbpqualification.AddItem "PP1"
cmbpqualification.AddItem "PP2"
cmbpqualification.AddItem "KG1"
cmbpqualification.AddItem "KG2"
cmbpqualification.AddItem "CLASS 1"
cmbpqualification.AddItem "CLASS 2"
cmbpqualification.AddItem "CLASS 3"
cmbpqualification.AddItem "CLASS 4"
cmbpqualification.AddItem "CLASS 5"
cmbpqualification.AddItem "CLASS 6"
cmbpqualification.AddItem "CLASS 7"
cmbpqualification.AddItem "MIDDEL"
cmbpqualification.AddItem "MATRIC"
cmbpqualification.AddItem "FIRST YEAR"
cmbpqualification.AddItem "INTERMEDIT"
cmbpqualification.AddItem "B.COM"
cmbpqualification.AddItem "B.S.C"
cmbpqualification.AddItem "MASTER"
cmbpqualification.AddItem "C.A"
cmbpqualification.Text = "None"

' FILL COMBO BOX ADDMISSION CLASS
With rssubjects
If .State = closed Then
Call subjectsopen
End If
.AbsolutePosition = 1
While .EOF = False
cmbclass.AddItem .Fields("LevelCode")
.MoveNext
Wend
End With

'fill combo box shift
cmbshift.AddItem "Morning"
cmbshift.AddItem "After Non"
cmbshift.AddItem "Evening"
cmbshift.Text = "Morning"

'fill combo box class timing
cmbclasstime.AddItem "9 to 11"
cmbclasstime.AddItem "11 to 1"
cmbclasstime.AddItem "1 to 3"
cmbclasstime.AddItem "3 to 5"
cmbclasstime.AddItem "5 to 7"
cmbclasstime.AddItem "7 to 9"
cmbclasstime.Text = "9 to 11"

'fill combo box status
cmbstatus.AddItem "Available"
cmbstatus.AddItem "Dropped"
cmbstatus.AddItem "Locked"
cmbstatus.AddItem "Batch Clear"
cmbstatus.Text = "Available"

End Sub

Private Sub picc_DblClick()
On Error Resume Next
frmplaceholder.Show
End Sub
Private Sub cmdfirst_Click()
rsSTUDENT.MoveFirst
dg.Bookmark = rsSTUDENT.AbsolutePosition
Call showdata
End Sub

Private Sub cmdlast_Click()
rsSTUDENT.MoveLast
With rsSTUDENT
If .EOF Then
MsgBox "This is the Last Recorods"
.MoveLast
End If
End With
dg.Bookmark = rsSTUDENT.AbsolutePosition
Call showdata
End Sub

Private Sub cmdnext_Click()
rsSTUDENT.MoveNext
With rsSTUDENT
If .EOF Then
MsgBox "This is the Last Recorods"
.MoveLast
End If
End With
dg.Bookmark = rsSTUDENT.AbsolutePosition
Call showdata
End Sub

Private Sub cmdpre_Click()
rsSTUDENT.MovePrevious
With rsSTUDENT
If .BOF Then
MsgBox "This is the First Recorods"
.MoveFirst
End If
End With
dg.Bookmark = rsSTUDENT.AbsolutePosition
Call showdata
End Sub

Public Sub showdata()
With rsSTUDENT
If .RecordCount > 0 Then
txtgrno.Text = .Fields(0)
txtname.Text = .Fields(1)
txtfname.Text = .Fields(2)
txtaddress.Text = .Fields(3)
txtdob.Value = .Fields(4)
txttel.Text = .Fields(5)
txtmobile.Text = .Fields(6)
cmbage.Text = .Fields(7)
cmbsex.Text = .Fields(8)
cmbreligion.Text = .Fields(9)
cmbpqualification.Text = .Fields(10)
txtpinstitution.Text = .Fields(11)
cmbclass.Text = .Fields(12)
cmbshift.Text = .Fields(13)
cmbclasstime.Text = .Fields(14)
txtpict.Text = .Fields(15)
If txtpict.Text = "" Then
On Error Resume Next
picc.Picture = Nothing
Else
picc.Picture = LoadPicture(App.Path & "\picture\" + txtpict.Text)
End If
cmbstatus.Text = .Fields(16)
End If
End With
End Sub
Public Sub chkdata(chk)

chk = "False"

If Trim(txtgrno.Text) = "" Then
    MsgBox "Fill Gr No Is Not Empty", vbInformation, "Empty Box"
    txtgrno.SetFocus
    chk = "True"
    Exit Sub
End If

If Trim(txtname.Text) = "" Then
    MsgBox "Fill Name Box Is Not Empty", vbInformation, "Empty Box"
    txtname.SetFocus
    chk = "True"
    Exit Sub
End If


If Trim(txtfname.Text) = "" Then
    MsgBox "Fill Father Name Box Is Not Empty", vbInformation, "Empty Box"
    txtfname.SetFocus
    chk = "True"
    Exit Sub
End If

If Trim(txtaddress.Text) = "" Then
    MsgBox "Fill Address Box Is Not Empty", vbInformation, "Empty Box"
    txtname.SetFocus
    chk = "True"
    Exit Sub
End If

If Trim(txtdob.Value) = "" Then
    MsgBox "Fill Date Of Birth Box Is Not Empty", vbInformation, "Empty Box"
    txtdob.SetFocus
    chk = "True"
    Exit Sub
End If

If Trim(txttel.Text) = "" Then
    MsgBox "Fill Contect No Box Is Not Empty", vbInformation, "Empty Box"
    txttel.SetFocus
    chk = "True"
    Exit Sub
End If

If Trim(txtmobile.Text) = "" Then
    MsgBox "Fill Mobile No Box Is Not Empty", vbInformation, "Empty Box"
    txtmobile.SetFocus
    chk = "True"
    Exit Sub
End If

If Trim(cmbage.Text) = "" Then
    MsgBox "Fill Age Box Is Not Empty", vbInformation, "Empty Box"
    cmbage.SetFocus
    chk = "True"
    Exit Sub
End If

If Trim(cmbsex.Text) = "" Then
    MsgBox "Fill Gender Box Is Not Empty", vbInformation, "Empty Box"
    cmbsex.SetFocus
    chk = "True"
    Exit Sub
End If

If Trim(cmbreligion.Text) = "" Then
    MsgBox "Fill Religion Name Box Is Not Empty", vbInformation, "Empty Box"
    cmbreligion.SetFocus
    chk = "True"
    Exit Sub
End If

If Trim(cmbpqualification.Text) = "" Then
    MsgBox "Fill Previous Qualification Box Is Not Empty", vbInformation, "Empty Box"
    cmbpqualification.SetFocus
    chk = "True"
    Exit Sub
End If

If Trim(txtpinstitution.Text) = "" Then
    MsgBox "Fill Previous Institute Name Box Is Not Empty", vbInformation, "Empty Box"
    txtpinstitution.SetFocus
    chk = "True"
    Exit Sub
End If

If Trim(cmbclass.Text) = "" Then
    MsgBox "Fill Class Name Box Is Not Empty", vbInformation, "Empty Box"
    cmbclass.SetFocus
    chk = "True"
    Exit Sub
End If

If Trim(cmbshift.Text) = "" Then
    MsgBox "Fill Shift Box Is Not Empty", vbInformation, "Empty Box"
    cmbshift.SetFocus
    chk = "True"
    Exit Sub
End If

If Trim(cmbclasstime.Text) = "" Then
    MsgBox "Fill Class Time Box Is Not Empty", vbInformation, "Empty Box"
    cmbclasstime.SetFocus
    chk = "True"
    Exit Sub
End If

If Trim(cmbstatus.Text) = "" Then
    MsgBox "Fill Student Status Box Is Not Empty", vbInformation, "Empty Box"
    cmbstatus.SetFocus
    chk = "True"
    Exit Sub
End If

'If txtpict.Text = "" Then
'MsgBox "Fill Name Box Is Not Empty", vbInformation, "Empty Box"
'txtname.SetFocus
'Exit Sub
'End If

'If picc.Picture = LoadPicture("") Then
'MsgBox "Fill Name Box Is Not Empty", vbInformation, "Empty Box"
'txtname.SetFocus
'Exit Sub
'End If
End Sub


Private Sub txtname_GotFocus()
If Trim(txtgrno.Text) = "" Then
    MsgBox "Please Write First GrNo In Numeric Value", vbInformation, "G R N O"
    txtgrno.Text = ""
    txtgrno.SetFocus
    Exit Sub
ElseIf Not IsNumeric(txtgrno.Text) Then
    MsgBox "Please Write GrNo In Numeric Value", vbInformation, "G R N O"
    txtgrno.Text = ""
    txtgrno.SetFocus
    Exit Sub
Else
    Dim CurrentRec As Variant
    With rsSTUDENT
    If .RecordCount > 0 Then
    CurrentRec = .Bookmark
    .MoveFirst
    .Find "GRNO='" & txtgrno & "'"
        If .EOF Then
            txtname.SetFocus
            .Bookmark = CurrentRec
        Else
            MsgBox "This GrNo is Alredy Exits", vbCritical, "Duplicate GrNo"
            cmdadd.Enabled = True
            cmdUpdate.Enabled = False
            cmddelete.Enabled = True
            cmdCancel.Enabled = False
            cmdedit.Enabled = True
            cmdprint.Enabled = True
            cmdfind.Enabled = True
            cmdfirst.Enabled = True
            cmdpre.Enabled = True
            cmdnext.Enabled = True
            cmdlast.Enabled = True
            Call showdata
            txtgrno.SetFocus
            Exit Sub
        End If
    End If
    End With
End If
End Sub
