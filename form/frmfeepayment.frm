VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form frmfeepayment 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Fee Payment"
   ClientHeight    =   7455
   ClientLeft      =   2475
   ClientTop       =   2460
   ClientWidth     =   9255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   9255
   Begin MSDataGridLib.DataGrid dg 
      Height          =   1335
      Left            =   240
      TabIndex        =   35
      Top             =   1320
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   2355
      _Version        =   393216
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
      Left            =   6120
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   31
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox txtlevelcode 
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
      Left            =   6120
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   29
      Top             =   3840
      Width           =   1815
   End
   Begin VB.ComboBox cmbgrno 
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
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   3360
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
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   3840
      Width           =   1575
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
      Height          =   360
      Left            =   6120
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   20
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox txtreceiptno 
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
      MaxLength       =   15
      TabIndex        =   0
      Top             =   2842
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   240
      TabIndex        =   24
      Top             =   120
      Width           =   8895
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   3000
         Picture         =   "frmfeepayment.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   3720
         Picture         =   "frmfeepayment.frx":09EA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   4440
         Picture         =   "frmfeepayment.frx":13D4
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   5160
         Picture         =   "frmfeepayment.frx":1956
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   5880
         Picture         =   "frmfeepayment.frx":2340
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton cmdfind 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   6600
         Picture         =   "frmfeepayment.frx":2D2A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   8040
         Picture         =   "frmfeepayment.frx":396C
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton cmdprint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   7320
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmfeepayment.frx":4147
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton cmdfirst 
         BackColor       =   &H80000009&
         Caption         =   "First"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmfeepayment.frx":4D8A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton cmdpre 
         BackColor       =   &H80000009&
         Caption         =   "Previous"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   840
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmfeepayment.frx":59CD
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton cmdnext 
         BackColor       =   &H80000009&
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmfeepayment.frx":6610
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton cmdlast 
         BackColor       =   &H80000009&
         Caption         =   "Last"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   2280
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmfeepayment.frx":7253
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   700
      End
   End
   Begin VB.TextBox txtremarks 
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
      MaxLength       =   30
      TabIndex        =   7
      Top             =   6240
      Width           =   4575
   End
   Begin VB.TextBox txtcontectno 
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
      MaxLength       =   15
      TabIndex        =   6
      Top             =   5760
      Width           =   1575
   End
   Begin TDBDate6Ctl.TDBDate txtdate 
      Bindings        =   "frmfeepayment.frx":7E96
      Height          =   360
      Left            =   4560
      TabIndex        =   1
      Top             =   2835
      Width           =   3330
      _Version        =   65536
      _ExtentX        =   5874
      _ExtentY        =   635
      Calendar        =   "frmfeepayment.frx":7EA1
      Caption         =   "frmfeepayment.frx":7FB9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmfeepayment.frx":8030
      Keys            =   "frmfeepayment.frx":804E
      Spin            =   "frmfeepayment.frx":80AC
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   16777215
      BorderStyle     =   1
      BtnPositioning  =   1
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
   Begin TDBNumber6Ctl.TDBNumber txtmonthfee 
      Height          =   360
      Left            =   1800
      TabIndex        =   5
      Top             =   4800
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   635
      Calculator      =   "frmfeepayment.frx":80D4
      Caption         =   "frmfeepayment.frx":80F4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmfeepayment.frx":8150
      Keys            =   "frmfeepayment.frx":816E
      Spin            =   "frmfeepayment.frx":81B8
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   ""
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#,###,###,##0.00"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   -999999999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   32636929
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtadmissionfee 
      Height          =   360
      Left            =   1800
      TabIndex        =   4
      Top             =   4320
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   635
      Calculator      =   "frmfeepayment.frx":81E0
      Caption         =   "frmfeepayment.frx":8200
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmfeepayment.frx":825C
      Keys            =   "frmfeepayment.frx":827A
      Spin            =   "frmfeepayment.frx":82C4
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   ""
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#,###,###,##0.00"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   -999999999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   1
      ValueVT         =   180224001
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label lbltotalamount 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5460
      TabIndex        =   38
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Monthly Fee  :"
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
      TabIndex        =   37
      Top             =   4853
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Admission Fee :"
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
      Left            =   240
      TabIndex        =   36
      Top             =   4380
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount In Words :"
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
      TabIndex        =   34
      Top             =   6780
      Width           =   1575
   End
   Begin VB.Label lblwords 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      TabIndex        =   33
      Top             =   6720
      Width           =   6855
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Class Name :"
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
      Left            =   4560
      TabIndex        =   32
      Top             =   4380
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Level Code :"
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
      Left            =   4560
      TabIndex        =   30
      Top             =   3900
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Month Of :"
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
      TabIndex        =   28
      Top             =   3893
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Name :"
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
      Left            =   4560
      TabIndex        =   27
      Top             =   3413
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Receipt Amount "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   7
      Left            =   4800
      TabIndex        =   26
      Top             =   4920
      Width           =   3135
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt No :"
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
      TabIndex        =   25
      Top             =   2895
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      X1              =   240
      X2              =   9000
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks :"
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
      TabIndex        =   23
      Top             =   6300
      Width           =   975
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Contect  No :"
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
      TabIndex        =   22
      Top             =   5820
      Width           =   1335
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
      TabIndex        =   21
      Top             =   3413
      Width           =   735
   End
End
Attribute VB_Name = "frmfeepayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsfeespayment As New ADODB.Recordset
Dim rsfeepayment As New ADODB.Recordset
Dim CurrentRec As Variant
Dim rsSTUDENT As New ADODB.Recordset
Dim rsprint As New ADODB.Recordset


Private Sub cmbgrno_Click()
With rsfeespayment
If .State = closed Then
.Open
End If
End With
If rsSTUDENT.State = 1 Then
rsSTUDENT.Close
End If
rsSTUDENT.Open "Select * from student_details where GRNO='" & cmbgrno.Text & "'", cn, adOpenDynamic, adLockOptimistic
With rsSTUDENT
If .RecordCount <= 0 Then
txtname.Text = ""
txtlevelcode.Text = ""
txtdescr.Text = ""
Else
txtname.Text = UCase(.Fields("NAME"))
txtlevelcode.Text = .Fields("LEVELCODE")
txtdescr.Text = UCase(.Fields("DESCR"))
.Close
End If
End With

End Sub


Private Sub cmdAdd_Click()

cmdAdd.Enabled = False
cmdUpdate.Enabled = True
cmdDelete.Enabled = False
cmdCancel.Enabled = True
cmdEdit.Enabled = False
cmdprint.Enabled = False
cmdfirst.Enabled = False
cmdpre.Enabled = False
cmdnext.Enabled = False
cmdlast.Enabled = False


' all fields blank for new entery
cmbgrno.Text = ""
txtreceiptno.Text = ""
txtdate.Value = Date
cmbmonth.Text = ""
txtlevelcode.Text = ""
txtadmissionfee.Text = ""
txtmonthfee.Text = ""
txtcontectno.Text = ""
txtremarks.Text = ""
txtdescr.Text = ""
txtname.Text = ""
lblwords.Caption = ""
lbltotalamount.Caption = ""
dg.Refresh
txtreceiptno.SetFocus
End Sub

Private Sub cmdCancel_Click()

cmbgrno.Text = ""
txtreceiptno.Text = ""
txtdate.Value = Date
cmbmonth.Text = ""
txtlevelcode.Text = ""
txtadmissionfee.Text = ""
txtmonthfee.Text = ""
txtcontectno.Text = ""
txtremarks.Text = ""
txtdescr.Text = ""
txtname.Text = ""
lblwords.Caption = ""
lbltotalamount.Caption = ""

MsgBox "Record has been Cancel"

rsfeespayment.AbsolutePosition = 1
CurrentRec = rsfeespayment.AbsolutePosition
dg.Refresh
Call showdata
cmdCancel.Enabled = False
cmdUpdate.Enabled = False
cmdAdd.Enabled = True
cmdDelete.Enabled = True
cmdEdit.Enabled = True
cmdprint.Enabled = True
cmdfirst.Enabled = True
cmdpre.Enabled = True
cmdnext.Enabled = True
cmdlast.Enabled = True

txtreceiptno.SetFocus
End Sub


Private Sub cmdDelete_Click()

'With rsfeespayment
If MsgBox("Are you sure you want to delete this recorde", vbYesNo + vbCritical, "Warning") = vbYes Then
receipt = txtreceiptno.Text
If rsfeepayment.State = 1 Then rsfeepayment.Close
If rsfeepayment.State = 0 Then
rsfeepayment.Open "select * from feepayment where receiptno='" & receipt & "'", cn, adOpenStatic, adLockOptimistic
rsfeepayment.Delete
MsgBox "Record has been deleted"
rsfeepayment.Close
End If

End If

With rsfeespayment

rsfeespayment.AbsolutePosition = 1
CurrentRec = .AbsolutePosition
dg.Refresh
Call showdata
dg.Refresh
End With
cmdUpdate.Enabled = False
cmdAdd.Enabled = True
cmdDelete.Enabled = True
cmdEdit.Enabled = True
cmdCancel.Enabled = False
End Sub

Private Sub cmdEdit_Click()
With rsfeespayment
.Fields("RECEIPTNO") = txtreceiptno.Text
.Fields("DATE") = txtdate.Value
.Fields("FEEMONTH") = cmbmonth.Text
.Fields("GRNO") = cmbgrno.Text
.Fields("LEVELCODE") = txtlevelcode.Text
.Fields("ADMISSIONFEE") = Val(txtadmissionfee.Text)
.Fields("MONTHLYFEE") = Val(txtmonthfee.Text)
.Fields("AMOUNT") = Val(lbltotalamount.Caption)
.Fields("CONTECTNO") = txtcontectno.Text
.Fields("REMARKS") = UCase(txtremarks.Text)
.Fields("AMOUNTWORDS") = lblwords.Caption
.Fields("USERNAME") = pubusername
.Update
MsgBox "Record has been Edit successfully"
dg.Refresh
End With
End Sub

Private Sub cmdexit_Click()
Me.Hide
End Sub


Private Sub cmdfind_Click()
frmfindreceipt.Show
frmfindreceipt.txtfind.Text = ""
frmfindreceipt.txtfind1.Text = ""
frmfindreceipt.txtfind2.Text = ""
frmfindreceipt.txtfind2.SetFocus
End Sub

Private Sub cmdprint_Click()
With rsprint
    If .State = adStateOpen Then .Close
    .Open "SELECT * from feespayment where receiptno='" & txtreceiptno & "'", cn, adOpenForwardOnly, adLockReadOnly
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
End Sub


Private Sub cmdUpdate_Click()
Call chkdata(chk)
If chk = "False" Then
With rsfeespayment
.AddNew
.Fields("RECEIPTNO") = UCase(txtreceiptno.Text)
.Fields("DATE") = txtdate.Value
.Fields("FEEMONTH") = cmbmonth.Text
.Fields("GRNO") = cmbgrno.Text
.Fields("LEVELCODE") = txtlevelcode.Text
.Fields("ADMISSIONFEE") = Val(txtadmissionfee.Text)
.Fields("MONTHLYFEE") = Val(txtmonthfee.Text)
.Fields("AMOUNT") = Val(lbltotalamount.Caption)
.Fields("CONTECTNO") = txtcontectno.Text
.Fields("REMARKS") = UCase(txtremarks.Text)
.Fields("AMOUNTWORDS") = lblwords.Caption
.Fields("USERNAME") = pubusername
.Update
MsgBox "Record has been Saved successfully"
dg.Refresh
End With
cmdUpdate.Enabled = False
cmdAdd.Enabled = True
cmdDelete.Enabled = True
cmdEdit.Enabled = True
cmdCancel.Enabled = False
cmdfirst.Enabled = True
cmdpre.Enabled = True
cmdnext.Enabled = True
cmdlast.Enabled = True
Else
cmdUpdate.Enabled = True
cmdAdd.Enabled = False
cmdDelete.Enabled = False
cmdEdit.Enabled = False
cmdCancel.Enabled = True
End If
End Sub


Private Sub Form_Load()
Call conopen

If rsfeespayment.State = 0 Then
rsfeespayment.Open "select * from feespayment", cn, adOpenStatic, adLockOptimistic
End If
If rsSTUDENT.State = 0 Then
rsSTUDENT.Open "Select * from student_details", cn, adOpenDynamic, adLockOptimistic
End If
Set dg.DataSource = rsfeespayment

cmdAdd.Enabled = True
cmdUpdate.Enabled = False
cmdCancel.Enabled = False
cmdDelete.Enabled = True
cmdEdit.Enabled = True
cmdfirst.Enabled = True
cmdpre.Enabled = True
cmdnext.Enabled = True
cmdlast.Enabled = True


With rsfeespayment
If .State = closed Then
.Open
End If
If .RecordCount > 0 Then
txtreceiptno.Text = .Fields("RECEIPTNO")
txtdate.Value = .Fields("DATE")
cmbmonth.Text = .Fields("FEEMONTH")
cmbgrno.Text = .Fields("GRNO")
txtlevelcode.Text = .Fields("LEVELCODE")
txtadmissionfee.Text = Format(.Fields("ADMISSIONFEE"), "###,###,##0.00")
txtmonthfee.Text = Format(.Fields("MONTHLYFEE"), "###,###,##0.00")
lbltotalamount.Caption = Format(.Fields("AMOUNT"), "###,###,##0.00")
txtcontectno.Text = .Fields("CONTECTNO")
txtremarks.Text = .Fields("REMARKS")
txtname.Text = .Fields("NAME")
txtdescr.Text = .Fields("DESCR")
lblwords.Caption = .Fields("AMOUNTWORDS")
End If
End With

'fill combo box grno
With rsSTUDENT
If .State = closed Then
.Open
End If
.AbsolutePosition = 1
While .EOF = False
cmbgrno.AddItem .Fields("GRNO")
.MoveNext
Wend
End With

'fill combo box month

cmbmonth.AddItem "January", 0
cmbmonth.AddItem "February", 1
cmbmonth.AddItem "March", 2
cmbmonth.AddItem "April", 3
cmbmonth.AddItem "May", 4
cmbmonth.AddItem "June", 5
cmbmonth.AddItem "July", 6
cmbmonth.AddItem "August", 7
cmbmonth.AddItem "September", 8
cmbmonth.AddItem "October", 9
cmbmonth.AddItem "November", 10
cmbmonth.AddItem "December", 11
End Sub

Private Sub cmdfirst_Click()
With rsfeespayment
If .RecordCount > 0 Then
rsfeespayment.MoveFirst
CurrentRec = .AbsolutePosition
Call showdata
End If
End With
End Sub

Private Sub cmdlast_Click()
With rsfeespayment
If .RecordCount > 0 Then
rsfeespayment.MoveLast
CurrentRec = .AbsolutePosition
If .EOF Then
MsgBox "This is the Last Recorods"
.MoveLast
CurrentRec = .AbsolutePosition
End If
Call showdata
End If
End With
End Sub

Private Sub cmdnext_Click()
With rsfeespayment
If .RecordCount > 0 Then
rsfeespayment.MoveNext
CurrentRec = .AbsolutePosition
If .EOF Then
MsgBox "This is the Last Recorods"
.MoveLast
CurrentRec = .AbsolutePosition
End If
Call showdata
End If
End With
End Sub
Private Sub cmdpre_Click()
With rsfeespayment
If .RecordCount > 0 Then
rsfeespayment.MovePrevious
CurrentRec = .AbsolutePosition
If .BOF Then
MsgBox "This is the First Recorods"
.MoveFirst
CurrentRec = .AbsolutePosition
End If
Call showdata
End If
End With
End Sub

Public Sub showdata()
Set dg.DataSource = Nothing
With rsfeespayment
If .State = 1 Then
.Close
End If
rsfeespayment.Open "select * from feespayment", cn, adOpenDynamic, adLockOptimistic
Set dg.DataSource = rsfeespayment
.AbsolutePosition = CurrentRec
If .RecordCount > 0 Then
txtreceiptno.Text = .Fields("RECEIPTNO")
txtdate.Value = .Fields("DATE")
cmbmonth.Text = .Fields("FEEMONTH")
cmbgrno.Text = .Fields("GRNO")
txtlevelcode.Text = .Fields("LEVELCODE")
txtadmissionfee.Text = Format(.Fields("ADMISSIONFEE"), "###,###,##0.00")
txtmonthfee.Text = Format(.Fields("MONTHLYFEE"), "###,###,##0.00")
lbltotalamount.Caption = Format(.Fields("AMOUNT"), "###,###,##0.00")
txtcontectno.Text = .Fields("CONTECTNO")
txtremarks.Text = .Fields("REMARKS")
txtname.Text = .Fields("NAME")
txtdescr.Text = .Fields("DESCR")
lblwords.Caption = cNumToWord(Format(lbltotalamount.Caption, "###,###,##0.00"))
dg.Refresh
End If
End With
CurrentRec = 0
End Sub

Private Sub lbltotalamount_Change()
lblwords.Caption = cNumToWord(Format(lbltotalamount.Caption, "###,###,##0.00"))
End Sub


'Private Sub txtmonthfee_Change()
'lblwords.Caption = cNumToWord(txtmonthfee.Text)
'End Sub


Private Sub txtdate_GotFocus()
If Trim(txtreceiptno.Text) = "" Then
    MsgBox "Please Write First Receipt No", vbInformation, "Receipt No"
    txtreceiptno.Text = ""
    txtreceiptno.SetFocus
    Exit Sub
Else
    Dim CurrentRec As Variant
    With rsfeespayment
    If .RecordCount > 0 Then
    CurrentRec = .Bookmark
    .MoveFirst
    .Find "RECEIPTNO='" & txtreceiptno.Text & "'"
        If .EOF Then
            txtdate.SetFocus
            .Bookmark = CurrentRec
        Else
            MsgBox "This Receipt No is Alredy Exits", vbCritical, "Duplicate Receipt No"
            txtreceiptno.SetFocus
        End If
    End If
        End With
End If
End Sub
Public Sub chkdata(chk)
chk = "False"

If Trim(txtreceiptno.Text) = "" Then
MsgBox "Receipt No Is Empty", vbInformation, "Receipt No"
txtreceiptno.SetFocus
chk = "True"
Exit Sub

ElseIf Trim(txtdate.Value) = "" Then
MsgBox "Date Is Empty", vbInformation, "Date"
txtdate.SetFocus
chk = "True"
Exit Sub

ElseIf Trim(cmbmonth.Text) = "" Then
MsgBox "Month Is Empty", vbInformation, "Month"
cmbmonth.SetFocus
chk = "True"
Exit Sub

ElseIf Trim(cmbgrno.Text) = "" Then
MsgBox "GR NO Is Empty", vbInformation, "GR NO"
cmbgrno.SetFocus
chk = "True"
Exit Sub

ElseIf Trim(txtmonthfee.Text) = "" Then
MsgBox "Amount Is Empty", vbInformation, "Amount"
txtmonthfee.SetFocus
chk = "True"
Exit Sub
End If

End Sub

Private Sub txtmonthfee_LostFocus()
lbltotalamount.Caption = Format(Val(txtmonthfee.Text) + Val(txtadmissionfee.Text), "###,###,##0.00")
End Sub
