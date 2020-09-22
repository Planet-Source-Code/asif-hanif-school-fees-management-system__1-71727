VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Begin VB.Form frmsubjects 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Class & Subjects"
   ClientHeight    =   7725
   ClientLeft      =   1320
   ClientTop       =   2100
   ClientWidth     =   10185
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   10185
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Charges"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4440
      TabIndex        =   82
      Top             =   1320
      Width           =   5655
      Begin VB.TextBox txttotal 
         Alignment       =   1  'Right Justify
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
         Left            =   4320
         LinkTimeout     =   70
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   91
         Top             =   1320
         Width           =   1215
      End
      Begin TDBNumber6Ctl.TDBNumber txtaddmission 
         Height          =   360
         Left            =   1680
         TabIndex        =   38
         Top             =   240
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   635
         Calculator      =   "frmsubjects.frx":0000
         Caption         =   "frmsubjects.frx":0020
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmsubjects.frx":007C
         Keys            =   "frmsubjects.frx":009A
         Spin            =   "frmsubjects.frx":00E4
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
         ValueVT         =   62914561
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtmonthly 
         Height          =   360
         Left            =   1680
         TabIndex        =   39
         Top             =   600
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   635
         Calculator      =   "frmsubjects.frx":010C
         Caption         =   "frmsubjects.frx":012C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmsubjects.frx":0188
         Keys            =   "frmsubjects.frx":01A6
         Spin            =   "frmsubjects.frx":01F0
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
         ValueVT         =   62914561
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txttution 
         Height          =   360
         Left            =   1680
         TabIndex        =   40
         Top             =   960
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   635
         Calculator      =   "frmsubjects.frx":0218
         Caption         =   "frmsubjects.frx":0238
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmsubjects.frx":0294
         Keys            =   "frmsubjects.frx":02B2
         Spin            =   "frmsubjects.frx":02FC
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
         ValueVT         =   62914561
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtexamination 
         Height          =   360
         Left            =   1680
         TabIndex        =   41
         Top             =   1320
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   635
         Calculator      =   "frmsubjects.frx":0324
         Caption         =   "frmsubjects.frx":0344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmsubjects.frx":03A0
         Keys            =   "frmsubjects.frx":03BE
         Spin            =   "frmsubjects.frx":0408
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
         ValueVT         =   62914561
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtcomputer 
         Height          =   360
         Left            =   4320
         TabIndex        =   42
         Top             =   240
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   635
         Calculator      =   "frmsubjects.frx":0430
         Caption         =   "frmsubjects.frx":0450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmsubjects.frx":04AC
         Keys            =   "frmsubjects.frx":04CA
         Spin            =   "frmsubjects.frx":0514
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
         ValueVT         =   62914561
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtlab 
         Height          =   360
         Left            =   4320
         TabIndex        =   43
         Top             =   600
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   635
         Calculator      =   "frmsubjects.frx":053C
         Caption         =   "frmsubjects.frx":055C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmsubjects.frx":05B8
         Keys            =   "frmsubjects.frx":05D6
         Spin            =   "frmsubjects.frx":0620
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
         ValueVT         =   62914561
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtother 
         Height          =   360
         Left            =   4320
         TabIndex        =   44
         Top             =   960
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   635
         Calculator      =   "frmsubjects.frx":0648
         Caption         =   "frmsubjects.frx":0668
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmsubjects.frx":06C4
         Keys            =   "frmsubjects.frx":06E2
         Spin            =   "frmsubjects.frx":072C
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
         ValueVT         =   62914561
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Fees :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   90
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Others Fees :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   89
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Lab. Fees :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   88
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Computer Fees :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   87
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Examination Fees :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   1373
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tution Fees :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   1013
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Monthly  Fees :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   653
         Width           =   1575
      End
      Begin VB.Label lbladdmition 
         BackStyle       =   0  'Transparent
         Caption         =   "Addmission Fees :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   293
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Subject's"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   240
      TabIndex        =   63
      Top             =   3360
      Width           =   9855
      Begin VB.TextBox txtsubject 
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
         Index           =   11
         Left            =   6000
         MaxLength       =   30
         TabIndex        =   35
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox txtmax 
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
         Index           =   11
         Left            =   8160
         MaxLength       =   30
         TabIndex        =   36
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtmin 
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
         Index           =   11
         Left            =   9000
         MaxLength       =   30
         TabIndex        =   37
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtsubject 
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
         Index           =   10
         Left            =   6000
         MaxLength       =   30
         TabIndex        =   32
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox txtmax 
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
         Index           =   10
         Left            =   8160
         MaxLength       =   30
         TabIndex        =   33
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtmin 
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
         Index           =   10
         Left            =   9000
         MaxLength       =   30
         TabIndex        =   34
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtsubject 
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
         Index           =   9
         Left            =   6000
         MaxLength       =   30
         TabIndex        =   29
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtmax 
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
         Index           =   9
         Left            =   8160
         MaxLength       =   30
         TabIndex        =   30
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtmin 
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
         Index           =   9
         Left            =   9000
         MaxLength       =   30
         TabIndex        =   31
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtsubject 
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
         Index           =   8
         Left            =   6000
         MaxLength       =   30
         TabIndex        =   26
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtmax 
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
         Index           =   8
         Left            =   8160
         MaxLength       =   30
         TabIndex        =   27
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtmin 
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
         Index           =   8
         Left            =   9000
         MaxLength       =   30
         TabIndex        =   28
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtsubject 
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
         Index           =   7
         Left            =   6000
         MaxLength       =   30
         TabIndex        =   23
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtmax 
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
         Index           =   7
         Left            =   8160
         MaxLength       =   30
         TabIndex        =   24
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtmin 
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
         Index           =   7
         Left            =   9000
         MaxLength       =   30
         TabIndex        =   25
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtsubject 
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
         Index           =   6
         Left            =   6000
         MaxLength       =   30
         TabIndex        =   20
         Top             =   585
         Width           =   1935
      End
      Begin VB.TextBox txtmax 
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
         Index           =   6
         Left            =   8160
         MaxLength       =   30
         TabIndex        =   21
         Top             =   585
         Width           =   735
      End
      Begin VB.TextBox txtmin 
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
         Index           =   6
         Left            =   9000
         MaxLength       =   30
         TabIndex        =   22
         Top             =   585
         Width           =   735
      End
      Begin VB.TextBox txtsubject 
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
         Index           =   5
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   17
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox txtmax 
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
         Index           =   5
         Left            =   3240
         MaxLength       =   30
         TabIndex        =   18
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtmin 
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
         Index           =   5
         Left            =   4080
         MaxLength       =   30
         TabIndex        =   19
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtsubject 
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
         Index           =   4
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   14
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox txtmax 
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
         Index           =   4
         Left            =   3240
         MaxLength       =   30
         TabIndex        =   15
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtmin 
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
         Index           =   4
         Left            =   4080
         MaxLength       =   30
         TabIndex        =   16
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtsubject 
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
         Index           =   3
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   11
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtmax 
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
         Index           =   3
         Left            =   3240
         MaxLength       =   30
         TabIndex        =   12
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtmin 
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
         Index           =   3
         Left            =   4080
         MaxLength       =   30
         TabIndex        =   13
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtsubject 
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
         Index           =   2
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   8
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtmax 
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
         Index           =   2
         Left            =   3240
         MaxLength       =   30
         TabIndex        =   9
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtmin 
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
         Index           =   2
         Left            =   4080
         MaxLength       =   30
         TabIndex        =   10
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtsubject 
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
         Index           =   1
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtmax 
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
         Index           =   1
         Left            =   3240
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtmin 
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
         Index           =   1
         Left            =   4080
         MaxLength       =   30
         TabIndex        =   7
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtsubject 
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
         Index           =   0
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   2
         Top             =   585
         Width           =   1935
      End
      Begin VB.TextBox txtmax 
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
         Index           =   0
         Left            =   3240
         MaxLength       =   30
         TabIndex        =   3
         Top             =   585
         Width           =   735
      End
      Begin VB.TextBox txtmin 
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
         Index           =   0
         Left            =   4080
         MaxLength       =   30
         TabIndex        =   4
         Top             =   585
         Width           =   735
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
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
         Left            =   9240
         TabIndex        =   79
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
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
         Left            =   8400
         TabIndex        =   78
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject :"
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
         Left            =   5040
         TabIndex        =   77
         Top             =   3060
         Width           =   855
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject :"
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
         Left            =   5040
         TabIndex        =   76
         Top             =   2580
         Width           =   855
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject :"
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
         Left            =   5040
         TabIndex        =   75
         Top             =   2100
         Width           =   855
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject :"
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
         Left            =   5040
         TabIndex        =   74
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject :"
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
         Left            =   5040
         TabIndex        =   73
         Top             =   1133
         Width           =   855
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject :"
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
         Left            =   5040
         TabIndex        =   72
         Top             =   638
         Width           =   855
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject :"
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
         Left            =   120
         TabIndex        =   71
         Top             =   3060
         Width           =   855
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject :"
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
         Left            =   120
         TabIndex        =   70
         Top             =   2580
         Width           =   855
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject :"
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
         Left            =   120
         TabIndex        =   69
         Top             =   2100
         Width           =   855
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject :"
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
         Left            =   120
         TabIndex        =   68
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject :"
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
         Left            =   120
         TabIndex        =   67
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject :"
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
         Left            =   120
         TabIndex        =   66
         Top             =   645
         Width           =   855
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
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
         Left            =   3420
         TabIndex        =   65
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
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
         Left            =   4260
         TabIndex        =   64
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   3720
      TabIndex        =   62
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
         Picture         =   "frmsubjects.frx":0754
         Style           =   1  'Graphical
         TabIndex        =   51
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
         Picture         =   "frmsubjects.frx":1397
         Style           =   1  'Graphical
         TabIndex        =   52
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
         Picture         =   "frmsubjects.frx":1FD9
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   240
      TabIndex        =   61
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
         Picture         =   "frmsubjects.frx":27B4
         Style           =   1  'Graphical
         TabIndex        =   53
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
         Picture         =   "frmsubjects.frx":33F7
         Style           =   1  'Graphical
         TabIndex        =   54
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
         Picture         =   "frmsubjects.frx":403A
         Style           =   1  'Graphical
         TabIndex        =   55
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
         Picture         =   "frmsubjects.frx":4C7D
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   120
         Width           =   735
      End
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
      Left            =   1800
      LinkTimeout     =   70
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1402
      Width           =   1455
   End
   Begin VB.TextBox txtdescr 
      DataMember      =   "SUBJECTS"
      DataSource      =   "de"
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
      TabIndex        =   1
      Top             =   1845
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   6240
      TabIndex        =   59
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton cmdedit 
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
         Picture         =   "frmsubjects.frx":58C0
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   2760
         Width           =   1335
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
         Left            =   1560
         Picture         =   "frmsubjects.frx":62AA
         Style           =   1  'Graphical
         TabIndex        =   48
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
         Left            =   840
         Picture         =   "frmsubjects.frx":6C94
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save"
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
         Picture         =   "frmsubjects.frx":7216
         Style           =   1  'Graphical
         TabIndex        =   45
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
         Picture         =   "frmsubjects.frx":7C00
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Label lbllnoofstudent 
      BackColor       =   &H00FFFFFF&
      Caption         =   "No Of Student :"
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
      Left            =   240
      TabIndex        =   81
      Top             =   2460
      Width           =   1455
   End
   Begin VB.Label lblnoofstudent 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   80
      Top             =   2400
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      X1              =   240
      X2              =   10080
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      X1              =   240
      X2              =   10080
      Y1              =   1200
      Y2              =   1200
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
      Index           =   1
      Left            =   240
      TabIndex        =   60
      Top             =   1905
      Width           =   1215
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
      Index           =   0
      Left            =   240
      TabIndex        =   58
      Top             =   1455
      Width           =   1215
   End
End
Attribute VB_Name = "frmsubjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim countstudent As New ADODB.Recordset
Dim rssubjects As New ADODB.Recordset
Dim CurrentRec As Variant
Dim rsprint As New ADODB.Recordset



Private Sub cmdAdd_Click()
cmdadd.Enabled = False
cmdCancel.Enabled = True
cmddelete.Enabled = False
cmdUpdate.Enabled = True
cmdEdit.Enabled = False

' all fields blank for new entery
 txtlevelcode.Text = ""
 txtdescr.Text = ""
' all subject blank for new entery
 i = 0
 For i = 0 To 11
 txtsubject(i).Text = ""
 Next i
 ' all maximum marks blank for new entery
 i = 0
 For i = 0 To 11
 txtmax(i).Text = ""
 Next i
 ' all minimum marks blank for new entery
 i = 0
 For i = 0 To 11
 txtmin(i).Text = ""
 Next i
' all carges column blank for new entery
txtaddmission.Text = ""
txtmonthly.Text = ""
txttution.Text = ""
txtexamination.Text = ""
txtcomputer.Text = ""
txtlab.Text = ""
txtother.Text = ""
txttotal.Text = ""

txtlevelcode.SetFocus
End Sub

Private Sub cmdCancel_Click()
cmdadd.Enabled = True
cmdCancel.Enabled = False
cmddelete.Enabled = True
cmdUpdate.Enabled = False
cmdEdit.Enabled = True

'empty subject box

a = 0
For a = 0 To 11
txtsubject(a).Text = ""
Next a

'empty maximum marks box

b = 0
For b = 0 To 11
txtmax(b).Text = ""
Next b

'empty minimum marks box

c = 0
For c = 0 To 11
txtmin(c).Text = ""
Next c

' all carges column blank for new entery
txtaddmission.Text = ""
txtmonthly.Text = ""
txttution.Text = ""
txtexamination.Text = ""
txtcomputer.Text = ""
txtlab.Text = ""
txtother.Text = ""
txttotal.Text = ""

MsgBox "Record has been Cancel Update"
If rssubjects.RecordCount > 0 Then
rssubjects.AbsolutePosition = 1
Call showdata
End If
txtlevelcode.SetFocus
End Sub


Private Sub cmdDelete_Click()
cmdadd.Enabled = True
cmdCancel.Enabled = False
cmddelete.Enabled = True
cmdUpdate.Enabled = False
cmdEdit.Enabled = True
With rssubjects
If .RecordCount > 0 Then
If MsgBox("Are you sure you want to delete this recorde", vbYesNo + vbCritical, "Warning") = vbYes Then
rssubjects.Delete
MsgBox "Record has been deleted"
End If
rssubjects.AbsolutePosition = 1
Call showdata
End If
End With
End Sub

Private Sub cmdEdit_Click()
cmdadd.Enabled = True
cmdCancel.Enabled = False
cmddelete.Enabled = True
cmdUpdate.Enabled = False
cmdEdit.Enabled = True
With rssubjects
If .RecordCount > 0 Then
Call subject
.Update
MsgBox "Record has been Edit successfully"
End If
End With
End Sub

Private Sub cmdexit_Click()
Me.Hide
End Sub


Private Sub cmdfind_Click()
frmfindlevel.Show
frmfindlevel.txtfind = ""
frmfindlevel.txtfind1 = ""
frmfindlevel.txtfind.SetFocus
End Sub

Private Sub cmdprint_Click()
With rsprint
    If .State = adStateOpen Then .Close
'    .Open "SELECT * from student_details where GRNO='" & txtgrno & "'", cn, adOpenForwardOnly, adLockReadOnly
.Open "select * from subjects where levelcode='" & txtlevelcode & "'", cn, adOpenForwardOnly, adLockOptimistic
    If .RecordCount = 0 Then
        MsgBox "No Records were found", vbInformation, "Report"
        Exit Sub
    End If
    Set rptclass.DataSource = rsprint
    rptclass.Sections("section4").Controls("lblschoolname").Caption = schoolname
    rptclass.Sections("section4").Controls("lbladdress1").Caption = address1
    rptclass.Sections("section4").Controls("lbladdress2").Caption = address2
    rptclass.Sections("section4").Controls("lblusername").Caption = pubusername
    rptclass.Sections("section4").Controls("lbldate").Caption = Now
    rptclass.Sections("section1").Controls("txtnostudent").Caption = lblnoofstudent.Caption
    rptclass.Sections("section1").Controls("lblnostudent").Caption = lbllnoofstudent.Caption
    
    rptclass.Show vbModal
End With

End Sub

Private Sub cmdUpdate_Click()
Call chkdata(chk)
If chk = "True" Then
cmdadd.Enabled = False
cmdCancel.Enabled = True
cmddelete.Enabled = False
cmdUpdate.Enabled = True
cmdEdit.Enabled = False
Exit Sub
End If
rssubjects.AddNew
With rssubjects
Call subject
.Update
MsgBox "Record has been Saved successfully"
End With
cmdadd.Enabled = True
cmdCancel.Enabled = False
cmddelete.Enabled = True
cmdUpdate.Enabled = False
cmdEdit.Enabled = True
End Sub


Private Sub Form_Load()
Call conopen
If rssubjects.State = closed Then
rssubjects.Open "select * from subjects order by levelcode", cn, adOpenDynamic, adLockOptimistic
End If
cmdadd.Enabled = True
cmdCancel.Enabled = False
cmddelete.Enabled = True
cmdUpdate.Enabled = False
cmdEdit.Enabled = True

With rssubjects
If .RecordCount > 0 Then
Call showdata
End If
End With
End Sub

Private Sub cmdfirst_Click()
With rssubjects
If .RecordCount > 0 Then
rssubjects.MoveFirst
CurrentRec = .AbsolutePosition
Call showdata
End If
End With
End Sub

Private Sub cmdlast_Click()
With rssubjects
If .RecordCount > 0 Then
.MoveLast
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
With rssubjects
If .RecordCount > 0 Then
.MoveNext
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
With rssubjects
If .RecordCount > 0 Then
.MovePrevious
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

Public Sub subject()
With rssubjects
.Fields("LEVELCODE") = txtlevelcode.Text
.Fields("DESCR") = UCase(txtdescr.Text)

' subjects

.Fields("SUB01") = txtsubject(0).Text
.Fields("SUB02") = txtsubject(1).Text
.Fields("SUB03") = txtsubject(2).Text
.Fields("SUB04") = txtsubject(3).Text
.Fields("SUB05") = txtsubject(4).Text
.Fields("SUB06") = txtsubject(5).Text
.Fields("SUB07") = txtsubject(6).Text
.Fields("SUB08") = txtsubject(7).Text
.Fields("SUB09") = txtsubject(8).Text
.Fields("SUB10") = txtsubject(9).Text
.Fields("SUB11") = txtsubject(10).Text
.Fields("SUB12") = txtsubject(11).Text

' maximum marks

.Fields("MAX01") = Val(txtmax(0).Text)
.Fields("MAX02") = Val(txtmax(1).Text)
.Fields("MAX03") = Val(txtmax(2).Text)
.Fields("MAX04") = Val(txtmax(3).Text)
.Fields("MAX05") = Val(txtmax(4).Text)
.Fields("MAX06") = Val(txtmax(5).Text)
.Fields("MAX07") = Val(txtmax(6).Text)
.Fields("MAX08") = Val(txtmax(7).Text)
.Fields("MAX09") = Val(txtmax(8).Text)
.Fields("MAX10") = Val(txtmax(9).Text)
.Fields("MAX11") = Val(txtmax(10).Text)
.Fields("MAX12") = Val(txtmax(11).Text)

'minimum marks

.Fields("MIN01") = Val(txtmin(0).Text)
.Fields("MIN02") = Val(txtmin(1).Text)
.Fields("MIN03") = Val(txtmin(2).Text)
.Fields("MIN04") = Val(txtmin(3).Text)
.Fields("MIN05") = Val(txtmin(4).Text)
.Fields("MIN06") = Val(txtmin(5).Text)
.Fields("MIN07") = Val(txtmin(6).Text)
.Fields("MIN08") = Val(txtmin(7).Text)
.Fields("MIN09") = Val(txtmin(8).Text)
.Fields("MIN10") = Val(txtmin(9).Text)
.Fields("MIN11") = Val(txtmin(10).Text)
.Fields("MIN12") = Val(txtmin(11).Text)

' fee charges

.Fields("addmissionfee") = Val(txtaddmission.Text)
.Fields("monthlyfee") = Val(txtmonthly.Text)
.Fields("tutionfee") = Val(txttution.Text)
.Fields("examinationfee") = Val(txtexamination.Text)
.Fields("computerfee") = Val(txtcomputer.Text)
.Fields("labfee") = Val(txtlab.Text)
.Fields("otherfee") = Val(txtother.Text)
.Fields("totalfee") = txttotal.Text

End With
End Sub
Public Sub showdata()
With rssubjects
If .RecordCount > 0 Then
txtlevelcode.Text = .Fields("LEVELCODE")
txtdescr.Text = .Fields("DESCR")

' subjects

 txtsubject(0).Text = .Fields("SUB01")
 txtsubject(1).Text = .Fields("SUB02")
 txtsubject(2).Text = .Fields("SUB03")
 txtsubject(3).Text = .Fields("SUB04")
 txtsubject(4).Text = .Fields("SUB05")
 txtsubject(5).Text = .Fields("SUB06")
 txtsubject(6).Text = .Fields("SUB07")
 txtsubject(7).Text = .Fields("SUB08")
 txtsubject(8).Text = .Fields("SUB09")
 txtsubject(9).Text = .Fields("SUB10")
 txtsubject(10).Text = .Fields("SUB11")
 txtsubject(11).Text = .Fields("SUB12")

' maximum marks

txtmax(0).Text = .Fields("MAX01")
txtmax(1).Text = .Fields("MAX02")
txtmax(2).Text = .Fields("MAX03")
txtmax(3).Text = .Fields("MAX04")
txtmax(4).Text = .Fields("MAX05")
txtmax(5).Text = .Fields("MAX06")
txtmax(6).Text = .Fields("MAX07")
txtmax(7).Text = .Fields("MAX08")
txtmax(8).Text = .Fields("MAX09")
txtmax(9).Text = .Fields("MAX10")
txtmax(10).Text = .Fields("MAX11")
txtmax(11).Text = .Fields("MAX12")

'minimum marks

txtmin(0).Text = .Fields("MIN01")
txtmin(1).Text = .Fields("MIN02")
txtmin(2).Text = .Fields("MIN03")
txtmin(3).Text = .Fields("MIN04")
txtmin(4).Text = .Fields("MIN05")
txtmin(5).Text = .Fields("MIN06")
txtmin(6).Text = .Fields("MIN07")
txtmin(7).Text = .Fields("MIN08")
txtmin(8).Text = .Fields("MIN09")
txtmin(9).Text = .Fields("MIN10")
txtmin(10).Text = .Fields("MIN11")
txtmin(11).Text = .Fields("MIN12")

' charges fee

txtaddmission.Text = Format(.Fields("addmissionfee"), "###,###0.00")
txtmonthly.Text = Format(.Fields("monthlyfee"), "###,###0.00")
txttution.Text = Format(.Fields("tutionfee"), "###,###0.00")
txtexamination.Text = Format(.Fields("examinationfee"), "###,###0.00")
txtcomputer.Text = Format(.Fields("computerfee"), "###,###0.00")
txtlab.Text = Format(.Fields("labfee"), "###,###0.00")
txtother.Text = Format(.Fields("otherfee"), "###,###0.00")
txttotal.Text = Format(.Fields("totalfee"), "###,###0.00")
End If
End With
CurrentRec = 0
End Sub

Public Sub chargessum()
txttotal.Text = Format(Round(Val(txtaddmission.Text), 0) + Round(Val(txtmonthly.Text), 0) + Round(Val(txttution.Text), 0) + Round(Val(txtexamination.Text), 0) + Round(Val(txtcomputer.Text), 0) + Round(Val(txtlab.Text), 0) + Round(Val(txtother.Text), 0), "###,###,##0.00")
End Sub
Private Sub txtaddmission_LostFocus()
Call chargessum
End Sub

Private Sub txtdescr_GotFocus()
If Trim(txtlevelcode.Text) = "" Then
    MsgBox "Please Write First Level Code", vbInformation, "Level Code"
    txtlevelcode.Text = ""
    txtlevelcode.SetFocus
    Exit Sub
Else
    Dim CurrentRec As Variant
    With rssubjects
    If .RecordCount > 0 Then
    CurrentRec = .Bookmark
    .MoveFirst
    .Find "levelcode='" & txtlevelcode.Text & "'"
        If .EOF Then
            txtdescr.SetFocus
            .Bookmark = CurrentRec
        Else
            MsgBox "This Level Code is Alredy Exits", vbCritical, "Duplicate Level Code"
            txtlevelcode.SetFocus
        End If
    End If
        End With
End If
End Sub

Private Sub txtlevelcode_Change()
With countstudent
If .State = 1 Then
.Close
End If
countstudent.Open "Select * from countstu where levelcode='" & txtlevelcode.Text & "'", cn, adOpenDynamic, adLockOptimistic
On Error Resume Next
If .RecordCount <= 0 Then
lblnoofstudent.Caption = 0
Else
lblnoofstudent.Caption = countstudent.Fields("code")
End If
countstudent.Close
End With
End Sub
Private Sub txtmonthly_LostFocus()
Call chargessum
End Sub
Private Sub txttution_LostFocus()
Call chargessum
End Sub
Private Sub txtexamination_LostFocus()
Call chargessum
End Sub
Private Sub txtcomputer_LostFocus()
Call chargessum
End Sub
Private Sub txtlab_LostFocus()
Call chargessum
End Sub
Private Sub txtother_LostFocus()
Call chargessum
End Sub
Public Sub chkdata(chk)

chk = "False"

If Trim(txtlevelcode.Text) = "" Then
    MsgBox "Fill Level Code", vbInformation, "Empty Box"
    txtlevelcode.SetFocus
    chk = "True"
    Exit Sub
End If

If Trim(txtdescr.Text) = "" Then
    MsgBox "Fill Level / Class Name", vbInformation, "Empty Box"
    txtdescr.SetFocus
    chk = "True"
    Exit Sub
End If

i = 0
For i = 0 To 11
If Trim(txtsubject(i).Text) = "" Then
    MsgBox "Fill Subject this Box is Empty", vbInformation, "Empty Box"
    txtsubject(i).Text = ""
    txtsubject(i).SetFocus
    chk = "True"
    Exit Sub
End If
Next i

i = 0
For i = 0 To 11
If Trim(txtmax(i).Text) = "" Then
    MsgBox "Fill Maximum Number this Box is Empty", vbInformation, "Empty Box"
    txtmax(i).Text = ""
    txtmax(i).SetFocus
    chk = "True"
    Exit Sub
ElseIf IsNumeric(txtmax(i).Text) = False Then
    MsgBox "Please Type any numeric value"
    txtmax(i).Text = ""
    txtmax(i).SetFocus
    chk = "True"
    Exit Sub
End If

Next i

i = 0
For i = 0 To 11
If Trim(txtmin(i).Text) = "" Then
    MsgBox "Fill Minimum Number this Box is Empty", vbInformation, "Empty Box"
    txtmin(i).Text = ""
    txtmin(i).SetFocus
    chk = "True"
    Exit Sub
ElseIf IsNumeric(txtmin(i).Text) = False Then
    MsgBox "Please Type any numeric value"
    txtmin(i).Text = ""
    txtmin(i).SetFocus
    chk = "True"
    Exit Sub
End If
Next i


If Trim(txtaddmission.Text) = "" Then
    MsgBox "Fill Admission Fee", vbInformation, "Empty Box"
    txtaddmission.SetFocus
    chk = "True"
    Exit Sub
End If

If Trim(txtmonthly.Text) = "" Then
    MsgBox "Fill Monthly Fee", vbInformation, "Empty Box"
    txtmonthly.SetFocus
    chk = "True"
    Exit Sub
End If


If Trim(txttution.Text) = "" Then
    MsgBox "Fill Tution Fee", vbInformation, "Empty Box"
    txttution.SetFocus
    chk = "True"
    Exit Sub
End If

If Trim(txtexamination.Text) = "" Then
    MsgBox "Fill Examination Fee", vbInformation, "Empty Box"
    txtexamination.SetFocus
    chk = "True"
    Exit Sub
End If

If Trim(txtcomputer.Text) = "" Then
    MsgBox "Fill Computer Fee", vbInformation, "Empty Box"
    txtcomputer.SetFocus
    chk = "True"
    Exit Sub
End If

If Trim(txtlab.Text) = "" Then
    MsgBox "Fill Lab Fee", vbInformation, "Empty Box"
    txtlab.SetFocus
    chk = "True"
    Exit Sub
End If

If Trim(txtother.Text) = "" Then
    MsgBox "Fill Other Fee", vbInformation, "Empty Box"
    txtother.SetFocus
    chk = "True"
    Exit Sub
End If

End Sub
