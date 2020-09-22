VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ColorBars ActiveX Control Demo"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   360
      Left            =   3150
      TabIndex        =   23
      ToolTipText     =   "Exit Demo"
      Top             =   4905
      Width           =   900
   End
   Begin VB.Frame Frame3 
      Caption         =   "And, of course, Progress Bar"
      Height          =   840
      Left            =   75
      TabIndex        =   27
      Top             =   3990
      Width           =   4065
      Begin VB.CheckBox cmdPauseProgress 
         Caption         =   "||"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3570
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Pause Progress Bar"
         Top             =   285
         Width           =   270
      End
      Begin VB.CommandButton cmdStartProgress 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3030
         TabIndex        =   20
         ToolTipText     =   "Start Progress Bar"
         Top             =   285
         Width           =   270
      End
      Begin VB.CommandButton cmdStopProgress 
         Caption         =   "g"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   6
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3300
         TabIndex        =   21
         ToolTipText     =   "Stop Progress Bar"
         Top             =   285
         Width           =   270
      End
      Begin prjColorBar.ColorBar ColorBar5 
         Height          =   285
         Left            =   120
         Top             =   330
         Width           =   2580
         _ExtentX        =   2117
         _ExtentY        =   318
         ForeColor       =   33023
         GradientEndColor=   65280
         GradientMidColor=   16777215
         GradientStartColor=   16776960
         Locked          =   0   'False
         UseGradient     =   -1  'True
      End
      Begin VB.Timer tmrProgress 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   4245
         Top             =   390
      End
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   4530
      Top             =   165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "ColorBars As Channel Levels:"
      Height          =   1080
      Left            =   75
      TabIndex        =   26
      Top             =   2850
      Width           =   4065
      Begin VB.CommandButton cmdStartLevel 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3030
         TabIndex        =   17
         ToolTipText     =   "Start Channel Levels"
         Top             =   420
         Width           =   270
      End
      Begin VB.CommandButton cmdStopLevel 
         Caption         =   "g"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   6
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3300
         TabIndex        =   18
         ToolTipText     =   "Stop Channel Levels"
         Top             =   420
         Width           =   270
      End
      Begin VB.CheckBox cmdPauseLevel 
         Caption         =   "||"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3570
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Pause Channel Levels"
         Top             =   420
         Width           =   270
      End
      Begin VB.Timer tmrLevel 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4080
         Top             =   300
      End
      Begin prjColorBar.ColorBar ColorBar3 
         Height          =   135
         Index           =   0
         Left            =   1335
         Top             =   420
         Width           =   1620
         _ExtentX        =   2117
         _ExtentY        =   318
         Locked          =   0   'False
         UseGradient     =   -1  'True
         UsePeaks        =   -1  'True
      End
      Begin prjColorBar.ColorBar ColorBar3 
         Height          =   135
         Index           =   1
         Left            =   1335
         Top             =   690
         Width           =   1620
         _ExtentX        =   2117
         _ExtentY        =   318
         Locked          =   0   'False
         UseGradient     =   -1  'True
         UsePeaks        =   -1  'True
      End
      Begin prjColorBar.ColorBar ColorBar4 
         Height          =   555
         Index           =   0
         Left            =   105
         Top             =   240
         Width           =   135
         _ExtentX        =   318
         _ExtentY        =   2117
         ForeColor       =   0
         Locked          =   0   'False
         Orientation     =   1
         PeakColor       =   0
         Segmented       =   -1  'True
         UsePeaks        =   -1  'True
      End
      Begin prjColorBar.ColorBar ColorBar4 
         Height          =   555
         Index           =   1
         Left            =   375
         Top             =   240
         Width           =   135
         _ExtentX        =   318
         _ExtentY        =   2117
         ForeColor       =   0
         Locked          =   0   'False
         Orientation     =   1
         PeakColor       =   0
         Segmented       =   -1  'True
         UsePeaks        =   -1  'True
      End
      Begin VB.Label lblHorzR 
         Caption         =   "Right:"
         Height          =   195
         Left            =   795
         TabIndex        =   16
         Top             =   660
         Width           =   465
      End
      Begin VB.Label lblHorzL 
         Caption         =   "Left:"
         Height          =   195
         Left            =   780
         TabIndex        =   15
         Top             =   390
         Width           =   375
      End
      Begin VB.Label lblVertR 
         Caption         =   "R"
         Height          =   195
         Left            =   390
         TabIndex        =   14
         Top             =   810
         Width           =   120
      End
      Begin VB.Label lblVertL 
         Caption         =   "L"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   810
         Width           =   120
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ColorBars As Visualization:"
      Height          =   2730
      Left            =   60
      TabIndex        =   24
      Top             =   60
      Width           =   4065
      Begin VB.CheckBox chkSeg 
         Caption         =   "Use Segments"
         Height          =   420
         Left            =   3000
         TabIndex        =   28
         ToolTipText     =   "Show solid ColorBar or in segments"
         Top             =   720
         Width           =   1005
      End
      Begin VB.PictureBox picPeak 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   2145
         ScaleHeight     =   180
         ScaleWidth      =   330
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Choose the color for the peaks"
         Top             =   1815
         Width           =   360
      End
      Begin VB.PictureBox picEnd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         ScaleHeight     =   180
         ScaleWidth      =   330
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Choose gradient color for top of the ColorBar"
         Top             =   2385
         Width           =   360
      End
      Begin VB.PictureBox picMid 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         ScaleHeight     =   180
         ScaleWidth      =   330
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Choose gradient color for middle of the ColorBar"
         Top             =   2100
         Width           =   360
      End
      Begin VB.PictureBox picStart 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         ScaleHeight     =   180
         ScaleWidth      =   330
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Choose gradient color for bottom of the ColorBar"
         Top             =   1815
         Width           =   360
      End
      Begin VB.CheckBox cmdPause 
         Caption         =   "||"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3570
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Pause Visualization"
         Top             =   285
         Width           =   270
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "g"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   6
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3300
         TabIndex        =   1
         ToolTipText     =   "Stop Visualization"
         Top             =   285
         Width           =   270
      End
      Begin VB.CheckBox chkPeaks 
         Caption         =   "Use Peaks"
         Height          =   240
         Left            =   2145
         TabIndex        =   4
         ToolTipText     =   "Display peak levels"
         Top             =   1425
         Width           =   1080
      End
      Begin VB.Timer tmrVis 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4185
         Top             =   585
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3030
         TabIndex        =   0
         ToolTipText     =   "Start Visualization"
         Top             =   285
         Width           =   270
      End
      Begin VB.Frame fraVis 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Visualization"
         ForeColor       =   &H00FFFFFF&
         Height          =   900
         Left            =   120
         TabIndex        =   25
         Top             =   285
         Width           =   2790
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   0
            Left            =   60
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   1
            Left            =   120
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   2
            Left            =   180
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   3
            Left            =   240
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   4
            Left            =   300
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   5
            Left            =   360
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   6
            Left            =   420
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   7
            Left            =   480
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   8
            Left            =   540
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   9
            Left            =   600
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   10
            Left            =   660
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   11
            Left            =   720
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   12
            Left            =   780
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   13
            Left            =   840
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   14
            Left            =   900
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   15
            Left            =   960
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   16
            Left            =   1020
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   17
            Left            =   1080
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   18
            Left            =   1140
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   19
            Left            =   1200
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   20
            Left            =   1260
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   21
            Left            =   1320
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   22
            Left            =   1380
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   23
            Left            =   1440
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   24
            Left            =   1500
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   25
            Left            =   1560
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   26
            Left            =   1620
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   27
            Left            =   1680
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   28
            Left            =   1740
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   29
            Left            =   1800
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   30
            Left            =   1860
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   31
            Left            =   1920
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   32
            Left            =   1980
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   33
            Left            =   2040
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   34
            Left            =   2100
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   35
            Left            =   2160
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   36
            Left            =   2220
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   37
            Left            =   2280
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   38
            Left            =   2340
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   39
            Left            =   2400
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   40
            Left            =   2460
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   41
            Left            =   2520
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   42
            Left            =   2580
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   43
            Left            =   2640
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
         Begin prjColorBar.ColorBar ColorBar2 
            Height          =   810
            Index           =   44
            Left            =   2700
            Top             =   75
            Width           =   30
            _ExtentX        =   318
            _ExtentY        =   2117
            ForeColor       =   16777215
            GradientEndColor=   16777215
            GradientMidColor=   16776960
            GradientStartColor=   65535
            Orientation     =   1
            PeakColor       =   65535
            UseGradient     =   -1  'True
         End
      End
      Begin VB.CheckBox chkGradient 
         Caption         =   "Use Gradient Colors"
         Height          =   240
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Enable/Disable Gradient Colors"
         Top             =   1425
         Value           =   1  'Checked
         Width           =   1800
      End
      Begin prjColorBar.ColorBar ColorBar1 
         Height          =   45
         Index           =   0
         Left            =   120
         Top             =   1710
         Width           =   1800
         _ExtentX        =   2117
         _ExtentY        =   318
         GradientEndColor=   4210752
         GradientMidColor=   33023
         GradientStartColor=   0
         Locked          =   0   'False
         Segmented       =   -1  'True
         UseGradient     =   -1  'True
         Value           =   100
      End
      Begin prjColorBar.ColorBar ColorBar1 
         Height          =   45
         Index           =   1
         Left            =   2145
         Top             =   1710
         Width           =   1815
         _ExtentX        =   2117
         _ExtentY        =   318
         GradientEndColor=   4210752
         GradientMidColor=   33023
         GradientStartColor=   0
         Locked          =   0   'False
         Segmented       =   -1  'True
         UseGradient     =   -1  'True
         Value           =   100
      End
      Begin VB.Label Label1 
         Caption         =   "Peak Color"
         Height          =   210
         Left            =   2550
         TabIndex        =   12
         Top             =   1830
         Width           =   870
      End
      Begin VB.Label lblEnd 
         Caption         =   "End Gradient Color"
         Height          =   210
         Left            =   525
         TabIndex        =   10
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label lblMid 
         Caption         =   "Mid Gradient Color"
         Height          =   210
         Left            =   525
         TabIndex        =   8
         Top             =   2115
         Width           =   1455
      End
      Begin VB.Label lblStart 
         Caption         =   "Start Gradient Color"
         Height          =   210
         Left            =   525
         TabIndex        =   6
         Top             =   1830
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bIsPlaying As Boolean

Private Sub chkGradient_Click()
     Dim lLoop As Long
     If chkGradient Then
          lblStart.Caption = "Start Gradient Color"
          picMid.Visible = True
          lblMid.Visible = True
          picEnd.Visible = True
          lblEnd.Visible = True
          For lLoop = 0 To ColorBar2.UBound
               With ColorBar2(lLoop)
                    .UseGradient = True
                    .GradientStartColor = picStart.BackColor
                    .GradientMidColor = picMid.BackColor
                    .GradientEndColor = picEnd.BackColor
               End With
          Next
     Else
          lblStart.Caption = "ColorBar Color"
          picMid.Visible = False
          lblMid.Visible = False
          picEnd.Visible = False
          lblEnd.Visible = False
          For lLoop = 0 To ColorBar2.UBound
               With ColorBar2(lLoop)
                    .UseGradient = False
                    .ForeColor = picStart.BackColor
               End With
          Next
     End If
End Sub ' chkGradient_Click

Private Sub chkPeaks_Click()
     Dim lLoop As Long
     If chkPeaks Then
          For lLoop = 0 To ColorBar2.UBound
               ColorBar2(lLoop).UsePeaks = True
          Next
     Else
          For lLoop = 0 To ColorBar2.UBound
               ColorBar2(lLoop).UsePeaks = False
          Next
     End If
End Sub ' chkPeaks_Click

Private Sub chkSeg_Click()
     Dim lLoop As Long
     If chkSeg Then
          For lLoop = 0 To ColorBar2.UBound
               ColorBar2(lLoop).Segmented = True
          Next
     Else
          For lLoop = 0 To ColorBar2.UBound
               ColorBar2(lLoop).Segmented = False
          Next
     End If
End Sub ' chkSeg_Click

Private Sub cmdExit_Click()
     Unload Me
End Sub ' cmdExit_Click

Private Sub cmdPause_Click()
     ' if checked
     If cmdPause Then
          ' disable timer
          tmrVis.Enabled = False
     Else
          ' enable timer
          tmrVis.Enabled = True
     End If
End Sub ' cmdPause_Click

Private Sub cmdPauseLevel_Click()
     If cmdPauseLevel Then
          ' disable timer
          tmrLevel.Enabled = False
     Else
          ' enable timer
          tmrLevel.Enabled = True
     End If
End Sub ' cmdPauseLevel_Click

Private Sub cmdPauseProgress_Click()
     If cmdPauseProgress Then
          ' disable timer
          tmrProgress.Enabled = False
     Else
          ' enable timer
          tmrProgress.Enabled = True
     End If
End Sub ' cmdPauseProgress_Click

Private Sub cmdStart_Click()
     ' stop other demos
'     cmdStopLevel = True
'     cmdStopProgress = True
     ' set global
     bIsPlaying = True
     ' start timer
     tmrVis.Enabled = True
     ' disable start button
     cmdStart.Enabled = False
     ' enable pause and stop buttons
     cmdStop.Enabled = True
     cmdPause.Enabled = True
     ' make sure cmdPause is not checked (graphical checkbox)
     cmdPause = False
End Sub ' cmdStart_Click

Private Sub cmdStartLevel_Click()
     ' stop other demos
'     cmdStop = True
'     cmdStopProgress = True
     ' Start timer
     tmrLevel.Enabled = True
     ' disable start button
     cmdStartLevel.Enabled = False
     ' enable pause and stop buttons
     cmdStopLevel.Enabled = True
     cmdPauseLevel.Enabled = True
     ' make sure cmdPause is not checked (graphical checkbox)
     cmdPauseLevel = False
End Sub ' cmdStartLevel_Click

Private Sub cmdStartProgress_Click()
     ' stop other demos
'     cmdStopLevel = True
'     cmdStop = True
     ' Start timer
     tmrProgress.Enabled = True
     ' disable start button
     cmdStartProgress.Enabled = False
     ' enable pause and stop buttons
     cmdStopProgress.Enabled = True
     cmdPauseProgress.Enabled = True
     ' make sure cmdPause is not checked (graphical checkbox)
     cmdPauseProgress = False
End Sub ' cmdStartProgress_Click

Private Sub cmdStop_Click()
     Dim lLoop As Long
     ' set global
     bIsPlaying = False
     ' if cmdPause is checked then uncheck it (graphical checkbox)
     cmdPause = False
     ' disable pause and stop buttons
     cmdPause.Enabled = False
     cmdStop.Enabled = False
     ' enable start button
     cmdStart.Enabled = True
     ' disable timer
     tmrVis.Enabled = False
     ' loop through colorbar array and set value to 0
     For lLoop = 0 To ColorBar2.UBound
          ColorBar2(lLoop).Value = False
     Next
End Sub ' cmdStop_Click

Private Sub cmdStopLevel_Click()
     tmrLevel.Enabled = False
     ColorBar3(0).Value = False
     ColorBar3(1).Value = False
     ColorBar4(0).Value = False
     ColorBar4(1).Value = False
     cmdStartLevel.Enabled = True
End Sub ' cmdStopLevel_Click

Private Sub cmdStopProgress_Click()
     tmrProgress.Enabled = False
     ColorBar5.Value = False
     cmdStartProgress.Enabled = True
End Sub ' cmdStopProgress_Click

Private Sub Form_Load()
     picStart.BackColor = ColorBar2(0).GradientStartColor
     picMid.BackColor = ColorBar2(0).GradientMidColor
     picEnd.BackColor = ColorBar2(0).GradientEndColor
     picPeak.BackColor = ColorBar2(0).PeakColor
     chkPeaks = 1
     chkSeg = 1
End Sub ' Form_Load

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     tmrLevel.Enabled = False
     DoEvents
     tmrProgress.Enabled = False
     DoEvents
     tmrVis.Enabled = False
     DoEvents
End Sub ' Form_QueryUnload

Private Sub Form_Unload(Cancel As Integer)
     Set frmTest = Nothing
     End
End Sub ' Form_Unload

Private Sub picEnd_Click()
     Dim lLoop As Long
     cdlg.CancelError = True
     On Error GoTo errHandler
     ' Display the Color Dialog box
     cdlg.ShowColor
     ' set picturebox backcolor
     picEnd.BackColor = cdlg.Color
     ' Loop through and set gradient start color
     For lLoop = 0 To ColorBar2.UBound
          ColorBar2(lLoop).GradientEndColor = cdlg.Color
     Next
errHandler:
End Sub ' picEnd_Click

Private Sub picMid_Click()
     Dim lLoop As Long
     cdlg.CancelError = True
     On Error GoTo errHandler
     ' Display the Color Dialog box
     cdlg.ShowColor
     ' set picturebox backcolor
     picMid.BackColor = cdlg.Color
     ' Loop through and set gradient start color
     For lLoop = 0 To ColorBar2.UBound
          ColorBar2(lLoop).GradientMidColor = cdlg.Color
     Next
errHandler:
End Sub ' picMid_Click

Private Sub picPeak_Click()
     Dim lLoop As Long
     cdlg.CancelError = True
     On Error GoTo errHandler
     ' Display the Color Dialog box
     cdlg.ShowColor
     ' set picturebox backcolor
     picPeak.BackColor = cdlg.Color
     ' Loop through and set gradient start color
     For lLoop = 0 To ColorBar2.UBound
          ColorBar2(lLoop).PeakColor = cdlg.Color
     Next
errHandler:
End Sub ' picPeak_Click

Private Sub picStart_Click()
     Dim lLoop As Long
     cdlg.CancelError = True
     On Error GoTo errHandler
     ' Display the Color Dialog box
     cdlg.ShowColor
     ' set picturebox backcolor
     picStart.BackColor = cdlg.Color
     ' Loop through and set gradient start color
     If chkGradient Then
          For lLoop = 0 To ColorBar2.UBound
               ColorBar2(lLoop).GradientStartColor = cdlg.Color
          Next
     Else
          For lLoop = 0 To ColorBar2.UBound
               ColorBar2(lLoop).ForeColor = cdlg.Color
          Next
     End If
errHandler:
End Sub ' picStart_Click

Private Sub tmrLevel_Timer()
     Dim lLoop As Long
     For lLoop = 0 To ColorBar4.UBound
          ColorBar4(lLoop).Value = (100 * Rnd) + 1
     Next
     For lLoop = 0 To ColorBar3.UBound
          ColorBar3(lLoop).Value = (100 * Rnd) + 1
     Next
End Sub ' tmrLevel_Timer

Private Sub tmrProgress_Timer()
     Dim lLoop As Long
     For lLoop = 0 To 100
          ColorBar5.Value = lLoop
          DoEvents
     Next
     For lLoop = 100 To 0 Step -1
          ColorBar5.Value = lLoop
          DoEvents
     Next
End Sub ' tmrProgress_Timer

Private Sub TmrVis_Timer()
     Dim lLoop As Long
     For lLoop = 0 To ColorBar2.UBound
          ' had to put this here because of something curious that
          ' is happening.  When the timer is disabled, it seems to
          ' want to continue firing the timer event a few times...
          If bIsPlaying Then _
               ColorBar2(lLoop).Value = (100 * Rnd) + 1
     Next
End Sub ' TmrVis_Timer
