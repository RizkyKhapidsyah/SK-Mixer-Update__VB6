VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form JamsMixer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mixer"
   ClientHeight    =   2415
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10875
   DrawWidth       =   4
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   9270
      Top             =   225
   End
   Begin VB.Frame Frame1 
      Caption         =   "VU Meter"
      Height          =   2310
      Left            =   9900
      TabIndex        =   94
      Top             =   45
      Width           =   960
      Begin VB.TextBox VuLeftLevel 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   96
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   285
      End
      Begin VB.TextBox VuRightLevel 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   540
         TabIndex        =   95
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   285
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   1590
         Left            =   180
         TabIndex        =   1
         Top             =   630
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   2805
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Max             =   3275
         Orientation     =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   1590
         Left            =   540
         TabIndex        =   0
         Top             =   630
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   2805
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Max             =   3275
         Orientation     =   1
         Scrolling       =   1
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   " Bass"
      Height          =   2325
      Index           =   11
      Left            =   45
      TabIndex        =   89
      Top             =   45
      Width           =   705
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   108
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   555
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   11
         Left            =   90
         TabIndex        =   90
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   345
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1770
         Index           =   11
         Left            =   30
         TabIndex        =   91
         Top             =   510
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   3122
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   11
         Left            =   90
         TabIndex        =   93
         Top             =   2340
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "Treble"
      Height          =   2325
      Index           =   10
      Left            =   765
      TabIndex        =   86
      Top             =   45
      Width           =   705
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   107
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   555
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   10
         Left            =   90
         TabIndex        =   88
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   525
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1770
         Index           =   10
         Left            =   45
         TabIndex        =   87
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   3122
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   10
         Left            =   45
         TabIndex        =   92
         Top             =   2340
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "Pc spk"
      Height          =   2325
      Index           =   9
      Left            =   8370
      TabIndex        =   78
      Top             =   45
      Width           =   705
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   106
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   555
      End
      Begin VB.CheckBox SBMLink 
         BackColor       =   &H8000000B&
         Height          =   210
         Index           =   9
         Left            =   90
         TabIndex        =   83
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture11 
         Height          =   330
         Index           =   9
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   80
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton MuteOn 
            BackColor       =   &H000000FF&
            Height          =   150
            Index           =   9
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   82
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton MuteOff 
            BackColor       =   &H000000FF&
            Height          =   150
            Index           =   9
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   81
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   9
         Left            =   90
         TabIndex        =   79
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1230
         Index           =   9
         Left            =   45
         TabIndex        =   84
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   9
         Left            =   45
         TabIndex        =   85
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "I25in"
      Height          =   2325
      Index           =   8
      Left            =   7605
      TabIndex        =   70
      Top             =   45
      Width           =   705
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   105
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   555
      End
      Begin VB.CheckBox SBMLink 
         BackColor       =   &H8000000B&
         Height          =   210
         Index           =   8
         Left            =   90
         TabIndex        =   75
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture11 
         Height          =   330
         Index           =   8
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   72
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton MuteOn 
            BackColor       =   &H000000FF&
            Height          =   150
            Index           =   8
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton MuteOff 
            BackColor       =   &H000000FF&
            Height          =   150
            Index           =   8
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   8
         Left            =   90
         TabIndex        =   71
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1230
         Index           =   8
         Left            =   30
         TabIndex        =   76
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   8
         Left            =   45
         TabIndex        =   77
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "MIDI"
      Height          =   2325
      Index           =   7
      Left            =   6840
      TabIndex        =   62
      Top             =   45
      Width           =   705
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   104
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   555
      End
      Begin VB.CheckBox SBMLink 
         BackColor       =   &H8000000B&
         Height          =   210
         Index           =   7
         Left            =   90
         TabIndex        =   67
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture11 
         Height          =   330
         Index           =   7
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   64
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton MuteOn 
            BackColor       =   &H000000FF&
            Height          =   150
            Index           =   7
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton MuteOff 
            BackColor       =   &H000000FF&
            Height          =   150
            Index           =   7
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   7
         Left            =   90
         TabIndex        =   63
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1230
         Index           =   7
         Left            =   30
         TabIndex        =   68
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   7
         Left            =   45
         TabIndex        =   69
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "TAD"
      Height          =   2325
      Index           =   6
      Left            =   6075
      TabIndex        =   54
      Top             =   45
      Width           =   705
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   103
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   555
      End
      Begin VB.CheckBox SBMLink 
         BackColor       =   &H8000000B&
         Height          =   210
         Index           =   6
         Left            =   90
         TabIndex        =   59
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture11 
         Height          =   330
         Index           =   6
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   56
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton MuteOn 
            BackColor       =   &H000000FF&
            Height          =   150
            Index           =   6
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton MuteOff 
            BackColor       =   &H000000FF&
            Height          =   150
            Index           =   6
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   6
         Left            =   90
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1230
         Index           =   6
         Left            =   30
         TabIndex        =   60
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   6
         Left            =   45
         TabIndex        =   61
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "Wave"
      Height          =   2325
      Index           =   5
      Left            =   5310
      TabIndex        =   46
      Top             =   45
      Width           =   705
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   102
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   555
      End
      Begin VB.CheckBox SBMLink 
         BackColor       =   &H8000000B&
         Height          =   210
         Index           =   5
         Left            =   90
         TabIndex        =   51
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture11 
         Height          =   330
         Index           =   5
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   48
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton MuteOn 
            BackColor       =   &H000000FF&
            Height          =   150
            Index           =   5
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton MuteOff 
            BackColor       =   &H000000FF&
            Height          =   150
            Index           =   5
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   5
         Left            =   90
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1230
         Index           =   5
         Left            =   30
         TabIndex        =   52
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   5
         Left            =   45
         TabIndex        =   53
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "AUX"
      Height          =   2325
      Index           =   4
      Left            =   4545
      TabIndex        =   38
      Top             =   45
      Width           =   705
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   101
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   555
      End
      Begin VB.CheckBox SBMLink 
         BackColor       =   &H8000000B&
         Height          =   210
         Index           =   4
         Left            =   90
         TabIndex        =   43
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture11 
         Height          =   330
         Index           =   4
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   40
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton MuteOn 
            BackColor       =   &H000000FF&
            Height          =   150
            Index           =   4
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton MuteOff 
            BackColor       =   &H000000FF&
            Height          =   150
            Index           =   4
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   4
         Left            =   90
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1230
         Index           =   4
         Left            =   30
         TabIndex        =   44
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   4
         Left            =   45
         TabIndex        =   45
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "Mic"
      Height          =   2325
      Index           =   3
      Left            =   3780
      TabIndex        =   30
      Top             =   45
      Width           =   705
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   100
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   555
      End
      Begin VB.CheckBox SBMLink 
         BackColor       =   &H8000000B&
         Height          =   210
         Index           =   3
         Left            =   90
         TabIndex        =   35
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture11 
         Height          =   330
         Index           =   3
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   32
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton MuteOn 
            BackColor       =   &H000000FF&
            Height          =   150
            Index           =   3
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton MuteOff 
            BackColor       =   &H000000FF&
            Height          =   150
            Index           =   3
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   3
         Left            =   90
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1230
         Index           =   3
         Left            =   30
         TabIndex        =   36
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   3
         Left            =   45
         TabIndex        =   37
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "Line In"
      Height          =   2325
      Index           =   2
      Left            =   3015
      TabIndex        =   22
      Top             =   45
      Width           =   705
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   99
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   555
      End
      Begin VB.CheckBox SBMLink 
         BackColor       =   &H8000000B&
         Height          =   210
         Index           =   2
         Left            =   90
         TabIndex        =   27
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture11 
         Height          =   330
         Index           =   2
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   24
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton MuteOn 
            BackColor       =   &H000000FF&
            Height          =   150
            Index           =   2
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton MuteOff 
            BackColor       =   &H000000FF&
            Height          =   150
            Index           =   2
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   2
         Left            =   90
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1230
         Index           =   2
         Left            =   45
         TabIndex        =   28
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   2
         Left            =   45
         TabIndex        =   29
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "Cd vol"
      Height          =   2325
      Index           =   1
      Left            =   2250
      TabIndex        =   14
      Top             =   45
      Width           =   705
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   98
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   555
      End
      Begin VB.CheckBox SBMLink 
         BackColor       =   &H8000000B&
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   19
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture11 
         Height          =   330
         Index           =   1
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   16
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton MuteOn 
            BackColor       =   &H000000FF&
            Height          =   150
            Index           =   1
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton MuteOff 
            BackColor       =   &H000000FF&
            Height          =   150
            Index           =   1
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1230
         Index           =   1
         Left            =   30
         TabIndex        =   20
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   1
         Left            =   45
         TabIndex        =   21
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "Master"
      Height          =   2325
      Index           =   0
      Left            =   1485
      TabIndex        =   6
      Top             =   45
      Width           =   705
      Begin VB.TextBox TXTVol 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   97
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   555
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin VB.PictureBox Picture11 
         Height          =   330
         Index           =   0
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   8
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton MuteOff 
            BackColor       =   &H000000FF&
            Height          =   150
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
         Begin VB.OptionButton MuteOn 
            BackColor       =   &H000000FF&
            Height          =   150
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
      End
      Begin VB.CheckBox SBMLink 
         BackColor       =   &H8000000B&
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   1800
         Width           =   240
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1230
         Index           =   0
         Left            =   30
         TabIndex        =   12
         Top             =   540
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   0
         Left            =   45
         TabIndex        =   13
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "SBM"
      Height          =   2325
      Left            =   9135
      TabIndex        =   2
      Top             =   45
      Width           =   705
      Begin VB.CommandButton Command1 
         Caption         =   "----I----"
         Height          =   150
         Left            =   90
         TabIndex        =   109
         ToolTipText     =   "Balance = 0"
         Top             =   2115
         Width           =   555
      End
      Begin VB.CheckBox Check11 
         BackColor       =   &H8000000B&
         Height          =   210
         Left            =   90
         TabIndex        =   3
         ToolTipText     =   "Links all selected faders to SBM fader"
         Top             =   1830
         Width           =   225
      End
      Begin ComctlLib.Slider Slider6 
         Height          =   1260
         Left            =   45
         TabIndex        =   4
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2223
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   1
         TickFrequency   =   3265
         Value           =   32768
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "SBM"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   300
         TabIndex        =   5
         Top             =   1830
         Width           =   375
      End
   End
End
Attribute VB_Name = "JamsMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is an update from my cd/mixer I have taken the cd player away
'and put the slider/mutes/balance controls into arrays.
'this makes it alittle more harder to understand but it is better/faster coding.
'If you wish to comment or get the VU meters working for me email me at
'micracom2@hotmail.com and I shall see if I can help.

Option Explicit
Dim volR As Long
Dim volL As Long
Dim volume As Long
Dim mute As MIXERCONTROL
Dim unmute As MIXERCONTROL
Dim ONOFF As MIXERCONTROL
Dim hmixer As Long             ' mixer handle
Dim VolCtrl As MIXERCONTROL    ' master volume control
Dim WavCtrl As MIXERCONTROL    ' wave output volume control
Dim CDVol As MIXERCONTROL      ' CD Volume
Dim LineVol As MIXERCONTROL    ' Line/In Volume
Dim MICROPHONE As MIXERCONTROL ' Microphone Volume
Dim PCSPEAKER As MIXERCONTROL    ' PcSpeaker Volume
Dim AUXVol As MIXERCONTROL     ' Auxillary Volume
Dim TELEPHONE As MIXERCONTROL     ' TAD-In Volume
Dim MIDIVol As MIXERCONTROL    ' Midi Volume
Dim I25InVol As MIXERCONTROL   ' I25In Volume
Dim Treble As MIXERCONTROL
Dim Bass As MIXERCONTROL
Dim rc As Long                 ' return code
Dim ok As Boolean              ' boolean return code

Private Sub Check11_Click()
Dim I As Integer 'Select all check box's
For I = 0 To 9
If Check11.Value = 1 Then
SBMLink(I).Value = 1
End If
If Check11.Value = 0 Then
SBMLink(I).Value = 0
End If
Next I
End Sub

Private Sub Command1_Click()
Dim A As Integer
For A = 0 To 9
BalanceControl(A).Value = 0
Next A
End Sub

Private Sub Form_Load()
Command2_Click 'Get Mixer Settings
End Sub

Function Errora()
    MsgBox "An error has ocurred."
End Function

Private Sub Timer1_Timer()
ProgressBar1.Value = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_SIGNEDMETER, VolCtrl) + 200
ProgressBar2.Value = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_SIGNEDMETER, VolCtrl) + 200
End Sub

Private Sub VolumeControl_Scroll(I As Integer)
    volume = 65535 - CLng(VolumeControl(I).Value)
    TXTVolumeControl(I).Text = volume
    BalanceControl_Scroll (I)
End Sub
Private Sub BalanceControl_Scroll(I As Integer)
    volR = TXTVolumeControl(I).Text * (IIf(BalanceControl(I) >= 0, 1, (100 + BalanceControl(I)) / 100))
    volL = TXTVolumeControl(I).Text * (IIf(BalanceControl(I) <= 0, 1, (100 - BalanceControl(I)) / 100))
Select Case (I)
Case 0
    SetPANControl hmixer, VolCtrl, volL, volR ' Stereo Mixer Control
    TXTVol = (volume / 6553)
Case 1
    SetPANControl hmixer, CDVol, volL, volR
    Text1 = (volume / 6553)
Case 2
    SetPANControl hmixer, LineVol, volL, volR
    Text2 = (volume / 6553)
Case 3
    SetVolumeControl hmixer, MICROPHONE, volume ' Mono Single Channel device
    Text3 = (volume / 6553)
Case 4
    SetPANControl hmixer, AUXVol, volL, volR
    Text4 = (volume / 6553)
Case 5
    SetPANControl hmixer, WavCtrl, volL, volR
    Text5 = (volume / 6553)
Case 6
    SetVolumeControl hmixer, TELEPHONE, volume
    Text6 = (volume / 6553)
Case 7
    SetPANControl hmixer, MIDIVol, volL, volR
    Text7 = (volume / 6553)
Case 8
    SetPANControl hmixer, I25InVol, volL, volR
    Text8 = (volume / 6553)
Case 9
    SetVolumeControl hmixer, PCSPEAKER, volume
    Text9 = (volume / 6553)
Case 10
    SetVolumeControl hmixer, Treble, volume
    Text10 = (volume / 6553)
Case 11
    SetVolumeControl hmixer, Bass, volume
    Text11 = (volume / 6553)
End Select
End Sub

Private Sub Command2_Click()
    'Open the mixer with deviceID 0.
    rc = mixerOpen(hmixer, 0, 0, 0, 0)
    If ((MMSYSERR_NOERROR <> rc)) Then
        MsgBox "Couldn't open the mixer please check if a audio mixer is installed then retry."
        Exit Sub
    End If
Dim I As Integer
For I = 0 To 11
Select Case I
Case 0
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_VOLUME, VolCtrl)
    If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, VolCtrl)
        TXTVol = (volume / 6553)
    End If
Case 1
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol, MIXERCONTROL_CONTROLTYPE_VOLUME, WavCtrl)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, WavCtrl)
        Text1 = (volume / 6553)
    End If
Case 2
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, MIXERCONTROL_CONTROLTYPE_VOLUME, MICROPHONE)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, MICROPHONE)
        Text2 = (volume / 6553)
    End If
Case 3
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_CDVol, MIXERCONTROL_CONTROLTYPE_VOLUME, CDVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, CDVol)
        Text3 = (volume / 6553)
    End If
Case 4
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_src_AUXVol, MIXERCONTROL_CONTROLTYPE_VOLUME, AUXVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, AUXVol)
        Text4 = (volume / 6553)
    End If
Case 5
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE, MIXERCONTROL_CONTROLTYPE_VOLUME, TELEPHONE)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, TELEPHONE)
        Text5 = (volume / 6553)
    End If
Case 6
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MIDIVol, MIXERCONTROL_CONTROLTYPE_VOLUME, MIDIVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, MIDIVol)
        Text6 = (volume / 6553)
    End If
Case 7
        ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER, MIXERCONTROL_CONTROLTYPE_VOLUME, PCSPEAKER)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, PCSPEAKER)
        Text7 = (volume / 6553)
    End If
Case 8
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_I25InVol, MIXERCONTROL_CONTROLTYPE_VOLUME, I25InVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, I25InVol)
        Text8 = (volume / 6553)
    End If
Case 9
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_LINEVol, MIXERCONTROL_CONTROLTYPE_VOLUME, LineVol)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, LineVol)
        Text9 = (volume / 6553)
    End If
Case 10
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_BASS, Bass)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, Bass)
        Text10 = (volume / 6553)
    End If
Case 11
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_TREBLE, Treble)
        If (ok = True) Then
        volume = GetVolumeControlValue(hmixer, Treble)
        Text11 = (volume / 6553)
    End If

End Select
        If volume <> -1 Then
            TXTVolumeControl(I) = volume
            VolumeControl(I) = 65535 - volume
        End If
Next I

End Sub
Private Sub MuteOn_Click(M As Integer) 'Mute on controls
Select Case M
Case 0
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
Case 1
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_CDVol, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
Case 2
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_LINEVol, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
Case 3
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
Case 4
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_src_AUXVol, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
Case 5
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
Case 6
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
Case 7
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MIDIVol, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
Case 8
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_I25InVol, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
Case 9
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
End Select
SetMuteControl hmixer, mute, 1
End Sub
Private Sub MuteOff_Click(MOff As Integer) 'Mute of contols
Select Case MOff
Case 0
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
Case 1
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_CDVol, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
Case 2
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_LINEVol, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
Case 3
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
Case 4
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_src_AUXVol, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
Case 5
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
Case 6
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
Case 7
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MIDIVol, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
Case 8
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_I25InVol, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
Case 9
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
End Select
unSetMuteControl hmixer, unmute, 1
End Sub
Private Sub Slider6_scroll()
Dim I As Integer
For I = 0 To 9
If SBMLink(I).Value = 1 Then
     SetPANControl hmixer, VolCtrl, volL, volR
     VolumeControl(I) = Slider6.Value
End If
VolumeControl_Scroll (I)
Next I
End Sub

