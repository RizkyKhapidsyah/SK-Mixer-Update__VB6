VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form JamsMixer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mixer"
   ClientHeight    =   4575
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6330
   DrawWidth       =   4
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   9270
      Top             =   225
   End
   Begin VB.Frame VolFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Master"
      ForeColor       =   &H80000008&
      Height          =   3615
      Index           =   0
      Left            =   2610
      TabIndex        =   0
      Top             =   750
      Width           =   1095
      Begin VB.TextBox TXTVol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   420
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   240
         Width           =   315
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   2190
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   3863
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   3265
         Value           =   32768
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   2370
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.OptionButton MuteOn 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mute off"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   6
         ToolTipText     =   "Unmute"
         Top             =   3330
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.OptionButton MuteOff 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mute On"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   5
         ToolTipText     =   "Mute"
         Top             =   3120
         Width           =   975
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   2670
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Balance"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   4
         Top             =   2880
         Width           =   510
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "The code is still large but all is needed for this to work. I hope this helps"
      Height          =   225
      Left            =   60
      TabIndex        =   9
      Top             =   390
      Width           =   6165
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Just Volume Control"
      Height          =   225
      Left            =   1650
      TabIndex        =   8
      Top             =   90
      Width           =   2925
   End
End
Attribute VB_Name = "JamsMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'micracom2@hotmail.com (Jamie Pocock)

Option Explicit
Dim volR As Long
Dim volL As Long
Dim volume As Long
Dim mute As MIXERCONTROL
Dim unmute As MIXERCONTROL
Dim ONOFF As MIXERCONTROL
Dim hmixer As Long             ' mixer handle
Dim VolCtrl As MIXERCONTROL    ' master volume control

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
    volume = 65535 - CLng(VolumeControl(0).Value)
    TXTVolumeControl(0).Text = volume
    BalanceControl_Scroll (0)
End Sub
Private Sub BalanceControl_Scroll(I As Integer)
    volR = TXTVolumeControl(0).Text * (IIf(BalanceControl(0) >= 0, 1, (100 + BalanceControl(0)) / 100))
    volL = TXTVolumeControl(0).Text * (IIf(BalanceControl(0) <= 0, 1, (100 - BalanceControl(0)) / 100))

    SetPANControl hmixer, VolCtrl, volL, volR ' Stereo Mixer Control
    TXTVol = (volume / 6553)

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


End Select
        If volume <> -1 Then
            TXTVolumeControl(0) = volume
            VolumeControl(0) = 65535 - volume
        End If
Next I

End Sub
Private Sub MuteOn_Click(M As Integer) 'Mute on controls
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
SetMuteControl hmixer, mute, 1
End Sub
Private Sub MuteOff_Click(MOff As Integer) 'Mute of contols
ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
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

