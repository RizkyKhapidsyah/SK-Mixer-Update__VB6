Attribute VB_Name = "Nebular1"
Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As Long, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long

Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)
       

Public Const MAXPNAMELEN = 32
Public Const MMSYSERR_NOERROR = 0
Public Const MIXER_SHORT_NAME_CHARS = 16
Public Const MIXER_LONG_NAME_CHARS = 64
Public Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Public Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2& ' separate left-right volume control
Public Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
Public Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Public Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Public Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Public Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Public Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000

Public Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Public Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Public Const MIXERCONTROL_CONTROLTYPE_MUTE = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
Public Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Public Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)

Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long


''''''''''''''''''''''''''''''''''''''''''''''''''''
Type MIXERCAPS
     wMid As Integer                   '  manufacturer id
     wPid As Integer                   '  product id
     vDriverVersion As Long            '  version of the driver
     szPname As String * MAXPNAMELEN   '  product name
     fdwSupport As Long             '  misc. support bits
     cDestinations As Long          '  count of destinations
End Type

' Mixer line types
Type Target
     dwType As Long                 '  MIXERLINE_TARGETTYPE_xxxx
     dwDeviceID As Long             '  target device ID of device type
     wMid As Integer                   '  of target device
     wPid As Integer                   '       "
     vDriverVersion As Long            '       "
     szPname As String * MAXPNAMELEN
End Type

Type MIXERLINE
     cbStruct As Long               '  size of MIXERLINE structure
     dwDestination As Long          '  zero based destination index
     dwSource As Long               '  zero based source index (if source)
     dwLineID As Long               '  unique line id for mixer device
     fdwLine As Long                '  state/information about line
     dwUser As Long                 '  driver specific information
     dwComponentType As Long        '  component type line connects to
     cChannels As Long              '  number of channels line supports
     cConnections As Long           '  number of connections (possible)
     cControls As Long              '  number of controls at this line
     szShortName As String * MIXER_SHORT_NAME_CHARS
     szName As String * MIXER_LONG_NAME_CHARS
     lpTarget As Target
End Type

' MM Control types
Type MIXERLINECONTROLS
     cbStruct As Long         '  size in Byte of MIXERLINECONTROLS
     dwLineID As Long         '  line id (from MIXERLINE.dwLineID)
     dwControl As Long        '  used with MIXER_GETLINECONTROLSF_ONEBYTYPE or MIXER_GETLINECONTROLSF_ONEBYID
     cControls As Long        '  count of controls pmxctrl points to
     cbmxctrl As Long         '  size in Byte of _one_ MIXERCONTROL
     pamxctrl As Long         '  pointer to first MIXERCONTROL array
End Type

Type MIXERCONTROL
     cbStruct As Long           '  size in Byte of MIXERCONTROL
     dwControlID As Long        '  unique control id for mixer device
     dwControlType As Long      '  MIXERCONTROL_CONTROLTYPE_xxx
     fdwControl As Long         '  MIXERCONTROL_CONTROLF_xxx
     cMultipleItems As Long     '  if MIXERCONTROL_CONTROLF_MULTIPLE set
     szShortName(1 To MIXER_SHORT_NAME_CHARS) As Byte
     szName(1 To MIXER_LONG_NAME_CHARS) As Byte
     Bounds(1 To 6) As Long
     Metrics(1 To 6) As Long
End Type

Type MIXERCONTROLDETAILS
     cbStruct As Long       '  size in Byte of MIXERCONTROLDETAILS
     dwControlID As Long    '  control id to get/set details on
     cChannels As Long      '  number of channels in paDetails array
     item As Long                           ' hwndOwner or cMultipleItems
     cbDetails As Long      '  size of _one_ details_XX struct
     paDetails As Long      '  pointer to array of details_XX structs
End Type

Type MIXERCONTROLDETAILS_BOOLEAN
     fValue As Long
End Type

Type MIXERCONTROLDETAILS_LISTTEXT
     dwParam1 As Long
     dwParam2 As Long
     szName As String * MIXER_LONG_NAME_CHARS
End Type

Type MIXERCONTROLDETAILS_SIGNED
     lValue As Long
End Type

Type MIXERCONTROLDETAILS_UNSIGNED
     dwValue As Long
End Type

Function GetMixerControl(ByVal hmixer As Long, ByVal componentType As Long, ByVal ctrlType As Long, _
                        ByRef mxc As MIXERCONTROL) As Boolean
                        
' This function attempts to obtain a mixer control. Returns True if successful.
   Dim mxlc As MIXERLINECONTROLS
   Dim mxl As MIXERLINE
   Dim hmem As Long
   Dim rc As Long
       
   mxl.cbStruct = Len(mxl)
   mxl.dwComponentType = componentType
   ' Obtain a line corresponding to the component type
   rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
   If (MMSYSERR_NOERROR = rc) Then
       mxlc.cbStruct = Len(mxlc)
       mxlc.dwLineID = mxl.dwLineID
       mxlc.dwControl = ctrlType
       mxlc.cControls = 1
       mxlc.cbmxctrl = Len(mxc)
       ' Allocate a buffer for the control
       hmem = GlobalAlloc(&H40, Len(mxc))
       mxlc.pamxctrl = GlobalLock(hmem)
       mxc.cbStruct = Len(mxc)
       ' Get the control
       rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
       If (MMSYSERR_NOERROR = rc) Then
           GetMixerControl = True
           ' Copy the control into the destination structure
           CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
       Else
           GetMixerControl = False
       End If
       GlobalFree (hmem)
       Exit Function
   End If
   GetMixerControl = False
End Function


Function GetVolumeControlValue(ByVal hmixer As Long, mxc As MIXERCONTROL) As Long
'This function Gets the value for a volume control. Returns True if successful
    Dim mxcd As MIXERCONTROLDETAILS
    Dim vol As MIXERCONTROLDETAILS_UNSIGNED
    mxcd.cbStruct = Len(mxcd)
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cChannels = 1
    mxcd.item = 0
    mxcd.cbDetails = Len(vol)
    mxcd.paDetails = 0
    hmem = GlobalAlloc(&H40, Len(vol))
    mxcd.paDetails = GlobalLock(hmem)
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr vol, mxcd.paDetails, Len(vol)
    GlobalFree (hmem)
    If (MMSYSERR_NOERROR = rc) Then
       GetVolumeControlValue = vol.dwValue
    Else
        GetVolumeControlValue = -1
    End If
End Function



Function SetPANControl(ByVal hmixer As Long, mxc As MIXERCONTROL, ByVal volL As Long, ByVal volR As Long) As Boolean
'This function sets the value for a volume control. Returns True if successful
   Dim mxcd As MIXERCONTROLDETAILS
   Dim vol(1) As MIXERCONTROLDETAILS_UNSIGNED
   mxcd.item = mxc.cMultipleItems
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cbStruct = Len(mxcd)
   mxcd.cbDetails = Len(vol(1))
   ' Allocate a buffer for the control value buffer
   mxcd.cChannels = 2
   hmem = GlobalAlloc(&H40, Len(vol(1)))
   mxcd.paDetails = GlobalLock(hmem)
   vol(1).dwValue = volR
   vol(0).dwValue = volL
   ' Copy the data into the control value buffer
   CopyPtrFromStruct mxcd.paDetails, vol(1).dwValue, Len(vol(0)) * mxcd.cChannels
   CopyPtrFromStruct mxcd.paDetails, vol(0).dwValue, Len(vol(1)) * mxcd.cChannels
   ' Set the control value
   rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
   GlobalFree (hmem)
   If (MMSYSERR_NOERROR = rc) Then
       SetPANControl = True
   Else
       SetPANControl = False
   End If
   
End Function
Function unSetMuteControl(ByVal hmixer As Long, mxc As MIXERCONTROL, ByVal unmute As Long) As Boolean
   Dim mxcd As MIXERCONTROLDETAILS
   Dim vol As MIXERCONTROLDETAILS_UNSIGNED
   mxcd.cbStruct = Len(mxcd)
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cChannels = 1
   mxcd.item = 0
   mxcd.cbDetails = Len(vol)
   hmem = GlobalAlloc(&H40, Len(vol))
   mxcd.paDetails = GlobalLock(hmem)
   vol.dwValue = unmute
   CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
   rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
   GlobalFree (hmem)
   If (MMSYSERR_NOERROR = rc) Then
       unSetMuteControl = True
   Else
       unSetMuteControl = False
   End If
End Function

Function SetMuteControl(ByVal hmixer As Long, mxc As MIXERCONTROL, mute As Boolean) As Boolean
   Dim mxcd As MIXERCONTROLDETAILS
   Dim vol As MIXERCONTROLDETAILS_UNSIGNED
   mxcd.cbStruct = Len(mxcd)
   mxcd.dwControlID = mxc.dwControlID
   mxcd.cChannels = 1
   mxcd.item = 0
   mxcd.cbDetails = Len(vol)
   hmem = GlobalAlloc(&H40, Len(vol))
   mxcd.paDetails = GlobalLock(hmem)
   vol.dwValue = volume
   CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
   rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
   GlobalFree (hmem)
   If (MMSYSERR_NOERROR = rc) Then
       SetMuteControl = True
   Else
       SetMuteControl = False
   End If
End Function

