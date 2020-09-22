VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MIDILib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Data structure used to store information about each MIDI device.
Private Type MIDIOUTCAPS
    wMid As Integer
    wPid As Integer
    vDriverVersion As Long 'Version of MIDI device driver.
    szPname As String * 32 'Name of MIDI device.
    wTechnology As Integer
    wVoices As Integer
    wNotes As Integer
    wChannelMask As Integer
    dwSupport As Long
End Type

'API declarations used for MIDI.
'-------------------------------
'Closes the connection to a MIDI device (hMidiOut).
Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long

'Gets the caption (name) to a specified MIDI device (uDeviceID).
'The device info is dumped into the lpCaps structure.
Private Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long

'Gets the total number of MIDI devices available.
Private Declare Function midiOutGetNumDevs Lib "winmm" () As Integer

'Opens a connection to the specified MIDI device (uDeviceID).
'The variable passed to the lphMidiOut argument will receive the open MIDI device's handle.
'This device handle is used to send messages to the MIDI device (to play sounds, etc.).
Private Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long

'Sends a message to the specified MIDI device's handle (hMidiOut).
Private Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long

'Maximum volume value (0-127).
Private Const MIDI_MAX_VOLUME As Integer = 127

'MIDI instrument constants:
'--------------------------
Private Const INST_ACOUSTIC_GRAND As Long = 0
Private Const INST_BRIGHT_ACOUSTIC As Long = 1
Private Const INST_ELECTRIC_GRAND As Long = 2
Private Const INST_HONKY_TONK As Long = 3
Private Const INST_ELECTRIC_PIANO_1 As Long = 4
Private Const INST_ELECTRIC_PIANO_2 As Long = 5
Private Const INST_HARPSICHORD As Long = 6
Private Const INST_CLAV As Long = 7
Private Const INST_CELESTA As Long = 8
Private Const INST_GLOCKENSPIEL As Long = 9
Private Const INST_MUSIC_BOX As Long = 10
Private Const INST_VIBRAPHONE As Long = 11
Private Const INST_MARIMBA As Long = 12
Private Const INST_XYLOPHONE As Long = 13
Private Const INST_TUBULAR_BELLS As Long = 14
Private Const INST_DULCIMER As Long = 15
Private Const INST_DRAWBAR_ORGAN As Long = 16
Private Const INST_PERCUSSIVE_ORGAN As Long = 17
Private Const INST_ROCK_ORGAN As Long = 18
Private Const INST_CHURCH_ORGAN As Long = 19
Private Const INST_REED_ORGAN As Long = 20
Private Const INST_ACCORDIAN As Long = 21
Private Const INST_HARMONICA As Long = 22
Private Const INST_TANGO_ACCORDIAN As Long = 23
Private Const INST_ACOUSTIC_GUITAR_NYLON As Long = 24
Private Const INST_ACOUSTIC_GUITAR_STEEL As Long = 25
Private Const INST_ELECTRIC_GUITAR_JAZZ As Long = 26
Private Const INST_ELECTRIC_GUITAR_CLEAN As Long = 27
Private Const INST_ELECTRIC_GUITAR_MUTED As Long = 28
Private Const INST_OVERDRIVEN_GUITAR As Long = 29
Private Const INST_DISTORTION_GUITAR As Long = 30
Private Const INST_GUITAR_HARMONICS As Long = 31
Private Const INST_ACOUSTIC_BASS As Long = 32
Private Const INST_ELECTRIC_BASS_FINGER As Long = 33
Private Const INST_ELECTRIC_BASS_PICK As Long = 34
Private Const INST_FRETLESS_BASS As Long = 35
Private Const INST_SLAP_BASS_1 As Long = 36
Private Const INST_SLAP_BASS_2 As Long = 37
Private Const INST_SYNTH_BASS_1 As Long = 38
Private Const INST_SYNTH_BASS_2 As Long = 39
Private Const INST_VIOLIN As Long = 40
Private Const INST_VIOLA As Long = 41
Private Const INST_CELLO As Long = 42
Private Const INST_CONTRABASS As Long = 43
Private Const INST_TREMELO_STRINGS As Long = 44
Private Const INST_PIZZICATO_STRINGS As Long = 45
Private Const INST_ORCHESTRAL_STRINGS As Long = 46
Private Const INST_TIMPANI As Long = 47
Private Const INST_STRING_ENSEMBLE_1 As Long = 48
Private Const INST_STRING_ENSEMBLE_2 As Long = 49
Private Const INST_SYNTH_STRINGS_1 As Long = 50
Private Const INST_SYNTH_STRINGS_2 As Long = 51
Private Const INST_CHOIR_AAHS As Long = 52
Private Const INST_VOICE_OOHS As Long = 53
Private Const INST_SYNTH_VOICE As Long = 54
Private Const INST_ORCHESTRA_HIT As Long = 55
Private Const INST_TRUMPET As Long = 56
Private Const INST_TROMBONE As Long = 57
Private Const INST_TUBA As Long = 58
Private Const INST_MUTED_TRUMPET As Long = 59
Private Const INST_FRENCH_HORN As Long = 60
Private Const INST_BRASS_SECTION As Long = 61
Private Const INST_SYNTH_BRASS_1 As Long = 62
Private Const INST_SYNTH_BRASS_2 As Long = 63
Private Const INST_SOPRANO_SAX As Long = 64
Private Const INST_ALTO_SAX As Long = 65
Private Const INST_TENOR_SAX As Long = 66
Private Const INST_BARITONE_SAX As Long = 67
Private Const INST_OBOE As Long = 68
Private Const INST_ENGLISH_HORN As Long = 69
Private Const INST_BASSOON As Long = 70
Private Const INST_CLARINET As Long = 71
Private Const INST_PICCOLO As Long = 72
Private Const INST_FLUTE As Long = 73
Private Const INST_RECORDER As Long = 74
Private Const INST_PAN_FLUTE As Long = 75
Private Const INST_BLOWN_BOTTLE As Long = 76
Private Const INST_SKAKUHACHI As Long = 77
Private Const INST_WHISTLE As Long = 78
Private Const INST_OCARINA As Long = 79
Private Const INST_LEAD_1_SQUARE As Long = 80
Private Const INST_LEAD_2_SAWTOOTH As Long = 81
Private Const INST_LEAD_3_CALLIOPE As Long = 82
Private Const INST_LEAD_4_CHIFF As Long = 83
Private Const INST_LEAD_5_CHARANG As Long = 84
Private Const INST_LEAD_6_VOICE As Long = 85
Private Const INST_LEAD_7_FIFTHS As Long = 86
Private Const INST_LEAD_8_BASS_LEAD As Long = 87
Private Const INST_PAD_1_NEW_AGE As Long = 88
Private Const INST_PAD_2_WARM As Long = 89
Private Const INST_PAD_3_POLYSYNTH As Long = 90
Private Const INST_PAD_4_CHOIR As Long = 91
Private Const INST_PAD_5_BOWED As Long = 92
Private Const INST_PAD_6_METALLIC As Long = 93
Private Const INST_PAD_7_HALO As Long = 94
Private Const INST_PAD_8_SWEEP As Long = 95
Private Const INST_FX_1_RAIN As Long = 96
Private Const INST_FX_2_SOUNDTRACK As Long = 97
Private Const INST_FX_3_CRYSTAL As Long = 98
Private Const INST_FX_4_ATMOSPHERE As Long = 99
Private Const INST_FX_5_BRIGHTNESS As Long = 100
Private Const INST_FX_6_GOBLINS As Long = 101
Private Const INST_FX_7_ECHOS As Long = 102
Private Const INST_FX_8_SCIFI As Long = 103
Private Const INST_SITAR As Long = 104
Private Const INST_BANJO As Long = 105
Private Const INST_SHAMISEN As Long = 106
Private Const INST_KOTO As Long = 107
Private Const INST_KALIMBA As Long = 108
Private Const INST_BAGPIPE As Long = 109
Private Const INST_FIDDLE As Long = 110
Private Const INST_SHANAI As Long = 111
Private Const INST_TINKLE_BELL As Long = 112
Private Const INST_AGOGO As Long = 113
Private Const INST_STEEL_DRUMS As Long = 114
Private Const INST_WOODBLOCK As Long = 115
Private Const INST_TAIKO_DRUM As Long = 116
Private Const INST_MELODIC_TOM As Long = 117
Private Const INST_SYNTH_DRUM As Long = 118
Private Const INST_REVERSE_CYMBAL As Long = 119
Private Const INST_GUITAR_FRET_NOISE As Long = 120
Private Const INST_BREATH_NOISE As Long = 121
Private Const INST_SEASHORE As Long = 122
Private Const INST_BIRD_TWEET As Long = 123
Private Const INST_TELEPHONE_RING As Long = 124
Private Const INST_HELICOPTER As Long = 125
Private Const INST_APPLAUSE As Long = 126
Private Const INST_GUNSHOT As Long = 127

Private intBaseNote As Integer 'Root or starting note.
Private intVolume As Integer 'Current MIDI playback volume.
Private intChannel As Integer 'Current MIDI channel (default is: 1).
Private intInstrument As Integer 'Current MIDI instrument (0-127).

Private lonMIDIHand As Long 'Currently open MIDI device handle.
Private lonDevID As Long 'Current MIDI device being used.
Private lonDevCount As Long 'Total number of MIDI devices available.
Private lonRet As Long 'Temporary variable to store API return values.
Private lonMsg As Long 'Output message being sent to the MidiOutShortMsg() API function.

Private bolDevOpen As Boolean 'MIDI device currently open?

'Returns the note values for each string on a guitar.
Public Function GuitarNotes() As Integer()
'Ret(0) = E1
'Ret(1) = A2
'Ret(2) = D3
'Ret(3) = G4
'Ret(4) = B5
'Ret(5) = E6
Dim intRet(5) As Integer

intRet(0) = 40 'First string.
intRet(1) = 45 'Second string.
intRet(2) = 50 'Third string.
intRet(3) = 55 'Fourth string.
intRet(4) = 59 'Fifth string.
intRet(5) = 64 'Sixth string.

GuitarNotes = intRet()
End Function

'Starts playing a note by sending a message to MidiOutShortMsg().
Private Sub StartNote(ByVal NoteValue As Integer)
lonMsg = &H90 + ((intBaseNote + NoteValue) * &H100) + (intVolume * &H10000) + intChannel
midiOutShortMsg lonMIDIHand, lonMsg
End Sub

'Stops playing a note by sending a message to MidiOutShortMsg().
Private Sub StopNote(ByVal NoteValue As Integer)
lonMsg = &H80 + ((intBaseNote + NoteValue) * &H100) + intChannel
midiOutShortMsg lonMIDIHand, lonMsg
End Sub

'Plays a note for the specified duration (in seconds).
Public Sub PlayNote(ByVal NoteValue As Integer, ByVal Duration As Integer)
StartNote NoteValue 'Start playing the note.
PauseNow Duration 'Continue playing until PauseNow routine finishes.
StopNote NoteValue 'Pause routine finished, stop the note.
End Sub

'Properties:
'-----------
Public Property Get DeviceID() As Long
DeviceID = lonDevID
End Property

Public Property Get Instrument() As Integer
Instrument = intInstrument
End Property

'Update the instrument by sending the appropriate message to the midiOutShortMsg() function.
Private Sub UpdateInstrument()
If bolDevOpen = True Then 'Check if the device is open first.
    lonMsg = (intInstrument * 256) + &HC0 + intChannel + (0 * 256) * 256
    midiOutShortMsg lonMIDIHand, lonMsg
End If
End Sub

Public Property Let Instrument(ByVal NewValue As Integer)
'Validate the instrument value (0-127).
If Not NewValue < 0 And Not NewValue > 127 Then
    intInstrument = NewValue
    UpdateInstrument
End If
End Property

Public Property Get BaseNote() As Integer
BaseNote = intBaseNote
End Property

Public Property Let BaseNote(ByVal NewValue As Integer)
intBaseNote = NewValue
End Property

Public Property Get Volume() As Integer
Volume = intVolume
End Property

Public Property Let Volume(ByVal NewValue As Integer)
'Validate the volume value (0-127).
If Not NewValue < 0 And Not NewValue > 127 Then
    intVolume = NewValue
End If
End Property

Public Property Get Channel() As Integer
Channel = intChannel
End Property

Public Property Let Channel(ByVal NewValue As Integer)
intChannel = NewValue
End Property

Public Property Get DeviceOpen() As Boolean
DeviceOpen = bolDevOpen
End Property

Private Sub InitMIDIVars()
'Set default values.
intChannel = 1
intBaseNote = 36
intInstrument = 0
intVolume = MIDI_MAX_VOLUME
lonRet = 0
End Sub

'Closes the connection to a MIDI device (DeviceID).
'Returns: TRUE if successful.
'ALWAYS USE THIS BEFORE CLOSING YOUR PROGRAM TO PREVENT ERRORS IN WINDOWS.
Public Function CloseMIDI(Optional ByVal DeviceID As Long = -1, Optional ByRef RetVal As Long) As Boolean
If bolDevOpen = True Then 'Check if the device is open.
    lonRet = midiOutClose(lonMIDIHand) 'Close the device.
    
    If lonRet = 0 Then 'Check if it was successful.
        lonMIDIHand = 0 'Successful, reset MIDI handle variable.
        bolDevOpen = False 'Device closed.
    End If
    
    RetVal = lonRet 'Return the API function return value for debugging purposes.
    CloseMIDI = (lonRet = 0) 'Successful = (ReturnValue = 0).
End If
End Function

'Opens a connection to specified MIDI device (DeviceID).
Public Function ConnectMIDI(Optional ByVal DeviceID As Long = -1, Optional ByRef RetVal As Long) As Boolean
CloseMIDI lonDevID 'Close any previous connections before proceeding.
InitMIDIVars 'Initialize the MIDI variables.

'Open a connection to the MIDI device (DeviceID).
'The handle to the open device will be stored in: lonMIDIHand.
'This handle will be used when sending messages to the MIDI device via midiOutShortMsg().

'It is recommended to use the default MIDI device on all computers.
'The MIDI device has a DeviceID value of -1 and is called "MIDI Mapper".
'You can also use DeviceID: 0 which is Microsoft GW Synth or something.
lonRet = midiOutOpen(lonMIDIHand, DeviceID, 0, 0, 0)

If lonRet = 0 Then 'Check if successful.
    lonDevID = DeviceID 'Successful, store the current device ID being used.
    bolDevOpen = True 'Device is now open.
End If

RetVal = lonRet 'Return the API function return value for debugging purposes.
ConnectMIDI = (lonRet = 0) 'Successful = (ReturnValue = 0).
End Function

'"Safely" returns the upper-boundaries of a string array (without raising an error).
Private Function SafeUBoundStr(StringArray() As String) As Long
On Error GoTo ErrorHandler

SafeUBoundStr = UBound(StringArray())

ErrorHandler:
End Function

'Retrieves all of the available MIDI devices on the computer.
'Returns each device name as an item in the string array.
Public Function GetMIDIDevices() As String()
Dim strRet() As String, lonLoop As Long
Dim udtCap As MIDIOUTCAPS

'Get total number of devices.
lonDevCount = (midiOutGetNumDevs - 1)
'Prepare string for first MIDI device (MIDI Mapper).
ReDim strRet(0) As String
'Store the MIDI Mapper device in the new string array item.
strRet(0) = "MIDI Mapper"

'Loop through each device.
For lonLoop = 0 To lonDevCount
    'Get the current device (lonLoop)'s info.
    'The info is dumped into the udtCap data structure.
    lonRet = midiOutGetDevCaps(lonLoop, udtCap, Len(udtCap))
    'Create a new string array item for the current device.
    'Use keyword: Preserve so it won't delete the other array items.
    ReDim Preserve strRet(lonLoop + 1) As String
    'Store the current device's name in the new string array item.
    strRet(lonLoop + 1) = udtCap.szPname
Next lonLoop

'Loop finished, return final string array.
GetMIDIDevices = strRet()
End Function

'Class was just initialized/created.
Private Sub Class_Initialize()
InitMIDIVars 'Initialize default MIDI values.
End Sub

Private Sub Class_Terminate()
CloseMIDI lonDevID 'Make sure we disconnect from the MIDI device.
End Sub

'Simple pause routine that works around the midnight-bug.
'Used for note durations.
Private Sub PauseNow(ByVal Seconds As Long)
Dim varStart As Variant, varNow As Variant
Dim varDiff As Variant

varStart = Format(Now(), "mm/dd/yyyy hh:nn:ss AM/PM")

Do While varDiff < Seconds
    varNow = Format(Now(), "mm/dd/yyyy hh:nn:ss AM/PM")
    varDiff = DateDiff("s", varStart, varNow)
    DoEvents
Loop
End Sub

