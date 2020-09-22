VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Easy Guitar Tuner"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7005
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin EasyGuitarTuner.XPButton cmdDefaults 
      Height          =   375
      Left            =   5400
      TabIndex        =   20
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Load Defaults"
      ForeColor       =   -2147483647
      BeginProperty FontType {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.StatusBar objStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5175
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDur 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   19
      Text            =   "2"
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox picString 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Index           =   5
      Left            =   3000
      ScaleHeight     =   30
      ScaleWidth      =   2415
      TabIndex        =   17
      Top             =   4800
      Width           =   2415
   End
   Begin VB.PictureBox picString 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Index           =   4
      Left            =   3000
      ScaleHeight     =   30
      ScaleWidth      =   2655
      TabIndex        =   16
      Top             =   4200
      Width           =   2655
   End
   Begin VB.PictureBox picString 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Index           =   3
      Left            =   3000
      ScaleHeight     =   30
      ScaleWidth      =   3015
      TabIndex        =   15
      Top             =   3600
      Width           =   3015
   End
   Begin VB.PictureBox picString 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Index           =   2
      Left            =   3000
      ScaleHeight     =   30
      ScaleWidth      =   3255
      TabIndex        =   14
      Top             =   3000
      Width           =   3255
   End
   Begin VB.PictureBox picString 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Index           =   1
      Left            =   3000
      ScaleHeight     =   30
      ScaleWidth      =   3495
      TabIndex        =   13
      Top             =   2400
      Width           =   3495
   End
   Begin VB.PictureBox picString 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Index           =   0
      Left            =   3000
      ScaleHeight     =   30
      ScaleWidth      =   3735
      TabIndex        =   12
      Top             =   1800
      Width           =   3735
   End
   Begin EasyGuitarTuner.XPButton cmdNote 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      Caption         =   "     E-1 (Thinnest String)"
      ForeColor       =   -2147483647
      BeginProperty FontType {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmMain.frx":0A02
      IconMask        =   "frmMain.frx":0D54
      CaptionAlignment=   0
   End
   Begin VB.ComboBox cmbTune 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   2775
   End
   Begin VB.ComboBox cmbInst 
      Height          =   315
      ItemData        =   "frmMain.frx":10A6
      Left            =   1920
      List            =   "frmMain.frx":10A8
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin EasyGuitarTuner.XPButton cmdNote 
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      Caption         =   "     B-2 (Second String)"
      ForeColor       =   -2147483647
      BeginProperty FontType {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmMain.frx":10AA
      IconMask        =   "frmMain.frx":13FC
      CaptionAlignment=   0
   End
   Begin EasyGuitarTuner.XPButton cmdNote 
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      Caption         =   "     G-3 (Third String)"
      ForeColor       =   -2147483647
      BeginProperty FontType {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmMain.frx":174E
      IconMask        =   "frmMain.frx":1AA0
      CaptionAlignment=   0
   End
   Begin EasyGuitarTuner.XPButton cmdNote 
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      Caption         =   "     D-4 (Fourth String)"
      ForeColor       =   -2147483647
      BeginProperty FontType {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmMain.frx":1DF2
      IconMask        =   "frmMain.frx":2144
      CaptionAlignment=   0
   End
   Begin EasyGuitarTuner.XPButton cmdNote 
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      Caption         =   "     A-5 (Fifth String)"
      ForeColor       =   -2147483647
      BeginProperty FontType {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmMain.frx":2496
      IconMask        =   "frmMain.frx":27E8
      CaptionAlignment=   0
   End
   Begin EasyGuitarTuner.XPButton cmdNote 
      Height          =   495
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      Caption         =   "     E-6 (Thickest String)"
      ForeColor       =   -2147483647
      BeginProperty FontType {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmMain.frx":2B3A
      IconMask        =   "frmMain.frx":2E8C
      CaptionAlignment=   0
   End
   Begin VB.PictureBox picNeck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   4680
      Picture         =   "frmMain.frx":31DE
      ScaleHeight     =   3375
      ScaleWidth      =   2175
      TabIndex        =   5
      Top             =   1680
      Width           =   2175
   End
   Begin EasyGuitarTuner.XPButton cmdHelp 
      Height          =   375
      Left            =   5400
      TabIndex        =   21
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Help"
      ForeColor       =   -2147483647
      BeginProperty FontType {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmMain.frx":A5C4
      IconMask        =   "frmMain.frx":A916
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      X1              =   120
      X2              =   6840
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      Caption         =   "Note duration:          seconds."
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   18
      Top             =   840
      Width           =   2580
   End
   Begin VB.Image imgDisplay 
      Height          =   240
      Index           =   2
      Left            =   120
      Picture         =   "frmMain.frx":AC68
      Top             =   840
      Width           =   240
   End
   Begin VB.Image imgDisplay 
      Height          =   240
      Index           =   1
      Left            =   120
      Picture         =   "frmMain.frx":B66A
      Top             =   480
      Width           =   240
   End
   Begin VB.Image imgDisplay 
      Height          =   240
      Index           =   0
      Left            =   120
      Picture         =   "frmMain.frx":C06C
      Top             =   120
      Width           =   240
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      Caption         =   "Desired tuning:"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   1320
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      Caption         =   "Instrument:"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   1020
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type EGT_TUNE
    strName As String
    intNotes(5) As Integer
End Type

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const TUNING_COUNT As Integer = 36

Private Const APPNAME As String = "Easy Guitar Tuner"
Private Const SEC As String = "Main"

Private udtTuning(1 To TUNING_COUNT) As EGT_TUNE

Private objMIDI As New MIDILib

Private intCurNotes(5) As Integer

Private Sub SaveFormSettings()
SaveSetting APPNAME, SEC, "Left", Me.Left
SaveSetting APPNAME, SEC, "Top", Me.Top
SaveSetting APPNAME, SEC, "WS", Me.WindowState
SaveSetting APPNAME, SEC, "Instrument", cmbInst.Text
SaveSetting APPNAME, SEC, "Tuning", cmbTune.Text
SaveSetting APPNAME, SEC, "Dur", txtDur.Text
End Sub

Private Sub ReadFormSettings()
On Error Resume Next

Dim strLeft As String, strTop As String
Dim strWS As String, strInst As String
Dim strTune As String, strDur As String

strLeft = GetSetting(APPNAME, SEC, "Left", "")
strTop = GetSetting(APPNAME, SEC, "Top", "")
strWS = GetSetting(APPNAME, SEC, "WS", "")
strInst = GetSetting(APPNAME, SEC, "Instrument", "")
strTune = GetSetting(APPNAME, SEC, "Tuning", "")
strDur = GetSetting(APPNAME, SEC, "Dur", "")

If Len(strLeft) > 0 And IsNumeric(strLeft) = True Then
    Me.Left = Val(strLeft)
End If

If Len(strTop) > 0 And IsNumeric(strTop) = True Then
    Me.Top = Val(strTop)
End If

If Len(strWS) > 0 And IsNumeric(strWS) = True Then
    
    If Val(strWS) = 0 Or Val(strWS) = 1 Then
        Me.WindowState = Val(strWS)
    End If

End If

If Len(strInst) > 0 Then
    cmbInst.Text = strInst
End If

If Len(strTune) > 0 Then
    cmbTune.Text = strTune
End If

If Len(strDur) > 0 And IsNumeric(strDur) = True Then
    
    If Len(strDur) < 2 And Val(strDur) > 0 Then
        txtDur.Text = strDur
    End If

End If
End Sub

Private Sub cmbInst_Click()
Dim lonInst As Long

lonInst = ValInstrument(cmbInst.Text)
objMIDI.Instrument = lonInst
End Sub

Private Sub cmbTune_Click()
Dim lonUDTInd As Long

lonUDTInd = (cmbTune.ListIndex + 1)

If lonUDTInd > 0 Then
    
    With udtTuning(lonUDTInd)
        intCurNotes(0) = .intNotes(0)
        intCurNotes(1) = .intNotes(1)
        intCurNotes(2) = .intNotes(2)
        intCurNotes(3) = .intNotes(3)
        intCurNotes(4) = .intNotes(4)
        intCurNotes(5) = .intNotes(5)
    End With

End If
End Sub

Private Sub cmdDefaults_Click()
Dim msgRet As VbMsgBoxResult

msgRet = MsgBox("Are you sure you want to load the default settings?", vbQuestion + vbYesNo, "Load Defaults")

If msgRet = vbYes Then
    On Error Resume Next
    
    cmbInst.Text = cmbInst.List(1)
    cmbTune.Text = cmbTune.List(0)
    txtDur.Text = "2"
    
    cmbInst_Click
    cmbTune_Click
End If
End Sub

Private Sub cmdHelp_Click()
LaunchHelp
End Sub

Private Sub cmdNote_Click(Index As Integer)
If Len(txtDur.Text) = 0 Then
    MsgBox "Enter a note duration in seconds", vbCritical, "Note Duration Required"
    txtDur.SetFocus
    Exit Sub
ElseIf IsNumeric(txtDur.Text) = False Then
    MsgBox "Enter a numeric value for the note duration", vbCritical, "Invalid Note Duration"
    txtDur.SetFocus
    txtDur.SelStart = 0
    txtDur.SelLength = Len(txtDur.Text)
    Exit Sub
End If

objMIDI.PlayNote intCurNotes(Index), Val(txtDur.Text)
End Sub

Private Sub Form_Load()
Dim bolRet As Boolean

HCenter Me, Screen
VCenter Me, Screen
LoadTunings
ListTunings
LoadInstruments
ReadFormSettings

bolRet = objMIDI.ConnectMIDI

If bolRet = False Then
    MsgBox "Error connecting to MIDI Mapper device", vbCritical, "Critical Error"
    Exit Sub
End If

cmbInst_Click
cmbTune_Click
objMIDI.BaseNote = 0
End Sub

Private Sub ResetTunings()
Dim intLoop As Integer, intNote As Integer

For intLoop = 1 To TUNING_COUNT
    
    With udtTuning(intLoop)
        
        For intNote = 0 To 5
            .intNotes(intNote) = 0
        Next intNote
        
        .strName = ""
    End With

Next intLoop
End Sub

Private Sub LoadInstruments()
With cmbInst
    .Clear
    .AddItem "Acoustic Guitar (Nylon)"
    .AddItem "Acoustic Guitar (Steel)"
    .AddItem "Electric Guitar (Jazz)"
    .AddItem "Electric Guitar (Clean)"
    .AddItem "Electric Guitar (Muted)"
    .AddItem "Overdriven Guitar"
    .AddItem "Distortion Guitar"
    On Error Resume Next
    .Text = .List(1)
End With
End Sub

Private Sub LaunchHelp()
On Error Resume Next

Dim bytData() As Byte, lonFF As Long
Dim strPath As String

strPath = App.Path & "\Help.html"
Kill strPath

bytData() = LoadResData("HELP.HTML", "CUSTOM")
lonFF = FreeFile

Open strPath For Binary Access Write As #lonFF
    Put #lonFF, , bytData()
Close #lonFF

strPath = App.Path & "\Fig1.0.gif"
Kill strPath

bytData() = LoadResData("FIG10", "CUSTOM")
lonFF = FreeFile

Open strPath For Binary Access Write As #lonFF
    Put #lonFF, , bytData()
Close #lonFF

strPath = App.Path & "\Fig1.1.gif"
Kill strPath

bytData() = LoadResData("FIG11", "CUSTOM")
lonFF = FreeFile

Open strPath For Binary Access Write As #lonFF
    Put #lonFF, , bytData()
Close #lonFF

DoEvents

ShellExecute Me.hwnd, "open", App.Path & "\Help.html", vbNullString, vbNullString, 1
End Sub

Private Function ValInstrument(ByVal InstName As String) As Long
Dim strLCase As String

strLCase = LCase$(InstName)

Select Case strLCase
    
    Case "acoustic guitar (nylon)"
        ValInstrument = 24
    Case "acoustic guitar (steel)"
        ValInstrument = 25
    Case "electric guitar (jazz)"
        ValInstrument = 26
    Case "electric guitar (clean)"
        ValInstrument = 27
    Case "electric guitar (muted)"
        ValInstrument = 28
    Case "overdriven guitar"
        ValInstrument = 29
    Case "distortion guitar"
        ValInstrument = 30
        
End Select
End Function

Private Sub ListTunings()
Dim intLoop As Integer

With cmbTune
    .Clear
    
    For intLoop = 1 To TUNING_COUNT
        .AddItem udtTuning(intLoop).strName
    Next intLoop
    
    On Error Resume Next
    
    .Text = .List(0)

End With
End Sub

Private Sub LoadTunings()
Dim intLoop As Integer

ResetTunings

For intLoop = 1 To TUNING_COUNT
    
    With udtTuning(intLoop)
        
        Select Case intLoop
            
            Case 1 'Standard tuning.
                .intNotes(0) = 64
                .intNotes(1) = 59
                .intNotes(2) = 55
                .intNotes(3) = 50
                .intNotes(4) = 45
                .intNotes(5) = 40
                
                .strName = "Standard Tuning"
            
            Case 2 'Bass-Standard.
                .intNotes(0) = 47
                .intNotes(1) = 43
                .intNotes(2) = 38
                .intNotes(3) = 33
                .intNotes(4) = 28
                .intNotes(5) = 35
                .strName = "Bass Standard"
            
            Case 3 'Dropped-D.
                .intNotes(0) = 64
                .intNotes(1) = 59
                .intNotes(2) = 55
                .intNotes(3) = 50
                .intNotes(4) = 45
                .intNotes(5) = 38
                .strName = "Dropped-D"
            
            Case 4 'Double Dropped-D.
                .intNotes(0) = 64
                .intNotes(1) = 59
                .intNotes(2) = 55
                .intNotes(3) = 50
                .intNotes(4) = 45
                .intNotes(5) = 38
                .strName = "Double Dropped-D"
            
            Case 5 'Down 1/2 Step.
                .intNotes(0) = 63
                .intNotes(1) = 58
                .intNotes(2) = 54
                .intNotes(3) = 49
                .intNotes(4) = 44
                .intNotes(5) = 39
                .strName = "Down ½ Step"
            
            Case 6 'Down 1 (Whole) Step.
                .intNotes(0) = 62
                .intNotes(1) = 57
                .intNotes(2) = 53
                .intNotes(3) = 48
                .intNotes(4) = 43
                .intNotes(5) = 38
                .strName = "Down 1 (Whole) Step"
            
            Case 7 'Down 1 1/2 Steps.
                .intNotes(0) = 61
                .intNotes(1) = 56
                .intNotes(2) = 52
                .intNotes(3) = 47
                .intNotes(4) = 42
                .intNotes(5) = 37
                .strName = "Down 1½ Steps"
            
            Case 8 'Down 2 Steps.
                .intNotes(0) = 60
                .intNotes(1) = 55
                .intNotes(2) = 51
                .intNotes(3) = 46
                .intNotes(4) = 41
                .intNotes(5) = 36
                .strName = "Down 2 Steps"
            
            Case 9 'Open D.
                .intNotes(0) = 62
                .intNotes(1) = 57
                .intNotes(2) = 54
                .intNotes(3) = 50
                .intNotes(4) = 45
                .intNotes(5) = 38
                .strName = "Open D"
            
            Case 10 'Open G.
                .intNotes(0) = 62
                .intNotes(1) = 59
                .intNotes(2) = 55
                .intNotes(3) = 50
                .intNotes(4) = 43
                .intNotes(5) = 38
                .strName = "Open G"
            
            Case 11 'Open C (Type 1).
                .intNotes(0) = 64
                .intNotes(1) = 60
                .intNotes(2) = 55
                .intNotes(3) = 48
                .intNotes(4) = 43
                .intNotes(5) = 36
                .strName = "Open C (Type 1)"
            
            Case 12 'Open C (Type 2).
                .intNotes(0) = 64
                .intNotes(1) = 60
                .intNotes(2) = 55
                .intNotes(3) = 52
                .intNotes(4) = 43
                .intNotes(5) = 36
                .strName = "Open C (Type 2)"
            
            Case 13 'Open E.
                .intNotes(0) = 64
                .intNotes(1) = 59
                .intNotes(2) = 56
                .intNotes(3) = 52
                .intNotes(4) = 47
                .intNotes(5) = 40
                .strName = "Open E"
            
            Case 14 'Open A.
                .intNotes(0) = 64
                .intNotes(1) = 61
                .intNotes(2) = 57
                .intNotes(3) = 52
                .intNotes(4) = 49
                .intNotes(5) = 45
                .strName = "Open A"
            
            Case 15 'Cross-Note.
                .intNotes(0) = 62
                .intNotes(1) = 57
                .intNotes(2) = 53
                .intNotes(3) = 50
                .intNotes(4) = 45
                .intNotes(5) = 38
                .strName = "Cross-Note"
            
            Case 16 'D Modal (Type 1).
                .intNotes(0) = 62
                .intNotes(1) = 57
                .intNotes(2) = 54
                .intNotes(3) = 50
                .intNotes(4) = 45
                .intNotes(5) = 38
                .strName = "D Modal (Type 1)"
            
            Case 17 'D Modal (Type 2).
                .intNotes(0) = 64
                .intNotes(1) = 59
                .intNotes(2) = 55
                .intNotes(3) = 50
                .intNotes(4) = 43
                .intNotes(5) = 38
                .strName = "D Modal (Type 2)"
            
            Case 18 'D Modal (Type 3).
                .intNotes(0) = 64
                .intNotes(1) = 59
                .intNotes(2) = 55
                .intNotes(3) = 50
                .intNotes(4) = 43
                .intNotes(5) = 36
                .strName = "D Modal (Type 3)"
            
            Case 19 'D Modal (Type 4).
                .intNotes(0) = 62
                .intNotes(1) = 57
                .intNotes(2) = 55
                .intNotes(3) = 50
                .intNotes(4) = 45
                .intNotes(5) = 38
                .strName = "D Modal (Type 4)"
            
            Case 20 'Fourths.
                .intNotes(0) = 65
                .intNotes(1) = 60
                .intNotes(2) = 55
                .intNotes(3) = 50
                .intNotes(4) = 45
                .intNotes(5) = 40
                .strName = "Fourths"
            
            Case 21 'Lute (Type 1).
                .intNotes(0) = 64
                .intNotes(1) = 59
                .intNotes(2) = 54
                .intNotes(3) = 50
                .intNotes(4) = 45
                .intNotes(5) = 40
                .strName = "Lute (Type 1)"
            
            Case 22 'Lute (Type 2).
                .intNotes(0) = 64
                .intNotes(1) = 57
                .intNotes(2) = 54
                .intNotes(3) = 50
                .intNotes(4) = 45
                .intNotes(5) = 40
                .strName = "Lute (Type 2)"
                
            Case 23 'Big City.
                .intNotes(0) = 69
                .intNotes(1) = 57
                .intNotes(2) = 54
                .intNotes(3) = 50
                .intNotes(4) = 45
                .intNotes(5) = 38
                .strName = "Big City"
                
            Case 24 'D Wahine.
                .intNotes(0) = 62
                .intNotes(1) = 59
                .intNotes(2) = 54
                .intNotes(3) = 50
                .intNotes(4) = 45
                .intNotes(5) = 38
                .strName = "D Wahine"
                
            Case 25 'D Minor.
                .intNotes(0) = 62
                .intNotes(1) = 57
                .intNotes(2) = 53
                .intNotes(3) = 50
                .intNotes(4) = 45
                .intNotes(5) = 38
                .strName = "D Minor"
            
            Case 26 'D Modal (Type 5).
                .intNotes(0) = 62
                .intNotes(1) = 57
                .intNotes(2) = 50
                .intNotes(3) = 50
                .intNotes(4) = 45
                .intNotes(5) = 38
                .strName = "D Modal (Type 5)"
            
            Case 27 'G6.
                .intNotes(0) = 64
                .intNotes(1) = 59
                .intNotes(2) = 55
                .intNotes(3) = 50
                .intNotes(4) = 43
                .intNotes(5) = 38
                .strName = "G6"
            
            Case 28 'G Minor.
                .intNotes(0) = 62
                .intNotes(1) = 58
                .intNotes(2) = 55
                .intNotes(3) = 50
                .intNotes(4) = 43
                .intNotes(5) = 38
                .strName = "G Minor"
            
            Case 29 'C6 (Type 1)."
                .intNotes(0) = 64
                .intNotes(1) = 59
                .intNotes(2) = 55
                .intNotes(3) = 48
                .intNotes(4) = 43
                .intNotes(5) = 36
                .strName = "C6 (Type 1)"
            
            Case 30 'Bron-Y-Aur.
                .intNotes(0) = 64
                .intNotes(1) = 60
                .intNotes(2) = 55
                .intNotes(3) = 48
                .intNotes(4) = 45
                .intNotes(5) = 36
                .strName = "Bron-Y-Aur"
            
            Case 31 'Parvardigar.
                .intNotes(0) = 62
                .intNotes(1) = 60
                .intNotes(2) = 55
                .intNotes(3) = 48
                .intNotes(4) = 43
                .intNotes(5) = 36
                .strName = "Parvardigar"
            
            Case 32 'Bruce Palmer Modal.
                .intNotes(0) = 64
                .intNotes(1) = 59
                .intNotes(2) = 52
                .intNotes(3) = 52
                .intNotes(4) = 47
                .intNotes(5) = 40
                .strName = "Bruce Palmer Modal"
            
            Case 33 'New Standard.
                .intNotes(0) = 67
                .intNotes(1) = 64
                .intNotes(2) = 57
                .intNotes(3) = 50
                .intNotes(4) = 43
                .intNotes(5) = 36
                .strName = "New Standard"
            
            Case 34 'Low C.
                .intNotes(0) = 62
                .intNotes(1) = 57
                .intNotes(2) = 55
                .intNotes(3) = 50
                .intNotes(4) = 43
                .intNotes(5) = 36
                .strName = "Low C"
            
            Case 35 'G Modal.
                .intNotes(0) = 62
                .intNotes(1) = 60
                .intNotes(2) = 55
                .intNotes(3) = 50
                .intNotes(4) = 43
                .intNotes(5) = 38
                .strName = "G Modal"
            
            Case 36 'A Minor.
                .intNotes(0) = 64
                .intNotes(1) = 60
                .intNotes(2) = 57
                .intNotes(3) = 52
                .intNotes(4) = 48
                .intNotes(5) = 45
                .strName = "A Minor"
                
        End Select
    
    End With

Next intLoop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SaveFormSettings
objMIDI.CloseMIDI
End Sub

Private Sub txtDur_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr$(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub
