VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDSSBuff8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------
' Author: Brian Black  BlackBsd@Mountain.Net
' DirectSoundSecondaryBuffer8 Class Source Code File
' License: Currently None
'----------------------------------------------------------------------
Option Explicit
Option Base 0


'DirectSound objects
Private DSBuffer As DirectSoundSecondaryBuffer8
Private ds3dBuffer As DirectSound3DBuffer8

Private SoundBuffer() As DirectSoundSecondaryBuffer8


'DirectSound Structures
Private Format As WAVEFORMATEX
Private dsBuf As DSBUFFERDESC
Private SoundCursors As DSCURSORS
Private vPos As D3DVECTOR
'Private dsbpn(0) As DSBPOSITIONNOTIFY

'Local Variables
Private bInterrupt As Boolean
Private bLooping As Boolean
Private bPositional As Boolean
Private BufferActive As Boolean
Private EndEvent As Long
Private BufferLoaded As Boolean

'For Writing To The Sound Buffer''''''''''''''
'Flag To Tell If I Have Wrote To The buffer yet
Private FirstWrite As Boolean
'This Tells Where The Data I Previously Wrote Is
Private LastDataPosition As Long

Private Sub Class_Initialize()
    
    bInterrupt = False
    bLooping = False
    bPositional = False
    
    If (DSBuffer Is Nothing) Then
        BufferActive = False
    Else
        BufferActive = True
    End If
    
    FirstWrite = True
    ReDim SoundBuffer(0)
    
End Sub

Private Sub Class_Terminate()
    StopBuffer
    Set ds3dBuffer = Nothing
    Set DSBuffer = Nothing
End Sub

Public Property Get DSBuffDesc() As DSBUFFERDESC
    DSBuffDesc = dsBuf
End Property

Public Property Let DSBuffDesc(ByRef NewBuffDesc As DSBUFFERDESC)
    dsBuf = NewBuffDesc
End Property

Public Property Get DSBObj() As DirectSoundSecondaryBuffer8
    Set DSBObj = DSBuffer
End Property

Public Property Set DSBObj(ByRef dssbNewValue As DirectSoundSecondaryBuffer8)
    Set DSBuffer = dssbNewValue
End Property

Public Property Get Interrupt() As Boolean
    Interrupt = bInterrupt
End Property

Public Property Let Interrupt(ByVal bNewValue As Boolean)
    bInterrupt = bNewValue
End Property

Public Property Get Looping() As Boolean
    Looping = bLooping
End Property

Public Property Let Looping(ByVal bNewValue As Boolean)
    bLooping = bNewValue
End Property

Public Property Get Positional() As Boolean
    Positional = bPositional
End Property

Public Property Let Positional(ByVal bNewValue As Boolean)
    bPositional = bNewValue
End Property

Public Property Get Frequency() As Long
    Frequency = DSBuffer.GetFrequency
End Property

Public Property Let Frequency(ByVal NewFrequency As Long)
    DSBuffer.SetFrequency NewFrequency
End Property

Public Property Get Volume() As Long
    Volume = DSBuffer.GetVolume
End Property

Public Property Let Volume(ByVal NewVolume As Long)
    DSBuffer.SetVolume NewVolume
End Property

Public Property Get Pan() As Long
    Pan = DSBuffer.GetPan
End Property

Public Property Let Pan(ByVal NewPan As Long)
    DSBuffer.SetPan NewPan
End Property

Public Property Get WavFormat() As WAVEFORMATEX
    DSBuffer.GetFormat WavFormat
End Property

Public Property Get BufferCurrentlyActive() As Boolean
    BufferCurrentlyActive = BufferActive
End Property

Public Property Get BufferCurrentlyLoaded() As Boolean
    BufferCurrentlyLoaded = BufferLoaded
End Property

Public Property Get Status() As CONST_DSBSTATUSFLAGS
    If Not (DSBuffer Is Nothing) Then
        Status = DSBuffer.GetStatus
    End If
End Property

Public Property Get Channels() As Long
    DSBuffer.GetFormat Format
    Channels = Format.nChannels
End Property

Public Property Get Samples() As Long
    DSBuffer.GetFormat Format
    Samples = Format.lSamplesPerSec
End Property

Public Property Get BitsPerSample() As Long
    DSBuffer.GetFormat Format
    BitsPerSample = Format.nBitsPerSample
End Property

Public Property Get Size() As Long
    DSBuffer.GetFormat Format
    Size = Format.nSize
End Property

Public Sub SetPosition(ByVal X As Long, ByVal Y As Long, ByVal z As Long)
    vPos.X = X
    vPos.Y = Y
    vPos.z = z
End Sub

Public Sub CreateBuffer(ByRef objDS As clsDirectSound8)
    If bPositional Then
        'Set Buffer Parameters To Accept Positional Data
        dsBuf.lFlags = DSBCAPS_CTRL3D Or DSBCAPS_CTRLVOLUME
        dsBuf.guid3DAlgorithm = GUID_DS3DALG_HRTF_LIGHT
    Else
        'Set Buffer Parameter For A Default Buffer
'        dsBuf.lFlags = DSBCAPS_STATIC Or DSBCAPS_CTRLFREQUENCY Or _
            DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or _
            DSBCAPS_GETCURRENTPOSITION2 Or DSBCAPS_GLOBALFOCUS
    End If
    
    'Create The Buffer
    Set DSBuffer = Nothing
    Set DSBuffer = objDS.DSObj.CreateSoundBuffer(dsBuf)
    
    'Set Positional Sound Environment
    If bPositional Then
        Set ds3dBuffer = Nothing
        Set ds3dBuffer = DSBuffer.GetDirectSound3DBuffer
        ds3dBuffer.SetConeAngles DS3D_MINCONEANGLE, 100, DS3D_IMMEDIATE
        ds3dBuffer.SetConeOutsideVolume -400, DS3D_IMMEDIATE
    End If
End Sub

Public Sub LoadSound(ByRef objDS As clsDirectSound8, ByVal sFilename As String)
    If bPositional Then
        'set buffer parameters to accept positional data
        dsBuf.lFlags = DSBCAPS_CTRL3D Or DSBCAPS_CTRLVOLUME
        dsBuf.guid3DAlgorithm = GUID_DS3DALG_HRTF_LIGHT
    Else
        'set buffer parameter for normal playback
        dsBuf.lFlags = DSBCAPS_STATIC Or DSBCAPS_CTRLFREQUENCY Or _
            DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or _
            DSBCAPS_GETCURRENTPOSITION2 Or DSBCAPS_GLOBALFOCUS
    End If
    
    'load wave file into DirectSound buffer
    Set DSBuffer = Nothing
    Set DSBuffer = objDS.DSObj.CreateSoundBufferFromFile(sFilename, dsBuf)
    
    If Not (DSBuffer Is Nothing) Then
        'buffer has the sound loaded into it
        BufferLoaded = True
    Else
        BufferLoaded = False
    End If
    
    'set positional sound environment
    If bPositional Then
        Set ds3dBuffer = Nothing
        Set ds3dBuffer = DSBuffer.GetDirectSound3DBuffer
        ds3dBuffer.SetConeAngles DS3D_MINCONEANGLE, 100, DS3D_IMMEDIATE
        ds3dBuffer.SetConeOutsideVolume -400, DS3D_IMMEDIATE
    End If
End Sub

Public Sub PlayAural(ByRef objDS As clsDirectSound8, ByRef ByteBuffer() As Byte, ByRef DataSize As Long)
'Static SoundBuffer() As DirectSoundSecondaryBuffer8
Dim TempDSBuf As DSBUFFERDESC
Dim ArrayFull As Boolean
Dim i As Integer

    TempDSBuf = dsBuf
    TempDSBuf.lBufferBytes = UBound(ByteBuffer()) + 1
    
    
    'Check To See If The Buffer is Being Used
    For i = 0 To UBound(SoundBuffer())
        If (SoundBuffer(i) Is Nothing) Then
            
            'Set SoundBuffer(i) = Nothing
            Set SoundBuffer(i) = objDS.DSObj.CreateSoundBuffer(TempDSBuf)
           
            'Write Data For Playback At Spot 0
            SoundBuffer(i).WriteBuffer 0, DataSize, ByteBuffer(0), DSBLOCK_DEFAULT 'DSBLOCK_ENTIREBUFFER 'DSBLOCK_DEFAULT
        
            'Now Play The Sound
            SoundBuffer(i).Play DSBPLAY_DEFAULT
            ArrayFull = False
            Exit For
        ElseIf ((SoundBuffer(i).GetStatus And DSBSTATUS_PLAYING) <> DSBSTATUS_PLAYING) Then
            
            Set SoundBuffer(i) = Nothing
            Set SoundBuffer(i) = objDS.DSObj.CreateSoundBuffer(TempDSBuf)
            
            'Write Data For Playback At Spot 0
            SoundBuffer(i).WriteBuffer 0, DataSize, ByteBuffer(0), DSBLOCK_DEFAULT 'DSBLOCK_ENTIREBUFFER 'DSBLOCK_DEFAULT
        
            'Now Play The Sound
            SoundBuffer(i).Play DSBPLAY_DEFAULT
            ArrayFull = False
            Exit For
        Else
            'buffer is playing
            ArrayFull = True
        End If
        
    Next i
    
    
    'If Array Is Full Then Redim The Array To Be 1 Larger
    If (ArrayFull) Then
        
        ReDim SoundBuffer(UBound(SoundBuffer) + 1)
        
        Debug.Print "Number Of Sound Buffers"; UBound(SoundBuffer) + 1
        
        Set SoundBuffer(UBound(SoundBuffer)) = Nothing
        Set SoundBuffer(UBound(SoundBuffer)) = objDS.DSObj.CreateSoundBuffer(TempDSBuf)
           
        'Write Data For Playback At Spot 0
        SoundBuffer(UBound(SoundBuffer)).WriteBuffer 0, DataSize, ByteBuffer(0), DSBLOCK_DEFAULT
        
        'Now Play The Sound
        SoundBuffer(UBound(SoundBuffer)).Play DSBPLAY_DEFAULT
    End If
        
End Sub

Public Sub PlayBuffer()
    'Stop Currently Playing Waves?
    If bInterrupt Then
        StopBuffer
    End If
    
    'Set Position Of Sound Source
    If bPositional Then
        ds3dBuffer.SetPosition vPos.X, vPos.Y, vPos.z, DS3D_IMMEDIATE
    End If
    
    'Loop The Sound?
    If bLooping Then
        DSBuffer.Play DSBPLAY_LOOPING
        BufferActive = True
    Else
        DSBuffer.Play DSBPLAY_DEFAULT
        BufferActive = True
    End If
End Sub

Public Sub StopBuffer()
    DSBuffer.Stop
    BufferActive = False
    DSBuffer.SetCurrentPosition 0
End Sub

Public Sub WriteBuff(ByRef ByteBuffer() As Byte)
Dim DS As Long
'''Write Data
'What i need to do is make sure that i can write the the area of the buffer
'and if so then do it.

    'Get A DataSize
    DS = (UBound(ByteBuffer) - 1)
    
    'Get Current Cursor Positions
    DSBuffer.GetCurrentPosition SoundCursors
    
    If ((LastDataPosition >= SoundCursors.lPlay) And _
            (LastDataPosition <= SoundCursors.lWrite)) Then
        'This Means That The Last Sound Position Is In Between .play and
        '.write cursors and cant be touched. This Means "Bleeding IS Occuring!!"
'        Debug.Print "Bleeding IS Occuring!!"
    
    ElseIf ((LastDataPosition < SoundCursors.lPlay) And _
            (LastDataPosition < SoundCursors.lWrite)) Then
        'This Means That The Last Sound Position Is In Behind .play and
        '.write cursors. This Means "Bleeding HAS Occured!!"
'        Debug.Print "Bleeding HAS Occured!!"
    Else
        'This Means The Last Positions
    
    End If
    
    If Not (FirstWrite) Then
        'Write Data For Playback At The End Of LastDataPosition
        DSBuffer.WriteBuffer LastDataPosition, DS, Buffer(0), DSBLOCK_DEFAULT
                
        'Set The New Sound Cursor Position
        If (DS + LastDataPosition - dsBuf.lBufferBytes) >= 0 Then
            'The CursorPosition Has Wraped Around So ..
            LastDataPosition = DS + LastDataPosition - dsBuf.lBufferBytes
        Else
            'The Cursor Has Not Wraped Around
            LastDataPosition = (LastDataPosition + DS)
        End If
    Else
        'Set The Flag
        FirstWrite = False
        
        'Get Current Cursor Positions
        DSBuffer.GetCurrentPosition SoundCursors
        
        'Write Data For Playback At The Current Write Spot
        DSBuffer.WriteBuffer SoundCursors.lWrite, DS, Buffer(0), DSBLOCK_DEFAULT
        
        'Set The New Sound Cursor Position
        If (DS + SoundCursors.lWrite - dsBuf.lBufferBytes) >= 0 Then
            'The CursorPosition Has Wraped Around So ..
            LastDataPosition = DS + SoundCursors.lWrite - dsBuf.lBufferBytes
        Else
            'The Cursor Has Not Wraped Around
            LastDataPosition = (SoundCursors.lWrite + DS)
        End If
    
    End If
End Sub

Public Sub SetCurrentPosition(ByVal NewPosition As Long)
    DSBuffer.SetCurrentPosition NewPosition
End Sub

Public Sub GetCurrentPosition(ByRef Cursors As DSCURSORS)
    Cursors = DSBuffer.GetCurrentPosition
End Sub

'Public Sub SetNotifyPositions(ByVal nElements, ByRef PSA() As DSBPOSITIONNOTIFY)
'    DSBuffer.SetNotificationPositions nElements, PSA()
'End Sub
