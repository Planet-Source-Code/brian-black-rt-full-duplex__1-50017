VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDSCBuff8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------
'/*
' * License:  FreeBsd Style
' * Author: Brian Black  BlackBsd@Mountain.Net
' * DirectSoundCaptureBuffer8 Class Source Code File
' * All rights reserved.
' *
' * Redistribution and use in source and binary forms, with or without
' * modification, are permitted provided that the following conditions
' * are met:
' * 1. Redistributions of source code must retain the above copyright
' *    notice, this list of conditions and the following disclaimer.
' * 2. Redistributions in binary form must reproduce the above copyright
' *    notice, this list of conditions and the following disclaimer in the
' *    documentation and/or other materials provided with the distribution.
' *
' * THIS SOFTWARE IS PROVIDED BY THE AUTHOR AND CONTRIBUTORS ``AS IS'' AND
' * ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
' * IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
' * ARE DISCLAIMED.  IN NO EVENT SHALL THE AUTHOR OR CONTRIBUTORS BE LIABLE
' * FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
' * DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS
' * OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
' * HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHERIN CONTRACT, STRICT
' * LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY
' * OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THEPOSSIBILITY OF
' * SUCH DAMAGE.
' *
' * $FreeBSD$
' */
' The objective i am trying to achieve is to have the ability to control a
' direct sound 8 capture buffer and any properties associated.
'----------------------------------------------------------------------

Option Explicit
Option Base 0

'Class Functions'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This Is A Windowz Function To Zero Out Areas Of Memory                                 '
Private Declare Sub RtlZeroMemory Lib "kernel32" (dest As Any, ByVal cbSize As Long)    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Class Objects Used Internal'''''''''''''''''''''''''''''''''''''
'DirectSoundCapture Buffer Objects                              '
Private DSCBuffer As DirectSoundCaptureBuffer8                  '
'DirectX Object                                                 '
Private objDX As New DirectX8                                   '
'frmDX8 Object                                                  '
Private WithEvents frmDX As frmDX8                              '
Attribute frmDX.VB_VarHelpID = -1
'The Handle To The frmDX8 Object                                '
Private XEventHandle As Long                                    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Class Structures Used Internal''''''''''''''''''''''''''''''''''
'DirectSoundCapture Structures                                  '
Private BuffWavFormat As WAVEFORMATEX                           '
Private DSCBuffDesc As DSCBUFFERDESC                            '
Private CaptureCursors As DSCURSORS                             '
'Position Notify Structure, used with frmDX Object              '
Private NewNotifyPosition(2) As DSBPOSITIONNOTIFY               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''Local Variables''''''''''''''''''''''''''''''''''''''''''''''''
'Could Be Used To Tell When To Interupt Capture Or PlayBack     '
Private bInterrupt As Boolean                                   '
'Used To Tell Weather To Loop The Playing/Capturing Buffer      '
Private bLooping As Boolean                                     '
'Used To Tell Weather The Buffer Is Currently Playing/Capturing '
'Private BufferActive As Boolean                                 '
'Flag For The First Time Reading The Capture Buffer             '
Private FirstRead As Boolean                                    '
'Value Telling Where The Current Write Cursor Position Is       '
Private CurrentWriteCursorPosition As Long                      '
'Value Telling Where The Last Write Cursor Position Is          '
Private LastWriteCursorPosition As Long                         '
'Value For The Current Data Size When Read From The Buffer      '
Private CurrentReadDataSize As Long                             '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Initialize()
    'Set Up Some General Properties
    bInterrupt = False
    bLooping = False
    FirstRead = True
    
    'Now The Buffer Is Not Set Up Yet So Just Set The Value.
'    BufferActive = False
    
    'Set Up The Variables That Handle The Set Notification Positions
    Set frmDX = New frmDX8
    
    'Create The Handle That Will Be Used
    XEventHandle = objDX.CreateEvent(frmDX)
    
End Sub

Private Sub Class_Terminate()
    'Make sure that the buffer is not playing
    If BufferCurrentlyActive Then
        StopCapture
    End If
    
    'Release the resources used by the DirectSoundCaptureBuffer, DSCBuffer.
    Set DSCBuffer = Nothing
    
    'Now Destroy The Event Handle That Was Created In The Class_Init() Event
    objDX.DestroyEvent (XEventHandle)
    
    'Release the resources, if any used, for the DirectX object, objDX.
    Set objDX = Nothing
    
    'Now Release The Resources For The Form
    Set frmDX = Nothing
End Sub

Public Property Get CaptureBuffDesc() As DSCBUFFERDESC
    'This Property Sets The Internal BufferDesc structure
    'equal to the return value.
    CaptureBuffDesc = DSCBuffDesc
End Property

Public Property Let CaptureBuffDesc(ByRef NewBuffDesc As DSCBUFFERDESC)
    'This Property Sets The Internal BufferDesc structure
    'To The Structure That Gets Passed Into It.
    DSCBuffDesc = NewBuffDesc
End Property

Public Property Get DSBObj() As DirectSoundCaptureBuffer8
    'Set The Return Value Equal To The Internal DSCBuffer.
    Set DSBObj = DSCBuffer
End Property

Public Property Set DSBObj(ByRef NewDSCBObject As DirectSoundCaptureBuffer8)
    'Set The Internal DSCBuffer Equal To The Object Passed In.
    Set DSCBuffer = NewDSCBObject
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

'I Think I Should Maybe Just Give The Caller The Wave Format Parameters
'or maybe keep allowing indivudal queries such as channels, samples, ...
'
Public Property Get CaptureBufferWaveFormat() As WAVEFORMATEX
    
    'set the format to the internal format
    Set CaptureBufferWaveFormat = BuffWavFormat
    
End Property

Public Property Get Channels() As Long
    'Gets The Number Of Channels The Capture Buffer Is Supporting
    Channels = BuffWavFormat.nChannels
End Property

Public Property Get Samples() As Long
    'Retrieves The Samples Per Second Of The Capture Buffer
    Samples = BuffWavFormat.lSamplesPerSec
End Property

Public Property Get BitsPerSample() As Long
    'Retrieves The Bits Per Sample Of The Capture Buffer
    BitsPerSample = BuffWavFormat.nBitsPerSample
End Property

Public Property Get Size() As Long
    'Retrieves The Size Of The Buffer
    Size = BuffWavFormat.nSize
End Property

Public Property Get BufferCurrentlyActive() As Boolean
Dim BufferStatusFlags As Long

'    'Get The Boolean Value For The State Of The Buffer
'    BufferCurrentlyActive = BufferActive
    
    'Use The Get Status Call To Determine Weather We Are Playing Or Not
    BufferStatusFlags = DSCBuffer.GetStatus
    
    'This Value Returns Either 0, 1 or 3 for stoped, playing
    'and playing with looping respectively.
    
    'If Buffer Is Not 0 Then It Must Be Playing
    If Not (BufferStatusFlags = 0) Then
        BufferCurrentlyActive = True
    End If
    
End Property

Public Sub CreateBuffer(ByRef objDC As clsDirectCapture8)
    'Maybe Have Some Stuff Up Here For The Buff Desc Structure
    
    'Create The Buffer
    Set DSCBuffer = Nothing
    Set DSCBuffer = objDC.DSCObj.CreateCaptureBuffer(DSCBuffDesc)
    
    'Upon Creating The Buffer I should set the buffwavformat
    'structure accordingly.
    DSCBuffer.GetFormat BuffWavFormat
    
    If Err Then
        ReleaseResources
        End
    End If
End Sub

Public Sub StartCapture()
    'This Sub Will Start The Capture Buffer.
    'Loop The Sound?
    If bLooping Then
        DSCBuffer.Start DSBPLAY_LOOPING
'        BufferActive = True
    Else
        DSCBuffer.Start DSBPLAY_DEFAULT
'        BufferActive = True
    End If
End Sub

Public Sub StopCapture()
    'This Will Stop The Capture Buffer
    DSCBuffer.Stop
'    BufferActive = False
End Sub

Public Sub GetCurrentPosition(ByRef Cursors As DSCURSORS)
    'This Sub Will Set The Structure Passed In To The
    'Current Cursor Positions Of The Capture Buffer
    Cursors = DSCBuffer.GetCurrentPosition
End Sub

Public Function ReadBuff(ByRef ByteBuffer() As Byte, ByRef StartPosition As Long) As Long
'This Function Will Will Cause The Class Object To Read Its Currently Capturing Buffer
'From "StartPosition" Up To The Most Current, Storing Data Into ByteBuffer().

'Used When Trying To Read In The New Cursor Positions 2 Times Used below.
Dim DS As Long          'DataSize

'Use These Vars For Debug''''''''
Dim TempLastPosition As Long    '
'''''''''''''''''''''''''''''''''

'Place The StartPosition Into A Temporary Variable To Be Used Later
'TempLastPosition = LastWriteCursorPosition
TempLastPosition = StartPosition

    DSCBuffer.GetCurrentPosition CaptureCursors
    
'    If Not (FirstRead) Then
        '''''''''''''''''''Read The Capture Buffer''''''''''''''''''''''''''''''''''''''
        'Get Some General Information About The Currently Capturing Capture Buffer, Such
        'As Cursor Positions.
        
        'Capture Buffer Record Cursor Position And
        'The Other Cursor That We Will Not Be Using.
'        DSCBuffer.GetCurrentPosition CaptureCursors
    
        'Maybe Not Use This Because We Can Refer To The "CaptureCursors.lWrite" Value.
        'Position Up To Which Data Is Valid
        CurrentWriteCursorPosition = CaptureCursors.lWrite
    
        'Calculate A Data Size
        'Get The Data Size = CurrentSpot - StartPosition
         CurrentReadDataSize = CurrentWriteCursorPosition - StartPosition
    
        If (CurrentReadDataSize < 0) Then
            'Knowing That We Are Setting Up The Notify Positions To Occur 2 Times Per Buffer,
            'This Must Mean The Buffer Cursor Wraped Around
            'Get A New Datasize = BufferSize - LastCursorPosition + NewPosition
            CurrentReadDataSize = DSCBuffDesc.lBufferBytes - _
                StartPosition + CurrentWriteCursorPosition
        
        ElseIf (CurrentReadDataSize = 0) Then
            'Again This Will Probally Not Happen Since The Notify Positions Are
            'Set Up Midway Between The Circular Buffer. Now That Said If Datasize
            'Is 0 Then The Difference From The Last Cursor Position
            'And The Current Position Is Equal To The Size Of The Buffer.
            'So We Have To Make Sure The Byte Array Is As Big
            'As The Size Of The Buffer.
           CurrentReadDataSize = DSCBuffDesc.lBufferBytes
        End If
    
        'Now CurrentReadDataSize Is A Positive Number, > 0.
        
    '    'Read in the new data up to the newest write cursor position, possibly
    '    'making DataSize a few bytes larger.
    '
    '    'Get New Capture Buffer Record Cursor Position
    '    DSCBuffer.GetCurrentPosition CaptureCursors
    '
    '    'Now Finalize The Actuall Size Of DataSize
    '    If ((CaptureCursors.lWrite - CurrentWriteCursorPosition) < 0) Then
    '        'New Cursor wraped also
    '        DS = (DSCBuffDesc.lBufferBytes - _
    '            CurrentWriteCursorPosition + CaptureCursors.lWrite)
    '        CurrentReadDataSize = (CurrentReadDataSize + DS)
    '    Else
    '        CurrentReadDataSize = (CurrentReadDataSize + _
    '            (CaptureCursors.lWrite - CurrentWriteCursorPosition))
    '    End If
    '
    
        ReDim ByteBuffer(CurrentReadDataSize - 1)
        RtlZeroMemory ByteBuffer(0), (UBound(ByteBuffer()) - 1)
        
        'Read Data From Capture Device Into Buffer
        DSCBuffer.ReadBuffer _
            StartPosition, _
            CurrentReadDataSize, _
            ByteBuffer(0), _
            DSCBLOCK_DEFAULT
        
        'Set The New CaptureLastCursorPosition
        StartPosition = CurrentWriteCursorPosition 'CaptureCursors.lWrite
        
        'Done, Return The New Size Of The ByteBuffer
        ReadBuff = (UBound(ByteBuffer()) + 1)
    
'    Else
'        FirstRead = Not (FirstRead)
'        ReadBuff = -1
'    End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    'Print Out The Results To Debug Window
'
'    'Buffer Size
'    Debug.Print "Buffer Size = "; DSCBuffDesc.lBufferBytes
'
'    'Print Last Position
'    Debug.Print "Last Cursor Position = "; TempLastPosition
'
'    'Current Datasize
'
'
'    Debug.Print "Data Size = "; CurrentReadDataSize
'
'    'Position Up To Which data Was Valid
'    Debug.Print "New Cursor Position ="; CaptureCursors.lWrite
'
'    'Now The OldPosition + DataSize Should = NewPos
'    If (CurrentReadDataSize + TempLastPosition - DSCBuffDesc.lBufferBytes) >= 0 Then
'        'The CursorPosition Has Wraped Around So ..
'        Debug.Print "New Position Should be = "; _
'            (CurrentReadDataSize + TempLastPosition - DSCBuffDesc.lBufferBytes)
'    Else
'        'The Cursor Has Not Wraped Around
'        Debug.Print "New Position Should be = "; _
'            (CurrentReadDataSize + TempLastPosition)
'    End If
'    Debug.Print
'    Debug.Print

End Function

Private Sub frmDX_DXCallBack(ByRef EventID As Long)
Dim TempDataSize As Long

    'The Buffer Has Hit Where we set a notification position.
    'we know this because the eventid is = to the xenventhandle (the frmdx8)
    
    If EventID = XEventHandle Then

'        DoEvents
        
        'Get The Size in bytes of the Pilots Capture Buffer, Also
        'put the data into the "Buffer" variable
        TempDataSize = Me.ReadBuff(Buffer, RecLCP1)

'        Debug.Print "Xevent occured  " & Time
        Debug.Print "Capture Buffer Size Is: "; DSCBuffDesc.lBufferBytes
        Debug.Print "Data Size Should Be Around: "; (DSCBuffDesc.lBufferBytes / 2)
        
        Debug.Print "Capture Cursor Capture Positions Is: "; CaptureCursors.lPlay
        Debug.Print "Capture Cursor lWrite Positions Is And Last Cursor Position: "; CaptureCursors.lWrite
        
'        Debug.Print "Last Cursor Position: "; RecLCP1
        Debug.Print "DS ="; TempDataSize
        Debug.Print
        Debug.Print
        
'        DoEvents
        
        'Test The Value
        If (TempDataSize > 2) Then
            'Got Data So Write Into the pilots Sound Buffer.
            clsDSB1.PlayAural clsDS1, Buffer, TempDataSize
        End If
    
    End If

End Sub

Public Sub EnableNotifyPosition()
'This Should Pass In The Required Variables
Dim BufferSliceSizeThird As Integer

    'Lets Try To Set Up 3 Notify Positions
     
    BufferSliceSizeThird = Int(DSCBuffDesc.lBufferBytes / 3)
    
    With NewNotifyPosition(0)
        .hEventNotify = XEventHandle
        'Set It Up To Go Off MidWay In The Buffer
        .lOffset = 0 '(DSCBuffDesc.lBufferBytes / 2)
    End With
    
    With NewNotifyPosition(1)
        .hEventNotify = XEventHandle
        .lOffset = BufferSliceSizeThird
    End With
    
    With NewNotifyPosition(2)
        .hEventNotify = XEventHandle
        .lOffset = (BufferSliceSizeThird * 2)
    End With

    DSCBuffer.SetNotificationPositions 3, NewNotifyPosition()
End Sub
