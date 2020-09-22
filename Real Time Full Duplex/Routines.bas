Attribute VB_Name = "modRoutines"
Option Explicit

Public Sub Main()
    
    Load frmMain
    
    If (InitDirectSoundObject) Then
        If (InitDirectCaptureObject) Then
            CreateSoundBuffer
            CreateCaptureBuffer
        End If
    End If
    
    frmMain.Show
    
'    bStartRec1Click = True
'    frmMain.tmrRecord_Timer
    
    
End Sub

Public Function InitDirectSoundObject() As Boolean
On Error GoTo FailedDirectSoundInit
    
    Set DSEnum = DX.GetDSEnum
    frmMain.lblNumOfDevices.Caption = DSEnum.GetCount - 1
    
    '''Set Up DirectSoundDevice
    If (DSEnum.GetCount) > 1 Then
        clsDS1.Startup frmMain.hWnd, DSEnum.GetGuid(2)
        blnSoundBoard1 = True
        InitDirectSoundObject = True
    End If

    Exit Function

FailedDirectSoundInit:
    InitDirectSoundObject = False
End Function

Public Function InitDirectCaptureObject() As Boolean
On Error GoTo FailedDirectCaptureInit

    Set DSCEnum = DX.GetDSCaptureEnum
    
    '''''''''''''''''''''''''''''''''''''''''''''
    'Now We Can Create The Capture Object       '
    'So We Can Make A Capture Buffer From It    '
    '''''''''''''''''''''''''''''''''''''''''''''
    If (DSCEnum.GetCount) > 1 Then
        clsDSC1.Startup DSCEnum.GetGuid(2)
        blnCaptureBoard1 = True
        InitDirectCaptureObject = True
    End If

    Exit Function

FailedDirectCaptureInit:
    InitDirectCaptureObject = False
End Function

Public Sub CreateSoundBuffer()
    'We Have To Create A Sound Buffer So the Output Of the capture buffer
    'Can Be Placed Into It.
    
    
    'First Create A Sound Buffer Description
    DSDesc.fxFormat = CreateWaveFormatEx(44100, 1, 16)
    
    'A Two Second Buffer 'Half Second Sound Buffer.
    'The buffer size is generally shown in bytes or samples
    DSDesc.lBufferBytes = (DSDesc.fxFormat.lAvgBytesPerSec / 2)
    
    'Set The Flags
    DSDesc.lFlags = DSBCAPS_STATIC Or DSBCAPS_CTRLFREQUENCY Or _
            DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or _
            DSBCAPS_GETCURRENTPOSITION2 Or DSBCAPS_GLOBALFOCUS
    
    
    'Now Assign The Buffer Desc We Just Created To A Direct Sound Buffer
    'Object's Buffer Desc Structure.
    clsDSB1.DSBuffDesc = DSDesc
    
    'Now Create The Actuall Sound Buffer Using the Direct Sound Object We Created
    'In The InitDirectSoundObject Function
    clsDSB1.CreateBuffer clsDS1

End Sub

Public Sub CreateCaptureBuffer()
    'Now Create A Capture Buffer So The Data Gets Captured into the buffer
    'and can be used later.
            
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Lets Get The Capabilities For Our                          '
        'Capture Object That We Created In InitDirectCaptureObject  '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        clsDSC1.GetCaptureCaps CaptureCaps
    
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Get The Capture Format We Will Support In The Sample   '
        'And Populate The Buffer Desc Structure                 '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If (CaptureCaps.lFormats And WAVE_FORMAT_4M16) Then
            'First Create A Capture Buffer Description
            DSCDesc.fxFormat = CreateWaveFormatEx(44100, 1, 16)
            'A Two Second Buffer 'Half Second Second Capture Buffer
            DSCDesc.lBufferBytes = (DSCDesc.fxFormat.lAvgBytesPerSec / 2)
            'Dont Forget To Set Some Flags
            DSCDesc.lFlags = DSCBCAPS_WAVEMAPPED
            
            ReDim Buffer(DSCDesc.lBufferBytes - 1)
        Else
            MsgBox "Could Not Get The Capture Capibilities We Need On This Card.", _
                    vbOKOnly Or vbCritical, "Exiting."
            ReleaseResources
            End
        End If
            
        'Set The Objects Structure Equal To What We Just Made
        clsDSCB1.CaptureBuffDesc = DSCDesc
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Create the Capture Buffer From Object clsDSC1      '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''
        clsDSCB1.CreateBuffer clsDSC1
End Sub

Public Sub AssociateCaptureAndSoundBuffers()

End Sub

Public Function CreateWaveFormatEx(Hz As Long, Channels As Integer, BITS As Integer) As WAVEFORMATEX
    'Create a WaveFormatEX Structure Using The Vars We Provide
    With CreateWaveFormatEx
        .nFormatTag = WAVE_FORMAT_PCM
        .nChannels = Channels
        .lSamplesPerSec = Hz
        .nBitsPerSample = BITS
        .nBlockAlign = ((.nChannels * .nBitsPerSample) / 8)
        .lAvgBytesPerSec = (.lSamplesPerSec * .nBlockAlign)
        .nSize = 0
    End With
End Function

Public Sub ReleaseResources()
  
  Set clsDS1 = Nothing
  
  Set clsDSC1 = Nothing
  
  Set clsDSB1 = Nothing
  
  Set clsDSCB1 = Nothing
 
  Set DX = Nothing

End Sub
