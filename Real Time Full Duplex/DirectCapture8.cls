VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDirectCapture8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------
' Author: Brian Black  BlackBsd@Mountain.Net
' DirectCapture8 Class Source Code File
' License: Currently None
'----------------------------------------------------------------------
Option Explicit
Option Base 0

'DirectSoundCapture Objects
Dim objDX As New DirectX8
Dim objDSC As DirectSoundCapture8

'Program Variables
Private DSCCAPS As DSCCAPS

Private Sub Class_Initialize()
    'Create The DirectX8 Object
    Set objDX = New DirectX8
End Sub

Private Sub Class_Terminate()
    Set objDSC = Nothing
    Set objDX = Nothing
End Sub

Public Function GetCaptureCaps(ByRef Caps As DSCCAPS)
    If Not (objDSC Is Nothing) Then
        objDSC.GetCaps Caps
    End If
End Function

Public Property Get DSCObj() As DirectSoundCapture8
    Set DSCObj = objDSC
End Property

Public Property Get DXObj() As DirectX8
    Set DXObj = objDX
End Property

Public Sub Startup(Optional ByRef GUID As String = "")
    'Create The DirectSound8 Object Using The Guid Or Not
    If Not (GUID = "") Then
        Set objDSC = objDX.DirectSoundCaptureCreate(GUID)
    Else
        Set objDSC = objDX.DirectSoundCaptureCreate("")
    End If
End Sub
