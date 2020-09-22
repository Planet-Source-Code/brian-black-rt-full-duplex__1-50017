VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDirectSound8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------
' Author: Brian Black  BlackBsd@Mountain.Net
' DirectSound8 Class Source Code File
' License: Currently None
'----------------------------------------------------------------------
Option Explicit
Option Base 0

Private objDX As DirectX8
Private objDS As DirectSound8

Private Sub Class_Initialize()
    'Create The DirectX8 Object
    Set objDX = New DirectX8

End Sub

Private Sub Class_Terminate()
    Set objDS = Nothing
    Set objDX = Nothing
End Sub

Public Property Get DSObj() As DirectSound8
    Set DSObj = objDS
End Property

Public Property Get DXObj() As DirectX8
    Set DXObj = objDX
End Property

Public Sub Startup(ByRef lWindowHandle As Long, Optional ByRef GUID As String = "")
    
    'Create The DirectSound8 Object Using The Guid Or Not
    If Not (GUID = "") Then
        'If Guid Is not "" then use that id to create the sound on
        Set objDS = objDX.DirectSoundCreate(GUID)
    Else
        'This Means to create the sound device on the OS Default Device.
        Set objDS = objDX.DirectSoundCreate("")
    End If
    
    'Set The Priority Level For DirectSound to the priority level.
    'Applications with this cooperative level can call the
    'DirectSoundPrimaryBuffer8.SetFormat method.
    objDS.SetCooperativeLevel lWindowHandle, DSSCL_PRIORITY
    
End Sub
