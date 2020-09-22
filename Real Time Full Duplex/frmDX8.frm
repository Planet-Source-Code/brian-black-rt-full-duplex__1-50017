VERSION 5.00
Begin VB.Form frmDX8 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmDX8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Event DXCallBack(ByRef EventID As Long)

Implements DirectXEvent8
 
Dim gDX As New DirectX8
Dim FormsEventHandle As Long
Dim dsbpn(0) As DSBPOSITIONNOTIFY

Private Sub Form_Load()
    FormsEventHandle = gDX.CreateEvent(Me)
End Sub

Private Sub DirectXEvent8_DXCallback(ByVal EventID As Long)
'    If EventID = EndEvent Then
        RaiseEvent DXCallBack(EventID)
'    End If
End Sub
