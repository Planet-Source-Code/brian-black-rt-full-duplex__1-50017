VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   3555
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrRecord 
      Left            =   480
      Top             =   2160
   End
   Begin VB.CommandButton cmdStopRecord 
      Caption         =   "Stop Record"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdStartRecord 
      Caption         =   "Start Record"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Number Of Sound Devices"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblNumOfDevices 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   2055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStartRecord_Click()
'check to c if the buffer capturing if not then start it

    If Not (clsDSCB1.BufferCurrentlyActive) Then
        clsDSCB1.Looping = True
        clsDSCB1.EnableNotifyPosition
        clsDSCB1.StartCapture
    End If
    
'    If Not (clsDSC1.DSCObj Is Nothing) Then
'        bStartRec1Click = True
'        bStopRec1Click = False
'    Else
'        MsgBox "Check To Make Sure Sound Card 1 Is Working Correctly", _
'                vbCritical, "Error."
'        bStartRec1Click = False
'        bStopRec1Click = True
'        bCaptureRunning1 = False
'        If (clsDS1.DSObj Is Nothing) Then
'            bStartRec1Click = False
'            bStopRec1Click = True
'            bSoundRunning1 = False
'        End If
'    End If
End Sub

Private Sub cmdStopRecord_Click()
    bStartRec1Click = False
    bStopRec1Click = True
End Sub

Private Sub Form_Load()
    
'    With Me.tmrRecord
'        .Interval = 45
'        .Enabled = True
'    End With
    
End Sub

Public Sub tmrRecord_Timer()
Dim TempDataSize As Long
    
    If bStartRec1Click = True Then
        If Not (bCaptureRunning1) Then
            clsDSCB1.Looping = True
            clsDSCB1.StartCapture
            bCaptureRunning1 = True
        End If
        
        If Not (bSoundRunning1) Then
            clsDSB1.Looping = True
            clsDSB1.PlayBuffer
            bSoundRunning1 = True
        End If
    
        'Get The Size in bytes of the Pilots Capture Buffer, Also
        'put the data into the "Buffer" variable
        TempDataSize = clsDSCB1.ReadBuff(Buffer, RecLCP1)
        
        'Test The Value
        If (TempDataSize > 0) Then
            'Got Data So Write Into the pilots Sound Buffer.
'            clsDSB1.WriteBuff Buffer
            clsDSB1.PlayAural clsDS1, Buffer, TempDataSize
        ElseIf (TempDataSize < 0) Then
            'This is the first read beign performed on the Capture Buffer
        End If
    
    ElseIf (bStopRec1Click) Then
        
        'Stop The Buffers
        If (bCaptureRunning1) Then
            clsDSCB1.StopCapture
            bCaptureRunning1 = False
        End If
        
        If (bSoundRunning1) Then
            clsDSB1.StopBuffer
            bSoundRunning1 = False
        End If
    
    End If

End Sub
