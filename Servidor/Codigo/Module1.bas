Attribute VB_Name = "Module1"
Option Explicit
 
Public DaleLoop As Boolean
 
Type tMainLoop
    MaxInt As Long
    LastCheck As Long
End Type
 
Private Const NumTimers As Byte = 2 '//Aca la cantidad de timers.
Public MainLoops(1 To NumTimers) As tMainLoop
 
Public Enum eTimers
    GameTimer = 1
    packetResend
End Enum
 
 
Public Sub MainLoop()
    Dim LoopC As Long
    'TODO: El loop hay que hacer que se haga solo cuando corresponda
    initIntervals
   
   
    Do While DaleLoop = True
        For LoopC = 1 To NumTimers
            With MainLoops(LoopC)
                If GetTickCount - .LastCheck >= .MaxInt Then
                    Call MakeProcces(LoopC)
                End If
            End With
            DoEvents
        Next LoopC
        DoEvents
    Loop
End Sub
 
Private Sub initIntervals()
 
    MainLoops(eTimers.GameTimer).MaxInt = 40
    MainLoops(eTimers.packetResend).MaxInt = 10
 
 
   
End Sub
 
Private Sub MakeProcces(ByVal index As Integer)
    Select Case index
        Case eTimers.GameTimer
            Call frmMain.GameTimer_Timer
 
        Case eTimers.packetResend
            Call frmMain.packetResend_Timer
               
    End Select
    MainLoops(index).LastCheck = GetTickCount
End Sub
