VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Member variables IDs for the in-code timers.

Private mTimerID As Long
Private mRunning As Boolean
Private mInterval As Long

' public tick event for an object of this class

Public Event Tick()


'---------------------------------------------------------------------------------------
' Procedure : TimerID
' Author    : beededea & chatGPT (with corrections by me)
' Date      : 17/12/2025
' Purpose   : property to obtain the timer ID number
'---------------------------------------------------------------------------------------
'
Friend Property Get TimerID() As Long
    On Error GoTo TimerID_Error

    TimerID = mTimerID

    On Error GoTo 0
    Exit Property

TimerID_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TimerID of Class Module cwTenShillings"
End Property

'---------------------------------------------------------------------------------------
' Procedure : TimerID
' Author    : beededea & chatGPT (with corrections by me)
' Date      : 17/12/2025
' Purpose   : property to set the timer ID number
'---------------------------------------------------------------------------------------
'
Friend Property Let TimerID(ByVal v As Long)
    On Error GoTo TimerID_Error

    mTimerID = v

    On Error GoTo 0
    Exit Property

TimerID_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TimerID of Class Module cwTenShillings"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Start
' Author    : beededea & chatGPT (with corrections by me)
' Date      : 17/12/2025
' Purpose   : start a new timer according to supplied interval
'---------------------------------------------------------------------------------------
'
Public Sub Start(ByVal IntervalMS As Long)
    On Error GoTo Start_Error

    If mRunning Then StopTimer
    mInterval = IntervalMS
    TimerManager.StartTimer Me, IntervalMS
    mRunning = True

    On Error GoTo 0
    Exit Sub

Start_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Start of Class Module clsTimer"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : StopTimer
' Author    : beededea & chatGPT (with corrections by me)
' Date      : 17/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub StopTimer()
    On Error GoTo StopTimer_Error

    If mRunning Then
        TimerManager.StopTimer Me
        mRunning = False
    End If

    On Error GoTo 0
    Exit Sub

StopTimer_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure StopTimer of Class Module clsTimer"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : RaiseTick
' Author    : beededea & chatGPT (with corrections by me)
' Date      : 17/12/2025
' Purpose   : raise a tick event for an object of this class
'---------------------------------------------------------------------------------------
'
Public Sub RaiseTick()
    On Error GoTo RaiseTick_Error

    If mRunning Then
        RaiseEvent Tick
    End If

    On Error GoTo 0
    Exit Sub

RaiseTick_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RaiseTick of Class Module clsTimer"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : beededea & chatGPT (with corrections by me)
' Date      : 17/12/2025
' Purpose   : terminate the current object of this class
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    On Error GoTo Class_Terminate_Error

    StopTimer

    On Error GoTo 0
    Exit Sub

Class_Terminate_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Terminate of Class Module clsTimer"
End Sub



