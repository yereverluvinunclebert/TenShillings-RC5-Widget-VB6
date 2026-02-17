VERSION 5.00
Begin VB.Form frmTimer 
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmTimer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer sleepTimer 
      Interval        =   3000
      Left            =   90
      Tag             =   "stores and compares the last time to see if the PC has slept"
      Top             =   1560
   End
   Begin VB.Timer unhideTimer 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   90
      Top             =   1095
   End
   Begin VB.Timer tmrScreenResolution 
      Enabled         =   0   'False
      Interval        =   4500
      Left            =   90
      Top             =   615
   End
   Begin VB.Timer revealWidgetTimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   90
      Top             =   135
   End
   Begin VB.Label Label4 
      Caption         =   "sleeptimer for testing awake from sleep"
      Height          =   195
      Left            =   705
      TabIndex        =   4
      Top             =   1650
      Width           =   3645
   End
   Begin VB.Label Label3 
      Caption         =   "Note: this invisible form is also the container for the large 128x128px project icon"
      Height          =   435
      Left            =   285
      TabIndex        =   3
      Top             =   3120
      Width           =   4125
   End
   Begin VB.Label Label2 
      Caption         =   "if the unhide setting is set by another process it will unhide the widget"
      Height          =   195
      Left            =   705
      TabIndex        =   2
      Top             =   1170
      Width           =   3645
   End
   Begin VB.Label Label1 
      Caption         =   "ScreenResolutionTimer for handling rotation of the screen"
      Height          =   195
      Left            =   705
      TabIndex        =   1
      Top             =   735
      Width           =   3570
   End
   Begin VB.Label Label 
      Caption         =   "revealWidgetTimer for revealing after a hide."
      Height          =   195
      Left            =   690
      TabIndex        =   0
      Top             =   270
      Width           =   3480
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmTimer
' Author    : beededea
' Date      : 25/10/2024
' Purpose   : holds all the VB6 timers AND the events and startup of the in-code timers used by the program
'---------------------------------------------------------------------------------------

'@IgnoreModule ModuleWithoutFolder
Option Explicit

' create in-code timer vars as class timers enabling the tick event for each instance
Private WithEvents timer1 As clsTimer
Attribute timer1.VB_VarHelpID = -1
Private WithEvents timer2 As clsTimer
Attribute timer2.VB_VarHelpID = -1

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : beededea
' Date      : 17/12/2025
' Purpose   : The load event for the form is only used to start the in-code timers
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
    
    On Error GoTo Form_Load_Error

    Set timer1 = New clsTimer
    Set timer2 = New clsTimer

    ' rename the timers as you see fit
    
    'timer1.Start 5000 ' 5 seconds
    'timer2.Start 90000 ' 90 seconds

    On Error GoTo 0
    Exit Sub

Form_Load_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form frmTimer"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : timer1_Tick
' Author    : beededea
' Date      : 17/12/2025
' Purpose   : Sample event for timer-in-code no. 1
'---------------------------------------------------------------------------------------
'
Public Sub timer1_Tick()
    On Error GoTo timer1_Tick_Error

    timer1.StopTimer
    
    Debug.Print "1 second"
   ' MsgBox "Tick 5 seconds"

    On Error GoTo 0
    Exit Sub

timer1_Tick_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure timer1_Tick of Form frmTimer"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : timer2_Tick
' Author    : beededea
' Date      : 17/12/2025
' Purpose   : Sample event for timer-in-code no. 2
'---------------------------------------------------------------------------------------
'
Public Sub timer2_Tick()
    On Error GoTo timer2_Tick_Error

    timer2.StopTimer
    
    Debug.Print "90"
    'MsgBox "Tick 90 seconds"

    On Error GoTo 0
    Exit Sub

timer2_Tick_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure timer2_Tick of Form frmTimer"
End Sub


' ====================================================================
' from this point onward the code is solely for the VB6 on-form timers
' ====================================================================


'---------------------------------------------------------------------------------------
' Procedure : revealWidgetTimer_Timer
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : revealWidgetTimer for revealing after a hide.
'---------------------------------------------------------------------------------------
'
Private Sub revealWidgetTimer_Timer()
    Static revealWidgetTimerCount As Integer
    
    On Error GoTo revealWidgetTimer_Timer_Error

    revealWidgetTimerCount = revealWidgetTimerCount + 1
    If revealWidgetTimerCount >= (giMinutesToHide * 12) Then
        revealWidgetTimerCount = 0

        fMain.TenShillingsForm.Visible = True
        revealWidgetTimer.Enabled = False
        gsWidgetHidden = "0"
        sPutINISetting "Software\TenShillings", "widgetHidden", gsWidgetHidden, gsSettingsFile
    End If

    On Error GoTo 0
    Exit Sub

revealWidgetTimer_Timer_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure revealWidgetTimer_Timer of Form frmTimer"
            Resume Next
          End If
    End With
End Sub




'---------------------------------------------------------------------------------------
' Procedure : tmrScreenResolution_Timer
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : ScreenResolutionTimer for handling rotation of the screen
'             in tablet mode or a resolution change
'             possibly due to an old game in full screen mode.
'---------------------------------------------------------------------------------------
'
Private Sub tmrScreenResolution_Timer()

'    Dim resizeProportion As Single: resizeProportion = 0
    
    On Error GoTo tmrScreenResolution_Timer_Error

    glPhysicalScreenHeightPixels = GetDeviceCaps(Me.hDC, VERTRES)
    glPhysicalScreenWidthPixels = GetDeviceCaps(Me.hDC, HORZRES)
    
    glVirtualScreenWidthPixels = fVirtualScreenWidth(True)
    glVirtualScreenHeightPixels = fVirtualScreenHeight(True)

    ' calls a routine that tests for a change in the monitor upon which the form sits, if so, resizes
    'Call resizeLocateRCFormByMoveToNewMonitor
    
    ' will be used to check for orientation changes
    If (glOldPhysicalScreenHeightPixels <> glPhysicalScreenHeightPixels) Or (glOldPhysicalScreenWidthPixels <> glPhysicalScreenWidthPixels) Then
        
        ' move/hide onto/from the main screen and position per orientation portrait/landscape
        Call mainScreen
'
        'store the resolution change
        glOldPhysicalScreenHeightPixels = glPhysicalScreenHeightPixels
        glOldPhysicalScreenWidthPixels = glPhysicalScreenWidthPixels
    End If

    On Error GoTo 0
    Exit Sub

tmrScreenResolution_Timer_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrScreenResolution_Timer of Form frmTimer"
            Resume Next
          End If
    End With
End Sub



'---------------------------------------------------------------------------------------
' Procedure : unhideTimer_Timer
' Author    : beededea
' Date      : 13/05/2023
' Purpose   : if the unhide setting is set by another process it will unhide the widget
'---------------------------------------------------------------------------------------
'
Private Sub unhideTimer_Timer()
    
    On Error GoTo unhideTimer_Timer_Error

    gsUnhide = fGetINISetting("Software\TenShillings", "unhide", gsSettingsFile)

    If gsUnhide = "true" Then
        fMain.TenShillingsForm.Visible = True
        sPutINISetting "Software\TenShillings", "unhide", vbNullString, gsSettingsFile
    End If

    On Error GoTo 0
    Exit Sub

unhideTimer_Timer_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure unhideTimer_Timer of Form frmTimer"
            Resume Next
          End If
    End With
End Sub



'---------------------------------------------------------------------------------------
' Procedure : sleepTimer_Timer
' Author    : beededea
' Date      : 21/04/2021
' Purpose   : timer that stores the last time the timer was run
' if the current time is greater than the last time stored by more than 30 seconds we can assume the system
' has been sent to sleep, if the two are significantly different then we reorganise the dock
'---------------------------------------------------------------------------------------
'
Private Sub sleepTimer_Timer()
    Dim strTimeNow As Date: strTimeNow = #1/1/2000 12:00:00 PM#  'set a variable to compare for the NOW time
    Dim lngSecondsGap As Double: lngSecondsGap = 0  ' set a variable for the difference in time
    
    Static strTimeThen As Date
    
    On Error GoTo sleepTimer_Timer_Error

    If strTimeThen = "00:00:00" Then strTimeThen = Now(): Exit Sub
    sleepTimer.Enabled = False
    
    strTimeNow = Now()
    
    lngSecondsGap = DateDiff("s", strTimeThen, strTimeNow)
    strTimeThen = Now()

    If lngSecondsGap > 60 Then
      
        gbThisWidgetAvailable = True
        ' Call ' screenWrite("system has just woken up from a sleep at " & Now() & vbCrLf & "updating digital widgets... ")
        
        'tenShillingsOverlay.BaseDate = Now()
        'gsTriggerDigitalWidgetPopulation = True
        
        fMain.TenShillingsForm.Refresh
        
'        If gsNumericDisplayRotation = "1" Then
'            tenShillingsOverlay.TmrDigitRotatorTicking = True
'        End If
'
'        '  clear any existing weekday indicator after a wake from sleep
'        If tenShillingsOverlay.weekdayToggleEnabled = True Then
'            Call hideDayOfWeek
'            fMain.TenShillingsForm.Widgets(fDayOfWeek).Widget.Alpha = 1
'        End If
        
        tenShillingsOverlay.Widget.Parent.Refresh
        
    End If
    
    sleepTimer.Enabled = True

    On Error GoTo 0
    Exit Sub

sleepTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sleepTimer_Timer of Form dock"

End Sub
