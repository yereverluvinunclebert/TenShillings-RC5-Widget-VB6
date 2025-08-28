VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cfWidget
' Author    : beededea
' Date      : 18/10/2024
' Purpose   : This is the class for the widget form and all the elements upon it.
'---------------------------------------------------------------------------------------

'@IgnoreModule IntegerDataType, ModuleWithoutFolder
Option Explicit

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

'Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
' (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
'  ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

' Declare widgetForm as an RC form to be available throughout by calling fWidget as defined as this class in modMain

Public WithEvents widgetForm As cWidgetForm
Attribute widgetForm.VB_VarHelpID = -1

' public vars and objects

Public collPSDNonUIElements As cCollection

Public FX As Long   ' available externally, to allow the widgetForm to be moved, need to add set/let for these public vars
Public FY As Long
Public FZ As Single

Private mGlowing As Boolean

' class private members for property assignment

Private mShowHelp As Boolean
Private mOpacity As String
Private pvtPSDWidth As Long
Private pvtPSDHeight As Long

' RC widgets

'Private WithEvents helpbottom As cWidgetBase

'local general vars used only within this class


'---------------------------------------------------------------------------------------
' Procedure : initialisePrivateVars
' Author    : beededea
' Date      : 12/02/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub initialisePrivateVars()
   On Error GoTo initialisePrivateVars_Error

    pvtPSDWidth = 0
    pvtPSDHeight = 0

   On Error GoTo 0
   Exit Sub

initialisePrivateVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialisePrivateVars of Class Module cfWidget"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   : Constructor (Initialise)
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
   
    On Error GoTo Class_Initialize_Error
    
    Call initialisePrivateVars
    
   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cfWidget"
End Sub





'---------------------------------------------------------------------------------------
' Procedure : widgetForm_DblClick
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : If a double-click command has been selected in the prefs then this will carry it out
'---------------------------------------------------------------------------------------
'
Private Sub widgetForm_DblClick()
    Dim userprof As String: userprof = vbNullString
    Dim thisCommand As String: thisCommand = vbNullString
    
    On Error GoTo widgetForm_DblClick_Error

    '    If gblIgnoreMouse = "1" Then Exit Sub

    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    If LTrim$(gblDblClickCommand) = vbNullString Then Exit Sub
    
    thisCommand = gblDblClickCommand
        
    If InStr(thisCommand, "%userprofile%") Then
        userprof = Environ$("USERPROFILE")
        thisCommand = Replace(thisCommand, "%userprofile%", userprof)
    End If
    
    ' .91 DAEB 08/12/2022 frmMain.frm SteamyDock responds to %systemroot% environment variables during runCommand
    If InStr(thisCommand, "%systemroot%") Then
        userprof = Environ$("SYSTEMROOT")
        thisCommand = Replace(thisCommand, "%systemroot%", userprof)
    End If
    
    If gblSHIFT_1 = True Then
        gblSHIFT_1 = False
        Call ShellExecute(widgetForm.hWnd, "Open", gblOpenFile, vbNullString, App.Path, 1)
    Else
        Call ShellExecute(widgetForm.hWnd, "runas", thisCommand, vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

widgetForm_DblClick_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure widgetForm_DblClick of Class Module cfWidget"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : widgetForm_KeyDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : get F5 and SHIFT keypresses and key downs for the main toggles H,A,S,M,P,T,D,W,B & 1-5
'---------------------------------------------------------------------------------------
'
Private Sub widgetForm_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
    On Error GoTo widgetForm_KeyDown_Error

    Call getKeyPress(KeyCode, Shift) ' will not catch Shift or CTRL keypresses on their own, must be combination with another key.

    On Error GoTo 0
    Exit Sub

widgetForm_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure widgetForm_KeyDown of Class Module cfWidget"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : widgetForm_MouseUp
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :    if this is a multi monitor system then:
'                we note the monitor ID at widgetForm form_load and store as the widgetFormMonitorID, for widgetPrefs widgetPrefsMonitorID
'                on widgetFromMouseUp (drag complete) we sample the monitor ID again -only works, for the prefs form we have to have a 200ms timer that tests for form .left and .top x,y change (no MOVED event and mouse UP does not work on the form itself when the titlebar dragged)
'                if the monitor has changed then sample the physical monitor resolution
'                if the resolution is different then calculate new size proportion
'                alter the size of the form to correspond to the monitor
'---------------------------------------------------------------------------------------
'
Private Sub widgetForm_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)

    On Error GoTo widgetForm_MouseUp_Error
    
    frmTimer.tmrScreenResolution.Interval = 4500

    ' save the form x,y position whenever the globe/form is dragged
    Call saveMainRCFormPosition

    Call positionRCFormByMonitorSize
    
    On Error GoTo 0
    Exit Sub

widgetForm_MouseUp_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure widgetForm_MouseUp of Class Module cfWidget"
            Resume Next
          End If
    End With
End Sub





'---------------------------------------------------------------------------------------
' Procedure : widgetForm_MouseMove
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : move the whole form on a mousDown and drag
'---------------------------------------------------------------------------------------
'
Private Sub widgetForm_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
  Static x0 As Single, y0 As Single: If Button = 0 Then x0 = x: y0 = y 'just store the offsets when no button is down
  
   On Error GoTo widgetForm_MouseMove_Error
    
    If overlayWidget.Locked = True Then Exit Sub
    If gblIgnoreMouse = "1" Then Exit Sub
    If gblMenuOccurred = True Then
        gblMenuOccurred = False
        Exit Sub
    End If

    If Button = vbLeftButton And Not widgetForm.ActiveWidget Is Nothing Then 'Form-Dragging (under certain conditions)

      FX = (x - x0) * widgetForm.WidgetRoot.Zoom + widgetForm.Left
      FY = (y - y0) * widgetForm.WidgetRoot.Zoom + widgetForm.Top
      widgetForm.Move FX, FY
      
      Debug.Print "FX " & FX
      Debug.Print "FY " & FY
      
    End If


   On Error GoTo 0
   Exit Sub

widgetForm_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure widgetForm_MouseMove of Class Module cfWidget"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : widgetForm_MouseWheel
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : Handles rotation of the mousewheel, with CTRL key pressed, making the widget smaller or larger. By default, rotate forward positive = smaller
'---------------------------------------------------------------------------------------
'
Private Sub widgetForm_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Single, ByVal Ypos As Single)
   Dim RotationDirection As Long: RotationDirection = 0
   
   On Error GoTo widgetForm_MouseWheel_Error

    Rotation = Rotation / 2

    If gblScrollWheelDirection = "1" Then
        If Rotation > 0 Then
            RotationDirection = 0 - Abs(Rotation)
        Else
            RotationDirection = Abs(Rotation)
        End If
    Else
        RotationDirection = Rotation
    End If
    
   'If MouseKeys = 8 Then Call AdjustZoom(FZ + RotationDirection / 2400)
   If MouseKeys = 8 Then widgetPrefs.WidgetSize = (FZ + RotationDirection / 2400) 'change the Zoom only, when the Ctrl-Key is down

   On Error GoTo 0
   Exit Sub

widgetForm_MouseWheel_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure widgetForm_MouseWheel of Class Module cfWidget"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : AdjustZoom
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   : Adjust the size of the whole widget, anchoring the widget top left, storing the values for later use
'---------------------------------------------------------------------------------------
'
Public Sub AdjustZoom(ByVal NewZoom As Single)
   On Error GoTo AdjustZoom_Error

    FZ = NewZoom
    If FZ < 0.05 Then FZ = 0.05 Else If FZ > 3 Then FZ = 3 'keep FZ within a sane ZoomInterval
    widgetForm.WidgetRoot.Zoom = FZ '* widgetForm.WidgetRoot.CurrentMonitor.Zoom '<-- if the multiplicator is activated, we'd have DPI-awareness on high-res monitors
    widgetForm.Move widgetForm.Left, widgetForm.Top, pvtPSDWidth * widgetForm.WidgetRoot.Zoom, pvtPSDHeight * widgetForm.WidgetRoot.Zoom
    
    ' when resized from the widget scroll up/down it needs to write it back so the size is correct when prefs opened.
    gblWidgetSize = CStr(NewZoom * 100)
    
    If widgetPrefs.IsLoaded Then widgetPrefs.sliWidgetSize.Value = Val(gblWidgetSize) ' ensure that the prefs size slider matches when zoom is modified without using the slider
    
    If widgetMonitorStruct.IsPrimary = True Then
        gblWidgetPrimaryHeightRatio = CStr(NewZoom)
    Else
        gblWidgetSecondaryHeightRatio = CStr(NewZoom)
    End If

   On Error GoTo 0
   Exit Sub

AdjustZoom_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AdjustZoom of Class Module cfWidget"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : widgetForm_MouseDown
' Author    :
' Date      : 09/05/2023
' Purpose   : A click on the widget form, test screen resolution and handle right mouse click for the menu
'---------------------------------------------------------------------------------------
'
Private Sub widgetForm_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo widgetForm_MouseDown_Error

    'essential code to test for monitor resolution before and after a mouse-down pick up and drag to another monitor
    frmTimer.tmrScreenResolution.Enabled = False
    frmTimer.tmrScreenResolution.Interval = 200
    frmTimer.tmrScreenResolution.Enabled = True
    
    If Button = vbRightButton Then
        gblMenuOccurred = True
        Call menuForm.PopupMenu(menuForm.mnuMainMenu)
    Else
        If gblIgnoreMouse = "1" Then Exit Sub
        ' so we pass it through
    End If

   On Error GoTo 0
   Exit Sub

widgetForm_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure widgetForm_MouseDown of Class Module cfWidget"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsVisible
' Author    : beededea
' Date      : 08/05/2023
'
' Purpose   : Useful property to determine by value as to whether this form is loaded,
'             allowing external checks to the form to determine whether it is loaded,
'             WITHOUT activating the form automatically.
'---------------------------------------------------------------------------------------
'
Public Property Get IsVisible() As Boolean
    On Error GoTo IsVisible_Error

    If widgetForm.WindowState = vbNormal Then
        IsVisible = widgetForm.Visible
    Else
        IsVisible = False
    End If

    On Error GoTo 0
    Exit Property

IsVisible_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsVisible of Class Module cfWidget"
            Resume Next
          End If
    End With
End Property



'---------------------------------------------------------------------------------------
' Property  : Opacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property to determine (by value) the opacity of the elements within this class
'---------------------------------------------------------------------------------------
'
Public Property Get opacity() As String
   On Error GoTo opacityGet_Error

   opacity = mOpacity

   On Error GoTo 0
   Exit Property

opacityGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Opacity of Class Module cfWidget"
End Property
'---------------------------------------------------------------------------------------
' Property  : Opacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property to determine (by value) the opacity of the elements within this class
'---------------------------------------------------------------------------------------
'
Public Property Let opacity(ByVal newValue As String)
   On Error GoTo opacityLet_Error

   If mOpacity <> newValue Then mOpacity = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

opacityLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Opacity of Class Module cfWidget"
End Property





