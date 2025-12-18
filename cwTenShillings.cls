VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwTenShillings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule IntegerDataType, ModuleWithoutFolder
'---------------------------------------------------------------------------------------
' Module    : cwTenShillings
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : create a fully click-through Widget, which allows for all kind of
'             things to render on top of the PSD-generated Widgets created in cTenShillingsWidget
'---------------------------------------------------------------------------------------

Option Explicit


Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

'Private TenShillingsxo As Integer
'Private TenShillingsyo  As Integer

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

' member variables for locally holding property values

Private mZoom As Double
Private mZoomDirection As String
Private mSkewDegrees As Double
Private mLocked As Boolean
Private mHidden As Boolean
Private mImageWidth As Long
Private mImageHeight As Long
Private mOpacity As Double
Private mSampleDelay As Long

'---------------------------------------------------------------------------------------
' Procedure : initialisePrivateVars
' Author    : beededea
' Date      : 12/02/2025
' Purpose   : initialise vars
'---------------------------------------------------------------------------------------
'
Private Sub initialisePrivateVars()

   On Error GoTo initialisePrivateVars_Error
   
    'TenShillingsxo = 0: TenShillingsyo = 0

   On Error GoTo 0
   Exit Sub

initialisePrivateVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialisePrivateVars of Class Module cwTenShillings"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize for cwTenShillings
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   : Constructor (Initialise) set the inital values of private vars, set up the timers
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()

    On Error GoTo Class_Initialize_Error
    
    Call initialisePrivateVars
    
    mZoom = 0
    
    Set W = Cairo.WidgetBase
        W.BackColor = -1  'we don't use any BackColor here, transparent
        W.ImplementsWheelMessages = True ' enables gsPlivate Sub W_MouseWheel
        
        mZoom = 1 ' the TenShillings size initially very small
        ZoomDirection = mZoomDirection
    
   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwTenShillings"
End Sub
    


'---------------------------------------------------------------------------------------
' Procedure : Widget
' Author    : olaf schmidt
' Date      : 21/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Widget() As cWidgetBase
   On Error GoTo Widget_Error

    Set Widget = W

   On Error GoTo 0
   Exit Property

Widget_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widget of Class Module cwTenShillings"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Widgets
' Author    : olaf schmidt
' Date      : 21/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Widgets() As cWidgets
   On Error GoTo Widgets_Error

    Set Widgets = W.Widgets

   On Error GoTo 0
   Exit Property

Widgets_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widgets of Class Module cwTenShillings"
End Property

'---------------------------------------------------------------------------------------
' Procedure : Zoom
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Zoom() As Double
   On Error GoTo ZoomGet_Error

    Zoom = mZoom

   On Error GoTo 0
   Exit Property

ZoomGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Zoom of Class Module cwTenShillings"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Zoom
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Zoom(ByVal newValue As Double)
   On Error GoTo ZoomLet_Error

    If mZoom <> newValue Then
      mZoom = newValue
    Else
      Exit Property
    End If
    
    If mZoom > 2 Then
        mZoom = 2
    ElseIf mZoom < 0.05 Then
        mZoom = 0.05 'keep FZ within a sane ZoomInterval
    End If
    
    If mZoom > gdResizeRestriction Then mZoom = gdResizeRestriction

    W.Refresh

   On Error GoTo 0
   Exit Property

ZoomLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Zoom of Class Module cwTenShillings"
End Property
'---------------------------------------------------------------------------------------
' Procedure : ZoomDirection
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get ZoomDirection() As String
   On Error GoTo ZoomDirectionGet_Error

  ZoomDirection = mZoomDirection

   On Error GoTo 0
   Exit Property

ZoomDirectionGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ZoomDirection of Class Module cwTenShillings"
End Property
'---------------------------------------------------------------------------------------
' Procedure : ZoomDirection
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let ZoomDirection(ByVal newValue As String)
   On Error GoTo ZoomDirectionLet_Error

   If mZoomDirection <> newValue Then mZoomDirection = newValue Else Exit Property
   W.Refresh

   On Error GoTo 0
   Exit Property

ZoomDirectionLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ZoomDirection of Class Module cwTenShillings"
End Property

'---------------------------------------------------------------------------------------
' Procedure : SkewDegrees
' Author    : beededea
' Date      : 05/09/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get SkewDegrees() As Double
    On Error GoTo SkewDegrees_Error

    SkewDegrees = mSkewDegrees

    On Error GoTo 0
    Exit Property

SkewDegrees_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SkewDegrees of Class Module cwTenShillings"
End Property
'---------------------------------------------------------------------------------------
' Procedure : SkewDegrees
' Author    : beededea
' Date      : 05/09/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let SkewDegrees(ByVal newValue As Double)
    On Error GoTo SkewDegrees_Error

    If mSkewDegrees <> newValue Then mSkewDegrees = newValue Else Exit Property
    W.Refresh

    On Error GoTo 0
    Exit Property

SkewDegrees_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SkewDegrees of Class Module cwTenShillings"
End Property

'---------------------------------------------------------------------------------------
' Procedure : opacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let opacity(ByVal newValue As Double)
   On Error GoTo opacityLet_Error

    If mOpacity <> newValue Then mOpacity = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

opacityLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure opacity of Class Module cwTenShillings"
End Property
'---------------------------------------------------------------------------------------
' Procedure : opacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get opacity() As Double
   On Error GoTo opacityGet_Error

    opacity = mOpacity

   On Error GoTo 0
   Exit Property

opacityGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure opacity of Class Module cwTenShillings"
End Property



'---------------------------------------------------------------------------------------
' Property  : Locked
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property used to store/set the lock state to lock the widget in place, the actual locking is done elsewhere
'---------------------------------------------------------------------------------------
'
Public Property Get Locked() As Boolean
   On Error GoTo LockedGet_Error

   Locked = mLocked

   On Error GoTo 0
   Exit Property

LockedGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwTenShillings"
End Property
'---------------------------------------------------------------------------------------
' Property  : Locked
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property used to store/set the lock state to lock the widget in place, the actual locking is done elsewhere
'---------------------------------------------------------------------------------------
'
Public Property Let Locked(ByVal newValue As Boolean)
   On Error GoTo LockedLet_Error

   If mLocked <> newValue Then mLocked = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

LockedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwTenShillings"
End Property





'---------------------------------------------------------------------------------------
' Property  : Hidden
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property to determine (by value)  the set the hidden value of the whole widget
'---------------------------------------------------------------------------------------
'
Public Property Get Hidden() As Boolean
   On Error GoTo HiddenGet_Error

   Hidden = mHidden

   On Error GoTo 0
   Exit Property

HiddenGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwTenShillings"
End Property

    


'---------------------------------------------------------------------------------------
' Property  : Hidden
' Author    : beededea
' Date      : 10/05/2023
' Purpose   : property to determine (by value)  the set the hidden value of the whole widget
'---------------------------------------------------------------------------------------
'
Public Property Let Hidden(ByVal newValue As Boolean)
   On Error GoTo HiddenLet_Error

    If mHidden <> newValue Then mHidden = newValue Else Exit Property
    If mHidden = True Then
        W.Visible = False
        
        W.Refresh
    Else
        W.Visible = True
        
        W.Refresh
    End If

   On Error GoTo 0
   Exit Property

HiddenLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwTenShillings"
End Property


'---------------------------------------------------------------------------------------
' Property  : imageWidth
' Author    : beededea
' Date      : 10/05/2023
' Purpose   : property to determine (by value)  the set the imageWidth value of the whole widget
'---------------------------------------------------------------------------------------
'
Public Property Let imageWidth(ByVal newValue As Long)
   On Error GoTo imageWidthLet_Error

    If mImageWidth <> newValue Then mImageWidth = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

imageWidthLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property imageWidth of Class Module cwTenShillings"
End Property

'---------------------------------------------------------------------------------------
' Property  : imageHeight
' Author    : beededea
' Date      : 10/05/2023
' Purpose   : property to determine (by value)  the set the imageHeight value of the whole widget
'---------------------------------------------------------------------------------------
'
Public Property Let imageHeight(ByVal newValue As Long)
   On Error GoTo imageHeightLet_Error

    If mImageHeight <> newValue Then mImageHeight = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

imageHeightLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property imageHeight of Class Module cwTenShillings"
End Property
    
'---------------------------------------------------------------------------------------
' Procedure : W_MouseWheel
' Author    : beededea
' Date      : 09/05/2023
' Purpose   : uses ScaleDrawings mZoom, mZoom
'---------------------------------------------------------------------------------------
'
Private Sub W_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal LineOffset As Long, ByVal xAbs As Single, ByVal yAbs As Single)
   On Error GoTo W_MouseWheel_Error

    If gsIgnoreMouse = "1" Then Exit Sub

    Select Case MouseKeys
    Case 8 'change the Zoom only, when the Ctrl-Key is down
        
        If ZoomDirection = "up" Then
            Me.Zoom = Me.Zoom - 0.02 * LineOffset
        Else
            Me.Zoom = Me.Zoom + 0.02 * LineOffset
        End If
        gsWidgetSize = CStr(TenShillingsWidget.Zoom * 100) ' store the value dynamically
        sPutINISetting "Software\TenShillings", "widgetSize", gsWidgetSize, gsSettingsFile
        
        ' saves the current ratios for the RC form
        Call saveRCFormCurrentSizeRatios
        
        If widgetPrefs.IsVisible = True Then
            widgetPrefs.sliWidgetSize.Value = Val(gsWidgetSize)
        End If
        W.Refresh
        
    Case Else
        SkewDegrees = SkewDegrees + LineOffset
        
        If SkewDegrees > 360 Then SkewDegrees = SkewDegrees - 360
        If SkewDegrees < 0 Then SkewDegrees = SkewDegrees + 360
         
        gsSkewDegrees = CStr(TenShillingsWidget.SkewDegrees) ' store the value dynamically
        sPutINISetting "Software\TenShillings", "skewDegrees", gsSkewDegrees, gsSettingsFile
        
        If widgetPrefs.IsVisible = True Then
            widgetPrefs.sliSkewDegrees.Value = Val(gsSkewDegrees)
        End If
    End Select

   On Error GoTo 0
   Exit Sub

W_MouseWheel_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_MouseWheel of Class Module cwTenShillings"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : W_MouseDown
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
   On Error GoTo W_MouseDown_Error

    If Button = vbRightButton Then
        Call menuForm.PopupMenu(menuForm.mnuMainMenu)
    Else
        If gsIgnoreMouse = "1" Then Exit Sub
        ' do whatever you want with a mouseDown here, not doing anything at the moment but we will...
    End If


   On Error GoTo 0
   Exit Sub

W_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_MouseDown of Class Module cwTenShillings"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : W_Paint
' Author    :
' Date      : 01/08/2023
' Purpose   : Paint event for RC ten shillings class
'             this only works with a fully square image
'             form width being 1894, image width being square 789x789px, a PNG with 5 pixels of padding left/right and 175px top/bottom
'             note: Default origin point is always top left
'             note: It is good practice to put all modifications to the graphics state between save()/restore() function call
'---------------------------------------------------------------------------------------
'
Private Sub W_Paint(CC As vbRichClient5.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
    Dim lReverseOrigin As Long: lReverseOrigin = 0
    Dim lImageXoffset As Long: lImageXoffset = 0
    Dim lImageYoffset As Long: lImageYoffset = 0
    Dim cXPadding As Long
    Dim cYPadding As Long
    Dim heightpadding As Long
    
    On Error GoTo W_Paint_Error
    
    ' set the top left origin point prior to rotation, nudging the image into the form a little 150px to prevent cropping on the left/upper parts when rotating
    cXPadding = 120
    cYPadding = 50
    
    lImageXoffset = mImageWidth + cXPadding ' the 150 padding will do for top and left gap
    heightpadding = (cYPadding * (gdResizeRestriction))
    lImageYoffset = mImageHeight + heightpadding
    
    ' save the current image state of a Cairo::Context created so far, saved to an internal stack of saved states, otherwise nothing will show
    ' This allows you to make temporary changes for a specific drawing operation and then restore the previous state later using CC.restore
    'CC.Save
    
    ' translate or slide the image over from the r/h side, the full size width of the image so that the top left of the image is at the top left of the form, complete with padding
    CC.TranslateDrawings lImageXoffset, lImageYoffset

    ' set the scaling ratio for all images
    CC.ScaleDrawings mZoom, mZoom

    ' set the rotation in degrees only for the coming render of the tenshillings image
    CC.RotateDrawingsDeg mSkewDegrees
    
    ' reverse the origin point prior to rotation, sliding the top left in the correct position
    lReverseOrigin = (mImageHeight / 2) - (mImageHeight) ' height or width does not matter, the image is square
    CC.TranslateDrawings lReverseOrigin, lReverseOrigin

    ' from the global ImageList render the tenshillings image in relation to the reversed translated origin above
    CC.RenderSurfaceContent "tenshillings", 0, 0, , , , mOpacity

    ' restores to the state saved by a preceding call to CC.Save
    'CC.Restore
    
    On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwTenShillings"
End Sub
 


