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

'Private TenShillingsxo As Integer
'Private TenShillingsyo  As Integer

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Private mZoom As Double
Private mZoomDirection As String
Private mSkewDegrees As Double

Private mLocked As Boolean ' property Locked
Private mHidden As Boolean ' property Hidden
Private mImageWidth As Long
Private mImageHeight As Long
Private mOpacity As Double





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
    'Dim intervalValue As Integer: intervalValue = 0
    
    On Error GoTo Class_Initialize_Error
    
    Call initialisePrivateVars
    
    mZoom = 0

    'TenShillingsxo = 15: TenShillingsyo = 27

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

  If mZoom <> newValue Then mZoom = newValue Else Exit Property
  If mZoom > 2 Then mZoom = 2 Else If mZoom < 0.05 Then mZoom = 0.05 'keep FZ within a sane ZoomInterval
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
' Author    : olaf schmidt
' Date      : 01/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub W_Paint(CC As vbRichClient5.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
    Dim reverseOrigin As Long: reverseOrigin = 0
    
    On Error GoTo W_Paint_Error
 
    reverseOrigin = (mImageHeight / 2) - (mImageHeight)
    
    If gsRichClientEnvironment = "RC6" Then
        ' Widgets are refreshed "from the ground up" in RC5 and not so in RC6
        ' extra code NOT needed when using RC5 but now required in RC6 - starts
        CC.Operator = CAIRO_OPERATOR_CLEAR
        CC.Paint 'clear the whole background of this Widgets underlying Surface via Clear-Operator
        CC.Operator = CAIRO_OPERATOR_OVER 'switch back to the default-Operator
        ' extra code not needed when using RC5 but now required in RC6 - ends
    End If
    
    ' note: default origin point is always top left, needs to be adjusted
    
    ' move the image into the form a little to prevent cropping when rotating
    CC.TranslateDrawings 150, 150

    ' set the origin point prior to rotation
    CC.TranslateDrawings mImageHeight, mImageHeight
    
    ' set the scaling ratio for all images
    CC.ScaleDrawings mZoom, mZoom
    
    ' save the current state created so far
    CC.Save
    
    ' set the rotation in degrees only for the coming render of the tenshillings image
    CC.RotateDrawingsDeg SkewDegrees
    
    ' reverse the origin point prior to rotation
    CC.TranslateDrawings reverseOrigin, reverseOrigin
    
    'render the tenshillings in relation to the reversed translated origin above
    CC.RenderSurfaceContent "tenshillings", 0, 0, , , , mOpacity
    
    ' restores to the state saved by a preceding call to CC.Save
    CC.Restore
    
    On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwTenShillings"
End Sub
 

