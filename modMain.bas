Attribute VB_Name = "modMain"
'---------------------------------------------------------------------------------------
' Module    : modMain
' Author    : beededea
' Date      : 22/01/2025
' Purpose   :
'---------------------------------------------------------------------------------------

'@IgnoreModule BooleanAssignedInIfElse, IntegerDataType, ModuleWithoutFolder


Option Explicit

'------------------------------------------------------ STARTS
' for SetWindowPos z-ordering
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOP As Long = 0 ' for SetWindowPos z-ordering
Public Const HWND_TOPMOST As Long = -1
Public Const HWND_BOTTOM As Long = 1
Private Const SWP_NOMOVE  As Long = &H2
Private Const SWP_NOSIZE  As Long = &H1
Public Const OnTopFlags  As Long = SWP_NOMOVE Or SWP_NOSIZE
'------------------------------------------------------ ENDS

'------------------------------------------------------ STARTS

' class objects instantiated
Public fMain As New cfMain
Public aboutWidget As cwAbout
Public helpWidget As cwHelp
Public licenceWidget As cwLicence
Public tenShillingsOverlay As cwTenShillingsOverlay

' any other private vars for public properties
Private m_sgsWidgetName As String
'------------------------------------------------------ ENDS


'---------------------------------------------------------------------------------------
' Procedure : Main
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : Program's entry point
'---------------------------------------------------------------------------------------
'
Private Sub Main()
   On Error GoTo Main_Error
    
   Call mainRoutine(False)

   On Error GoTo 0
   Exit Sub

Main_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Main of Module modMain"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : main_routine
' Author    : beededea
' Date      : 27/06/2023
' Purpose   : called by sub main() to allow this routine to be called elsewhere,
'             a reload for example.
'---------------------------------------------------------------------------------------
'
Public Sub mainRoutine(ByVal restart As Boolean)

    Dim extractCommand As String: extractCommand = vbNullString
    Dim licenceState As Integer: licenceState = 0

    On Error GoTo main_routine_Error
    
    ' initialise global vars
    Call initialiseGlobalVars
    
    gbStartupFlg = True
    gsWidgetName = "TenShillings Widget"
    
    extractCommand = Command$ ' capture any parameter passed, remove if a soft reload
    If restart = True Then extractCommand = vbNullString
    
    'Test for the coding environment and set a global variable to alter conditions throughout, mostly in text
    Call testForCodingEnvironment
    
    ' Test for the version of RichClient and set a global variable to alter conditions throughout, mostly in text, there is no new_c.version in RC5
    Call testForRichClientVersion
      
    ' Load the sounds into numbered buffers ready for playing, not currently required, this particular widget does not have complex sound usage
    'Call loadAsynchSoundFiles
    
    ' resolve VB6 sizing width bug
    Call determineScreenDimensions

    'add Resources to the global ImageList
    Call addImagesToImageList
    
    ' check the Windows version
    gbClassicThemeCapable = fTestClassicThemeCapable
  
    ' get this tool's entry in the trinkets settings file and assign the app.path
    Call getTrinketsFile
    
    ' get the location of this tool's settings file (appdata)
    Call getToolSettingsFile
    
    ' read the widget settings from the new configuration file
    Call readSettingsFile("Software\TenShillings", gsSettingsFile)
    
    ' validate the inputs of any data from the input settings file
    Call validateInputs

    ' check first usage via licence acceptance value and then set initial DPI awareness
    Call setAutomaticDPIState(licenceState)
            
    ' initialise and create the three main RC forms (widget, about and licence) on the current display
    Call createRCFormsOnCurrentDisplay
        
    ' place the form at the saved location and configure all the form elements
    Call makeVisibleFormElements
        
    ' run the functions that are also called at reload time.
    Call adjustMainControls(licenceState) ' this needs to be here after the initialisation of the Cairo forms and widgets
    
    ' move/hide onto/from the main screen
    Call mainScreen
        
    ' if the program is run in unhide mode, write the settings and exit
    Call handleUnhideMode(extractCommand)
    
    'load the preferences form but don't yet show it, speeds up access to the prefs via the menu
    Call loadPreferenceForm
    
    ' if the parameter states re-open prefs then shows the prefs
    If extractCommand = "prefs" Then Call makeProgramPreferencesAvailable

    'load the message form but don't yet show it, speeds up access to the message form when needed.
    Load frmMessage
    
    ' display licence screen on first usage
    Call showLicence(fLicenceState)
    
    ' make the prefs appear on the first time running
    Call checkFirstTime
 
    ' configure any global timers here
    Call configureTimers
    
    ' note the monitor primary at the preferences form_load and store as glOldwidgetFormMonitorPrimary
    Call identifyPrimaryMonitor
    
    ' make the busy sand timer invisible
    Call hideBusyTimer
            
    'subclassed the widget form to generate a balloon tooltip
    If Not InIde Then Call SubclassForm(fMain.TenShillingsForm.hwnd, ObjPtr(fMain.TenShillingsForm))
    
    ' end the startup by un-setting the start global flag
    gbStartupFlg = False
    gbReload = False
    
    ' RC message pump will auto-exit when Cairo Forms > 0 so we run it only when 0, this prevents message interruption
    ' when running twice on reload. Do not move this line.
    #If twinbasic Then
        Cairo.WidgetForms.EnterMessageLoop
    #Else
        If restart = False Then Cairo.WidgetForms.EnterMessageLoop
    #End If
    
    ' don't put anything here, place it above the Cairo.WidgetForms.EnterMessageLoop
    
    ' NOTE: the final act in startup is the form_resize_event that is triggered by the subclassed WM_EXITSIZEMOVE when the form is finally revealed
     
   On Error GoTo 0
   Exit Sub

main_routine_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure main_routine of Module modMain at "
    
End Sub
 
 

 
'---------------------------------------------------------------------------------------
' Procedure : testForRichClientVersion
' Author    : beededea
' Date      : 14/10/2025
' Purpose   : Test for the version of RichClient and set a global variable to alter conditions throughout, mostly in text, there is no new_c.version in RC5
'---------------------------------------------------------------------------------------
'
Private Sub testForRichClientVersion()

    On Error GoTo testForRichClientVersion_Error

    If fFExists(App.Path & "\BIN\vbRichClient5.dll") Then
        gsRichClientEnvironment = "RC5"
    ElseIf fFExists(App.Path & "\BIN\RC6.dll") Then
        gsRichClientEnvironment = "RC6"
    End If

    On Error GoTo 0
    Exit Sub

testForRichClientVersion_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure testForRichClientVersion of Module modMain"

End Sub
 
 
'---------------------------------------------------------------------------------------
' Procedure : testForCodingEnvironment
' Author    : beededea
' Date      : 14/10/2025
' Purpose   : Test for the coding environment and set a global variable to alter conditions throughout, mostly in text
'---------------------------------------------------------------------------------------
'
Private Sub testForCodingEnvironment()

    On Error GoTo testForCodingEnvironment_Error

    #If twinbasic Then
        gsCodingEnvironment = "TwinBasic"
    #Else
        gsCodingEnvironment = "VB6"
    #End If

    On Error GoTo 0
    Exit Sub

testForCodingEnvironment_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure testForCodingEnvironment of Module modMain"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : loadPreferenceForm
' Author    : beededea
' Date      : 20/02/2025
' Purpose   : load the preferences form but don't yet show it, speeds up access to the prefs via the menu
'---------------------------------------------------------------------------------------
'
Private Sub loadPreferenceForm()
        
   On Error GoTo loadPreferenceForm_Error

    If widgetPrefs.IsLoaded = False Then
        Load widgetPrefs
        gbPrefsFormResizedInCode = True
        Call widgetPrefs.PrefsFormResizeEvent
    End If

   On Error GoTo 0
   Exit Sub

loadPreferenceForm_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadPreferenceForm of Module modMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : setAutomaticDPIState
' Author    : beededea
' Date      : 20/02/2025
' Purpose   : check first usage via licence acceptance value and then set initial DPI awareness
'---------------------------------------------------------------------------------------
'
Private Sub setAutomaticDPIState(ByRef licenceState As Integer)
   On Error GoTo setAutomaticDPIState_Error

    licenceState = fLicenceState()
    If licenceState = 0 Then
        Call testDPIAndSetInitialAwareness ' determine High DPI awareness or not by default on first run
    Else
        Call setDPIaware ' determine the user settings for DPI awareness, for this program and all its forms.
    End If

   On Error GoTo 0
   Exit Sub

setAutomaticDPIState_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setAutomaticDPIState of Module modMain"
End Sub
 
'
'---------------------------------------------------------------------------------------
' Procedure : identifyPrimaryMonitor
' Author    : beededea
' Date      : 20/02/2025
' Purpose   : note the monitor primary at the main form_load and store as glOldwidgetFormMonitorPrimary - will be resampled regularly later and compared
'---------------------------------------------------------------------------------------
'
Private Sub identifyPrimaryMonitor()
    Dim widgetFormMonitorID As Long: widgetFormMonitorID = 0
    
    On Error GoTo identifyPrimaryMonitor_Error

    gWidgetMonitorStruct = cWidgetFormScreenProperties(fMain.TenShillingsForm, widgetFormMonitorID)
    glOldWidgetFormMonitorPrimary = gWidgetMonitorStruct.IsPrimary

    On Error GoTo 0
    Exit Sub

identifyPrimaryMonitor_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure identifyPrimaryMonitor of Module modMain"
End Sub
 

 


'---------------------------------------------------------------------------------------
' Procedure : checkFirstTime
' Author    : beededea
' Date      : 12/05/2023
' Purpose   : check for first time running, first time run shows prefs
'---------------------------------------------------------------------------------------
'
Private Sub checkFirstTime()

   On Error GoTo checkFirstTime_Error

    If gsFirstTimeRun = "true" Then
        Call makeProgramPreferencesAvailable
        gsFirstTimeRun = "false"
        sPutINISetting "Software\TenShillings", "firstTimeRun", gsFirstTimeRun, gsSettingsFile
    End If

   On Error GoTo 0
   Exit Sub

checkFirstTime_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkFirstTime of Module modMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : initialiseGlobalVars
' Author    : beededea
' Date      : 12/05/2023
' Purpose   : initialise global vars
'---------------------------------------------------------------------------------------
'
Private Sub initialiseGlobalVars()
      
    On Error GoTo initialiseGlobalVars_Error
    
    glMonitorCount = 0

    ' general
    gsStartup = vbNullString
    gsWidgetFunctions = vbNullString
    gsPointerAnimate = vbNullString
    gsSamplingInterval = vbNullString

    ' config
    gsWidgetTooltips = vbNullString
    gsPrefsTooltips = vbNullString
    
    gsShowTaskbar = vbNullString
    gsShowHelp = vbNullString

    
    gsDpiAwareness = vbNullString
    
    gsWidgetSize = vbNullString
    gsSkewDegrees = vbNullString
    gsScrollWheelDirection = vbNullString
    
    ' position
    gsAspectHidden = vbNullString
    gsWidgetPosition = vbNullString
    gsWidgetLandscape = vbNullString
    gsWidgetPortrait = vbNullString
    gsLandscapeFormHoffset = vbNullString
    gsLandscapeFormVoffset = vbNullString
    gsPortraitHoffset = vbNullString
    gsPortraitYoffset = vbNullString
    gsVLocationPercPrefValue = vbNullString
    gsHLocationPercPrefValue = vbNullString
    
    ' sounds
    gsEnableSounds = vbNullString
    
    ' development
    gsDebug = vbNullString
    gsDblClickCommand = vbNullString
    gsOpenFile = vbNullString
    gsDefaultVB6Editor = vbNullString
    gsDefaultTBEditor = vbNullString
         
    ' font
    gsClockFont = vbNullString
    gsWidgetFont = vbNullString
    gsPrefsFont = vbNullString
    gsPrefsFontSizeHighDPI = vbNullString
    gsPrefsFontSizeLowDPI = vbNullString
    gsPrefsFontItalics = vbNullString
    gsPrefsFontColour = vbNullString
    
    gsDisplayScreenFont = vbNullString
    gsDisplayScreenFontSize = vbNullString
    gsDisplayScreenFontItalics = vbNullString
    gsDisplayScreenFontColour = vbNullString
    
    ' window
    gsWindowLevel = vbNullString
    gsPreventDragging = vbNullString
    gsOpacity = vbNullString
    gsWidgetHidden = vbNullString
    gsHidingTime = vbNullString
    gsIgnoreMouse = vbNullString
    gsFormVisible = vbNullString
    
    gbMenuOccurred = False ' bool
    gsFirstTimeRun = vbNullString
    gsMultiMonitorResize = vbNullString
    
    ' general storage variables declared
    gsSettingsDir = vbNullString
    gsSettingsFile = vbNullString
    
    gsTrinketsDir = vbNullString
    gsTrinketsFile = vbNullString
    
    gsWidgetHighDpiXPos = vbNullString
    gsWidgetHighDpiYPos = vbNullString
    
    gsWidgetLowDpiXPos = vbNullString
    gsWidgetLowDpiYPos = vbNullString
    
    gsLastSelectedTab = vbNullString
    gsSkinTheme = vbNullString
    
    ' general variables declared
    'toolSettingsFile = vbNullString
    gbClassicThemeCapable = False
    glStoreThemeColour = 0
    'windowsVer = vbNullString
    
    ' vars to obtain correct screen width (to correct VB6 bug) STARTS
    glScreenTwipsPerPixelX = 0
    glScreenTwipsPerPixelY = 0
    glPhysicalScreenWidthTwips = 0
    glPhysicalScreenHeightTwips = 0
    glPhysicalScreenHeightPixels = 0
    glPhysicalScreenWidthPixels = 0
    
    glVirtualScreenHeightPixels = 0
    glVirtualScreenWidthPixels = 0
    
    glOldPhysicalScreenHeightPixels = 0
    glOldPhysicalScreenWidthPixels = 0
    
    gsPrefsPrimaryHeightTwips = vbNullString
    gsPrefsSecondaryHeightTwips = vbNullString
    gsWidgetPrimaryHeightRatio = vbNullString
    gsWidgetSecondaryHeightRatio = vbNullString
    
    gsMessageAHeightTwips = vbNullString
    gsMessageAWidthTwips = vbNullString
    
    ' key presses
    gbCTRL_1 = False
    gbSHIFT_1 = False
    
    ' other globals
    giDebugFlg = 0
    giMinutesToHide = 0
    gsAspectRatio = vbNullString
    gsCodingEnvironment = vbNullString
    gsRichClientEnvironment = vbNullString
    
    gdResizeRestriction = 0
    
   On Error GoTo 0
   Exit Sub

initialiseGlobalVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialiseGlobalVars of Module modMain"
    
End Sub

        
'---------------------------------------------------------------------------------------
' Procedure : addImagesToImageList
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : add image Resources to the global ImageList
'---------------------------------------------------------------------------------------
'
Private Sub addImagesToImageList()
    
    On Error GoTo addImagesToImageList_Error

'    add Resources to the global ImageList that are not being pulled from the PSD directly
    
    Cairo.ImageList.AddImage "about", App.Path & "\Resources\images\about.png"
    Cairo.ImageList.AddImage "licence", App.Path & "\Resources\images\frame.png"
    Cairo.ImageList.AddImage "help", App.Path & "\Resources\images\tenShillingsHelp.png"
    Cairo.ImageList.AddImage "frmIcon", App.Path & "\Resources\images\Icon.png"
    
    'add Resources to the global ImageList
    
    Cairo.ImageList.AddImage "tenshillings", App.Path & "\Resources\images\TenShillings.png"

    ' NOTE: cannot yet add the above images to the GDIP imageList yet as the Cairo functions extract their images directly from RC's own Cairo-based imageList
    
    
    ' addition of the Prefs tab Jpeg icon images to the GDIP imageList dictionary, previously used Cairo.ImageList.AddImage "filename"
    
    ' normal images
    thisImageList.AddImage "about-icon-dark", App.Path & "\Resources\images\about-icon-dark-1010.jpg"
    thisImageList.AddImage "about-icon-light", App.Path & "\Resources\images\about-icon-light-1010.jpg"
    thisImageList.AddImage "config-icon-dark", App.Path & "\Resources\images\config-icon-dark-1010.jpg"
    thisImageList.AddImage "config-icon-light", App.Path & "\Resources\images\config-icon-light-1010.jpg"
    thisImageList.AddImage "development-icon-light", App.Path & "\Resources\images\development-icon-light-1010.jpg"
    thisImageList.AddImage "development-icon-dark", App.Path & "\Resources\images\development-icon-dark-1010.jpg"
    thisImageList.AddImage "general-icon-dark", App.Path & "\Resources\images\general-icon-dark-1010.jpg"
    thisImageList.AddImage "general-icon-light", App.Path & "\Resources\images\general-icon-light-1010.jpg"
    thisImageList.AddImage "sounds-icon-light", App.Path & "\Resources\images\sounds-icon-light-1010.jpg"
    thisImageList.AddImage "sounds-icon-dark", App.Path & "\Resources\images\sounds-icon-dark-1010.jpg"
    thisImageList.AddImage "windows-icon-light", App.Path & "\Resources\images\windows-icon-light-1010.jpg"
    thisImageList.AddImage "windows-icon-dark", App.Path & "\Resources\images\windows-icon-dark-1010.jpg"
    thisImageList.AddImage "font-icon-dark", App.Path & "\Resources\images\font-icon-dark-1010.jpg"
    thisImageList.AddImage "font-icon-light", App.Path & "\Resources\images\font-icon-light-1010.jpg"
    thisImageList.AddImage "position-icon-light", App.Path & "\Resources\images\position-icon-light-1010.jpg"
    thisImageList.AddImage "position-icon-dark", App.Path & "\Resources\images\position-icon-dark-1010.jpg"
    
    
    ' clicked images
    thisImageList.AddImage "general-icon-dark-clicked", App.Path & "\Resources\images\general-icon-dark-600-clicked.jpg"
    thisImageList.AddImage "config-icon-dark-clicked", App.Path & "\Resources\images\config-icon-dark-600-clicked.jpg"
    thisImageList.AddImage "font-icon-dark-clicked", App.Path & "\Resources\images\font-icon-dark-600-clicked.jpg"
    thisImageList.AddImage "sounds-icon-dark-clicked", App.Path & "\Resources\images\sounds-icon-dark-600-clicked.jpg"
    thisImageList.AddImage "position-icon-dark-clicked", App.Path & "\Resources\images\position-icon-dark-600-clicked.jpg"
    thisImageList.AddImage "development-icon-dark-clicked", App.Path & "\Resources\images\development-icon-dark-600-clicked.jpg"
    thisImageList.AddImage "windows-icon-dark-clicked", App.Path & "\Resources\images\windows-icon-dark-600-clicked.jpg"
    thisImageList.AddImage "about-icon-dark-clicked", App.Path & "\Resources\images\about-icon-dark-600-clicked.jpg"
    thisImageList.AddImage "general-icon-light-clicked", App.Path & "\Resources\images\general-icon-light-600-clicked.jpg"
    thisImageList.AddImage "config-icon-light-clicked", App.Path & "\Resources\images\config-icon-light-600-clicked.jpg"
    thisImageList.AddImage "font-icon-light-clicked", App.Path & "\Resources\images\font-icon-light-600-clicked.jpg"
    thisImageList.AddImage "sounds-icon-light-clicked", App.Path & "\Resources\images\sounds-icon-light-600-clicked.jpg"
    thisImageList.AddImage "position-icon-light-clicked", App.Path & "\Resources\images\position-icon-light-600-clicked.jpg"
    thisImageList.AddImage "development-icon-light-clicked", App.Path & "\Resources\images\development-icon-light-600-clicked.jpg"
    thisImageList.AddImage "windows-icon-light-clicked", App.Path & "\Resources\images\windows-icon-light-600-clicked.jpg"
    thisImageList.AddImage "about-icon-light-clicked", App.Path & "\Resources\images\about-icon-light-600-clicked.jpg"
    
    ' load the icon images on the message form to the image list
    Call loadMessageIconImages
    
   On Error GoTo 0
   Exit Sub

addImagesToImageList_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure addImagesToImageList of Module modMain, an image has probably been accidentally deleted from the resources/images folder."

End Sub


'---------------------------------------------------------------------------------------
' Procedure : loadMessageIconImages
' Author    : beededea
' Date      : 22/06/2023
' Purpose   : set the icon images on the message form
'---------------------------------------------------------------------------------------
'
Private Sub loadMessageIconImages()
    
    Dim resourcePath As String: resourcePath = vbNullString
    
    On Error GoTo loadMessageIconImages_Error
    
    resourcePath = App.Path & "\resources\images"
    
    thisImageList.AddImage "windowsInformation1920", resourcePath & "\windowsInformation1920.jpg"
    thisImageList.AddImage "windowsOrangeExclamation1920", resourcePath & "\windowsOrangeExclamation1920.jpg"
    thisImageList.AddImage "windowsShieldQMark1920", resourcePath & "\windowsShieldQMark1920.jpg"
    thisImageList.AddImage "windowsCritical1920", resourcePath & "\windowsCritical1920.jpg"
    
   On Error GoTo 0
   Exit Sub

loadMessageIconImages_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadMessageIconImages of Form frmMessage"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : adjustMainControls
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : called at runtime and on restart, sets the characteristics of the widget, individual controls and menus
'---------------------------------------------------------------------------------------
'
Public Sub adjustMainControls(Optional ByVal licenceState As Integer)
   Dim thisEditor As String: thisEditor = vbNullString
   Dim bigScreen As Long: bigScreen = 3840
   
   On Error GoTo adjustMainControls_Error

    ' validate the inputs of any data from the input settings file
    Call validateInputs
    
    ' initial call just to obtain initial physical screen monitor ID
    Call resizeLocateRCFormByMoveToNewMonitor
        
    ' if the licenstate is 0 then the program is running for the first time, so pre-size the form to fit larger screens
    If licenceState = 0 Then
        ' the widget displays at 100% at a screen width of 3840 pixels
        If glPhysicalScreenWidthPixels >= bigScreen Then
            gsWidgetSize = CStr((glPhysicalScreenWidthPixels / bigScreen) * 100)
        End If
    End If
    
    menuForm.mnuAbout.Caption = "About TenShillings " & gsRichClientEnvironment & " Cairo " & gsCodingEnvironment & " widget"
    
'    ' Set the opacity of the widget, passing just this one global variable to a public property within the class
    tenShillingsOverlay.Opacity = Val(gsOpacity) / 100
    
    tenShillingsOverlay.SkewDegrees = CDbl(gsSkewDegrees)
    
    ' set the initial size
    If glMonitorCount > 1 And (LTrim$(gsMultiMonitorResize) = "1" Or LTrim$(gsMultiMonitorResize) = "2") Then
        If gWidgetMonitorStruct.IsPrimary = True Then
            tenShillingsOverlay.Zoom = (Val(gsWidgetPrimaryHeightRatio))
        Else
            tenShillingsOverlay.Zoom = (Val(gsWidgetSecondaryHeightRatio))
        End If
    Else
        tenShillingsOverlay.Zoom = Val(gsWidgetSize) / 100
    End If
    
    If gsWidgetFunctions = "1" Then
        menuForm.mnuSwitchOff.Checked = False
        menuForm.mnuTurnFunctionsOn.Checked = True
    Else
        menuForm.mnuSwitchOff.Checked = True
        menuForm.mnuTurnFunctionsOn.Checked = False
    End If
    
    If gsDebug = "1" Then
        #If twinbasic Then
            If gsDefaultTBEditor <> vbNullString Then thisEditor = gsDefaultTBEditor
        #Else
            If gsDefaultVB6Editor <> vbNullString Then thisEditor = gsDefaultVB6Editor
        #End If
        
        menuForm.mnuEditWidget.Caption = "Edit Widget using " & thisEditor
        menuForm.mnuEditWidget.Visible = True
    Else
        menuForm.mnuEditWidget.Visible = False
    End If
    
    '  fMain.TenShillingsForm.ShowInTaskbar = Not (gsShowTaskbar = "0") ' no!
    
    If gsShowTaskbar = "0" Then
        fMain.TenShillingsForm.ShowInTaskbar = False
    Else
        fMain.TenShillingsForm.ShowInTaskbar = True
    End If
    
    ' set the visibility and characteristics of the interactive areas
    ' the alpha is already set to zero for all layers found in the PSD, we now turn them back on as we require
        
    If gsDebug = "1" Then
        #If twinbasic Then
            If gsDefaultTBEditor <> vbNullString Then thisEditor = gsDefaultTBEditor
        #Else
            If gsDefaultVB6Editor <> vbNullString Then thisEditor = gsDefaultVB6Editor
        #End If
        
        menuForm.mnuEditWidget.Caption = "Edit Widget using " & thisEditor
        menuForm.mnuEditWidget.Visible = True
    Else
        menuForm.mnuEditWidget.Visible = False
    End If
        
   ' set the lock state of the widget
   If gsPreventDragging = "0" Then
        menuForm.mnuLockWidget.Checked = False
        tenShillingsOverlay.Locked = False
    Else
        menuForm.mnuLockWidget.Checked = True
        tenShillingsOverlay.Locked = True ' this is just here for continuity's sake, it is also set at the time the control is selected
    End If
    
    tenShillingsOverlay.Opacity = Val(gsOpacity) / 100

    ' set the z-ordering of the window
    Call setAlphaFormZordering
    
    ' set the tooltips on the main screen
    Call setRichClientTooltips
    
    ' set the hiding time for the hiding timer, can't read the minutes from comboxbox as the prefs isn't yet open
    Call setHidingTime

    If giMinutesToHide > 0 Then menuForm.mnuHideWidget.Caption = "Hide Widget for " & giMinutesToHide & " min."
    
    ' refresh the form in order to show the above changes immediately
    tenShillingsOverlay.Widget.Refresh

   On Error GoTo 0
   Exit Sub

adjustMainControls_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjustMainControls of Module modMain " _
        & " Most likely one of the layers above is named incorrectly."

End Sub

'---------------------------------------------------------------------------------------
' Procedure : setAlphaFormZordering
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : set the z-ordering of the window
'---------------------------------------------------------------------------------------
'
Public Sub setAlphaFormZordering()

   On Error GoTo setAlphaFormZordering_Error

    If Val(gsWindowLevel) = 0 Then
        Call SetWindowPos(fMain.TenShillingsForm.hwnd, HWND_BOTTOM, 0&, 0&, 0&, 0&, OnTopFlags)
    ElseIf Val(gsWindowLevel) = 1 Then
        Call SetWindowPos(fMain.TenShillingsForm.hwnd, HWND_TOP, 0&, 0&, 0&, 0&, OnTopFlags)
    ElseIf Val(gsWindowLevel) = 2 Then
        Call SetWindowPos(fMain.TenShillingsForm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, OnTopFlags)
    End If

   On Error GoTo 0
   Exit Sub

setAlphaFormZordering_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setAlphaFormZordering of Module modMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : readSettingsFile
' Author    : beededea
' Date      : 12/05/2020
' Purpose   : read the application's setting file and assign values to public vars
'---------------------------------------------------------------------------------------
'
Public Sub readSettingsFile(ByVal Location As String, ByVal gsSettingsFile As String)
    On Error GoTo readSettingsFile_Error

    If fFExists(gsSettingsFile) Then
        
        ' general
        gsStartup = fGetINISetting(Location, "startup", gsSettingsFile)
        gsWidgetFunctions = fGetINISetting(Location, "widgetFunctions", gsSettingsFile)
        gsPointerAnimate = fGetINISetting(Location, "pointerAnimate", gsSettingsFile)
        gsSamplingInterval = fGetINISetting(Location, "samplingInterval", gsSettingsFile)
        
        ' configuration
        gsWidgetTooltips = fGetINISetting(Location, "widgetTooltips", gsSettingsFile)
        gsPrefsTooltips = fGetINISetting(Location, "prefsTooltips", gsSettingsFile)
        
        gsShowTaskbar = fGetINISetting(Location, "showTaskbar", gsSettingsFile)
        gsShowHelp = fGetINISetting(Location, "showHelp", gsSettingsFile)
        gsDpiAwareness = fGetINISetting(Location, "dpiAwareness", gsSettingsFile)
        gsWidgetSize = fGetINISetting(Location, "widgetSize", gsSettingsFile)
        gsSkewDegrees = fGetINISetting(Location, "skewDegrees", gsSettingsFile)
        
        gsScrollWheelDirection = fGetINISetting(Location, "scrollWheelDirection", gsSettingsFile)
        
        ' position
        gsAspectHidden = fGetINISetting(Location, "aspectHidden", gsSettingsFile)
        gsWidgetPosition = fGetINISetting(Location, "widgetPosition", gsSettingsFile)
        gsWidgetLandscape = fGetINISetting(Location, "widgetLandscape", gsSettingsFile)
        gsWidgetPortrait = fGetINISetting(Location, "widgetPortrait", gsSettingsFile)
        gsLandscapeFormHoffset = fGetINISetting(Location, "landscapeHoffset", gsSettingsFile)
        gsLandscapeFormVoffset = fGetINISetting(Location, "landscapeYoffset", gsSettingsFile)
        gsPortraitHoffset = fGetINISetting(Location, "portraitHoffset", gsSettingsFile)
        gsPortraitYoffset = fGetINISetting(Location, "portraitYoffset", gsSettingsFile)
        gsVLocationPercPrefValue = fGetINISetting(Location, "vLocationPercPrefValue", gsSettingsFile)
        gsHLocationPercPrefValue = fGetINISetting(Location, "hLocationPercPrefValue", gsSettingsFile)

        ' font
        gsClockFont = fGetINISetting(Location, "clockFont", gsSettingsFile)
        gsWidgetFont = fGetINISetting(Location, "widgetFont", gsSettingsFile)
        gsPrefsFont = fGetINISetting(Location, "prefsFont", gsSettingsFile)
        gsPrefsFontSizeHighDPI = fGetINISetting(Location, "prefsFontSizeHighDPI", gsSettingsFile)
        gsPrefsFontSizeLowDPI = fGetINISetting(Location, "prefsFontSizeLowDPI", gsSettingsFile)
        gsPrefsFontItalics = fGetINISetting(Location, "prefsFontItalics", gsSettingsFile)
        gsPrefsFontColour = fGetINISetting(Location, "prefsFontColour", gsSettingsFile)
    
        gsDisplayScreenFont = fGetINISetting(Location, "displayScreenFont", gsSettingsFile)
        gsDisplayScreenFontSize = fGetINISetting(Location, "displayScreenFontSize", gsSettingsFile)
        gsDisplayScreenFontItalics = fGetINISetting(Location, "displayScreenFontItalics", gsSettingsFile)
        gsDisplayScreenFontColour = fGetINISetting(Location, "displayScreenFontColour", gsSettingsFile)
       
        ' sound
'        gsEnableSounds = fGetINISetting(Location, "enableSounds", gsSettingsFile)

        
        ' development
        gsDebug = fGetINISetting(Location, "debug", gsSettingsFile)
        gsDblClickCommand = fGetINISetting(Location, "dblClickCommand", gsSettingsFile)
        gsOpenFile = fGetINISetting(Location, "openFile", gsSettingsFile)
        gsDefaultVB6Editor = fGetINISetting(Location, "defaultVB6Editor", gsSettingsFile)
        gsDefaultTBEditor = fGetINISetting(Location, "defaultTBEditor", gsSettingsFile)
        
        ' other
        gsWidgetHighDpiXPos = fGetINISetting("Software\TenShillings", "widgetHighDpiXPos", gsSettingsFile)
        gsWidgetHighDpiYPos = fGetINISetting("Software\TenShillings", "widgetHighDpiYPos", gsSettingsFile)
        gsWidgetLowDpiXPos = fGetINISetting("Software\TenShillings", "widgetLowDpiXPos", gsSettingsFile)
        gsWidgetLowDpiYPos = fGetINISetting("Software\TenShillings", "widgetLowDpiYPos", gsSettingsFile)
        gsLastSelectedTab = fGetINISetting(Location, "lastSelectedTab", gsSettingsFile)
        gsSkinTheme = fGetINISetting(Location, "skinTheme", gsSettingsFile)
        
        ' window
        gsWindowLevel = fGetINISetting(Location, "windowLevel", gsSettingsFile)
        gsPreventDragging = fGetINISetting(Location, "preventDragging", gsSettingsFile)
        gsOpacity = fGetINISetting(Location, "opacity", gsSettingsFile)
        
        ' we do not want the widget to hide at startup
        gsWidgetHidden = "0"
        
        gsHidingTime = fGetINISetting(Location, "hidingTime", gsSettingsFile)
        gsIgnoreMouse = fGetINISetting(Location, "ignoreMouse", gsSettingsFile)
        gsFormVisible = fGetINISetting(Location, "formVisible", gsSettingsFile)
        
        gsMultiMonitorResize = fGetINISetting(Location, "multiMonitorResize", gsSettingsFile)
        gsFirstTimeRun = fGetINISetting(Location, "firstTimeRun", gsSettingsFile)

                           
        gsWidgetSecondaryHeightRatio = fGetINISetting(Location, "widgetSecondaryHeightRatio", gsSettingsFile)
        gsWidgetPrimaryHeightRatio = fGetINISetting(Location, "widgetPrimaryHeightRatio", gsSettingsFile)
        
        gsMessageAHeightTwips = fGetINISetting(Location, "messageAHeightTwips", gsSettingsFile)
        gsMessageAWidthTwips = fGetINISetting(Location, "messageAWidthTwips ", gsSettingsFile)
        
    End If

   On Error GoTo 0
   Exit Sub

readSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readSettingsFile of Module common2"

End Sub


    
'---------------------------------------------------------------------------------------
' Procedure : validateInputs
' Author    : beededea
' Date      : 17/06/2020
' Purpose   : validate the relevant entries from the settings.ini file in user appdata
'---------------------------------------------------------------------------------------
'
Public Sub validateInputs()
    
   On Error GoTo validateInputs_Error
            
        ' general
        If gsWidgetFunctions = vbNullString Then gsWidgetFunctions = "1" ' always turn
'        If gsAnimationInterval = vbNullString Then gsAnimationInterval = "130"
        If gsStartup = vbNullString Then gsStartup = "1"
        
        If gsPointerAnimate = vbNullString Then gsPointerAnimate = "0"
        If gsSamplingInterval = vbNullString Then gsSamplingInterval = "3"
        
        ' Configuration
        If gsWidgetTooltips = "False" Then gsWidgetTooltips = "0"
        If gsWidgetTooltips = vbNullString Then gsWidgetTooltips = "0"
        
        'If gsEnablePrefsTooltips = vbNullString Then gsEnablePrefsTooltips = "false"
        If gsPrefsTooltips = "False" Then gsPrefsTooltips = "0"
        If gsPrefsTooltips = vbNullString Then gsPrefsTooltips = "0"
        
        If gsShowTaskbar = vbNullString Then gsShowTaskbar = "0"
        If gsShowHelp = vbNullString Then gsShowHelp = "1"
'        If gsTogglePendulum = vbNullString Then gsTogglePendulum = "0"
'        If gs24HourWidgetMode = vbNullString Then gs24HourWidgetMode = "1"
'
        If gsDpiAwareness = vbNullString Then gsDpiAwareness = "0"
        If gsWidgetSize = vbNullString Then gsWidgetSize = "100"
        If gsSkewDegrees = vbNullString Then gsSkewDegrees = "0"
        
        If gsScrollWheelDirection = vbNullString Then gsScrollWheelDirection = "1"
'        If gsNumericDisplayRotation = vbNullString Then gsNumericDisplayRotation = "1"
               
        ' fonts
        If gsPrefsFont = vbNullString Then gsPrefsFont = "times new roman"
        If gsClockFont = vbNullString Then gsClockFont = gsPrefsFont
        If gsPrefsFontSizeHighDPI = vbNullString Then gsPrefsFontSizeHighDPI = "8"
        If gsPrefsFontSizeLowDPI = vbNullString Then gsPrefsFontSizeLowDPI = "8"
        If gsPrefsFontItalics = vbNullString Then gsPrefsFontItalics = "false"
        If gsPrefsFontColour = vbNullString Then gsPrefsFontColour = "0"

        If gsWidgetFont = vbNullString Then gsWidgetFont = gsPrefsFont

        If gsDisplayScreenFont = vbNullString Then gsDisplayScreenFont = "courier new"
        If gsDisplayScreenFont = "Courier  New" Then gsDisplayScreenFont = "courier new"
        If gsDisplayScreenFontSize = vbNullString Then gsDisplayScreenFontSize = "5"
        If gsDisplayScreenFontItalics = vbNullString Then gsDisplayScreenFontItalics = "false"
        If gsDisplayScreenFontColour = vbNullString Then gsDisplayScreenFontColour = "0"

        ' sounds
        
        If gsEnableSounds = vbNullString Then gsEnableSounds = "1"
'        If gsEnableTicks = vbNullString Then gsEnableTicks = "0"
'        If gsEnableChimes = vbNullString Then gsEnableChimes = "0"
'        If gsEnableAlarms = vbNullString Then gsEnableAlarms = "0"
'        If gsVolumeBoost = vbNullString Then gsVolumeBoost = "0"
        
        
        ' position
        If gsAspectHidden = vbNullString Then gsAspectHidden = "0"
        If gsWidgetPosition = vbNullString Then gsWidgetPosition = "0"
        If gsWidgetLandscape = vbNullString Then gsWidgetLandscape = "0"
        If gsWidgetPortrait = vbNullString Then gsWidgetPortrait = "0"
        If gsLandscapeFormHoffset = vbNullString Then gsLandscapeFormHoffset = vbNullString
        If gsLandscapeFormVoffset = vbNullString Then gsLandscapeFormVoffset = vbNullString
        If gsPortraitHoffset = vbNullString Then gsPortraitHoffset = vbNullString
        If gsPortraitYoffset = vbNullString Then gsPortraitYoffset = vbNullString
        If gsVLocationPercPrefValue = vbNullString Then gsVLocationPercPrefValue = vbNullString
        If gsHLocationPercPrefValue = vbNullString Then gsHLocationPercPrefValue = vbNullString
                
        ' development
        If gsDebug = vbNullString Then gsDebug = "0"
        If gsDblClickCommand = vbNullString And gsFirstTimeRun = "True" Then gsDblClickCommand = "mmsys.cpl"
        If gsOpenFile = vbNullString Then gsOpenFile = vbNullString
        If gsDefaultVB6Editor = vbNullString Then gsDefaultVB6Editor = vbNullString
        If gsDefaultTBEditor = vbNullString Then gsDefaultTBEditor = vbNullString
        
        ' window
        If gsWindowLevel = vbNullString Then gsWindowLevel = "1" 'WindowLevel", gsSettingsFile)
        If gsOpacity = vbNullString Then gsOpacity = "100"
        If gsWidgetHidden = vbNullString Then gsWidgetHidden = "0"
        If gsHidingTime = vbNullString Then gsHidingTime = "0"
        If gsIgnoreMouse = vbNullString Then gsIgnoreMouse = "0"
        If gsFormVisible = vbNullString Then gsFormVisible = "0"
        
        If gsPreventDragging = vbNullString Then gsPreventDragging = "0"
        If gsMultiMonitorResize = vbNullString Then gsMultiMonitorResize = "0"
        
        
        ' other
        If gsFirstTimeRun = vbNullString Then gsFirstTimeRun = "true"
        If gsLastSelectedTab = vbNullString Then gsLastSelectedTab = "general"
        If gsSkinTheme = vbNullString Then gsSkinTheme = "dark"
    
        
        If gsWidgetPrimaryHeightRatio = "" Then gsWidgetPrimaryHeightRatio = "1"
        If gsWidgetSecondaryHeightRatio = "" Then gsWidgetSecondaryHeightRatio = "1"
        
        
   On Error GoTo 0
   Exit Sub

validateInputs_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateInputs of form modMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : getTrinketsFile
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : get this tool's entry in the trinkets settings file and assign the app.path
'---------------------------------------------------------------------------------------
'
Private Sub getTrinketsFile()
    On Error GoTo getTrinketsFile_Error
    
    Dim iFileNo As Integer: iFileNo = 0
    
    gsTrinketsDir = fSpecialFolder(feUserAppData) & "\trinkets" ' just for this user alone
    gsTrinketsFile = gsTrinketsDir & "\" & gsWidgetName & ".ini"
        
    'if the folder does not exist then create the folder
    If Not fDirExists(gsTrinketsDir) Then
        MkDir gsTrinketsDir
    End If

    'if the settings.ini does not exist then create the file by copying
    If Not fFExists(gsTrinketsFile) Then

        iFileNo = FreeFile
        'open the file for writing
        Open gsTrinketsFile For Output As #iFileNo
        Write #iFileNo, App.Path & "\" & App.EXEName & ".exe"
        Write #iFileNo,
        Close #iFileNo
    End If
    
   On Error GoTo 0
   Exit Sub

getTrinketsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getTrinketsFile of Form modMain"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : getToolSettingsFile
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : get this tool's settings file and assign to a global var
'---------------------------------------------------------------------------------------
'
Private Sub getToolSettingsFile()
    On Error GoTo getToolSettingsFile_Error
    ''If giDebugFlg = 1  Then Debug.Print "%getToolSettingsFile"
    
    Dim iFileNo As Integer: iFileNo = 0
    
    gsSettingsDir = fSpecialFolder(feUserAppData) & "\TenShillings-" & gsRichClientEnvironment & "-Widget-" & gsCodingEnvironment & "" ' just for this user alone
    gsSettingsFile = gsSettingsDir & "\settings.ini"
        
    'if the folder does not exist then create the folder
    If Not fDirExists(gsSettingsDir) Then
        MkDir gsSettingsDir
    End If

    'if the settings.ini does not exist then create the file by copying
    If Not fFExists(gsSettingsFile) Then

        iFileNo = FreeFile
        'open the file for writing
        Open gsSettingsFile For Output As #iFileNo
        Close #iFileNo
    End If
    
   On Error GoTo 0
   Exit Sub

getToolSettingsFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getToolSettingsFile of Form modMain"

End Sub



'
'---------------------------------------------------------------------------------------
' Procedure : configureTimers
' Author    : beededea
' Date      : 07/05/2023
' Purpose   : configure any global timers here
'---------------------------------------------------------------------------------------
'
Private Sub configureTimers()

    On Error GoTo configureTimers_Error
    
'    gtOldSettingsModificationTime = FileDateTime(gsSettingsFile)

    frmTimer.tmrScreenResolution.Enabled = True
    frmTimer.unhideTimer.Enabled = True

    On Error GoTo 0
    Exit Sub

configureTimers_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure configureTimers of Module modMain"
            Resume Next
          End If
    End With
 
End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : setHidingTime
' Author    : beededea
' Date      : 07/05/2023
' Purpose   : set the hiding time for the hiding timer, can't read the minutes from comboxbox as the prefs isn't yet open
'---------------------------------------------------------------------------------------
'
Private Sub setHidingTime()
    
    On Error GoTo setHidingTime_Error

    If gsHidingTime = "0" Then giMinutesToHide = 1
    If gsHidingTime = "1" Then giMinutesToHide = 5
    If gsHidingTime = "2" Then giMinutesToHide = 10
    If gsHidingTime = "3" Then giMinutesToHide = 20
    If gsHidingTime = "4" Then giMinutesToHide = 30
    If gsHidingTime = "5" Then giMinutesToHide = 60

    On Error GoTo 0
    Exit Sub

setHidingTime_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setHidingTime of Module modMain"
            Resume Next
          End If
    End With

End Sub

'---------------------------------------------------------------------------------------
' Procedure : createRCFormsOnCurrentDisplay
' Author    : beededea
' Date      : 07/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub createRCFormsOnCurrentDisplay()
    Dim thisSrf As cCairoSurface
    Dim ImageHeight As Long: ImageHeight = 0
    Dim ImageWidth As Long: ImageWidth = 0
    
    On Error GoTo createRCFormsOnCurrentDisplay_Error
    
    If Cairo.ImageList.Exists("tenshillings") Then Set thisSrf = Cairo.ImageList("tenshillings")
    
    ImageWidth = thisSrf.Width
    ImageHeight = thisSrf.height

    With New_c.Displays(1) 'get the current Display
      Call fMain.initAndCreateTenShillingsForm(ImageWidth, ImageHeight, gsWidgetName)
    End With

    With New_c.Displays(1) 'get the current Display
      Call fMain.initAndCreateAboutForm(gsWidgetName)
    End With
    
    With New_c.Displays(1) 'get the current Display
      Call fMain.initAndCreateHelpForm(gsWidgetName)
    End With

    With New_c.Displays(1) 'get the current Display
      Call fMain.initAndCreateLicenceForm(gsWidgetName)
    End With
    
    On Error GoTo 0
    Exit Sub

createRCFormsOnCurrentDisplay_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createRCFormsOnCurrentDisplay of Module modMain"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : handleUnhideMode
' Author    : beededea
' Date      : 13/05/2023
' Purpose   : when run in 'unhide' mode it writes the settings file then exits, the other
'             running but hidden process will unhide itself by timer.
'---------------------------------------------------------------------------------------
'
Private Sub handleUnhideMode(ByVal thisUnhideMode As String)
    
    On Error GoTo handleUnhideMode_Error

    If thisUnhideMode = "unhide" Then     'parse the command line
        gsUnhide = "true"
        sPutINISetting "Software\TenShillings", "unhide", gsUnhide, gsSettingsFile
        Call thisForm_Unload
        End
    End If

    On Error GoTo 0
    Exit Sub

handleUnhideMode_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure handleUnhideMode of Module modMain"
            Resume Next
          End If
    End With
End Sub




' .74 DAEB 22/05/2022 rDIConConfig.frm Msgbox replacement that can be placed on top of the form instead as the middle of the screen, see Steamydock for a potential replacement?
'---------------------------------------------------------------------------------------
' Procedure : msgBoxA
' Author    : beededea
' Date      : 20/05/2022
' Purpose   : ans = msgBoxA("main message", vbOKOnly, "title bar message", False)
'---------------------------------------------------------------------------------------
'
Public Function msgBoxA(ByVal msgBoxPrompt As String, Optional ByVal msgButton As VbMsgBoxResult, Optional ByVal msgTitle As String, Optional ByVal msgShowAgainChkBox As Boolean = False, Optional ByRef msgContext As String = "none") As Integer
     
    ' set the defined properties of a form
    On Error GoTo msgBoxA_Error

    frmMessage.propMessage = msgBoxPrompt
    frmMessage.propTitle = msgTitle
    frmMessage.propShowAgainChkBox = msgShowAgainChkBox
    frmMessage.propButtonVal = msgButton
    frmMessage.propMsgContext = msgContext
    Call frmMessage.Display ' run a subroutine in the form that displays the form

    msgBoxA = frmMessage.propReturnedValue

    On Error GoTo 0
    Exit Function

msgBoxA_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure msgBoxA of Module mdlMain"
            Resume Next
          End If
    End With

End Function




''---------------------------------------------------------------------------------------
'' Procedure : loadAsynchSoundFiles
'' Author    : beededea
'' Date      : 27/01/2025
'' Purpose   : Load the sounds into numbered buffers ready for playing
''---------------------------------------------------------------------------------------
''
'Private Sub loadAsynchSoundFiles()
'
'   On Error GoTo loadAsynchSoundFiles_Error
''
''    LoadSoundFile 1, App.path & "\resources\sounds\belltoll-quiet.wav"
''    LoadSoundFile 2, App.path & "\resources\sounds\belltoll.wav"
''    LoadSoundFile 3, App.path & "\resources\sounds\belltollLong-quiet.wav"
''    LoadSoundFile 4, App.path & "\resources\sounds\belltollLong.wav"
''    LoadSoundFile 5, App.path & "\resources\sounds\fullchime-quiet.wav"
''    LoadSoundFile 6, App.path & "\resources\sounds\fullchime.wav"
''    LoadSoundFile 7, App.path & "\resources\sounds\halfchime-quiet.wav"
''    LoadSoundFile 8, App.path & "\resources\sounds\halfchime.wav"
''    LoadSoundFile 9, App.path & "\resources\sounds\quarterchime-quiet.wav"
''    LoadSoundFile 10, App.path & "\resources\sounds\quarterchime.wav"
''    LoadSoundFile 11, App.path & "\resources\sounds\threequarterchime-quiet.wav"
''    LoadSoundFile 12, App.path & "\resources\sounds\threequarterchime.wav"
''    LoadSoundFile 13, App.path & "\resources\sounds\ticktock-quiet.wav"
''    LoadSoundFile 14, App.path & "\resources\sounds\ticktock.wav"
''    LoadSoundFile 15, App.path & "\resources\sounds\zzzz-quiet.wav"
''    LoadSoundFile 16, App.path & "\resources\sounds\zzzz.wav"
''    LoadSoundFile 17, App.path & "\resources\sounds\till-quiet.wav"
''    LoadSoundFile 18, App.path & "\resources\sounds\till.wav"
'
'   On Error GoTo 0
'   Exit Sub
'
'loadAsynchSoundFiles_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure loadAsynchSoundFiles of Module modMain"
'
'End Sub


'---------------------------------------------------------------------------------------
'Procedure:   playAsynchSound
'Author:      beededea
' Date      : 27/01/2025
' Purpose   : requires minimal changes to replace playSound code in the rest of the program
'---------------------------------------------------------------------------------------
'
'Public Sub playAsynchSound(ByVal SoundFile As String)
'
'     Dim soundindex As Long: soundindex = 0
'
'     On Error GoTo playAsynchSound_Error
'
'     If SoundFile = App.Path & "\resources\sounds\belltoll-quiet.wav" Then soundindex = 1
'     If SoundFile = App.Path & "\resources\sounds\belltoll.wav" Then soundindex = 2
'     If SoundFile = App.Path & "\resources\sounds\belltollLong-quiet.wav" Then soundindex = 3
'     If SoundFile = App.Path & "\resources\sounds\belltollLong.wav" Then soundindex = 4
'     If SoundFile = App.Path & "\resources\sounds\fullchime-quiet.wav" Then soundindex = 5
'     If SoundFile = App.Path & "\resources\sounds\fullchime.wav" Then soundindex = 6
'     If SoundFile = App.Path & "\resources\sounds\halfchime-quiet.wav" Then soundindex = 7
'     If SoundFile = App.Path & "\resources\sounds\halfchime.wav" Then soundindex = 8
'     If SoundFile = App.Path & "\resources\sounds\quarterchime-quiet.wav" Then soundindex = 9
'     If SoundFile = App.Path & "\resources\sounds\quarterchime.wav" Then soundindex = 10
'     If SoundFile = App.Path & "\resources\sounds\threequarterchime-quiet.wav" Then soundindex = 11
'     If SoundFile = App.Path & "\resources\sounds\threequarterchime.wav" Then soundindex = 12
'     If SoundFile = App.Path & "\resources\sounds\ticktock-quiet.wav" Then soundindex = 13
'     If SoundFile = App.Path & "\resources\sounds\ticktock.wav" Then soundindex = 14
'     If SoundFile = App.Path & "\resources\sounds\zzzz-quiet.wav" Then soundindex = 15
'     If SoundFile = App.Path & "\resources\sounds\zzzz.wav" Then soundindex = 16
'     If SoundFile = App.Path & "\resources\sounds\till-quiet.wav" Then soundindex = 17
'     If SoundFile = App.Path & "\resources\sounds\till.wav" Then soundindex = 18
'
'     Call playSounds(soundindex) ' writes the wav files (previously stored in a memory buffer) and feeds that buffer to the waveOutWrite API
'
'   On Error GoTo 0
'   Exit Sub
'
'playAsynchSound_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure playAsynchSound of Module modMain"
'
'End Sub



''---------------------------------------------------------------------------------------
'' Procedure : stopAsynchSound
'' Author    : beededea
'' Date      : 27/01/2025
'' Purpose   : requires minimal changes to previous playSound code
''---------------------------------------------------------------------------------------
''
'Public Sub stopAsynchSound(ByVal SoundFile As String)
'
'     Dim soundindex As Long: soundindex = 0
'
'     On Error GoTo stopAsynchSound_Error
'
''     If SoundFile = App.path & "\resources\sounds\belltoll-quiet.wav" Then soundindex = 1
''     If SoundFile = App.path & "\resources\sounds\belltoll.wav" Then soundindex = 2
''     If SoundFile = App.path & "\resources\sounds\belltollLong-quiet.wav" Then soundindex = 3
''     If SoundFile = App.path & "\resources\sounds\belltollLong.wav" Then soundindex = 4
''     If SoundFile = App.path & "\resources\sounds\fullchime-quiet.wav" Then soundindex = 5
''     If SoundFile = App.path & "\resources\sounds\fullchime.wav" Then soundindex = 6
''     If SoundFile = App.path & "\resources\sounds\halfchime-quiet.wav" Then soundindex = 7
''     If SoundFile = App.path & "\resources\sounds\halfchime.wav" Then soundindex = 8
''     If SoundFile = App.path & "\resources\sounds\quarterchime-quiet.wav" Then soundindex = 9
''     If SoundFile = App.path & "\resources\sounds\quarterchime.wav" Then soundindex = 10
''     If SoundFile = App.path & "\resources\sounds\threequarterchime-quiet.wav" Then soundindex = 11
''     If SoundFile = App.path & "\resources\sounds\threequarterchime.wav" Then soundindex = 12
''     If SoundFile = App.path & "\resources\sounds\ticktock-quiet.wav" Then soundindex = 13
''     If SoundFile = App.path & "\resources\sounds\ticktock.wav" Then soundindex = 14
''     If SoundFile = App.path & "\resources\sounds\zzzz-quiet.wav" Then soundindex = 15
''     If SoundFile = App.path & "\resources\sounds\zzzz.wav" Then soundindex = 16
''     If SoundFile = App.path & "\resources\sounds\till-quiet.wav" Then soundindex = 17
''     If SoundFile = App.path & "\resources\sounds\till.wav" Then soundindex = 18
'
'     Call StopSound(soundindex)
'
'   On Error GoTo 0
'   Exit Sub
'
'stopAsynchSound_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure stopAsynchSound of Module modMain"
'
'End Sub


''---------------------------------------------------------------------------------------
'' Procedure : stopAllAsynchSounds
'' Author    : beededea
'' Date      : 04/02/2025
'' Purpose   : ONLY stops any WAV files currently playing in asynchronous mode.
''---------------------------------------------------------------------------------------
''
'Public Sub stopAllAsynchSounds()
'
'   On Error GoTo stopAllAsynchSounds_Error
'
''    Call StopSound(1)
''    Call StopSound(2)
''    Call StopSound(3)
''    Call StopSound(4)
''    Call StopSound(5)
''    Call StopSound(6)
''    Call StopSound(7)
''    Call StopSound(8)
''    Call StopSound(9)
''    Call StopSound(10)
''    Call StopSound(12)
''    Call StopSound(13)
''    Call StopSound(14)
''    Call StopSound(15)
''    Call StopSound(16)
''    Call StopSound(17)
''    Call StopSound(18)
'
'   On Error GoTo 0
'   Exit Sub
'
'stopAllAsynchSounds_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure stopAllAsynchSounds of Module modMain"
'
'End Sub
'
'
'
'' test open hardware monitor is running
'Private Sub checkMonitorIsRunning()
'
'
'End Sub



'---------------------------------------------------------------------------------------
' Procedure : gsWidgetName
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsWidgetName() As String

    On Error GoTo gsWidgetName_Error

    gsWidgetName = m_sgsWidgetName

    On Error GoTo 0
    Exit Property

gsWidgetName_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsWidgetName of Module modMain"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsWidgetName
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsWidgetName(ByVal sgsWidgetName As String)

    On Error GoTo gsWidgetName_Error

    m_sgsWidgetName = sgsWidgetName

    On Error GoTo 0
    Exit Property

gsWidgetName_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsWidgetName of Module modMain"

End Property
