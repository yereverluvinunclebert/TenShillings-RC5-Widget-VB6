Attribute VB_Name = "Module1"
'@IgnoreModule IntegerDataType, ModuleWithoutFolder

'---------------------------------------------------------------------------------------
' Module    : Module1
' Author    : beededea
' Date      : 27/04/2023
' Purpose   : Module for declaring any public and private constants, APIs and types, public subroutines and functions.
'---------------------------------------------------------------------------------------

Option Explicit

'------------------------------------------------------ STARTS
'constants used to choose a font via the system dialog window
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const GHND As Long = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Const LF_FACESIZE As Integer = 32
Private Const CF_INITTOLOGFONTSTRUCT  As Long = &H40&
Private Const CF_SCREENFONTS As Long = &H1

'type declaration used to choose a font via the system dialog window
Private Type FormFontInfo
  Name As String
  Weight As Integer
  Height As Integer
  UnderLine As Boolean
  Italic As Boolean
  Color As Long
End Type

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type FONTSTRUC
  lStructSize As Long
  hWnd As Long
  hDC As Long
  lpLogFont As Long
  iPointSize As Long
  flags As Long
  rgbColors As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
  hInstance As Long
  lpszStyle As String
  nFontType As Integer
  MISSING_ALIGNMENT As Integer
  nSizeMin As Long
  nSizeMax As Long
End Type

'Private Type ChooseColorStruct
'    lStructSize As Long
'    hwndOwner As Long
'    hInstance As Long
'    rgbResult As Long
'    lpCustColors As Long
'    flags As Long
'    lCustData As Long
'    lpfnHook As Long
'    lpTemplateName As String
'End Type
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS

'APIs used to choose a font via the system dialog window
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As FONTSTRUC) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
'------------------------------------------------------ ENDS



'------------------------------------------------------ STARTS
' API and enums for acquiring the special folder paths
Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long

Public Enum FolderEnum ' has to be public
    feCDBurnArea = 59 ' \Docs & Settings\User\Local Settings\Application Data\Microsoft\CD Burning
    feCommonAppData = 35 ' \Docs & Settings\All Users\Application Data
    feCommonAdminTools = 47 ' \Docs & Settings\All Users\Start Menu\Programs\Administrative Tools
    feCommonDesktop = 25 ' \Docs & Settings\All Users\Desktop
    feCommonDocs = 46 ' \Docs & Settings\All Users\Documents
    feCommonPics = 54 ' \Docs & Settings\All Users\Documents\Pictures
    feCommonMusic = 53 ' \Docs & Settings\All Users\Documents\Music
    feCommonStartMenu = 22 ' \Docs & Settings\All Users\Start Menu
    feCommonStartMenuPrograms = 23 ' \Docs & Settings\All Users\Start Menu\Programs
    feCommonTemplates = 45 ' \Docs & Settings\All Users\Templates
    feCommonVideos = 55 ' \Docs & Settings\All Users\Documents\My Videos
    feLocalAppData = 28 ' \Docs & Settings\User\Local Settings\Application Data
    feLocalCDBurning = 59 ' \Docs & Settings\User\Local Settings\Application Data\Microsoft\CD Burning
    feLocalHistory = 34 ' \Docs & Settings\User\Local Settings\History
    feLocalTempInternetFiles = 32 ' \Docs & Settings\User\Local Settings\Temporary Internet Files
    feProgramFiles = 38 ' \Program Files
    feProgramFilesCommon = 43 ' \Program Files\Common Files
    'feRecycleBin = 10 ' ???
    feUser = 40 ' \Docs & Settings\User
    feUserAdminTools = 48 ' \Docs & Settings\User\Start Menu\Programs\Administrative Tools
    feUserAppData = 26 ' \Docs & Settings\User\Application Data
    feUserCache = 32 ' \Docs & Settings\User\Local Settings\Temporary Internet Files
    feUserCookies = 33 ' \Docs & Settings\User\Cookies
    feUserDesktop = 16 ' \Docs & Settings\User\Desktop
    feUserDocs = 5 ' \Docs & Settings\User\My Documents
    feUserFavorites = 6 ' \Docs & Settings\User\Favorites
    feUserMusic = 13 ' \Docs & Settings\User\My Documents\My Music
    feUserNetHood = 19 ' \Docs & Settings\User\NetHood
    feUserPics = 39 ' \Docs & Settings\User\My Documents\My Pictures
    feUserPrintHood = 27 ' \Docs & Settings\User\PrintHood
    feUserRecent = 8 ' \Docs & Settings\User\Recent
    feUserSendTo = 9 ' \Docs & Settings\User\SendTo
    feUserStartMenu = 11 ' \Docs & Settings\User\Start Menu
    feUserStartMenuPrograms = 2 ' \Docs & Settings\User\Start Menu\Programs
    feUserStartup = 7 ' \Docs & Settings\User\Start Menu\Programs\Startup
    feUserTemplates = 21 ' \Docs & Settings\User\Templates
    feUserVideos = 14  ' \Docs & Settings\User\My Documents\My Videos
    feWindows = 36 ' \Windows
    feWindowFonts = 20 ' \Windows\Fonts
    feWindowsResources = 56 ' \Windows\Resources
    feWindowsSystem = 37 ' \Windows\System32
End Enum
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' APIs for useful functions START
Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
' APIs for useful functions END
'------------------------------------------------------ ENDS

'------------------------------------------------------ STARTS
' Constants and APIs for playing sounds non-asychronously
Public Const SND_ASYNC As Long = &H1             '  play asynchronously
Public Const SND_FILENAME  As Long = &H20000     '  name is a file name

Public Declare Function playSound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
'API Functions to read/write information from INI File
Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
    , ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long _
    , ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
    , ByVal lpString As Any, ByVal lpFileName As String) As Long
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
'constants and APIs defined for querying the registry
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Const HKEY_CURRENT_USER As Long = &H80000001
Private Const REG_SZ  As Long = 1                          ' Unicode nul terminated string

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Any, ByVal cbData As Long) As Long
'------------------------------------------------------ ENDS

'------------------------------------------------------ STARTS
' Enums defined for opening a common dialog box to select files without OCX dependencies
Private Enum FileOpenConstants
    'ShowOpen, ShowSave constants.
    cdlOFNAllowMultiselect = &H200&
    cdlOFNCreatePrompt = &H2000&
    cdlOFNExplorer = &H80000
    cdlOFNExtensionDifferent = &H400&
    cdlOFNFileMustExist = &H1000&
    cdlOFNHideReadOnly = &H4&
    cdlOFNLongNames = &H200000
    cdlOFNNoChangeDir = &H8&
    cdlOFNNoDereferenceLinks = &H100000
    cdlOFNNoLongNames = &H40000
    cdlOFNNoReadOnlyReturn = &H8000&
    cdlOFNNoValidate = &H100&
    cdlOFNOverwritePrompt = &H2&
    cdlOFNPathMustExist = &H800&
    cdlOFNReadOnly = &H1&
    cdlOFNShareAware = &H4000&
End Enum

' Types defined for opening a common dialog box to select files without OCX dependencies
Private Type OPENFILENAME
    lStructSize As Long    'The size of this struct (Use the Len function)
    hwndOwner As Long       'The hWnd of the owner window. The dialog will be modal to this window
    hInstance As Long            'The instance of the calling thread. You can use the App.hInstance here.
    lpstrFilter As String        'Use this to filter what files are showen in the dialog. Separate each filter with Chr$(0). The string also has to end with a Chr(0).
    lpstrCustomFilter As String  'The pattern the user has choosed is saved here if you pass a non empty string. I never use this one
    nMaxCustFilter As Long       'The maximum saved custom filters. Since I never use the lpstrCustomFilter I always pass 0 to this.
    nFilterIndex As Long         'What filter (of lpstrFilter) is showed when the user opens the dialog.
    lpstrFile As String          'The path and name of the file the user has chosed. This must be at least MAX_PATH (260) character long.
    nMaxFile As Long             'The length of lpstrFile + 1
    lpstrFileTitle As String     'The name of the file. Should be MAX_PATH character long
    nMaxFileTitle As Long        'The length of lpstrFileTitle + 1
    lpstrInitialDir As String    'The path to the initial path :) If you pass an empty string the initial path is the current path.
    lpstrTitle As String         'The caption of the dialog.
    flags As FileOpenConstants                'Flags. See the values in MSDN Library (you can look at the flags property of the common dialog control)
    nFileOffset As Integer       'Points to the what character in lpstrFile where the actual filename begins (zero based)
    nFileExtension As Integer    'Same as nFileOffset except that it points to the file extention.
    lpstrDefExt As String        'Can contain the extention Windows should add to a file if the user doesn't provide one (used with the GetSaveFileName API function)
    lCustData As Long            'Only used if you provide a Hook procedure (Making a Hook procedure is pretty messy in VB.
    lpfnHook As Long             'Pointer to the hook procedure.
    lpTemplateName As String     'A string that contains a dialog template resource name. Only used with the hook procedure.
End Type

'Private Type BROWSEINFO
'    hwndOwner As Long
'    pidlRoot As Long 'LPCITEMIDLIST
'    pszDisplayName As String
'    lpszTitle As String
'    ulFlags As Long
'    lpfn As Long  'BFFCALLBACK
'    lParam As Long
'    iImage As Long
'End Type

' vars defined for opening a common dialog box to select files without OCX dependencies
Private x_OpenFilename As OPENFILENAME

' APIs declared for opening a common dialog box to select files without OCX dependencies
Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (lpofn As OPENFILENAME) As Long
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' APIs, constants and types defined for determining the OS version
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type

Private Const VER_PLATFORM_WIN32s As Long = 0
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Private Const VER_PLATFORM_WIN32_NT As Long = 2
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' APIs, constants and types defined for determining existence of files and folders
Private Const OF_EXIST         As Long = &H4000
Private Const OFS_MAXPATHNAME  As Long = 128
Private Const HFILE_ERROR      As Long = -1
 
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type
     
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, _
                            lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
'------------------------------------------------------ ENDS
             
'------------------------------------------------------ STARTS
' private variables used within properties replacing global variables
'   avoids breaking encapsulation
'   allows validation during let
'   allows addition of logic to provide potential constraint of values during each LET

' general

Private m_sgsMultiCoreEnable As String
Private m_sWidgetSize As String
Private m_sgsStartup As String
Private m_sgsWidgetFunctions As String
Private m_sgsPointerAnimate As String
Private m_sgsSamplingInterval As String

' config

Private m_sgsSkewDegrees As String ' unique to the rotating widgets
Private m_sgsWidgetTooltips As String
Private m_sgsPrefsTooltips As String
Private m_sgsShowTaskbar As String
Private m_sgsShowHelp As String
Private m_sgsDpiAwareness As String
'Public gsWidgetSize As String
Private m_sgsScrollWheelDirection As String
Private m_sgsWidgetHighDpiXPos As String
Private m_sgsWidgetHighDpiYPos As String
Private m_sgsWidgetLowDpiXPos As String
Private m_sgsWidgetLowDpiYPos As String
       
' font

Private m_sgsClockFont As String
Private m_sgsWidgetFont As String
Private m_sgsPrefsFont As String
Private m_sgsPrefsFontSizeHighDPI As String
Private m_sgsPrefsFontSizeLowDPI As String
Private m_sgsPrefsFontItalics As String
Private m_sgsPrefsFontColour As String
Private m_sgsDisplayScreenFont As String
Private m_sgsDisplayScreenFontSize As String
Private m_sgsDisplayScreenFontItalics As String
Private m_sgsDisplayScreenFontColour As String

'------------------------------------------------------ ENDS



        

'------------------------------------------------------ STARTS
' global variables - mostly read from and written to settings.ini
' There are so many global vars because the old YWE javascript version of this program used global vars, this was a conversion.
' Note: In VB6 public variables used in class modules are treated as properties, passed by value, not by reference




' sounds

Public gsEnableSounds As String

' position

Public gsAspectRatio As String
Public gsAspectHidden As String
Public gsWidgetPosition As String
Public gsWidgetLandscape As String
Public gsWidgetPortrait As String
Public gsLandscapeFormHoffset As String
Public gsLandscapeFormVoffset As String
Public gsPortraitHoffset As String
Public gsPortraitYoffset As String
Public gsVLocationPercPrefValue As String
Public gsHLocationPercPrefValue As String

' development

Public gsDebug As String
Public gsDblClickCommand As String
Public gsOpenFile As String
Public gsDefaultVB6Editor As String
Public gsDefaultTBEditor As String
Public gsCodingEnvironment As String
Public gsRichClientEnvironment As String

' window

Public gsFormVisible As String ' unique to the rotating widgets

Public giMinutesToHide As Integer
Public gbWindowLevelWasChanged As Boolean
Public gsWindowLevel As String
Public gsPreventDragging As String
Public gsOpacity  As String
Public gsWidgetHidden  As String
Public gsHidingTime  As String
Public gsIgnoreMouse  As String
Public gbMenuOccurred As Boolean
Public gsFirstTimeRun  As String
Public gsMultiMonitorResize  As String

' vars to obtain actual correct screen width (to correct VB6 bug) twips
Public glPhysicalScreenWidthTwips As Long
Public glPhysicalScreenHeightTwips As Long
' pixels
Public glPhysicalScreenHeightPixels As Long
Public glPhysicalScreenWidthPixels As Long
Public glOldPhysicalScreenHeightPixels As Long
Public glOldPhysicalScreenWidthPixels As Long
Public glVirtualScreenHeightPixels As Long
Public glVirtualScreenWidthPixels As Long

' vars to obtain the virtual (multi-monitor) width twips
Public glVirtualScreenHeightTwips As Long
Public glVirtualScreenWidthTwips As Long

' vars stored for positioning the prefs form

Public glWidgetPrefsOldHeightTwips As Long
Public glWidgetPrefsOldWidthTwips As Long
Public gsPrefsHighDpiXPosTwips As String
Public gsPrefsHighDpiYPosTwips As String
Public gsPrefsLowDpiXPosTwips As String
Public gsPrefsLowDpiYPosTwips As String
Public gsPrefsPrimaryHeightTwips As String
Public gsPrefsSecondaryHeightTwips As String
Public gsWidgetPrimaryHeightRatio As String
Public gsWidgetSecondaryHeightRatio As String

Public glMonitorCount As Long
Public glOldPrefsFormMonitorPrimary As Long
Public glOldWidgetFormMonitorPrimary As Long

Public gbMsgBoxADynamicSizingFlg As Boolean
Public gbPrefsFormResizedInCode As Boolean

' General variables declared

Public gsSettingsDir As String
Public gsSettingsFile As String
Public gsTrinketsDir As String
Public gsTrinketsFile As String

Public gsMessageAHeightTwips  As String
Public gsMessageAWidthTwips   As String

Public gsMulticoreXPosTwips As String
Public gsMulticoreYPosTwips As String

' General variables declared

Public gsLastSelectedTab As String
Public gsSkinTheme As String
Public gsUnhide As String
Public gbClassicThemeCapable As Boolean
Public glStoreThemeColour As Long

' key presses
Public gbCTRL_1 As Boolean
Public gbSHIFT_1 As Boolean
Private pbDebugMode As Boolean ' .30 DAEB 03/03/2021 frmMain.frm replaced the inIDE function that used a variant to one without
Public giDebugFlg As Integer
Public gbStartupFlg As Boolean
Public gbThisWidgetAvailable As Boolean
Public gbReload As Boolean

'Public gtOldSettingsModificationTime  As Date

'------------------------------------------------------ ENDS





'---------------------------------------------------------------------------------------
' Procedure : fFExists
' Author    : RobDog888 https://www.vbforums.com/member.php?17511-RobDog888
' Date      : 19/07/2023
' Purpose   : Test for file existence using the OpenFile API
'---------------------------------------------------------------------------------------
'
Public Function fFExists(ByVal Fname As String) As Boolean
 
    Dim lRetVal As Long
    Dim OfSt As OFSTRUCT
    
    On Error GoTo fFExists_Error
    
    lRetVal = OpenFile(Fname, OfSt, OF_EXIST)
    If lRetVal <> HFILE_ERROR Then
        fFExists = True
    Else
        fFExists = False
    End If

   On Error GoTo 0
   Exit Function

fFExists_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fFExists of Module Module1"
    
End Function



'---------------------------------------------------------------------------------------
' Procedure : fDirExists
' Author    : zeezee https://www.vbforums.com/member.php?90054-zeezee
' Date      : 19/07/2023
' Purpose   : Test for file existence using the PathFileExists API
'---------------------------------------------------------------------------------------
'
Public Function fDirExists(ByVal pstrFolder As String) As Boolean
   On Error GoTo fDirExists_Error

    fDirExists = (PathFileExists(pstrFolder) = 1)
    If fDirExists Then fDirExists = (PathIsDirectory(pstrFolder) <> 0)

   On Error GoTo 0
   Exit Function

fDirExists_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fDirExists of Module Module1"
End Function

'
'---------------------------------------------------------------------------------------
' Procedure : fLicenceState
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : check the state of the licence
'---------------------------------------------------------------------------------------
'
Public Function fLicenceState() As Integer
    Dim slicence As String: slicence = "0"
    
    On Error GoTo fLicenceState_Error
    ''If giDebugFlg = 1  Then DebugPrint "%" & "fLicenceState"
    
    fLicenceState = 0
    ' read the tool's own settings file
    If fFExists(gsSettingsFile) Then ' does the tool's own settings.ini exist?
        slicence = fGetINISetting("Software\TenShillings", "licence", gsSettingsFile)
        ' if the licence state is not already accepted then display the licence form
        If slicence = "1" Then fLicenceState = 1
    End If

   On Error GoTo 0
   Exit Function

fLicenceState_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fLicenceState of Form common"

End Function
'---------------------------------------------------------------------------------------
' Procedure : showLicence
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : check the state of the licence
'---------------------------------------------------------------------------------------
'
Public Sub showLicence(ByVal licenceState As Integer)
    'Dim slicence As String: slicence = "0"
    On Error GoTo showLicence_Error
    ''If giDebugFlg = 1  Then DebugPrint "%" & "showLicence"
    
    ' if the licence state is not already accepted then display the licence form
    If licenceState = 0 Then
        'Call LoadFileToTB(frmLicence.txtLicenceTextBox, App.Path & "\Resources\txt\licence.txt", False)
        Call licenceSplash
    End If

   On Error GoTo 0
   Exit Sub

showLicence_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure showLicence of Form common"

End Sub
    
    
    

'---------------------------------------------------------------------------------------
' Procedure : setDPIaware
' Author    : beededea
' Date      : 29/10/2023
' Purpose   : This sets DPI awareness for the whole program incl. especially the VB6 forms, requires a program hard restart.
'---------------------------------------------------------------------------------------
'
Public Sub setDPIaware()
    On Error GoTo setDPIaware_Error
       
    If gsDpiAwareness = "1" Then
        If Not InIDE Then
            Cairo.SetDPIAwareness ' this way avoids the VB6 IDE shrinking (sadly, VB6 has a high DPI unaware IDE)
            gbMsgBoxADynamicSizingFlg = True
        End If
    End If

    On Error GoTo 0
    Exit Sub

setDPIaware_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setDPIaware of Module modMain"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : testDPIAndSetInitialAwareness
' Author    : beededea
' Date      : 29/10/2023
' Purpose   : if screen width in pixels is greater than 1960 then set high DPI by default
'---------------------------------------------------------------------------------------
'
Public Sub testDPIAndSetInitialAwareness()
    On Error GoTo testDPIAndSetInitialAwareness_Error

    'If fPixelsPerInchX() > 96 Then ' always seems to provide 96, no matter what I do.
    
     If glPhysicalScreenWidthPixels > 1960 Then
        gsDpiAwareness = "1"
        Call setDPIaware
    End If

    On Error GoTo 0
    Exit Sub

testDPIAndSetInitialAwareness_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure testDPIAndSetInitialAwareness of Module Module1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : LoadFileToTB
' Author    : https://www.vbforums.com/member.php?95578-Zach_VB6
' Date      : 26/08/2019
' Purpose   : Loads file specified by FilePath into textcontrol
'             (e.g., Text Box, Rich Text Box) specified by TxtBox
'---------------------------------------------------------------------------------------
'
Public Sub LoadFileToTB(ByVal TxtBox As Object, ByVal FilePath As String, Optional ByVal Append As Boolean = False)

    'If Append = true, then loaded text is appended to existing
    ' contents else existing contents are overwritten
    
    Dim iFile As Integer: iFile = 0
    Dim s As String: s = vbNullString
    
    On Error GoTo LoadFileToTB_Error

   ''If giDebugFlg = 1  Then msgbox "%" & LoadFileToTB

    If Dir$(FilePath) = vbNullString Then Exit Sub
    
    On Error GoTo ErrorHandler:
    s = TxtBox.Text
    
    iFile = FreeFile
    Open FilePath For Input As #iFile
    s = Input(LOF(iFile), #iFile)
    If Append Then
        TxtBox.Text = TxtBox.Text & s
    Else
        TxtBox.Text = s
    End If
    
ErrorHandler:
    If iFile > 0 Then Close #iFile

   On Error GoTo 0
   Exit Sub

LoadFileToTB_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadFileToTB of Form common"

End Sub


'
'---------------------------------------------------------------------------------------
' Procedure : fGetINISetting
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : Get the INI Setting from the File
'---------------------------------------------------------------------------------------
'
Public Function fGetINISetting(ByVal sHeading As String, ByVal sKey As String, ByRef sINIFileName As String) As String
   On Error GoTo fGetINISetting_Error
    Const cparmLen As Integer = 500 ' maximum no of characters allowed in the returned string
    Dim sReturn As String * cparmLen ' not going to initialise this with a 500 char string
    Dim sDefault As String * cparmLen
    Dim lLength As Long: lLength = 0

    lLength = GetPrivateProfileString(sHeading, sKey, sDefault, sReturn, cparmLen, sINIFileName)
    fGetINISetting = Mid$(sReturn, 1, lLength)

   On Error GoTo 0
   Exit Function

fGetINISetting_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fGetINISetting of module module1"
End Function

'
'---------------------------------------------------------------------------------------
' Procedure : sPutINISetting
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : Save a specific INI setting to a specific section of the filename supplied using a key to identify the setting
'---------------------------------------------------------------------------------------
'
Public Sub sPutINISetting(ByVal sHeading As String, ByVal sKey As String, ByVal sSetting As String, ByRef sINIFileName As String)

   On Error GoTo sPutINISetting_Error

    Dim unusedReturnValue As Long: unusedReturnValue = 0
    
    unusedReturnValue = WritePrivateProfileString(sHeading, sKey, sSetting, sINIFileName)

   On Error GoTo 0
   Exit Sub

sPutINISetting_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sPutINISetting of module module1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : writeRegistry
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : write to the registry
'---------------------------------------------------------------------------------------
'
Public Sub writeRegistry(ByRef hKey As Long, ByRef strPath As String, ByRef strvalue As String, ByRef strData As String)

    Dim keyhand As Long: keyhand = 0
    Dim unusedReturnValue As Long: unusedReturnValue = 0
    
    On Error GoTo writeRegistry_Error

    unusedReturnValue = RegCreateKey(hKey, strPath, keyhand)
    unusedReturnValue = RegSetValueEx(keyhand, strvalue, 0, REG_SZ, ByVal strData, Len(strData))
    unusedReturnValue = RegCloseKey(keyhand)

   On Error GoTo 0
   Exit Sub

writeRegistry_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeRegistry of module module1"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : fSpecialFolder
' Author    : si_the_geek vbforums
' Date      : 17/10/2019
' Purpose   : Returns the path to the specified special folder (AppData etc)
'---------------------------------------------------------------------------------------
'
Public Function fSpecialFolder(ByVal pfe As FolderEnum) As String
    Const MAX_PATH As Integer = 260
    Dim strPath As String: strPath = vbNullString
    Dim strBuffer As String: strBuffer = vbNullString
    
    On Error GoTo fSpecialFolder_Error

    strBuffer = Space$(MAX_PATH)
    If SHGetFolderPath(0, pfe, 0, 0, strBuffer) = 0 Then strPath = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)
    If Right$(strPath, 1) = "\" Then strPath = Left$(strPath, Len(strPath) - 1)
    fSpecialFolder = strPath

    On Error GoTo 0
    Exit Function

fSpecialFolder_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fSpecialFolder of Module Module1"
End Function

'---------------------------------------------------------------------------------------
' Procedure : addTargetfile
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : open a dialogbox to select a file as the target, normally a binary
'---------------------------------------------------------------------------------------
'
Public Sub addTargetFile(ByVal fieldValue As String, ByRef retFileName As String)
    Dim FilePath As String: FilePath = vbNullString
    Dim dialogInitDir As String: dialogInitDir = vbNullString
    Dim retfileTitle As String: retfileTitle = vbNullString
    Const x_MaxBuffer As Integer = 256
    
    ''If giDebugFlg = 1  Then Debug.Print "%" & "addTargetfile"
    
    On Error Resume Next
    
    dialogInitDir = App.Path
    
    ' set the default folder to the existing reference
    If Not fieldValue = vbNullString Then
        If fFExists(fieldValue) Then
            ' extract the folder name from the string
            FilePath = fGetDirectory(fieldValue)
            ' set the default folder to the existing reference
            dialogInitDir = FilePath 'start dir, might be "C:\" or so also
        ElseIf fDirExists(fieldValue) Then ' this caters for the entry being just a folder name
            ' set the default folder to the existing reference
            dialogInitDir = fieldValue 'start dir, might be "C:\" or so also
        Else
            dialogInitDir = App.Path 'start dir, might be "C:\" or so also
        End If
    End If
    
  With x_OpenFilename
'    .hwndOwner = Me.hWnd
    .hInstance = App.hInstance
    .lpstrTitle = "Select a File Target"
    .lpstrInitialDir = dialogInitDir
    
    .lpstrFilter = "Text Files" & vbNullChar & "*.txt" & vbNullChar & "All Files" & vbNullChar & "*.*" & vbNullChar & vbNullChar
    .nFilterIndex = 2
    
    .lpstrFile = String(x_MaxBuffer, 0)
    .nMaxFile = x_MaxBuffer - 1
    .lpstrFileTitle = .lpstrFile
    .nMaxFileTitle = x_MaxBuffer - 1
    .lStructSize = Len(x_OpenFilename)
  End With
  

  Call obtainOpenFileName(retFileName, retfileTitle) ' retfile will be buffered to 256 bytes

   On Error GoTo 0
   
   Exit Sub

End Sub
'---------------------------------------------------------------------------------------
' Procedure : fGetDirectory
' Author    : beededea
' Date      : 11/07/2019
' Purpose   : get the folder or directory path as a string not including the last backslash
'---------------------------------------------------------------------------------------
'
Public Function fGetDirectory(ByRef Path As String) As String

   On Error GoTo fGetDirectory_Error
   ''If giDebugFlg = 1  Then DebugPrint "%" & "fnGetDirectory"

    If InStrRev(Path, "\") = 0 Then
        fGetDirectory = vbNullString
        Exit Function
    End If
    fGetDirectory = Left$(Path, InStrRev(Path, "\") - 1)

   On Error GoTo 0
   Exit Function

fGetDirectory_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fGetDirectory of module module1"
End Function

'---------------------------------------------------------------------------------------
' Procedure : obtainOpenFileName
' Author    : beededea
' Date      : 02/09/2019
' Purpose   : using GetOpenFileName API returns file name and title, the filename will be buffered to 256 bytes
'---------------------------------------------------------------------------------------
'
Public Sub obtainOpenFileName(ByRef retFileName As String, ByRef retfileTitle As String)
   On Error GoTo obtainOpenFileName_Error
   ''If giDebugFlg = 1  Then Debug.Print "%obtainOpenFileName"

  If GetOpenFileName(x_OpenFilename) <> 0 Then
'    If x_OpenFilename.lpstrFile = "*.*" Then
'        'txtTarget.Text = savLblTarget
'    Else
        retfileTitle = x_OpenFilename.lpstrFileTitle
        retFileName = x_OpenFilename.lpstrFile
'    End If
  'Else
    'The CANCEL button was pressed
    'MsgBox "Cancel"
  End If

   On Error GoTo 0
   Exit Sub

obtainOpenFileName_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure obtainOpenFileName of module module1.bas"
End Sub





'
'---------------------------------------------------------------------------------------
' Procedure : GetWindowsVersion
' Author    :
' Date      : 28/05/2023
' Purpose   : Returns the version of Windows that the user is running
'             Be aware that if run in the VB6 IDE it may result in "Windows XP" regardless of the o/s you are running.
'             May also report Win 8 for Win 8 and above.
'---------------------------------------------------------------------------------------
'
Public Function GetWindowsVersion() As String
    Dim OSV As OSVERSIONINFO
    
    On Error GoTo GetWindowsVersion_Error

    OSV.OSVSize = Len(OSV)

    If GetVersionEx(OSV) = 1 Then
        Select Case OSV.PlatformID
            Case VER_PLATFORM_WIN32s
                GetWindowsVersion = "Win32s on Windows 3.1"
            Case VER_PLATFORM_WIN32_NT
                GetWindowsVersion = "Windows NT"
                
                Select Case OSV.dwVerMajor
                    Case 3
                        GetWindowsVersion = "Windows NT 3.5"
                    Case 4
                        GetWindowsVersion = "Windows NT 4.0"
                    Case 5
                        Select Case OSV.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Windows 2000"
                            Case 1
                                GetWindowsVersion = "Windows XP"
                            Case 2
                                GetWindowsVersion = "Windows Server 2003"
                        End Select
                    Case 6
                        Select Case OSV.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Windows Vista"
                            Case 1
                                GetWindowsVersion = "Windows 7"
                            Case 2
                                GetWindowsVersion = "Windows 8"
                            Case 3
                                GetWindowsVersion = "Windows 8.1"
                            Case 10
                                GetWindowsVersion = "Windows 10"
                        End Select
                End Select
        
            Case VER_PLATFORM_WIN32_WINDOWS:
                Select Case OSV.dwVerMinor
                    Case 0
                        GetWindowsVersion = "Windows 95"
                    Case 90
                        GetWindowsVersion = "Windows Me"
                    Case Else
                        GetWindowsVersion = "Windows 98"
                End Select
        End Select
    Else
        GetWindowsVersion = "Unable to identify your version of Windows."
    End If

   On Error GoTo 0
   Exit Function

GetWindowsVersion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetWindowsVersion of Module Module1"
End Function




'----------------------------------------
'Name: TestWinVer
'Description: Tests the multiplicity of Windows versions and returns some values, largely redundant now, might be used later for XP/ReactOS testing/running.
'----------------------------------------
Public Function fTestClassicThemeCapable() As Boolean
    Dim windowsVer As String
    '=================================
    '2000 / XP / NT / 7 / 8 / 10
    '=================================
    On Error GoTo fTestClassicThemeCapable_Error

    Dim ProgramFilesDir As String: ProgramFilesDir = vbNullString
    Dim strString As String: strString = vbNullString
    'Dim shortWindowsVer As String: shortWindowsVer = vbNullString
    
    fTestClassicThemeCapable = False
    windowsVer = vbNullString
    
    ' other variable assignments
    strString = fReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName")
    windowsVer = strString

    ' note that when running in compatibility mode the o/s will respond with "Windows XP"
    ' The IDE runs in compatibility mode so it may report the wrong working folder

    'Get the value of "ProgramFiles", or "ProgramFilesDir"
        
    windowsVer = GetWindowsVersion
    
    Select Case windowsVer
    Case "Windows NT 4.0"
        fTestClassicThemeCapable = True
        strString = fReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case "Windows 2000"
        fTestClassicThemeCapable = True
        strString = fReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case "Windows XP"
        fTestClassicThemeCapable = True
        strString = fReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProgramFilesDir")
    Case "Windows Server 2003"
        fTestClassicThemeCapable = True
        strString = fReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case "Windows Vista"
        fTestClassicThemeCapable = True
        strString = fReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case "Windows 7"
        fTestClassicThemeCapable = True
        strString = fReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case Else ' windows 8/10/11+
        fTestClassicThemeCapable = False
        strString = fReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProgramFilesDir")
    End Select

    ProgramFilesDir = strString
    If ProgramFilesDir = vbNullString Then ProgramFilesDir = "c:\program files (x86)" ' 64bit systems
    If Not fDirExists(ProgramFilesDir) Then
        ProgramFilesDir = "c:\program files" ' 32 bit systems
    End If
   
    On Error GoTo 0: Exit Function

fTestClassicThemeCapable_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fTestClassicThemeCapable of module module1"

End Function




'---------------------------------------------------------------------------------------
' Procedure : fReadRegistry
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : get a string from the registry
'---------------------------------------------------------------------------------------
'
Public Function fReadRegistry(ByRef hKey As Long, ByRef strPath As String, ByRef strvalue As String) As String

    Dim keyhand As Long: keyhand = 0
    Dim lResult As Long: lResult = 0
    Dim strBuf As String: strBuf = vbNullString
    Dim lDataBufSize As Long: lDataBufSize = 0
    Dim intZeroPos As Integer: intZeroPos = 0
    Dim unusedReturnValue As Integer: unusedReturnValue = 0

    Dim lValueType As Variant

    On Error GoTo fReadRegistry_Error

    unusedReturnValue = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strvalue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String$(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strvalue, 0&, 0&, ByVal strBuf, lDataBufSize)
        Dim ERROR_SUCCESS As Variant
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                fReadRegistry = Left$(strBuf, intZeroPos - 1)
            Else
                fReadRegistry = strBuf
            End If
        End If
    End If

   On Error GoTo 0
   Exit Function

fReadRegistry_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fReadRegistry of module module1"
End Function



'
'---------------------------------------------------------------------------------------
' Procedure : changeFont
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : select a font for the default form
'---------------------------------------------------------------------------------------
'
Public Sub changeFont(ByVal frm As Form, ByVal fntNow As Boolean, ByRef fntFont As String, ByRef fntSize As Integer, ByRef fntWeight As Integer, ByRef fntStyle As Boolean, ByRef fntColour As Long, ByRef fntItalics As Boolean, ByRef fntUnderline As Boolean, ByRef fntFontResult As Boolean)
    
   On Error GoTo changeFont_Error

    fntWeight = 0
    fntStyle = False
    'fntColour = 0
    'fntBold = False
    'fntUnderline = False
    fntFontResult = False
    
    'If giDebugFlg = 1  Then Debug.Print "%mnuFont_Click"

    displayFontSelector fntFont, fntSize, fntWeight, fntStyle, fntColour, fntItalics, fntUnderline, fntFontResult
    If fntFontResult = False Then Exit Sub

    If fntFont <> vbNullString And fntNow = True Then
        Call changeFormFont(frm, fntFont, Val(fntSize), fntWeight, fntStyle, fntItalics, fntColour)
    End If
    
   On Error GoTo 0
   Exit Sub

changeFont_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure changeFont of Module Module1"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : displayFontSelector
' Author    : beededea
' Date      : 29/02/2020
' Purpose   : display the font dialog selector
'---------------------------------------------------------------------------------------
'
Public Sub displayFontSelector(ByRef currFont As String, ByRef currSize As Integer, ByRef currWeight As Integer, ByVal currStyle As Boolean, ByRef currColour As Long, ByRef currItalics As Boolean, ByRef currUnderline As Boolean, ByRef fontResult As Boolean)

    Dim thisFont As FormFontInfo

    On Error GoTo displayFontSelector_Error

    With thisFont
      .Color = currColour
      .Height = currSize
      .Weight = currWeight
      '400     Font is normal.
      '700     Font is bold.
      .Italic = currItalics
      .UnderLine = currUnderline
      .Name = currFont
    End With
    
    fontResult = fDialogFont(thisFont)
    If fontResult = False Then Exit Sub
    
    ' some fonts have naming problems and the result is an empty font name field on the font selector
    If thisFont.Name = vbNullString Then thisFont.Name = "times new roman"
    If thisFont.Name = vbNullString Then Exit Sub
    
    With thisFont
        currFont = .Name
        currSize = .Height
        currWeight = .Weight
        currItalics = .Italic
        currUnderline = .UnderLine
        currColour = .Color
        'ctl = .Name & " - Size:" & .Height
    End With

   On Error GoTo 0
   Exit Sub

displayFontSelector_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayFontSelector of module module1"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : changeFormFont
' Author    : beededea
' Date      : 12/07/2019
' Purpose   : change the font throughout the whole form
'---------------------------------------------------------------------------------------
'
Public Sub changeFormFont(ByVal formName As Object, ByVal suppliedFont As String, ByVal suppliedSize As Integer, ByVal suppliedWeight As Integer, ByVal suppliedStyle As Boolean, ByVal suppliedItalics As Boolean, ByVal suppliedColour As Long)
    On Error GoTo changeFormFont_Error
        
    Dim Ctrl As Control
      
    ' loop through all the controls and identify the labels and text boxes
    For Each Ctrl In formName.Controls
        If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is textBox) Or (TypeOf Ctrl Is FileListBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Or (TypeOf Ctrl Is ListBox) Then
            If Ctrl.Name <> "lblDragCorner" And Ctrl.Name <> "txtDisplayScreenFont" Then
                If suppliedFont <> vbNullString Then Ctrl.Font.Name = suppliedFont
                If suppliedSize > 0 Then Ctrl.Font.Size = suppliedSize
                Ctrl.Font.Italic = suppliedItalics
            End If
            Select Case True
                Case (TypeOf Ctrl Is CommandButton)
                    ' stupif fecking VB6 will not let you change the font of the forecolour on a button!
                    'Ctrl.ForeColor = suppliedColour
                    ' do nothing
                Case Else
                    Ctrl.ForeColor = suppliedColour
            End Select
        End If
    Next
     
   On Error GoTo 0
   Exit Sub

changeFormFont_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure changeFormFont of module module1"
    
End Sub



'---------------------------------------------------------------------------------------
' Procedure : fDialogFont
' Author    : beededea
' Date      : 21/08/2020
' Purpose   : display the default windows dialog box that allows the user to select a font.
'             note: this is placed central screen by subclassing in modCentre.bas
'---------------------------------------------------------------------------------------
'
Public Function fDialogFont(ByRef f As FormFontInfo) As Boolean
      
    Dim logFnt As LOGFONT
    Dim ftStruc As FONTSTRUC
    Dim lLogFontAddress As Long: lLogFontAddress = 0
    Dim lMemHandle As Long: lMemHandle = 0
    Dim hWndAccessApp As Long: hWndAccessApp = 0
    
    Const LOGPIXELSY As Integer = 90        '  Logical pixels/inch in Y

    On Error GoTo fDialogFont_Error
    
    logFnt.lfWeight = f.Weight
    logFnt.lfItalic = f.Italic * -1
    logFnt.lfUnderline = f.UnderLine * -1
    logFnt.lfHeight = -fMulDiv(CLng(f.Height), GetDeviceCaps(GetDC(hWndAccessApp), LOGPIXELSY), 72)
    Call StringToByte(f.Name, logFnt.lfFaceName())
    ftStruc.rgbColors = f.Color
    ftStruc.lStructSize = Len(ftStruc)
    
    lMemHandle = GlobalAlloc(GHND, Len(logFnt))
    If lMemHandle = 0 Then
      fDialogFont = False
      Exit Function
    End If

    lLogFontAddress = GlobalLock(lMemHandle)
    If lLogFontAddress = 0 Then
      fDialogFont = False
      Exit Function
    End If
    
    CopyMemory ByVal lLogFontAddress, logFnt, Len(logFnt)
    ftStruc.lpLogFont = lLogFontAddress
    'ftStruc.flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT
    ftStruc.flags = CF_SCREENFONTS Or CF_INITTOLOGFONTSTRUCT
    If ChooseFont(ftStruc) = 1 Then
      CopyMemory logFnt, ByVal lLogFontAddress, Len(logFnt)
      f.Weight = logFnt.lfWeight
      f.Italic = CBool(logFnt.lfItalic)
      f.UnderLine = CBool(logFnt.lfUnderline)
      f.Name = fByteToString(logFnt.lfFaceName())
      f.Height = CLng(ftStruc.iPointSize / 10)
      f.Color = ftStruc.rgbColors
      fDialogFont = True
    Else
      fDialogFont = False
    End If

   On Error GoTo 0
   Exit Function

fDialogFont_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fDialogFont of Module module1"
End Function


'---------------------------------------------------------------------------------------
' Procedure : fMulDiv
' Author    :
' Date      : 21/08/2020
' Purpose   : Used in fDialogFont above, the fMulDiv function multiplies two 32-bit values and then divides the 64-bit result by a third 32-bit value.
'---------------------------------------------------------------------------------------
'
Private Function fMulDiv(ByVal In1 As Long, ByVal In2 As Long, ByVal In3 As Long) As Long
        
    Dim lngTemp As Long: lngTemp = 0
    On Error GoTo fMulDiv_Error
    
    On Error GoTo fMulDiv_err
    If In3 <> 0 Then
        lngTemp = In1 * In2
        lngTemp = lngTemp / In3
    Else
        lngTemp = -1
    End If

    fMulDiv = lngTemp
    Exit Function
fMulDiv_err:
    lngTemp = -1
    Resume fMulDiv_err

   On Error GoTo 0
   Exit Function

fMulDiv_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fMulDiv of Module module1"
End Function



'---------------------------------------------------------------------------------------
' Procedure : StringToByte
' Author    :
' Date      : 21/08/2020
' Purpose   : Used in fDialogFont above, converts a provided string to a byte array
'---------------------------------------------------------------------------------------
'
Private Sub StringToByte(ByVal InString As String, ByRef ByteArray() As Byte)
    
    Dim intLbound As Integer: intLbound = 0
    Dim intUbound As Integer: intUbound = 0
    Dim intLen As Integer: intLen = 0
    Dim intX As Integer: intX = 0
    
    On Error GoTo StringToByte_Error

    intLbound = LBound(ByteArray)
    intUbound = UBound(ByteArray)
    intLen = Len(InString)
    If intLen > intUbound - intLbound Then intLen = intUbound - intLbound
    For intX = 1 To intLen
        ByteArray(intX - 1 + intLbound) = Asc(Mid(InString, intX, 1))
    Next

   On Error GoTo 0
   Exit Sub

StringToByte_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure StringToByte of Module module1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fByteToString
' Author    :
' Date      : 21/08/2020
' Purpose   : Used in fDialogFont above, converts a byte array provided to a string
'---------------------------------------------------------------------------------------
'
Private Function fByteToString(ByRef aBytes() As Byte) As String
      
    Dim dwBytePoint As Long: dwBytePoint = 0
    Dim dwByteVal As Long: dwByteVal = 0
    Dim szOut As String: szOut = vbNullString
    
    On Error GoTo fByteToString_Error

    dwBytePoint = LBound(aBytes)
    While dwBytePoint <= UBound(aBytes) ' whileing and wending my way through the bytearrays >sigh<
      dwByteVal = aBytes(dwBytePoint)
      If dwByteVal = 0 Then
        fByteToString = szOut
        Exit Function
      Else
        szOut = szOut & Chr$(dwByteVal)
      End If
      dwBytePoint = dwBytePoint + 1
    Wend
    fByteToString = szOut

   On Error GoTo 0
   Exit Function

fByteToString_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fByteToString of Module module1"
End Function

'---------------------------------------------------------------------------------------
' Procedure : aboutClickEvent
' Author    : beededea
' Date      : 02/05/2023
' Purpose   : the public subroutine called by two forms, handling the menu about click event for both
'---------------------------------------------------------------------------------------
'
Public Sub aboutClickEvent()
    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo aboutClickEvent_Error
'    If gsVolumeBoost = "1" Then
'        fileToPlay = "till.wav"
'    Else
'        fileToPlay = "till-quiet.wav"
'    End If
    

    If gsEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        playSound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    ' The RC forms are measured in pixels so the positioning needs to pre-convert the twips into pixels
   
    fMain.aboutForm.Top = (glPhysicalScreenHeightPixels / 2) - (fMain.aboutForm.Height / 2)
    fMain.aboutForm.Left = (glPhysicalScreenWidthPixels / 2) - (fMain.aboutForm.Width / 2)
     
    fMain.aboutForm.Load
    fMain.aboutForm.Show
    
    'aboutWidget.opacity = 0
    aboutWidget.ShowMe = True
    aboutWidget.Widget.Refresh
    
    'fMain.aboutForm.Load
    'fMain.aboutForm.show
      
    If (fMain.aboutForm.WindowState = 1) Then
        fMain.aboutForm.WindowState = 0
    End If

   On Error GoTo 0
   Exit Sub

aboutClickEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure aboutClickEvent of Module Module1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : helpSplash
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub helpSplash()

    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo helpSplash_Error

    fileToPlay = "till.wav"
    If gsEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        playSound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

    fMain.helpForm.Top = (glPhysicalScreenHeightPixels / 2) - (fMain.helpForm.Height / 2)
    fMain.helpForm.Left = (glPhysicalScreenWidthPixels / 2) - (fMain.helpForm.Width / 2)
     
    'helpWidget.MyOpacity = 0
    helpWidget.ShowMe = True
    helpWidget.Widget.Refresh
    
    fMain.helpForm.Load
    fMain.helpForm.Show
    
     If (fMain.helpForm.WindowState = 1) Then
         fMain.helpForm.WindowState = 0
     End If

   On Error GoTo 0
   Exit Sub

helpSplash_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure helpSplash of Form menuForm"
     
End Sub
'---------------------------------------------------------------------------------------
' Procedure : licenceSplash
' Author    : beededea
' Date      : 03/08/2023
' Purpose   : a public subroutine called by two forms, handling the menu licence click event for both as well as some point in the startup main.
'---------------------------------------------------------------------------------------
'
Public Sub licenceSplash()

    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo licenceSplash_Error

'    If gsVolumeBoost = "1" Then
        fileToPlay = "till.wav"
'    Else
'        fileToPlay = "till-quiet.wav"
'    End If
    
    If gsEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        playSound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    
    fMain.licenceForm.Top = (glPhysicalScreenHeightPixels / 2) - (fMain.licenceForm.Height / 2)
    fMain.licenceForm.Left = (glPhysicalScreenWidthPixels / 2) - (fMain.licenceForm.Width / 2)
     
    'licenceWidget.opacity = 0
    'opacityflag = 0
    licenceWidget.ShowMe = True
    licenceWidget.Widget.Refresh
    
    fMain.licenceForm.Load
    fMain.licenceForm.Show

    ' the btnDecline_Click and btnAccept_Click are in modmain.bas
    
     If (fMain.licenceForm.WindowState = 1) Then
         fMain.licenceForm.WindowState = 0
     End If

   On Error GoTo 0
   Exit Sub

licenceSplash_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure licenceSplash of Form menuForm"
     
End Sub


'---------------------------------------------------------------------------------------
' Procedure : mnuCoffee_ClickEvent
' Author    : beededea
' Date      : 20/05/2023
' Purpose   : the public subroutine called by two forms, handling the specific menu click event for both
'---------------------------------------------------------------------------------------
'
Public Sub mnuCoffee_ClickEvent()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    On Error GoTo mnuCoffee_ClickEvent_Error
    
    answer = vbYes
    answerMsg = " Help support the creation of more widgets like this, DO send us a coffee! This button opens a browser window and connects to the Kofi donate page for this widget). Will you be kind and proceed?"
    answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Request to Donate a Kofi", True, "mnuCoffeeClickEvent")

    If answer = vbYes Then
        Call ShellExecute(menuForm.hWnd, "Open", "https://www.ko-fi.com/yereverluvinunclebert", vbNullString, App.Path, 1)
    End If

   On Error GoTo 0
   Exit Sub

mnuCoffee_ClickEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuCoffee_ClickEvent of Module Module1"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuSupport_ClickEvent
' Author    : beededea
' Date      : 20/05/2023
' Purpose   : the public subroutine called by two forms, handling the specific menu click event for both
'---------------------------------------------------------------------------------------
'
Public Sub mnuSupport_ClickEvent()

    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString

    On Error GoTo mnuSupport_ClickEvent_Error
    
    answer = vbYes
    answerMsg = "Visiting the support page - this button opens a browser window and connects to our Github issues page where you can send us a support query. Proceed?"
    answer = msgBoxA(answerMsg, vbExclamation + vbYesNo, "Request to Contact Support", True, "mnuSupportClickEvent")

    If answer = vbYes Then
        Call ShellExecute(menuForm.hWnd, "Open", "https://github.com/yereverluvinunclebert/TenShillings-" & gsRichClientEnvironment & "-Widget-" & gsCodingEnvironment & "/issues", vbNullString, App.Path, 1)
    End If

   On Error GoTo 0
   Exit Sub

mnuSupport_ClickEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuSupport_ClickEvent of Module Module1"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : mnuLicence_ClickEvent
' Author    : beededea
' Date      : 20/05/2023
' Purpose   : the public subroutine called by two forms, handling the specific menu click event for both
'---------------------------------------------------------------------------------------
'
Public Sub mnuLicence_ClickEvent()

   On Error GoTo mnuLicence_ClickEvent_Error
    
    Call licenceSplash

   On Error GoTo 0
   Exit Sub

mnuLicence_ClickEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mnuLicence_ClickEvent of Module Module1"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : setRichClientTooltips
' Author    : beededea
' Date      : 15/05/2023
' Purpose   : Set the tooltips using RC tooltip functionality only.
'             Note: there are also the balloon tooltips and standard VB tooltips set elsewhere.
'---------------------------------------------------------------------------------------
'
Public Sub setRichClientTooltips()
   On Error GoTo setRichClientTooltips_Error

    If gsWidgetTooltips = "1" Then
        TenShillingsWidget.Widget.ToolTip = "Use Mouse scrollwheel UP/DOWN to rotate, press CTRL at same time to resize. "
        aboutWidget.Widget.ToolTip = "Click on me to make me go away."
    Else
        TenShillingsWidget.Widget.ToolTip = vbNullString
        aboutWidget.Widget.ToolTip = vbNullString
   End If
    
   Call ChangeToolTipWidgetDefaultSettings(Cairo.ToolTipWidget.Widget)

   On Error GoTo 0
   Exit Sub

setRichClientTooltips_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setRichClientTooltips of Module Module1"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : ChangeToolTipWidgetDefaultSettings
' Author    : beededea
' Date      : 20/06/2023
' Purpose   : Set the size and characteristics of the RC tooltips
'---------------------------------------------------------------------------------------
'
Public Sub ChangeToolTipWidgetDefaultSettings(ByRef My_Widget As cWidgetBase)

   On Error GoTo ChangeToolTipWidgetDefaultSettings_Error

    With My_Widget
    
        .FontName = gsWidgetFont
        .FontSize = Val(gsPrefsFontSizeLowDPI)
    
    End With

   On Error GoTo 0
   Exit Sub

ChangeToolTipWidgetDefaultSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ChangeToolTipWidgetDefaultSettings of Module Module1"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : makeVisibleFormElements
' Author    : beededea
' Date      : 01/03/2023
' Purpose   : adjust Form Position on startup placing form onto Correct Monitor when placed off screen due to
'             monitor/resolution changes.
'---------------------------------------------------------------------------------------
'
Public Sub makeVisibleFormElements()


    Dim formLeftPixels As Long: formLeftPixels = 0
    Dim formTopPixels As Long: formTopPixels = 0
    
    On Error GoTo makeVisibleFormElements_Error

    'NOTE that when you position a widget you are positioning the form it is drawn upon.

    If gsDpiAwareness = "1" Then
        formLeftPixels = Val(gsWidgetHighDpiXPos)
        formTopPixels = Val(gsWidgetHighDpiYPos)
    Else
        formLeftPixels = Val(gsWidgetLowDpiXPos)
        formTopPixels = Val(gsWidgetLowDpiYPos)
    End If
    
    ' The RC forms are measured in pixels, whereas the native forms are in twips, do remember that...

    glMonitorCount = fGetMonitorCount
'    If glMonitorCount > 1 Then
'        Call SetFormOnMonitor(fMain.TenShillingsForm.hWnd, formLeftPixels, formTopPixels)
'    Else
        fMain.TenShillingsForm.Left = formLeftPixels
        fMain.TenShillingsForm.Top = formTopPixels
'    End If
    
    fMain.TenShillingsForm.Show

    On Error GoTo 0
    Exit Sub

makeVisibleFormElements_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure makeVisibleFormElements of Module Module1"
            Resume Next
          End If
    End With
        
End Sub


'---------------------------------------------------------------------------------------
' Procedure : getkeypress
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : getting a keypress from the keyboard
'---------------------------------------------------------------------------------------
'
Public Sub getKeyPress(ByVal KeyCode As Integer, ByVal Shift As Integer)
    'Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
   
    On Error GoTo getkeypress_Error

    If gbCTRL_1 Or gbSHIFT_1 Then
        gbCTRL_1 = False
        gbSHIFT_1 = False
    End If
    
    If Shift Then
        gbSHIFT_1 = True
    End If

    Select Case KeyCode
        Case vbKeyControl
            gbCTRL_1 = True
        Case vbKeyShift
            gbSHIFT_1 = True
        Case 82 ' R
            If Shift = 1 Then Call hardRestart

        Case 116 ' Performing a hard restart message box shift+F5
            If Shift = 1 Then
                'answer = vbYes
                answerMsg = "Performing a hard restart now, press OK."
                'answer =
                msgBoxA answerMsg, vbExclamation + vbOK, "Performing a hard restart", True, "getKeypressHardRestart1"
                Call hardRestart
            Else
                Call reloadProgram 'f5 refresh button as per all browsers
            End If
    End Select
 
    On Error GoTo 0
   Exit Sub

getkeypress_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getkeypress of Module module1"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : determineScreenDimensions
' Author    : beededea
' Date      : 18/09/2020
' Purpose   : VB6 has a bug - the screen width determination is incorrect, the API call below resolves this.
'             In addition, it often returns a faulty value when a full screen game runs, changing the resolution
'             This routine, sets up vars that will be used for checking orientation changes
'---------------------------------------------------------------------------------------
'
Public Sub determineScreenDimensions()

   On Error GoTo determineScreenDimensions_Error
   
    'If giDebugFlg = 1 Then msgbox "% sub determineScreenDimensions"

    ' only calling TwipsPerPixelX/Y functions once on startup
    glScreenTwipsPerPixelY = fTwipsPerPixelY
    glScreenTwipsPerPixelX = fTwipsPerPixelX
    
    glPhysicalScreenHeightPixels = GetDeviceCaps(menuForm.hDC, VERTRES) ' we use the name of any form that we don't mind being loaded at this point
    glPhysicalScreenWidthPixels = GetDeviceCaps(menuForm.hDC, HORZRES)

    glPhysicalScreenHeightTwips = glPhysicalScreenHeightPixels * glScreenTwipsPerPixelY
    glPhysicalScreenWidthTwips = glPhysicalScreenWidthPixels * glScreenTwipsPerPixelX
    
    glVirtualScreenHeightPixels = fVirtualScreenHeight(True)
    glVirtualScreenWidthPixels = fVirtualScreenWidth(True)

    glVirtualScreenHeightTwips = fVirtualScreenHeight(False)
    glVirtualScreenWidthTwips = fVirtualScreenWidth(False)
    
    glOldPhysicalScreenHeightPixels = glPhysicalScreenHeightPixels ' will be used to check for orientation changes
    glOldPhysicalScreenWidthPixels = glPhysicalScreenWidthPixels
    
   On Error GoTo 0
   Exit Sub

determineScreenDimensions_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & " in procedure determineScreenDimensions of Module Module1"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : mainScreen
' Author    : beededea
' Date      : 04/05/2023
' Purpose   : Function to move the widget itself onto the main screen if it has been moved off so far it cannot be seen - an accident that is possible by user positioning
'             calculate the current hlocation in % of the screen
'             - called on startup and by timer
'---------------------------------------------------------------------------------------
'
Public Sub mainScreen()

    On Error GoTo mainScreen_Error

    ' check for aspect ratio and determine whether it is in portrait or landscape mode
    If glPhysicalScreenWidthPixels > glPhysicalScreenHeightPixels Then
        gsAspectRatio = "landscape"
    Else
        gsAspectRatio = "portrait"
    End If
    
    ' check if the widget has a lock for the screen type.
    If gsAspectRatio = "landscape" Then
        If gsWidgetLandscape = "1" Then
            If gsLandscapeFormHoffset <> vbNullString Then
                fMain.TenShillingsForm.Left = Val(gsLandscapeFormHoffset)
                fMain.TenShillingsForm.Top = Val(gsLandscapeFormVoffset)
            End If
        End If
        If gsAspectHidden = "2" Then
            Debug.Print "Hiding the widget for landscape mode"
            fMain.TenShillingsForm.Visible = False
        End If
    End If
    
    ' check if the widget has a lock for the screen type.
    If gsAspectRatio = "portrait" Then
        If gsWidgetPortrait = "1" Then
            fMain.TenShillingsForm.Left = Val(gsPortraitHoffset)
            fMain.TenShillingsForm.Top = Val(gsPortraitYoffset)
        End If
        If gsAspectHidden = "1" Then
            Debug.Print "Hiding the widget for portrait mode"
            fMain.TenShillingsForm.Visible = False
        End If
    End If
    
    ' calculate the on screen widget's solid visible portion's X position in relation to the left screen edge
    Call checkScreenEdgeLeft
    
    ' calculate the on screen widget's solid visible portion's Y position in relation to the top screen edge
    Call checkScreenEdgeTop
            
    ' calculate the on screen widget's solid visible portion's X position in relation to the right screen edge
    Call checkScreenEdgeRight

    ' calculate the on screen widget's solid visible portion's X position in relation to the bottom screen edge
    Call checkScreenEdgeBottom

    ' calculate the current hlocation in % of the screen
    ' store the current hlocation in % of the screen
    If gsWidgetPosition = "1" Then
        gsHLocationPercPrefValue = CStr(fMain.TenShillingsForm.Left / glVirtualScreenWidthPixels * 100)
        gsVLocationPercPrefValue = CStr(fMain.TenShillingsForm.Top / glVirtualScreenHeightPixels * 100)
    End If

   On Error GoTo 0
   Exit Sub

mainScreen_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mainScreen of Module Module1"

End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : checkScreenEdgeBottom
' Author    : beededea
' Date      : 14/09/2025
' Purpose   : calculate the on screen widget's solid visible portion's X position in relation to the bottom screen edge
'---------------------------------------------------------------------------------------
'
Private Sub checkScreenEdgeBottom()

    Dim widgetCurrentHeightPx As Long: widgetCurrentHeightPx = 0
    Dim formMidPointY As Long: formMidPointY = 0
    Dim widgetTopY As Long: widgetTopY = 0
    'Dim widgetLeftX As Long: widgetLeftX = 0
    Dim screenEdge As Long: screenEdge = 0
    
    On Error GoTo checkScreenEdgeBottom_Error

    If (fMain.TenShillingsForm.Top + fMain.TenShillingsForm.Height) > glVirtualScreenHeightPixels Then  ' if any part of the form is off screen
        ' the widget height is divided by two as it is doubled earlier
        widgetCurrentHeightPx = (TenShillingsWidget.Widget.Height / 2 * TenShillingsWidget.Zoom) ' pixels
        formMidPointY = (fMain.TenShillingsForm.Height / 2) + fMain.TenShillingsForm.Top
        widgetTopY = formMidPointY '- (widgetCurrentHeightPx)
        screenEdge = 100 ' pixels from the edge
        
        ' if the widget itself is close to the left of the screen then reposition it back on screen
        If widgetTopY >= (glVirtualScreenHeightPixels - screenEdge) Then
            fMain.TenShillingsForm.Top = ((glVirtualScreenHeightPixels - screenEdge) - (widgetCurrentHeightPx / 2) - 300)
        End If
    End If

    On Error GoTo 0
    Exit Sub

checkScreenEdgeBottom_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkScreenEdgeBottom of Module Module1"
    
End Sub




'---------------------------------------------------------------------------------------
' Procedure : checkScreenEdgeRight
' Author    : beededea
' Date      : 14/09/2025
' Purpose   : calculate the on screen widget's solid visible portion's X position in relation to the right screen edge
'---------------------------------------------------------------------------------------
'
Private Sub checkScreenEdgeRight()

    Dim widgetCurrentWidthPx As Long: widgetCurrentWidthPx = 0
    Dim formMidPointX As Long: formMidPointX = 0
    'Dim widgetBottomY As Long: widgetBottomY = 0
    Dim widgetLeftX As Long: widgetLeftX = 0
    Dim screenEdge As Long: screenEdge = 0
    
    On Error GoTo checkScreenEdgeRight_Error

    If (fMain.TenShillingsForm.Left + fMain.TenShillingsForm.Width) > glVirtualScreenWidthPixels Then ' if any part of the form is off screen
        widgetCurrentWidthPx = (TenShillingsWidget.Widget.Width / 2 * TenShillingsWidget.Zoom) ' pixels
        formMidPointX = (fMain.TenShillingsForm.Width / 2) + fMain.TenShillingsForm.Left
        widgetLeftX = formMidPointX ' - widgetCurrentWidthPx / 2
        screenEdge = 100 ' pixels from the edge

        ' if the widget itself is close to the very left of the screen then reposition it back on screen
        If widgetLeftX >= glVirtualScreenWidthPixels - screenEdge Then
            fMain.TenShillingsForm.Left = (glVirtualScreenWidthPixels - screenEdge) - (fMain.TenShillingsForm.Width / 2) + (widgetCurrentWidthPx / 2) - 100
        End If
    End If

    On Error GoTo 0
    Exit Sub

checkScreenEdgeRight_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkScreenEdgeRight of Module Module1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : checkScreenEdgeLeft
' Author    : beededea
' Date      : 13/09/2025
' Purpose   : ' calculate the on screen widget's solid visible portion's X position in relation to the left screen edge
'---------------------------------------------------------------------------------------
'
Private Sub checkScreenEdgeLeft()

    Dim widgetCurrentWidthPx As Long: widgetCurrentWidthPx = 0
    Dim formMidPointX As Long: formMidPointX = 0
    'Dim widgetBottomY As Long: widgetBottomY = 0
    Dim widgetRightX As Long: widgetRightX = 0
    Dim screenEdge As Long: screenEdge = 0
    
    On Error GoTo checkScreenEdgeLeft_Error

    If fMain.TenShillingsForm.Left < 0 Then ' if any part of the form is off screen
        widgetCurrentWidthPx = (TenShillingsWidget.Widget.Width / 2 * TenShillingsWidget.Zoom) ' pixels
        formMidPointX = (fMain.TenShillingsForm.Width / 2) + fMain.TenShillingsForm.Left
        widgetRightX = formMidPointX '+ widgetCurrentWidthPx / 2
        screenEdge = 100 ' pixels from the edge
        
        ' if the widget itself is close to the very left of the screen then reposition it back on screen
        If widgetRightX <= screenEdge Then
            fMain.TenShillingsForm.Left = screenEdge - (((fMain.TenShillingsForm.Width / 2) + widgetCurrentWidthPx / 2)) + 100
        End If
    End If

    On Error GoTo 0
    Exit Sub

checkScreenEdgeLeft_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkScreenEdgeLeft of Module Module1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : checkScreenEdgeTop
' Author    : beededea
' Date      : 13/09/2025
' Purpose   : calculate the on screen widget's solid visible portion's Y position in relation to the top screen edge
'---------------------------------------------------------------------------------------
'
Private Sub checkScreenEdgeTop()
    Dim widgetCurrentHeightPx As Long: widgetCurrentHeightPx = 0
    Dim formMidPointY As Long: formMidPointY = 0
    Dim widgetBottomY As Long: widgetBottomY = 0
    'Dim widgetTopY As Long: widgetTopY = 0
    Dim screenEdge As Long: screenEdge = 0
    
    On Error GoTo checkScreenEdgeTop_Error

    If fMain.TenShillingsForm.Top < 0 Then ' if any part of the form is off screen
        widgetCurrentHeightPx = (TenShillingsWidget.Widget.Height / 2 * TenShillingsWidget.Zoom) ' pixels
        formMidPointY = (fMain.TenShillingsForm.Height / 2) + fMain.TenShillingsForm.Top
        widgetBottomY = formMidPointY + widgetCurrentHeightPx / 2
        screenEdge = 100 ' pixels from the edge
        
        ' if the widget itself is close to the left of the screen then reposition it back on screen
        If formMidPointY <= screenEdge Then
            fMain.TenShillingsForm.Top = (screenEdge - ((fMain.TenShillingsForm.Height / 2) + (widgetCurrentHeightPx / 2))) + 300
        End If
    End If

    On Error GoTo 0
    Exit Sub

checkScreenEdgeTop_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure checkScreenEdgeTop of Module Module1"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : thisForm_Unload
' Author    : beededea
' Date      : 18/08/2022
' Purpose   : the standard form unload routine called from several places
'---------------------------------------------------------------------------------------
'
Public Sub thisForm_Unload() ' name follows VB6 standard naming convention
    On Error GoTo Form_Unload_Error
    
    Call saveMainRCFormPosition
    
    Call unloadAllForms(True)

    On Error GoTo 0
    Exit Sub

Form_Unload_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Class Module module1"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : unloadAllForms
' Author    : beededea
' Date      : 28/06/2023
' Purpose   : unload all VB6 and RC6 forms
'---------------------------------------------------------------------------------------
'
Public Sub unloadAllForms(ByVal endItAll As Boolean)
    
   On Error GoTo unloadAllForms_Error
   
    ' empty all asynchronous sound buffers and release
    'Call FreeSound(ALL_SOUND_BUFFERS)
   
    ' stop all VB6 timers in the timer form
    frmTimer.revealWidgetTimer.Enabled = False
    frmTimer.tmrScreenResolution.Enabled = False
    frmTimer.unhideTimer.Enabled = False
    frmTimer.sleepTimer.Enabled = False
    
    ' stop all VB6 timers in the prefs form
    
    'widgetPrefs.tmrPrefsMonitorSaveHeight.Enabled = False
    widgetPrefs.themeTimer.Enabled = False
    widgetPrefs.tmrPrefsScreenResolution.Enabled = False
    widgetPrefs.tmrWritePositionAndSize.Enabled = False

    'unload the RC6 widgets on the RC6 forms first
    
    aboutWidget.Widgets.RemoveAll
    helpWidget.Widgets.RemoveAll
    fMain.TenShillingsForm.Widgets.RemoveAll
    
    ' unload the native VB6 forms
    
    Unload frmMessage
    Unload widgetPrefs
    Unload frmTimer
    Unload menuForm

    ' RC6's own method for killing forms
    
    fMain.aboutForm.Unload
    fMain.helpForm.Unload
    fMain.TenShillingsForm.Unload
    fMain.licenceForm.Unload
    
    ' remove all variable references to each RC form in turn
    
    Set fMain.aboutForm = Nothing
    Set fMain.helpForm = Nothing
    Set fMain.TenShillingsForm = Nothing
    Set fMain.licenceForm = Nothing
    
    ' remove all variable references to each VB6 form in turn
    
    Set widgetPrefs = Nothing
    Set frmTimer = Nothing
    Set menuForm = Nothing
    Set frmMessage = Nothing
    
    On Error Resume Next
    
    If endItAll = True Then End

   On Error GoTo 0
   Exit Sub

unloadAllForms_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure unloadAllForms of Module Module1"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : reloadProgram
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : called from several places to reload the program
'---------------------------------------------------------------------------------------
'
Public Sub reloadProgram()
    
    On Error GoTo reloadProgram_Error
    
    'TenShillingsWidget.ShowHelp = False ' needs to be set to false for the reload to reshow it, if enabled
    
    gbThisWidgetAvailable = False ' tell the ' screenWrite util that the widgetForm is no longer available to write console events to
    gbReload = True
    
    Call saveMainRCFormPosition
    
    Call unloadAllForms(False) ' unload forms but do not END
    
    ' this will call the routines as called by sub main() and initialise the program and RELOAD the RC6 forms.
    Call mainRoutine(True) ' sets the restart flag to avoid repriming the RC6 message pump.

    On Error GoTo 0
    Exit Sub

reloadProgram_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure reloadProgram of Module Module1"
            Resume Next
          End If
    End With

End Sub

'---------------------------------------------------------------------------------------
' Procedure : saveMainRCFormPosition
' Author    : beededea
' Date      : 04/08/2023
' Purpose   : called from several locations saves the widget X,Y positions in high or low DPI forms as well as the current size
'---------------------------------------------------------------------------------------
'
Public Sub saveMainRCFormPosition()

   On Error GoTo saveMainRCFormPosition_Error

    If gsDpiAwareness = "1" Then
        gsWidgetHighDpiXPos = CStr(fMain.TenShillingsForm.Left) ' saving in pixels
        gsWidgetHighDpiYPos = CStr(fMain.TenShillingsForm.Top)
        sPutINISetting "Software\TenShillings", "widgetHighDpiXPos", gsWidgetHighDpiXPos, gsSettingsFile
        sPutINISetting "Software\TenShillings", "widgetHighDpiYPos", gsWidgetHighDpiYPos, gsSettingsFile

    Else
        gsWidgetLowDpiXPos = CStr(fMain.TenShillingsForm.Left) ' saving in pixels
        gsWidgetLowDpiYPos = CStr(fMain.TenShillingsForm.Top)
        sPutINISetting "Software\TenShillings", "widgetLowDpiXPos", gsWidgetLowDpiXPos, gsSettingsFile
        sPutINISetting "Software\TenShillings", "widgetLowDpiYPos", gsWidgetLowDpiYPos, gsSettingsFile
    End If
    
    sPutINISetting "Software\TenShillings", "widgetPrimaryHeightRatio", gsWidgetPrimaryHeightRatio, gsSettingsFile
    sPutINISetting "Software\TenShillings", "widgetSecondaryHeightRatio", gsWidgetSecondaryHeightRatio, gsSettingsFile
    gsWidgetSize = CStr(TenShillingsWidget.Zoom * 100)
    gsSkewDegrees = CStr(TenShillingsWidget.SkewDegrees)
    
    sPutINISetting "Software\TenShillings", "widgetSize", gsWidgetSize, gsSettingsFile
    sPutINISetting "Software\TenShillings", "skewDegrees", gsSkewDegrees, gsSettingsFile

   On Error GoTo 0
   Exit Sub

saveMainRCFormPosition_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure saveMainRCFormPosition of Module Module1"
    
End Sub
    
 '---------------------------------------------------------------------------------------
' Procedure : saveMainRCFormSize
' Author    : beededea
' Date      : 04/08/2023
' Purpose   : called from several locations saves the widget X,Y positions in high or low DPI forms as well as the current size
'---------------------------------------------------------------------------------------
'
Public Sub saveMainRCFormSize()

   On Error GoTo saveMainRCFormSize_Error

    sPutINISetting "Software\TenShillings", "widgetPrimaryHeightRatio", gsWidgetPrimaryHeightRatio, gsSettingsFile
    sPutINISetting "Software\TenShillings", "widgetSecondaryHeightRatio", gsWidgetSecondaryHeightRatio, gsSettingsFile
    gsWidgetSize = CStr(TenShillingsWidget.Zoom * 100)
    gsSkewDegrees = CStr(TenShillingsWidget.SkewDegrees)
    sPutINISetting "Software\TenShillings", "widgetSize", gsWidgetSize, gsSettingsFile
    sPutINISetting "Software\TenShillings", "skewDegrees", gsSkewDegrees, gsSettingsFile

   On Error GoTo 0
   Exit Sub

saveMainRCFormSize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure saveMainRCFormSize of Module Module1"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : makeProgramPreferencesAvailable
' Author    : beededea
' Date      : 01/05/2023
' Purpose   : open the VB6 preference form in the correct position
'---------------------------------------------------------------------------------------
'
Public Sub makeProgramPreferencesAvailable()
    On Error GoTo makeProgramPreferencesAvailable_Error
    
    If widgetPrefs.IsVisible = False Then
    
        widgetPrefs.Visible = True
        widgetPrefs.Show  ' show it again
        widgetPrefs.SetFocus

        If widgetPrefs.WindowState = vbMinimized Then
            widgetPrefs.WindowState = vbNormal
        End If
        
        ' set the current position of the utility according to previously stored positions
        
        Call readPrefsPosition
        Call widgetPrefs.positionPrefsMonitor
        
    Else
        widgetPrefs.SetFocus
    End If
    
    widgetPrefs.btnSave.Enabled = False

   On Error GoTo 0
   Exit Sub

makeProgramPreferencesAvailable_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure makeProgramPreferencesAvailable of module module1.bas"
End Sub
    

'---------------------------------------------------------------------------------------
' Procedure : readPrefsPosition
' Author    : beededea
' Date      : 28/05/2023
' Purpose   : read the form X/Y params from the toolSettings.ini
'---------------------------------------------------------------------------------------
'
Public Sub readPrefsPosition()

    'Dim gPrefsMonitorStruct As UDTMonitor
    'Dim prefsFormMonitorID As Long: prefsFormMonitorID = 0
            
    On Error GoTo readPrefsPosition_Error

    If gsDpiAwareness = "1" Then
        gsPrefsHighDpiXPosTwips = fGetINISetting("Software\TenShillings", "prefsHighDpiXPosTwips", gsSettingsFile)
        gsPrefsHighDpiYPosTwips = fGetINISetting("Software\TenShillings", "prefsHighDpiYPosTwips", gsSettingsFile)
        
        ' if a current location not stored then position to the middle of the screen
        If gsPrefsHighDpiXPosTwips <> "" Then
            widgetPrefs.Left = Val(gsPrefsHighDpiXPosTwips)
        Else
            widgetPrefs.Left = glPhysicalScreenWidthTwips / 2 - widgetPrefs.Width / 2
        End If
        
        gsPrefsHighDpiXPosTwips = widgetPrefs.Left

        If gsPrefsHighDpiYPosTwips <> "" Then
            widgetPrefs.Top = Val(gsPrefsHighDpiYPosTwips)
        Else
            widgetPrefs.Top = Screen.Height / 2 - widgetPrefs.Height / 2
        End If
        
        gsPrefsHighDpiYPosTwips = widgetPrefs.Top
        
    Else
        gsPrefsLowDpiXPosTwips = fGetINISetting("Software\TenShillings", "prefsLowDpiXPosTwips", gsSettingsFile)
        gsPrefsLowDpiYPosTwips = fGetINISetting("Software\TenShillings", "prefsLowDpiYPosTwips", gsSettingsFile)
              
        ' if a current location not stored then position to the middle of the screen
        If gsPrefsLowDpiXPosTwips <> "" Then
            widgetPrefs.Left = Val(gsPrefsLowDpiXPosTwips)
        Else
            widgetPrefs.Left = glPhysicalScreenWidthTwips / 2 - widgetPrefs.Width / 2
        End If
        
        gsPrefsLowDpiXPosTwips = widgetPrefs.Left

        If gsPrefsLowDpiYPosTwips <> "" Then
            widgetPrefs.Top = Val(gsPrefsLowDpiYPosTwips)
        Else
            widgetPrefs.Top = Screen.Height / 2 - widgetPrefs.Height / 2
        End If
        
        gsPrefsLowDpiYPosTwips = widgetPrefs.Top
    End If
        
    gsPrefsPrimaryHeightTwips = fGetINISetting("Software\TenShillings", "prefsPrimaryHeightTwips", gsSettingsFile)
    gsPrefsSecondaryHeightTwips = fGetINISetting("Software\TenShillings", "prefsSecondaryHeightTwips", gsSettingsFile)
        
   ' on very first install this will be zero, then size of the prefs as a proportion of the screen size
    If gsPrefsPrimaryHeightTwips = "" Then
        If Screen.Height > gdPrefsStartHeight * 2 Then
            gsPrefsPrimaryHeightTwips = Screen.Height / 2
        Else
            gsPrefsPrimaryHeightTwips = gdPrefsStartHeight
        End If
    End If
    
   On Error GoTo 0
   Exit Sub

readPrefsPosition_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readPrefsPosition of Module Module1"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : writePrefsPositionAndSize
' Author    : beededea
' Date      : 28/05/2023
' Purpose   : save the current X and y position of this form to allow repositioning when restarting
'             also the height of the form on a per monitor basis, called when closing and via a timer
'---------------------------------------------------------------------------------------
'
Public Sub writePrefsPositionAndSize()
         
    On Error GoTo writePrefsPositionAndSize_Error

    If widgetPrefs.IsVisible = True And widgetPrefs.WindowState = vbNormal Then ' when vbMinimised the value = -48000  !
        If gsDpiAwareness = "1" Then
            gsPrefsHighDpiXPosTwips = CStr(widgetPrefs.Left)
            gsPrefsHighDpiYPosTwips = CStr(widgetPrefs.Top)
            
            ' now write those params to the toolSettings.ini
            sPutINISetting "Software\TenShillings", "prefsHighDpiXPosTwips", gsPrefsHighDpiXPosTwips, gsSettingsFile
            sPutINISetting "Software\TenShillings", "prefsHighDpiYPosTwips", gsPrefsHighDpiYPosTwips, gsSettingsFile
        Else
            gsPrefsLowDpiXPosTwips = CStr(widgetPrefs.Left)
            gsPrefsLowDpiYPosTwips = CStr(widgetPrefs.Top)
            
            ' now write those params to the toolSettings.ini
            sPutINISetting "Software\TenShillings", "prefsLowDpiXPosTwips", gsPrefsLowDpiXPosTwips, gsSettingsFile
            sPutINISetting "Software\TenShillings", "prefsLowDpiYPosTwips", gsPrefsLowDpiYPosTwips, gsSettingsFile
            
        End If
        
        If LTrim$(gsMultiMonitorResize) <> "2" Then Exit Sub

        If gPrefsMonitorStruct.IsPrimary = True Then
            gsPrefsPrimaryHeightTwips = Trim$(CStr(widgetPrefs.Height))
            sPutINISetting "Software\TenShillings", "prefsPrimaryHeightTwips", gsPrefsPrimaryHeightTwips, gsSettingsFile
        Else
            gsPrefsSecondaryHeightTwips = Trim$(CStr(widgetPrefs.Height))
            sPutINISetting "Software\TenShillings", "prefsSecondaryHeightTwips", gsPrefsSecondaryHeightTwips, gsSettingsFile
        End If
    End If
    
    On Error GoTo 0
   Exit Sub

writePrefsPositionAndSize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writePrefsPositionAndSize of Form widgetPrefs"
End Sub


''---------------------------------------------------------------------------------------
'' Procedure : settingsTimer_Timer
'' Author    : beededea
'' Date      : 03/03/2023
'' Purpose   : Checking the date/time of the settings.ini file meaning that another tool has edited the settings
''---------------------------------------------------------------------------------------
'' this has to be in a shared module and not in the prefs form as it will be running in the normal context woithout prefs showing.
'
'Private Sub settingsTimer_Timer()
'
'    gsUnhide = fGetINISetting("Software\TenShillings", "unhide", gsSettingsFile)
'
'    If gsUnhide = "true" Then
'        'TenShillingsWidget.Hidden = False
'        fMain.TenShillingsForm.Visible = True
'        sPutINISetting "Software\TenShillings", "unhide", vbNullString, gsSettingsFile
'    End If
'
'    On Error GoTo 0
'    Exit Sub
'
'settingsTimer_Timer_Error:
'
'    With Err
'         If .Number <> 0 Then
'            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure settingsTimer_Timer of Form module1.bas"
'            Resume Next
'          End If
'    End With
'End Sub






'---------------------------------------------------------------------------------------
' Procedure : toggleWidgetLock
' Author    : beededea
' Date      : 03/08/2023
' Purpose   : called from a few locations, toggles the lock state and saves it
'---------------------------------------------------------------------------------------
'
Public Sub toggleWidgetLock()
    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo toggleWidgetLock_Error

    fileToPlay = "lock.wav"
    
    If gsPreventDragging = "1" Then
        ' Call ' screenWrite("Widget lock released")
        menuForm.mnuLockWidget.Checked = False
        If widgetPrefs.IsLoaded = True Then widgetPrefs.chkPreventDragging.Value = 0
        gsPreventDragging = "0"
        TenShillingsWidget.Locked = False
    Else
        ' Call ' screenWrite("Widget locked in place")
        menuForm.mnuLockWidget.Checked = True
        If widgetPrefs.IsLoaded = True Then widgetPrefs.chkPreventDragging.Value = 1
        TenShillingsWidget.Locked = True
        gsPreventDragging = "1"
    End If
    
    fMain.TenShillingsForm.Refresh
    
    sPutINISetting "Software\TenShillings", "preventDragging", gsPreventDragging, gsSettingsFile
   
    If gsEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        playSound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    On Error GoTo 0
   Exit Sub

toggleWidgetLock_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure toggleWidgetLock of Module Module1"

End Sub




'---------------------------------------------------------------------------------------
' Procedure : SwitchOff
' Author    : beededea
' Date      : 03/08/2023
' Purpose   : Turns off the functionality for the whole program, stopping all timers and then saves the state
'---------------------------------------------------------------------------------------
'
Public Sub SwitchOff()

   On Error GoTo SwitchOff_Error

    menuForm.mnuSwitchOff.Checked = True
    menuForm.mnuTurnFunctionsOn.Checked = False
    
    gsWidgetFunctions = "0"
    sPutINISetting "Software\TenShillings", "widgetFunctions", gsWidgetFunctions, gsSettingsFile

   On Error GoTo 0
   Exit Sub

SwitchOff_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SwitchOff of Module Module1"

End Sub



'---------------------------------------------------------------------------------------
' Procedure : TurnFunctionsOn
' Author    : beededea
' Date      : 03/08/2023
' Purpose   : turns only the main timer on and then saves the state
'---------------------------------------------------------------------------------------
'
Public Sub TurnFunctionsOn()
    Dim fileToPlay As String: fileToPlay = vbNullString

    On Error GoTo TurnFunctionsOn_Error

    fileToPlay = "ting.wav"

    If gsEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        playSound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

    menuForm.mnuSwitchOff.Checked = False
    menuForm.mnuTurnFunctionsOn.Checked = True
    
    gsWidgetFunctions = "1"
    sPutINISetting "Software\TenShillings", "widgetFunctions", gsWidgetFunctions, gsSettingsFile

   On Error GoTo 0
   Exit Sub

TurnFunctionsOn_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TurnFunctionsOn of Form menuForm"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : hardRestart
' Author    : beededea
' Date      : 14/08/2023
' Purpose   : Perform a hard restart using an external program causing the prefs to run immediately afterward.
'---------------------------------------------------------------------------------------
'
Public Sub hardRestart()
    'Dim answer As VbMsgBoxResult: answer = vbNo
    Dim answerMsg As String: answerMsg = vbNullString
    Dim thisCommand As String: thisCommand = vbNullString
    
    On Error GoTo hardRestart_Error

    thisCommand = App.Path & "\TenShillings-Widget-Restart.exe"
    
    If fFExists(thisCommand) Then
        
        ' run the helper program that kills the current process and restarts it
        Call ShellExecute(widgetPrefs.hWnd, "open", thisCommand, "TenShillings-" & gsRichClientEnvironment & "-Widget-" & gsCodingEnvironment & ".exe prefs", "", 1)
    Else
        'answer = MsgBox(thisCommand & " is missing", vbOKOnly + vbExclamation)
        answerMsg = thisCommand & " is missing"
        ' answer =
        msgBoxA answerMsg, vbOKOnly + vbExclamation, "Restart Error Notification", False
    End If

   On Error GoTo 0
   Exit Sub

hardRestart_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure hardRestart of Module Module1"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : InIDE
' Author    :
' Date      : 09/02/2021
' Purpose   : checks whether the code is running in the VB6 IDE or not
'---------------------------------------------------------------------------------------
'
Public Function InIDE() As Boolean

   On Error GoTo InIDE_Error

    ' .30 DAEB 03/03/2021 frmMain.frm replaced the inIDE function that used a variant to one without
    ' This will only be done if in the IDE
    Debug.Assert InDebugMode
    If pbDebugMode Then
        InIDE = True
    End If

   On Error GoTo 0
   Exit Function

InIDE_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InIDE of Module Module1"
End Function


'---------------------------------------------------------------------------------------
' Procedure : InDebugMode
' Author    : beededea
' Date      : 02/03/2021
' Purpose   : ' .30 DAEB 03/03/2021 frmMain.frm replaced the inIDE function that used a variant to one without
'---------------------------------------------------------------------------------------
'
Private Function InDebugMode() As Boolean
   On Error GoTo InDebugMode_Error

    pbDebugMode = True
    InDebugMode = True

   On Error GoTo 0
   Exit Function

InDebugMode_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InDebugMode of Module Module1"
End Function


'---------------------------------------------------------------------------------------
' Procedure : clearAllMessageBoxRegistryEntries
' Author    : beededea
' Date      : 11/04/2023
' Purpose   : Clear all the message box "show again" entries in the registry
'---------------------------------------------------------------------------------------
'
Public Sub clearAllMessageBoxRegistryEntries()
    On Error GoTo clearAllMessageBoxRegistryEntries_Error

    SaveSetting App.EXEName, "Options", "Show message" & "mnuFacebookClick", 0
    SaveSetting App.EXEName, "Options", "Show message" & "mnuLatestClick", 0
    SaveSetting App.EXEName, "Options", "Show message" & "mnuSweetsClick", 0
    SaveSetting App.EXEName, "Options", "Show message" & "mnuWidgetsClick", 0
    SaveSetting App.EXEName, "Options", "Show message" & "mnuCoffeeClickEvent", 0
    SaveSetting App.EXEName, "Options", "Show message" & "mnuSupportClickEvent", 0
    SaveSetting App.EXEName, "Options", "Show message" & "chkDpiAwarenessRestart", 0
    SaveSetting App.EXEName, "Options", "Show message" & "chkDpiAwarenessAbnormal", 0
    SaveSetting App.EXEName, "Options", "Show message" & "chkEnableTooltipsClick", 0
    SaveSetting App.EXEName, "Options", "Show message" & "lblGitHubDblClick", 0
    SaveSetting App.EXEName, "Options", "Show message" & "sliOpacityClick", 0
    SaveSetting App.EXEName, "Options", "Show message" & " ", 0

    On Error GoTo 0
    Exit Sub

clearAllMessageBoxRegistryEntries_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure clearAllMessageBoxRegistryEntries of Form dock"
            Resume Next
          End If
    End With
    
End Sub


''---------------------------------------------------------------------------------------
'' Procedure : determineIconWidth
'' Author    : beededea
'' Date      : 02/10/2023
'' Purpose   :
''---------------------------------------------------------------------------------------
''
'Public Function determineIconWidth(ByRef thisForm As Form, ByVal thisDynamicSizingFlg As Boolean) As Long
'
'    Dim topIconWidth As Long: topIconWidth = 0
'
'    On Error GoTo determineIconWidth_Error
'
''    If thisDynamicSizingFlg = False Then
''        'Exit Function
''    End If
'
'    If thisForm.Width < 10500 Then
'        topIconWidth = 600 '40 pixels
'    End If
'
'    If thisForm.Width >= 10500 And thisForm.Width < 12000 Then
'        topIconWidth = 730
'    End If
'
'    If thisForm.Width >= 12000 And thisForm.Width < 13500 Then
'        topIconWidth = 834
'    End If
'
'    If thisForm.Width >= 13500 And thisForm.Width < 15000 Then
'        topIconWidth = 940
'    End If
'
'    If thisForm.Width >= 15000 Then
'        topIconWidth = 1010
'    End If
'    'topIconWidth = 2000
'    determineIconWidth = topIconWidth
'
'    On Error GoTo 0
'    Exit Function
'
'determineIconWidth_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure determineIconWidth of Form widgetPrefs"
'
'End Function

'---------------------------------------------------------------------------------------
' Procedure : ArrayString
' Author    : beededea
' Date      : 09/10/2023
' Purpose   : allows population of a string array from a comma separated string
'             VB6 normally creates a variant when assigning a comma separated string to a var with an undeclared type
'             this avoids that scenario.
'---------------------------------------------------------------------------------------
'
Public Function ArrayString(ParamArray tokens()) As String() ' always byval
    On Error GoTo ArrayString_Error
    
    Dim Arr() As String

    ReDim Arr(UBound(tokens)) As String
    Dim I As Long
    For I = 0 To UBound(tokens)
        Arr(I) = tokens(I)
    Next
    ArrayString = Arr

    On Error GoTo 0
    Exit Function

ArrayString_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ArrayString of Module Module1"
End Function



'---------------------------------------------------------------------------------------
' Procedure : hideBusyTimer
' Author    : beededea
' Date      : 03/01/2025
' Purpose   : hide all of the sand timer image widgets used to make an animated sand timer
'---------------------------------------------------------------------------------------
'
Public Sub hideBusyTimer()

   On Error GoTo hideBusyTimer_Error

'    fMain.TenShillingsForm.Widgets("busy1").Widget.Alpha = 0
'    fMain.TenShillingsForm.Widgets("busy2").Widget.Alpha = 0
'    fMain.TenShillingsForm.Widgets("busy3").Widget.Alpha = 0
'    fMain.TenShillingsForm.Widgets("busy4").Widget.Alpha = 0
'    fMain.TenShillingsForm.Widgets("busy5").Widget.Alpha = 0
'    fMain.TenShillingsForm.Widgets("busy6").Widget.Alpha = 0
'
'    fMain.TenShillingsForm.Widgets("busy1").Widget.Refresh
'    fMain.TenShillingsForm.Widgets("busy2").Widget.Refresh
'    fMain.TenShillingsForm.Widgets("busy3").Widget.Refresh
'    fMain.TenShillingsForm.Widgets("busy4").Widget.Refresh
'    fMain.TenShillingsForm.Widgets("busy5").Widget.Refresh
'    fMain.TenShillingsForm.Widgets("busy6").Widget.Refresh

   On Error GoTo 0
   Exit Sub

hideBusyTimer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure hideBusyTimer of Form widgetPrefs"
    
End Sub






'---------------------------------------------------------------------------------------
' Procedure : saveRCFormCurrentSizeRatios
' Author    : beededea
' Date      : 01/10/2025
' Purpose   : saves the current ratios for the RC form alone
'---------------------------------------------------------------------------------------
'
Public Sub saveRCFormCurrentSizeRatios()
    Dim resizeProportion As Double: resizeProportion = 0

    On Error GoTo saveRCFormCurrentSizeRatios_Error

    If LTrim$(gsMultiMonitorResize) = "2" Then
        If gWidgetMonitorStruct.IsPrimary Then
            gsWidgetPrimaryHeightRatio = TenShillingsWidget.Zoom
            sPutINISetting "Software\TenShillings", "widgetPrimaryHeightRatio", gsWidgetPrimaryHeightRatio, gsSettingsFile
        Else
            gsWidgetSecondaryHeightRatio = TenShillingsWidget.Zoom
            sPutINISetting "Software\TenShillings", "widgetSecondaryHeightRatio", gsWidgetSecondaryHeightRatio, gsSettingsFile
        End If
    End If

    On Error GoTo 0
    Exit Sub

saveRCFormCurrentSizeRatios_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure saveRCFormCurrentSizeRatios of Form widgetPrefs"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : savePrefsFormCurrentSize
' Author    : beededea
' Date      : 01/10/2025
' Purpose   : saves the absolute sizes for the Prefs form
'---------------------------------------------------------------------------------------
'
Public Sub savePrefsFormCurrentSize()
    Dim resizeProportion As Double: resizeProportion = 0

    On Error GoTo savePrefsFormCurrentSize_Error

    If LTrim$(gsMultiMonitorResize) = "2" Then
        ' now save the prefs form absolute sizes
        If gPrefsMonitorStruct.IsPrimary = True Then
            sPutINISetting "Software\TenShillings", "prefsPrimaryHeightTwips", gsPrefsPrimaryHeightTwips, gsSettingsFile
        Else
            sPutINISetting "Software\TenShillings", "prefsSecondaryHeightTwips", gsPrefsSecondaryHeightTwips, gsSettingsFile
        End If
    End If

    On Error GoTo 0
    Exit Sub

savePrefsFormCurrentSize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure savePrefsFormCurrentSize of Form widgetPrefs"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : saveMainFormsCurrentSizeAndRatios
' Author    : beededea
' Date      : 01/10/2025
' Purpose   : saves the current ratios for the RC form and the absolute sizes for the Prefs form
'---------------------------------------------------------------------------------------
'
Public Sub saveMainFormsCurrentSizeAndRatios()

    On Error GoTo saveMainFormsCurrentSizeAndRatios_Error

    If LTrim$(gsMultiMonitorResize) = "2" Then
    
        '  saves the current ratios for the RC form alone
        Call saveRCFormCurrentSizeRatios
        
        ' saves the absolute sizes for the Prefs form
        Call savePrefsFormCurrentSize

    End If

    On Error GoTo 0
    Exit Sub

saveMainFormsCurrentSizeAndRatios_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure saveMainFormsCurrentSizeAndRatios of Form widgetPrefs"
    
End Sub



'---------------------------------------------------------------------------------------
' Procedure : gsWidgetSize
' Author    : beededea
' Date      : 06/09/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsWidgetSize() As String

    On Error GoTo gsWidgetSize_Error

    gsWidgetSize = m_sWidgetSize

    On Error GoTo 0
    Exit Property

gsWidgetSize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsWidgetSize of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsWidgetSize
' Author    : beededea
' Date      : 06/09/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsWidgetSize(ByVal sgsWidgetSize As String)

    On Error GoTo gsWidgetSize_Error

    m_sWidgetSize = sgsWidgetSize

    On Error GoTo 0
    Exit Property

gsWidgetSize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsWidgetSize of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsStartup
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsStartup() As String

    On Error GoTo gsStartup_Error

    gsStartup = m_sgsStartup

    On Error GoTo 0
    Exit Property

gsStartup_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsStartup of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsStartup
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsStartup(ByVal sgsStartup As String)

    On Error GoTo gsStartup_Error

    m_sgsStartup = sgsStartup

    On Error GoTo 0
    Exit Property

gsStartup_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsStartup of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsWidgetFunctions
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsWidgetFunctions() As String

    On Error GoTo gsWidgetFunctions_Error

    gsWidgetFunctions = m_sgsWidgetFunctions

    On Error GoTo 0
    Exit Property

gsWidgetFunctions_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsWidgetFunctions of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsWidgetFunctions
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsWidgetFunctions(ByVal sgsWidgetFunctions As String)

    On Error GoTo gsWidgetFunctions_Error

    m_sgsWidgetFunctions = sgsWidgetFunctions

    On Error GoTo 0
    Exit Property

gsWidgetFunctions_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsWidgetFunctions of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsPointerAnimate
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsPointerAnimate() As String

    On Error GoTo gsPointerAnimate_Error

    gsPointerAnimate = m_sgsPointerAnimate

    On Error GoTo 0
    Exit Property

gsPointerAnimate_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsPointerAnimate of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsPointerAnimate
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsPointerAnimate(ByVal sgsPointerAnimate As String)

    On Error GoTo gsPointerAnimate_Error

    m_sgsPointerAnimate = sgsPointerAnimate

    On Error GoTo 0
    Exit Property

gsPointerAnimate_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsPointerAnimate of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsMultiCoreEnable
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsMultiCoreEnable() As String

    On Error GoTo gsMultiCoreEnable_Error

    gsMultiCoreEnable = m_sgsMultiCoreEnable

    On Error GoTo 0
    Exit Property

gsMultiCoreEnable_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsMultiCoreEnable of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsMultiCoreEnable
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsMultiCoreEnable(ByVal sgsMultiCoreEnable As String)

    On Error GoTo gsMultiCoreEnable_Error

    m_sgsMultiCoreEnable = sgsMultiCoreEnable

    On Error GoTo 0
    Exit Property

gsMultiCoreEnable_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsMultiCoreEnable of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsSamplingInterval
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsSamplingInterval() As String

    On Error GoTo gsSamplingInterval_Error

    gsSamplingInterval = m_sgsSamplingInterval

    On Error GoTo 0
    Exit Property

gsSamplingInterval_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsSamplingInterval of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsSamplingInterval
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsSamplingInterval(ByVal sgsSamplingInterval As String)

    On Error GoTo gsSamplingInterval_Error

    m_sgsSamplingInterval = sgsSamplingInterval

    On Error GoTo 0
    Exit Property

gsSamplingInterval_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsSamplingInterval of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsSkewDegrees
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsSkewDegrees() As String

    On Error GoTo gsSkewDegrees_Error

    gsSkewDegrees = m_sgsSkewDegrees

    On Error GoTo 0
    Exit Property

gsSkewDegrees_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsSkewDegrees of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsSkewDegrees
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsSkewDegrees(ByVal sgsSkewDegrees As String)

    On Error GoTo gsSkewDegrees_Error

    m_sgsSkewDegrees = sgsSkewDegrees

    On Error GoTo 0
    Exit Property

gsSkewDegrees_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsSkewDegrees of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsWidgetTooltips
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsWidgetTooltips() As String

    On Error GoTo gsWidgetTooltips_Error

    gsWidgetTooltips = m_sgsWidgetTooltips

    On Error GoTo 0
    Exit Property

gsWidgetTooltips_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsWidgetTooltips of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsWidgetTooltips
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsWidgetTooltips(ByVal sgsWidgetTooltips As String)

    On Error GoTo gsWidgetTooltips_Error

    m_sgsWidgetTooltips = sgsWidgetTooltips

    On Error GoTo 0
    Exit Property

gsWidgetTooltips_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsWidgetTooltips of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsPrefsTooltips
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsPrefsTooltips() As String

    On Error GoTo gsPrefsTooltips_Error

    gsPrefsTooltips = m_sgsPrefsTooltips

    On Error GoTo 0
    Exit Property

gsPrefsTooltips_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsPrefsTooltips of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsPrefsTooltips
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsPrefsTooltips(ByVal sgsPrefsTooltips As String)

    On Error GoTo gsPrefsTooltips_Error

    m_sgsPrefsTooltips = sgsPrefsTooltips

    On Error GoTo 0
    Exit Property

gsPrefsTooltips_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsPrefsTooltips of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsShowTaskbar
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsShowTaskbar() As String

    On Error GoTo gsShowTaskbar_Error

    gsShowTaskbar = m_sgsShowTaskbar

    On Error GoTo 0
    Exit Property

gsShowTaskbar_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsShowTaskbar of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsShowTaskbar
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsShowTaskbar(ByVal sgsShowTaskbar As String)

    On Error GoTo gsShowTaskbar_Error

    m_sgsShowTaskbar = sgsShowTaskbar

    On Error GoTo 0
    Exit Property

gsShowTaskbar_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsShowTaskbar of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsShowHelp
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsShowHelp() As String

    On Error GoTo gsShowHelp_Error

    gsShowHelp = m_sgsShowHelp

    On Error GoTo 0
    Exit Property

gsShowHelp_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsShowHelp of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsShowHelp
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsShowHelp(ByVal sgsShowHelp As String)

    On Error GoTo gsShowHelp_Error

    m_sgsShowHelp = sgsShowHelp

    On Error GoTo 0
    Exit Property

gsShowHelp_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsShowHelp of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsDpiAwareness
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsDpiAwareness() As String

    On Error GoTo gsDpiAwareness_Error

    gsDpiAwareness = m_sgsDpiAwareness

    On Error GoTo 0
    Exit Property

gsDpiAwareness_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsDpiAwareness of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsDpiAwareness
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsDpiAwareness(ByVal sgsDpiAwareness As String)

    On Error GoTo gsDpiAwareness_Error

    m_sgsDpiAwareness = sgsDpiAwareness

    On Error GoTo 0
    Exit Property

gsDpiAwareness_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsDpiAwareness of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsScrollWheelDirection
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsScrollWheelDirection() As String

    On Error GoTo gsScrollWheelDirection_Error

    gsScrollWheelDirection = m_sgsScrollWheelDirection

    On Error GoTo 0
    Exit Property

gsScrollWheelDirection_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsScrollWheelDirection of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsScrollWheelDirection
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsScrollWheelDirection(ByVal sgsScrollWheelDirection As String)

    On Error GoTo gsScrollWheelDirection_Error

    m_sgsScrollWheelDirection = sgsScrollWheelDirection

    On Error GoTo 0
    Exit Property

gsScrollWheelDirection_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsScrollWheelDirection of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsWidgetHighDpiXPos
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsWidgetHighDpiXPos() As String

    On Error GoTo gsWidgetHighDpiXPos_Error

    gsWidgetHighDpiXPos = m_sgsWidgetHighDpiXPos

    On Error GoTo 0
    Exit Property

gsWidgetHighDpiXPos_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsWidgetHighDpiXPos of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsWidgetHighDpiXPos
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsWidgetHighDpiXPos(ByVal sgsWidgetHighDpiXPos As String)

    On Error GoTo gsWidgetHighDpiXPos_Error

    m_sgsWidgetHighDpiXPos = sgsWidgetHighDpiXPos

    On Error GoTo 0
    Exit Property

gsWidgetHighDpiXPos_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsWidgetHighDpiXPos of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsWidgetHighDpiYPos
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsWidgetHighDpiYPos() As String

    On Error GoTo gsWidgetHighDpiYPos_Error

    gsWidgetHighDpiYPos = m_sgsWidgetHighDpiYPos

    On Error GoTo 0
    Exit Property

gsWidgetHighDpiYPos_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsWidgetHighDpiYPos of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsWidgetHighDpiYPos
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsWidgetHighDpiYPos(ByVal sgsWidgetHighDpiYPos As String)

    On Error GoTo gsWidgetHighDpiYPos_Error

    m_sgsWidgetHighDpiYPos = sgsWidgetHighDpiYPos

    On Error GoTo 0
    Exit Property

gsWidgetHighDpiYPos_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsWidgetHighDpiYPos of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsWidgetLowDpiXPos
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsWidgetLowDpiXPos() As String

    On Error GoTo gsWidgetLowDpiXPos_Error

    gsWidgetLowDpiXPos = m_sgsWidgetLowDpiXPos

    On Error GoTo 0
    Exit Property

gsWidgetLowDpiXPos_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsWidgetLowDpiXPos of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsWidgetLowDpiXPos
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsWidgetLowDpiXPos(ByVal sgsWidgetLowDpiXPos As String)

    On Error GoTo gsWidgetLowDpiXPos_Error

    m_sgsWidgetLowDpiXPos = sgsWidgetLowDpiXPos

    On Error GoTo 0
    Exit Property

gsWidgetLowDpiXPos_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsWidgetLowDpiXPos of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsWidgetLowDpiYPos
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsWidgetLowDpiYPos() As String

    On Error GoTo gsWidgetLowDpiYPos_Error

    gsWidgetLowDpiYPos = m_sgsWidgetLowDpiYPos

    On Error GoTo 0
    Exit Property

gsWidgetLowDpiYPos_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsWidgetLowDpiYPos of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsWidgetLowDpiYPos
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsWidgetLowDpiYPos(ByVal sgsWidgetLowDpiYPos As String)

    On Error GoTo gsWidgetLowDpiYPos_Error

    m_sgsWidgetLowDpiYPos = sgsWidgetLowDpiYPos

    On Error GoTo 0
    Exit Property

gsWidgetLowDpiYPos_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsWidgetLowDpiYPos of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsClockFont
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsClockFont() As String

    On Error GoTo gsClockFont_Error

    gsClockFont = m_sgsClockFont

    On Error GoTo 0
    Exit Property

gsClockFont_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsClockFont of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsClockFont
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsClockFont(ByVal sgsClockFont As String)

    On Error GoTo gsClockFont_Error

    m_sgsClockFont = sgsClockFont

    On Error GoTo 0
    Exit Property

gsClockFont_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsClockFont of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsWidgetFont
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsWidgetFont() As String

    On Error GoTo gsWidgetFont_Error

    gsWidgetFont = m_sgsWidgetFont

    On Error GoTo 0
    Exit Property

gsWidgetFont_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsWidgetFont of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsWidgetFont
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsWidgetFont(ByVal sgsWidgetFont As String)

    On Error GoTo gsWidgetFont_Error

    m_sgsWidgetFont = sgsWidgetFont

    On Error GoTo 0
    Exit Property

gsWidgetFont_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsWidgetFont of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsPrefsFont
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsPrefsFont() As String

    On Error GoTo gsPrefsFont_Error

    gsPrefsFont = m_sgsPrefsFont

    On Error GoTo 0
    Exit Property

gsPrefsFont_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsPrefsFont of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsPrefsFont
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsPrefsFont(ByVal sgsPrefsFont As String)

    On Error GoTo gsPrefsFont_Error

    m_sgsPrefsFont = sgsPrefsFont

    On Error GoTo 0
    Exit Property

gsPrefsFont_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsPrefsFont of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsPrefsFontSizeHighDPI
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsPrefsFontSizeHighDPI() As String

    On Error GoTo gsPrefsFontSizeHighDPI_Error

    gsPrefsFontSizeHighDPI = m_sgsPrefsFontSizeHighDPI

    On Error GoTo 0
    Exit Property

gsPrefsFontSizeHighDPI_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsPrefsFontSizeHighDPI of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsPrefsFontSizeHighDPI
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsPrefsFontSizeHighDPI(ByVal sgsPrefsFontSizeHighDPI As String)

    On Error GoTo gsPrefsFontSizeHighDPI_Error

    m_sgsPrefsFontSizeHighDPI = sgsPrefsFontSizeHighDPI

    On Error GoTo 0
    Exit Property

gsPrefsFontSizeHighDPI_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsPrefsFontSizeHighDPI of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsPrefsFontSizeLowDPI
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsPrefsFontSizeLowDPI() As String

    On Error GoTo gsPrefsFontSizeLowDPI_Error

    gsPrefsFontSizeLowDPI = m_sgsPrefsFontSizeLowDPI

    On Error GoTo 0
    Exit Property

gsPrefsFontSizeLowDPI_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsPrefsFontSizeLowDPI of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsPrefsFontSizeLowDPI
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsPrefsFontSizeLowDPI(ByVal sgsPrefsFontSizeLowDPI As String)

    On Error GoTo gsPrefsFontSizeLowDPI_Error

    m_sgsPrefsFontSizeLowDPI = sgsPrefsFontSizeLowDPI

    On Error GoTo 0
    Exit Property

gsPrefsFontSizeLowDPI_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsPrefsFontSizeLowDPI of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsPrefsFontItalics
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsPrefsFontItalics() As String

    On Error GoTo gsPrefsFontItalics_Error

    gsPrefsFontItalics = m_sgsPrefsFontItalics

    On Error GoTo 0
    Exit Property

gsPrefsFontItalics_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsPrefsFontItalics of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsPrefsFontItalics
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsPrefsFontItalics(ByVal sgsPrefsFontItalics As String)

    On Error GoTo gsPrefsFontItalics_Error

    m_sgsPrefsFontItalics = sgsPrefsFontItalics

    On Error GoTo 0
    Exit Property

gsPrefsFontItalics_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsPrefsFontItalics of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsPrefsFontColour
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsPrefsFontColour() As String

    On Error GoTo gsPrefsFontColour_Error

    gsPrefsFontColour = m_sgsPrefsFontColour

    On Error GoTo 0
    Exit Property

gsPrefsFontColour_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsPrefsFontColour of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsPrefsFontColour
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsPrefsFontColour(ByVal sgsPrefsFontColour As String)

    On Error GoTo gsPrefsFontColour_Error

    m_sgsPrefsFontColour = sgsPrefsFontColour

    On Error GoTo 0
    Exit Property

gsPrefsFontColour_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsPrefsFontColour of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsDisplayScreenFont
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsDisplayScreenFont() As String

    On Error GoTo gsDisplayScreenFont_Error

    gsDisplayScreenFont = m_sgsDisplayScreenFont

    On Error GoTo 0
    Exit Property

gsDisplayScreenFont_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsDisplayScreenFont of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsDisplayScreenFont
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsDisplayScreenFont(ByVal sgsDisplayScreenFont As String)

    On Error GoTo gsDisplayScreenFont_Error

    m_sgsDisplayScreenFont = sgsDisplayScreenFont

    On Error GoTo 0
    Exit Property

gsDisplayScreenFont_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsDisplayScreenFont of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsDisplayScreenFontSize
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsDisplayScreenFontSize() As String

    On Error GoTo gsDisplayScreenFontSize_Error

    gsDisplayScreenFontSize = m_sgsDisplayScreenFontSize

    On Error GoTo 0
    Exit Property

gsDisplayScreenFontSize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsDisplayScreenFontSize of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsDisplayScreenFontSize
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsDisplayScreenFontSize(ByVal sgsDisplayScreenFontSize As String)

    On Error GoTo gsDisplayScreenFontSize_Error

    m_sgsDisplayScreenFontSize = sgsDisplayScreenFontSize

    On Error GoTo 0
    Exit Property

gsDisplayScreenFontSize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsDisplayScreenFontSize of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsDisplayScreenFontItalics
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsDisplayScreenFontItalics() As String

    On Error GoTo gsDisplayScreenFontItalics_Error

    gsDisplayScreenFontItalics = m_sgsDisplayScreenFontItalics

    On Error GoTo 0
    Exit Property

gsDisplayScreenFontItalics_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsDisplayScreenFontItalics of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsDisplayScreenFontItalics
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsDisplayScreenFontItalics(ByVal sgsDisplayScreenFontItalics As String)

    On Error GoTo gsDisplayScreenFontItalics_Error

    m_sgsDisplayScreenFontItalics = sgsDisplayScreenFontItalics

    On Error GoTo 0
    Exit Property

gsDisplayScreenFontItalics_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsDisplayScreenFontItalics of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsDisplayScreenFontColour
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get gsDisplayScreenFontColour() As String

    On Error GoTo gsDisplayScreenFontColour_Error

    gsDisplayScreenFontColour = m_sgsDisplayScreenFontColour

    On Error GoTo 0
    Exit Property

gsDisplayScreenFontColour_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsDisplayScreenFontColour of Module Module1"

End Property

'---------------------------------------------------------------------------------------
' Procedure : gsDisplayScreenFontColour
' Author    : beededea
' Date      : 08/10/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let gsDisplayScreenFontColour(ByVal sgsDisplayScreenFontColour As String)

    On Error GoTo gsDisplayScreenFontColour_Error

    m_sgsDisplayScreenFontColour = sgsDisplayScreenFontColour

    On Error GoTo 0
    Exit Property

gsDisplayScreenFontColour_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gsDisplayScreenFontColour of Module Module1"

End Property
