VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule IntegerDataType, ModuleWithoutFolder
'---------------------------------------------------------------------------------------
' Module    : cfMain
' Author    : beededea
' Date      : 28/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public WithEvents aboutForm As cWidgetForm
Attribute aboutForm.VB_VarHelpID = -1
Public WithEvents helpForm As cWidgetForm
Attribute helpForm.VB_VarHelpID = -1
Public WithEvents licenceForm As cWidgetForm
Attribute licenceForm.VB_VarHelpID = -1
Public WithEvents TenShillingsForm As cWidgetForm
Attribute TenShillingsForm.VB_VarHelpID = -1

Public FX As Long   ' available externally, to allow the gaugeForm to be moved, need to add set/let for these public vars
Public FY As Long
Public FZ As Single

Public licenceFormX As Long
Public licenceFormY As Long
' Public licenceFormZ As Single

Private WithEvents btnAccept As cwButton
Attribute btnAccept.VB_VarHelpID = -1
Private WithEvents btnDecline As cwButton
Attribute btnDecline.VB_VarHelpID = -1
Private WithEvents textBox As cwTextBox
Attribute textBox.VB_VarHelpID = -1
Private lblDeclineOrNot As cwLabel
 
 
 
'---------------------------------------------------------------------------------------
' Procedure : initAndCreateWidgetForm
' Author    :
' Date      : 27/04/2023
' Purpose   : Initialise and create the tenShillings form only, add the tenShillings widget based on the tenShillings class, to the form
'---------------------------------------------------------------------------------------
'
Public Sub initAndCreateTenShillingsForm(ByVal x As Long, ByVal y As Long, Optional ByVal dx As Long, Optional ByVal dy As Long, Optional ByVal Caption As String)
    Dim thisSrf As cCairoSurface
    Dim imageHeight As Long: imageHeight = 0
    Dim imageWidth As Long: imageWidth = 0
    Dim tenShillingsFormWidth As Long: tenShillingsFormWidth = 0
    Dim tenShillingsFormHeight As Long: tenShillingsFormHeight = 0
    Dim tenShillingsWidgetWidth As Long: tenShillingsWidgetWidth = 0
    Dim tenShillingsWidgetHeight As Long: tenShillingsWidgetHeight = 0
    Dim tenShillingsWidgetZoomValue As Double: tenShillingsWidgetZoomValue = 1
    
    Const WFFStyleTool As Integer = 7 ' additional styles above the normal five  WFFStyleTool = 7 makes the taskbar button go away
    
    On Error GoTo initAndCreateTenShillingsForm_Error
    
    If Cairo.ImageList.Exists("tenshillings") Then Set thisSrf = Cairo.ImageList("tenshillings")
    
    imageWidth = thisSrf.Width
    imageHeight = thisSrf.Height
    
    If gblDpiAwareness = "0" Then
        tenShillingsFormWidth = imageWidth * 2.2 ' roughly twice the width of the main image plus a bit
        tenShillingsFormHeight = imageHeight * 2.2 '
        If tenShillingsFormWidth > gblVirtualScreenWidthPixels Then tenShillingsFormWidth = gblVirtualScreenWidthPixels
        If tenShillingsFormHeight > gblVirtualScreenHeightPixels Then tenShillingsFormHeight = gblVirtualScreenHeightPixels
    Else
        tenShillingsFormWidth = imageWidth * 2.2 ' roughly twice the width of the main image plus a bit
        tenShillingsFormHeight = imageHeight * 2.2 '
        If tenShillingsFormWidth > gblVirtualScreenWidthPixels Then
            tenShillingsFormWidth = gblVirtualScreenWidthPixels
            tenShillingsWidgetZoomValue = 1
        Else
            tenShillingsWidgetZoomValue = 1
        End If
        If tenShillingsFormHeight > gblVirtualScreenHeightPixels Then
            tenShillingsFormHeight = gblVirtualScreenHeightPixels
            tenShillingsWidgetZoomValue = 1
        Else
            tenShillingsWidgetZoomValue = 1
        End If
    End If
    tenShillingsWidgetWidth = tenShillingsFormWidth - 100
    tenShillingsWidgetHeight = tenShillingsFormHeight - 100
    
    ' create an invisible form using a Cairo Widget Form with the predefined caption and sizing
    Set TenShillingsForm = Cairo.WidgetForms.Create(WFFStyleTool, "TenShillings the " & Caption & " Widget", True, tenShillingsFormWidth, tenShillingsFormHeight)
    
    ' make the tenShillings form transparent
    TenShillingsForm.WidgetRoot.BackColor = -1
    
    ' create a new tenShillings widget with a name, position and size, first step of that is to run Class_Initialize
    Set TenShillingsWidget = TenShillingsForm.Widgets.Add(New cwTenShillings, "tenShillings", 1, 1, tenShillingsWidgetWidth, tenShillingsWidgetHeight)
    
    If gblDpiAwareness = "0" Then
        TenShillingsWidget.Zoom = tenShillingsWidgetZoomValue
    Else
        TenShillingsWidget.Zoom = tenShillingsWidgetZoomValue
    End If
    
   On Error GoTo 0
   Exit Sub

initAndCreateTenShillingsForm_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initAndCreateTenShillingsForm of Class Module cfMain"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : initAndCreateAboutForm
' Author    :
' Date      : 27/04/2023
' Purpose   : Initialise and create the about form only, add the about widget based on the about class, to the form
'---------------------------------------------------------------------------------------
'
Public Sub initAndCreateAboutForm(ByVal x As Long, ByVal y As Long, Optional ByVal dx As Long, Optional ByVal dy As Long, Optional ByVal Caption As String)
    
    Dim aboutFormWidth As Long: aboutFormWidth = 0
    Dim aboutFormHeight As Long: aboutFormHeight = 0
    Dim aboutWidgetWidth As Long: aboutWidgetWidth = 0
    Dim aboutWidgetHeight As Long: aboutWidgetHeight = 0
    Dim aboutWidgetZoomValue As Double: aboutWidgetZoomValue = 1
    
    Const WFFStyleTool As Integer = 7 ' additional styles above the normal five  WFFStyleTool = 7 makes the taskbar button go away
    
    On Error GoTo initAndCreateAboutForm_Error
    
    If gblDpiAwareness = "0" Then
        aboutFormWidth = 900
        aboutFormHeight = 850
        If aboutFormWidth > gblVirtualScreenWidthPixels Then aboutFormWidth = gblVirtualScreenWidthPixels
        If aboutFormHeight > gblVirtualScreenHeightPixels Then aboutFormHeight = gblVirtualScreenHeightPixels
    Else
        aboutFormWidth = 1800
        aboutFormHeight = 1600
        If aboutFormWidth > gblVirtualScreenWidthPixels Then
            aboutFormWidth = gblVirtualScreenWidthPixels
            aboutWidgetZoomValue = 1
        Else
            aboutWidgetZoomValue = 1.6
        End If
        If aboutFormHeight > gblVirtualScreenHeightPixels Then
            aboutFormHeight = gblVirtualScreenHeightPixels
            aboutWidgetZoomValue = 1
        Else
            aboutWidgetZoomValue = 1.6
        End If
    End If
    aboutWidgetWidth = aboutFormWidth - 100
    aboutWidgetHeight = aboutFormHeight - 100
    
    ' create an invisible form using a Cairo Widget Form with the predefined caption and sizing
    Set aboutForm = Cairo.WidgetForms.Create(WFFStyleTool, "About the " & Caption & " Widget", True, aboutFormWidth, aboutFormHeight)
    
    ' make the about form transparent
    aboutForm.WidgetRoot.BackColor = -1
    
    ' create a new about widget with a name, position and size, first step of that is to run Class_Initialize
    Set aboutWidget = aboutForm.Widgets.Add(New cwAbout, "about", 1, 1, aboutWidgetWidth, aboutWidgetHeight)
    
    If gblDpiAwareness = "0" Then
        aboutWidget.Zoom = aboutWidgetZoomValue
    Else
        aboutWidget.Zoom = aboutWidgetZoomValue
    End If
    
   On Error GoTo 0
   Exit Sub

initAndCreateAboutForm_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initAndCreateAboutForm of Class Module cfMain"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : initAndCreateHelpForm
' Author    :
' Date      : 27/04/2023
' Purpose   : Initialise and create the Help form only, add the Help widget based on the Help class, to the form
'---------------------------------------------------------------------------------------
'
Public Sub initAndCreateHelpForm(ByVal x As Long, ByVal y As Long, Optional ByVal dx As Long, Optional ByVal dy As Long, Optional ByVal Caption As String)
    
    Dim helpFormWidth As Long: helpFormWidth = 0
    Dim helpFormHeight As Long: helpFormHeight = 0
    Dim helpWidgetWidth As Long: helpWidgetWidth = 0
    Dim helpWidgetHeight As Long: helpWidgetHeight = 0
    Dim helpWidgetZoomValue As Double: helpWidgetZoomValue = 1
    
    Const WFFStyleTool As Integer = 7 ' additional styles above the normal five  WFFStyleTool = 7 makes the taskbar button go away
    
    On Error GoTo initAndCreateHelpForm_Error
    
    If gblDpiAwareness = "0" Then
        helpFormWidth = 1000
        helpFormHeight = 800
        If helpFormWidth > gblVirtualScreenWidthPixels Then helpFormWidth = gblVirtualScreenWidthPixels
        If helpFormHeight > gblVirtualScreenHeightPixels Then helpFormHeight = gblVirtualScreenHeightPixels
        helpWidgetZoomValue = 0.8
    Else
        helpFormWidth = 1600
        helpFormHeight = 1300
        If helpFormWidth > gblVirtualScreenWidthPixels Then
            helpFormWidth = gblVirtualScreenWidthPixels
            helpWidgetZoomValue = 1
        Else
            helpWidgetZoomValue = 1.4
        End If
        If helpFormHeight > gblVirtualScreenHeightPixels Then
            helpFormHeight = gblVirtualScreenHeightPixels
            helpWidgetZoomValue = 1
        Else
            helpWidgetZoomValue = 1.4
        End If
    End If
    helpWidgetWidth = helpFormWidth - 100
    helpWidgetHeight = helpFormHeight - 100
    
    ' create an invisible form using a Cairo Widget Form with the predefined caption and sizing
    Set helpForm = Cairo.WidgetForms.Create(WFFStyleTool, "Help the " & Caption & " Widget", True, helpFormWidth, helpFormHeight)
    
    ' make the help form transparent
    helpForm.WidgetRoot.BackColor = -1
    
    ' create a new help widget with a name, position and size, first step of that is to run Class_Initialize
    Set helpWidget = helpForm.Widgets.Add(New cwHelp, "help", 1, 1, helpWidgetWidth, helpWidgetHeight)
    
    If gblDpiAwareness = "0" Then
        helpWidget.Zoom = helpWidgetZoomValue
    Else
        helpWidget.Zoom = helpWidgetZoomValue
    End If
    
   On Error GoTo 0
   Exit Sub

initAndCreateHelpForm_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initAndCreateHelpForm of Class Module cfMain"
    
End Sub



'---------------------------------------------------------------------------------------
' Procedure : initAndCreateLicenceForm
' Author    :
' Date      : 27/04/2023
' Purpose   : Initialise and create the licence form only, add the licence widget based on the licence class, to the form
'             as well as two RC cwButton widgets for the accept and decline buttons
'---------------------------------------------------------------------------------------
'
Public Sub initAndCreateLicenceForm(ByVal x As Long, ByVal y As Long, Optional ByVal dx As Long, Optional ByVal dy As Long, Optional ByVal Caption As String)
    
    Dim licenceFormWidth As Long: licenceFormWidth = 0
    Dim licenceFormHeight As Long: licenceFormHeight = 0
    Dim licenceWidgetWidth As Long: licenceWidgetWidth = 0
    Dim licenceWidgetHeight As Long: licenceWidgetHeight = 0
    
    Dim licenceWidgetZoomValue As Double: licenceWidgetZoomValue = 1
    
    Const WFFStyleTool As Integer = 7 ' additional styles above the normal five  WFFStyleTool = 7 makes the taskbar button go away
    
    On Error GoTo initAndCreateLicenceForm_Error
    
    If gblDpiAwareness = "0" Then
        licenceFormWidth = 650
        licenceFormHeight = 550
        If licenceFormWidth > gblVirtualScreenWidthPixels Then licenceFormWidth = gblVirtualScreenWidthPixels
        If licenceFormHeight > gblVirtualScreenHeightPixels Then licenceFormHeight = gblVirtualScreenHeightPixels
        licenceWidgetZoomValue = 0.8
    Else
        licenceFormWidth = 750
        licenceFormHeight = 650
        If licenceFormWidth > gblVirtualScreenWidthPixels Then
            licenceFormWidth = gblVirtualScreenWidthPixels
            licenceWidgetZoomValue = 0.8
        Else
            licenceWidgetZoomValue = 1.5
        End If
        
        If licenceFormHeight > gblVirtualScreenHeightPixels Then
            licenceFormHeight = gblVirtualScreenHeightPixels
            licenceWidgetZoomValue = 0.8
        Else
            licenceWidgetZoomValue = 1.5
        End If
        
    End If
    licenceWidgetWidth = licenceFormWidth - 100
    licenceWidgetHeight = licenceFormHeight - 100
           
    ' create an invisible form using a Cairo Widget Form with the predefined caption and sizing
    Set licenceForm = Cairo.WidgetForms.Create(WFFStyleTool, Caption, True, licenceFormWidth, licenceFormHeight)
    
    ' make the licence form transparent
    licenceForm.WidgetRoot.BackColor = -1
         
    ' create a new licence class widget with a name, position and size
    Set licenceWidget = licenceForm.Widgets.Add(New cwLicence, "licence", 0, 0, licenceWidgetWidth, licenceWidgetHeight)
    
    If gblDpiAwareness = "0" Then
        licenceWidget.Zoom = licenceWidgetZoomValue
    Else
        licenceWidget.Zoom = licenceWidgetZoomValue
    End If

    ' When running in the IDE the RC6 DLLs are required to be registered. If for some reason they become unregistered then
    ' the following 'set' commands will generate an error, to handle this we have a temporary error trap for the IDE only.
    
    If InIDE Then On Error GoTo RC6NotRegistered_Error

    ' add two RC predefined button controls to the licence form
    If gblDpiAwareness = "1" Then
        Set btnDecline = licenceForm.Widgets.Add(New_W("cwButton"), "btnDecline", 352, 397, 70, 29)
        Set btnAccept = licenceForm.Widgets.Add(New_W("cwButton"), "btnAccept", 428, 397, 70, 29)
        Set textBox = licenceForm.Widgets.Add(New_W("cwTextBox"), "textBox", 50, 60, 440, 280)
        Set lblDeclineOrNot = licenceForm.Widgets.Add(New_W("cwLabel"), "lblDeclineOrNot", 50, 345, 440, 35)
    Else
        Set btnDecline = licenceForm.Widgets.Add(New_W("cwButton"), "btnDecline", 189, 210, 40, 18)
        Set btnAccept = licenceForm.Widgets.Add(New_W("cwButton"), "btnAccept", 228, 210, 40, 18)
        Set textBox = licenceForm.Widgets.Add(New_W("cwTextBox"), "textBox", 25, 30, 240, 155)
        Set lblDeclineOrNot = licenceForm.Widgets.Add(New_W("cwLabel"), "lblDeclineOrNot", 25, 188, 240, 20)
    End If
    
    On Error GoTo initAndCreateLicenceForm_Error
    
    If gblDpiAwareness = "1" Then
        licenceForm.Widgets("TextBox").Widget.FontSize = 11
        licenceForm.Widgets("btnDecline").Widget.FontSize = 10
        licenceForm.Widgets("btnAccept").Widget.FontSize = 10
        licenceForm.Widgets("lblDeclineOrNot").Widget.FontSize = 8
    Else
        licenceForm.Widgets("TextBox").Widget.FontSize = 8
        licenceForm.Widgets("btnDecline").Widget.FontSize = 7
        licenceForm.Widgets("btnAccept").Widget.FontSize = 7
        licenceForm.Widgets("lblDeclineOrNot").Widget.FontSize = 6
    End If
    
    licenceForm.Widgets("btnDecline").Caption = "Decline"
    licenceForm.Widgets("btnDecline").Widget.ForeColor = vbWhite
    licenceForm.Widgets("btnDecline").Widget.Alpha = 0.05
    licenceForm.Widgets("btnDecline").Widget.ToolTip = "Press this button to decline the terms of the GPL licence."
        
    licenceForm.Widgets("btnAccept").Caption = "Accept"
    licenceForm.Widgets("btnAccept").Widget.ForeColor = vbWhite
    licenceForm.Widgets("btnAccept").Widget.Alpha = 0.05
    licenceForm.Widgets("btnAccept").Widget.ToolTip = "Press this button to accept the terms of the GPL licence."
    
    licenceForm.Widgets("TextBox").MultiLine = True
    licenceForm.Widgets("TextBox").ScrollBars = vbVertical
    licenceForm.Widgets("TextBox").Widget.BackColor = &H4000&
    licenceForm.Widgets("TextBox").Widget.ForeColor = vbWhite
    
    Call LoadFileToTB(textBox, App.Path & "\resources\txt\licence.txt", False)

    licenceForm.Widgets("TextBox").Text = gblWidgetName & vbCrLf & vbCrLf & _
    "Copyright © 1992-2023 Dean Beedell" & vbCrLf & vbCrLf & _
    "This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public Licence as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version." & vbCrLf & vbCrLf & _
        "This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details." & vbCrLf & vbCrLf & _
        "You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA" & vbCrLf & vbCrLf & _
        "If you use this software in any way whatsoever then that implies acceptance of the licence. If you do not wish to comply with the licence terms then please remove the download, binary and source code from your systems immediately." & vbCrLf & vbCrLf & _
        "You may use any of my own imagery in your own creations but commercially only with my permission. In all other non-commercial cases I require a credit to the original artist using my name or one of my pseudonyms and a link to my site. With regard to the commercial use of incorporated images, permission and a licence would need to be obtained from the original owner and creator, ie. me."

    licenceForm.Widgets("lblDeclineOrNot").BorderWidth = 0
    licenceForm.Widgets("lblDeclineOrNot").WordWrap = True
    licenceForm.Widgets("lblDeclineOrNot").Alignment = vbCenter
    licenceForm.Widgets("lblDeclineOrNot").Widget.BackColor = -1
    licenceForm.Widgets("lblDeclineOrNot").Widget.ForeColor = vbWhite
    
    licenceForm.Widgets("lblDeclineOrNot").Caption = "If you do not agree with the terms set forth above, please click the " & _
        "Decline button below and destroy this widget and its documentation"
        
   On Error GoTo 0

   Exit Sub

initAndCreateLicenceForm_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initAndCreateLicenceForm_Error of Class Module cfMain"
    
    Exit Sub
   
RC6NotRegistered_Error:
    
    MsgBox "Error in procedure initAndCreateLicenceForm_Error of Class Module cfMain. vbRichClient5.dll is NOT registered, use the registration script to register vbRichClient5.dll."

End Sub
'---------------------------------------------------------------------------------------
' Procedure : btnAccept_Click
' Author    : beededea
' Date      : 20/09/2023
' Purpose   : Click event for this RC cwButton widget
'---------------------------------------------------------------------------------------
'
Private Sub btnAccept_Click()
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim fileToPlay As String: fileToPlay = vbNullString
    Dim slicence As String: slicence = "0"
    
    On Error GoTo btnAccept_Click_Error
    
    fileToPlay = "ting.wav"
    
    If gblEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        playSound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If

    answer = msgBoxA("User accepted licence, thankyou.", vbOKOnly + vbExclamation, "Licence Acceptance", True, "btnAcceptClick")
    
    slicence = "1"
    sPutINISetting "Software\TenShillings", "licence", slicence, gblSettingsFile
    
    licenceWidget.ShowMe = False
    licenceWidget.MyOpacity = 0
    licenceWidget.Widget.Refresh

    licenceForm.Visible = False
    
   On Error GoTo 0
   Exit Sub

btnAccept_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnAccept_Click of Class Module cfMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnDecline_Click
' Author    : beededea
' Date      : 20/09/2023
' Purpose   : Click event for this RC cwButton widget
'---------------------------------------------------------------------------------------
'
Private Sub btnDecline_Click()
    
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim slicence As String: slicence = "0"
        
    On Error GoTo btnDecline_Click_Error

    'answer = MsgBox("Licence declined, please destroy your copy of this widget and its documentation.", vbYes)
    answer = msgBoxA("Licence declined, please destroy your copy of this widget and its documentation. This widget will now unload.", vbOKOnly + vbExclamation, "Licence Declined", False)
    slicence = "0"
    sPutINISetting "Software\TenShillings", "licence", slicence, gblSettingsFile

    Call thisForm_Unload

   On Error GoTo 0
   Exit Sub

btnDecline_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnDecline_Click of Class Module cfMain"
End Sub







'---------------------------------------------------------------------------------------
' Procedure : TenShillingsForm_DblClick
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub TenShillingsForm_DblClick()
    Dim userprof As String: userprof = vbNullString
    Dim thisCommand As String: thisCommand = vbNullString
    
    On Error GoTo TenShillingsForm_DblClick_Error

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
        Call ShellExecute(fMain.TenShillingsForm.hWnd, "Open", gblOpenFile, vbNullString, App.Path, 1)
    Else
        Call ShellExecute(fMain.TenShillingsForm.hWnd, "runas", thisCommand, vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

TenShillingsForm_DblClick_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TenShillingsForm_DblClick of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub

 

'---------------------------------------------------------------------------------------
' Procedure : TenShillingsForm_MouseMove
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub TenShillingsForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Static x0 As Single, y0 As Single: If Button = 0 Then x0 = x: y0 = y 'just store the offsets when no button is down
   
    On Error GoTo TenShillingsForm_MouseMove_Error
    
    If Button = vbLeftButton And Not TenShillingsForm.ActiveWidget Is Nothing Then 'Form-Dragging (under certain conditions)
    
        If TenShillingsWidget.Locked = True Then Exit Sub
        If gblIgnoreMouse = "1" Then Exit Sub
        
        If gblMenuOccurred = True Then
            gblMenuOccurred = False
            Exit Sub
        End If

        FX = (x - x0) * TenShillingsForm.WidgetRoot.Zoom + TenShillingsForm.Left
        FY = (y - y0) * TenShillingsForm.WidgetRoot.Zoom + TenShillingsForm.Top
        TenShillingsForm.Move FX, FY
        
'        Debug.Print "FX " & FX
'        Debug.Print "FY " & FY

    End If

    On Error GoTo 0
    Exit Sub

TenShillingsForm_MouseMove_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TenShillingsForm_MouseMove of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure : TenShillingsForm_KeyDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : get F5 and SHIFT keypresses
'---------------------------------------------------------------------------------------
'
Private Sub TenShillingsForm_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
    On Error GoTo TenShillingsForm_KeyDown_Error

    Call getKeyPress(KeyCode, Shift)

    On Error GoTo 0
    Exit Sub

TenShillingsForm_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TenShillingsForm_KeyDown of Class Module cfMain"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : TenShillingsForm_MouseUp
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : save the form x,y position when ever the TenShillings/form is dragged
'---------------------------------------------------------------------------------------
'
Private Sub TenShillingsForm_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo TenShillingsForm_MouseUp_Error
    
    frmTimer.tmrScreenResolution.Interval = 4500

    ' save the form x,y position whenever the globe/form is dragged
    Call saveMainRCFormPosition

    Call positionRCFormByMonitorSize

    On Error GoTo 0
    Exit Sub

TenShillingsForm_MouseUp_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TenShillingsForm_MouseUp of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub
