VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImageList32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cImageList32
' Author    : The vast majority from beededea but the basic class boilerplate from chatGPT
' Date      : 04/01/2026
' Purpose   : replacement of the RichClient imageList with the scripting dictionary to store PNGs
'---------------------------------------------------------------------------------------

Option Explicit

Private mDict As Object ' Scripting.Dictionary or Cristian Buse's dictionary class

' ===========================
' Win32 / GDI
' ===========================
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObj As Long) As Long
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long


' ===========================
' GDI+
' ===========================
Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" _
    (ByVal FileName As Long, ByRef Bitmap As Long) As Long

Private Declare Function GdipDisposeImage Lib "gdiplus" _
    (ByVal Image As Long) As Long

Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" _
    (ByVal Bitmap As Long, ByRef hbmReturn As Long, ByVal background As Long) As Long
    
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlob&, ByVal fDeleteOnRelease As Long, ppstm As stdole.IUnknown) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal pStream As Long, Image As Long) As Long
Private Declare Function GdipGetImageWidth Lib "GdiPlus.dll" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "GdiPlus.dll" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal dx As Long, ByVal dy As Long, ByVal stride As Long, ByVal PixelFormat As Long, ByVal pScanData As Long, Image As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal img As Long, Context As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal Context As Long, ByVal PixOffsetMode As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "GdiPlus.dll" (ByVal Graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal Graphics As Long, ByVal SmoothingMode As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef imageattr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal enableFlag As Long, colourMatrix As Any, grayMatrix As Any, ByVal Flags As ColorMatrixFlags) As GpStatus
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal Context As Long, ByVal Image As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal Callback As Long, ByVal CallbackData As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal Graphics As Long) As Long
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, Size As Long) As GpStatus
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, encoders As Any) As GpStatus

' APIs image cropping
Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal Image As Long, ByRef PixelFormat As Long) As Long
Private Declare Function GdipCloneBitmapAreaI Lib "gdiplus" (ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal PixelFormat As Long, ByVal srcBitmap As Long, dstBitmap As Long) As GpStatus

' Windows+ constants END

Private Const PixelFormat32bppPARGB = &HE200B
Private Const PixelFormat32bppARGB = &H26200A

' global GDI+ Types START
Private Type BITMAPINFOHEADER
    Size As Long
    Width As Long
    Height As Long
    Planes As Integer
    BitCount As Integer
    Compression As Long
    SizeImage As Long
    XPelsPerMeter As Long
    YPelsPerMeter As Long
    ClrUsed As Long
    ClrImportant As Long
End Type

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Type GDIPLUS_STARTINPUT
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECTF
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmpHeader As BITMAPINFOHEADER
    bmpColors As RGBQUAD
End Type

Private Type CLSID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type

Private Type ImageCodecInfo
   ClassID As CLSID
   FormatID As CLSID
   CodecName As Long      ' String Pointer; const WCHAR*
   DllName As Long        ' String Pointer; const WCHAR*
   FormatDescription As Long ' String Pointer; const WCHAR*
   FilenameExtension As Long ' String Pointer; const WCHAR*
   MimeType As Long       ' String Pointer; const WCHAR*
   Flags As ImageCodecFlags   ' Should be a Long equivalent
   Version As Long
   SigCount As Long
   SigSize As Long
   SigPattern As Long      ' Byte Array Pointer; BYTE*
   SigMask As Long         ' Byte Array Pointer; BYTE*
End Type
' global GDI+ Types END

' .07 DAEB 19/04/2021 mdlMain.bas  added a new type link for determining shortcuts
Private Type Link
    Attributes As Long
    FileName As String
    Description As String
    RelPath As String
    WorkingDir As String
    Arguments As String
    CustomIcon As String
End Type


Private Type PICTDESC
    cbSizeOfStruct  As Long
    PicType         As Long
    hImage          As Long
    xExt            As Long
    yExt            As Long
End Type


Private Type GpMatrix
    matrix(5) As Single
End Type
 
Private Type ColorMatrix
    m(0 To 4, 0 To 4) As Single
End Type
' vars for GDI+ colour matrix ENDS


' global GDI+ Enums START
Private Enum GDIPLUS_ALIGNMENT
   StringAlignmentNear = 0
   StringAlignmentCenter = 1
   StringAlignmentFar = 2
End Enum

Private Enum GDIPLUS_FONTSTYLE
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum

Private Enum GDIPLUS_UNIT
    UnitWorld
    UnitDisplay
    UnitPixel
    UnitPoint
    UnitInch
    UnitDocument
    UnitMillimeter
End Enum

'Private Enum TASKBAR_POSITION
'    vbbottom
'    vbLeft
'    vbRight
'    vbtop
'End Enum

    
Private Enum TASKBAR_POSITION
    vbtop
    vbBottom
    vbLeft
    vbRight
End Enum
    

' NOTE: Enums evaluate to a Long
Private Enum GpStatus   ' aka Status
   Ok = 0
   GenericError = 1
   InvalidParameter = 2
   OutOfMemory = 3
   ObjectBusy = 4
   InsufficientBuffer = 5
   NotImplemented = 6
   Win32Error = 7
   WrongState = 8
   Aborted = 9
   FileNotFound = 10
   ValueOverflow = 11
   AccessDenied = 12
   UnknownImageFormat = 13
   FontFamilyNotFound = 14
   FontStyleNotFound = 15
   NotTrueTypeFont = 16
   UnsupportedGdiplusVersion = 17
   GdiplusNotInitialized = 18
   PropertyNotFound = 19
   PropertyNotSupported = 20
   ProfileNotFound = 21
End Enum

' Information flags about image codecs
Private Enum ImageCodecFlags
   ImageCodecFlagsEncoder = &H1
   ImageCodecFlagsDecoder = &H2
   ImageCodecFlagsSupportBitmap = &H4
   ImageCodecFlagsSupportVector = &H8
   ImageCodecFlagsSeekableEncode = &H10
   ImageCodecFlagsBlockingDecode = &H20

   ImageCodecFlagsBuiltin = &H10000
   ImageCodecFlagsSystem = &H20000
   ImageCodecFlagsUser = &H40000
End Enum

Private Enum eInterpolationMode
  ipmDefault = &H0
  ipmLow = &H1
  ipmHigh = &H2
  ipmBilinear = &H3
  ipmBicubic = &H4
  ipmNearestNeighbor = &H5
  ipmHighQualityBilinear = &H6
  ipmHighQualityBicubic = &H7
End Enum

'Private Enum WrapMode
'   WrapModeTile = &H8
'   WrapModeTileFlipX = &H9
'   WrapModeTileFlipY = &H10
'   WrapModeTileFlipXY = &H11
'   WrapModeClamp = &H12
'End Enum

Private Enum SmoothingModeEnum
    SmoothingModeDefault = 0&
    SmoothingModeHighSpeed = 1&
    SmoothingModeHighQuality = 2&
    SmoothingModeNone = 3&
    SmoothingModeAntiAlias8x4 = 4&
    SmoothingModeAntiAlias = 4&
    SmoothingModeAntiAlias8x8 = 5&
End Enum

' vars for GDI+ colour matrix STARTS
Private Enum ColorAdjustType
   ColorAdjustTypeDefault
   ColorAdjustTypeBitmap
   ColorAdjustTypeBrush
   ColorAdjustTypePen
   ColorAdjustTypeText
   ColorAdjustTypeCount
   ColorAdjustTypeAny
End Enum
 
Private Enum ColorMatrixFlags
   ColorMatrixFlagsDefault = 0
   ColorMatrixFlagsSkipGrays = 1
   ColorMatrixFlagsAltGray = 2
End Enum



'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : beededea
' Date      : 04/01/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    On Error GoTo Class_Initialize_Error

    Set mDict = CreateObject("Scripting.Dictionary")

    On Error GoTo 0
    Exit Sub

Class_Initialize_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cImageList32"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : beededea
' Date      : 04/01/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    On Error GoTo Class_Terminate_Error

    Clear
    Set mDict = Nothing

    On Error GoTo 0
    Exit Sub

Class_Terminate_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Terminate of Class Module cImageList32"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : beededea
' Date      : 04/01/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Count() As Long
    On Error GoTo Count_Error

    Count = mDict.Count

    On Error GoTo 0
    Exit Property

Count_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Count of Class Module cImageList32"
End Property

'---------------------------------------------------------------------------------------
' Procedure : Exists
' Author    : beededea
' Date      : 04/01/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Exists(ByVal Key As String) As Boolean
    On Error GoTo Exists_Error

    Exists = mDict.Exists(Key)

    On Error GoTo 0
    Exit Function

Exists_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Exists of Class Module cImageList32"
End Function

'---------------------------------------------------------------------------------------
' Procedure : Remove
' Author    : beededea
' Date      : 04/01/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Remove(ByVal Key As String)
    On Error GoTo Remove_Error

    If mDict.Exists(Key) Then
        DeleteObject mDict(Key)
        mDict.Remove Key
    End If

    On Error GoTo 0
    Exit Sub

Remove_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Remove of Class Module cImageList32"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Clear
' Author    : beededea
' Date      : 04/01/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub Clear()
    Dim k As Variant
    On Error GoTo Clear_Error

    For Each k In mDict.Keys
        DeleteObject mDict(k)
    Next
    mDict.RemoveAll

    On Error GoTo 0
    Exit Sub

Clear_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Clear of Class Module cImageList32"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetHandle
' Author    : beededea
' Date      : 04/01/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetHandle(ByVal Key As String) As Long
    On Error GoTo GetHandle_Error

    If Not mDict.Exists(Key) Then Err.Raise 5, , "Image key not found"
    GetHandle = mDict(Key)

    On Error GoTo 0
    Exit Function

GetHandle_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetHandle of Class Module cImageList32"
End Function

'---------------------------------------------------------------------------------------
' Procedure : AddPNG
' Author    : beededea
' Date      : 04/01/2026
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub AddPNG(ByVal Key As String, ByVal FilePath As String)
'    Dim gpBitmap As Long
'    Dim hBmp As Long
    Dim imageOpacity As Long: imageOpacity = 0

    On Error GoTo AddPNG_Error
    
    imageOpacity = 100

    If mDict.Exists(Key) Then Err.Raise 457, , "Key already exists"
    
    ' App.Path & "\busy-F1-32x32x24.png"
    If fFExists(FilePath) Then
        resizeAndLoadImgToDict mDict, Key, FilePath, 0, 0, (128), (128), , imageOpacity
    End If

'    If GdipCreateBitmapFromFile(StrPtr(FilePath), gpBitmap) <> 0 Then
'        Err.Raise 53, , "Failed to load PNG"
'    End If
'
'    ' background = 0 ? preserves alpha
'    If GdipCreateHBITMAPFromBitmap(gpBitmap, hBmp, 0) <> 0 Then
'        GdipDisposeImage gpBitmap
'        Err.Raise 7, , "Failed to create HBITMAP"
'    End If

'    GdipDisposeImage gpBitmap

'    mDict.Add Key, hBmp

    On Error GoTo 0
    Exit Sub

AddPNG_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AddPNG of Class Module cImageList32"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ReadBytesFromFile
' Author    : beededea
' Date      : 07/04/2020
' Purpose   : Credit to Olaf Schmidt
'---------------------------------------------------------------------------------------
'
Public Function ReadBytesFromFile(ByVal FileName As String) As Byte()
   On Error GoTo ReadBytesFromFile_Error

    Dim ab As Object
    
'  With CreateObject("ADODB.Stream")
'    .Open
'      .Type = 1 'adTypeBinary
'      .LoadFromFile filename
'      ReadBytesFromFile = .Read
'    .Close
'  End With

  Set ab = CreateObject("ADODB.Stream")
  
  With ab
    .Open
      .Type = 1 'adTypeBinary
      .LoadFromFile FileName
      ReadBytesFromFile = .Read
    .Close
  End With
  
   On Error GoTo 0
   Exit Function

ReadBytesFromFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ReadBytesFromFile of module mdlMain.bas"
End Function




'---------------------------------------------------------------------------------------
' Procedure : GetEncoderClsid
' Author    : beededea
' Date      : 21/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
' Built-in encoders for saving: (You can *try* to get other types also)
'   image/bmp
'   image/jpeg
'   image/gif
'   image/tiff
'   image/pngI received a
'
' Notes When Saving:
'The JPEG encoder supports the Transformation, Quality, LuminanceTable, and ChrominanceTable parameter categories.
'The TIFF encoder supports the Compression, ColorDepth, and SaveFlag parameter categories.
'The BMP, PNG, and GIF encoders no do not support additional parameters.
'
' Purpose:
'The function calls GetImageEncoders to get an array of ImageCodecInfo objects. If one of the
'ImageCodecInfo objects in that array represents the requested encoder, the function returns
'the index of the ImageCodecInfo object and copies the CLSID into the variable pointed to by
'pClsid. If the function fails, it returns –1.

Private Function GetEncoderClsid(strMimeType As String, ClassID As CLSID) As Long
   Dim num As Long
   Dim Size As Long
   Dim i As Long

   Dim ICI() As ImageCodecInfo
   Dim Buffer() As Byte
   
   On Error GoTo GetEncoderClsid_Error

   GetEncoderClsid = -1 'Failure flag

   ' Get the encoder array size
   Call GdipGetImageEncodersSize(num, Size)
   If Size = 0 Then Exit Function ' Failed!

   ' Allocate room for the arrays dynamically
   ReDim ICI(1 To num) As ImageCodecInfo
   ReDim Buffer(1 To Size) As Byte

   ' Get the array and string data
   Call GdipGetImageEncoders(num, Size, Buffer(1))
   ' Copy the class headers
   Call CopyMemory(ICI(1), Buffer(1), (Len(ICI(1)) * num))

   ' Loop through all the codecs
   For i = 1 To num
      ' Must convert the pointer into a usable string
      If StrComp(PtrToStrW(ICI(i).MimeType), strMimeType, vbTextCompare) = 0 Then
         ClassID = ICI(i).ClassID   ' Save the class id
         GetEncoderClsid = i        ' return the index number for success
         Exit For
      End If
   Next
   ' Free the memory
   Erase ICI
   Erase Buffer

   On Error GoTo 0
   Exit Function

GetEncoderClsid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetEncoderClsid of Module mdlMain"
End Function



'---------------------------------------------------------------------------------------
' Procedure : PtrToStrW
' Author    : From www.mvps.org/vbnet...i think
' Date      : 21/08/2020
' Purpose   : '   Dereferences an ANSI or Unicode string pointer
'   and returns a normal VB BSTR
'---------------------------------------------------------------------------------------
'
Public Function PtrToStrW(ByVal lpsz As Long) As String
    Dim sOut As String
    Dim lLen As Long

   On Error GoTo PtrToStrW_Error

    lLen = lstrlenW(lpsz)

    If (lLen > 0) Then
        sOut = StrConv(String$(lLen, vbNullChar), vbUnicode)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen * 2)
        PtrToStrW = StrConv(sOut, vbFromUnicode)
    End If

   On Error GoTo 0
   Exit Function

PtrToStrW_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PtrToStrW of Module mdlMain"
End Function

'---------------------------------------------------------------------------------------
' Procedure : createScaledImg
' Author    : Credit to Olaf Schmidt for the original
'             also to Joaquim https://www.vbforums.com/showthread.php?840601-RESOLVED-how-use-ColorMatrix
' Date      : 07/04/2020
' Purpose   : Creates the scaled image with quality and opacity attributes
'---------------------------------------------------------------------------------------
'
Public Function createScaledImg(SrcImg As Long, dxSrc As Long, dySrc As Long, dxDst As Long, dyDst As Long, opacity As Integer) As Long
    Dim img As Long
    Dim Ctx As Long
    Dim imgQuality As Long
    Dim SmoothingMode As Long
    'Dim stat As GpStatus
    Dim imgAttr As Long
    Dim clrMatrix As ColorMatrix
    Dim graMatrix As ColorMatrix
    Dim rDIconQuality As String: rDIconQuality = vbNullString

    On Error GoTo createScaledImg_Error

    imgAttr = &H11
    rDIconQuality = "2"
    
    'Setup the transformation matrix for alpha adjustment used in GdipSetImageAttributesColorMatrix below
    clrMatrix.m(0, 0) = 1
    clrMatrix.m(1, 1) = 1
    clrMatrix.m(2, 2) = 1
    clrMatrix.m(3, 3) = 1 * opacity / 100 ' 0.5 'Alpha transform (50%) ' cairo_image_surface_create_for_data (BGRA = alpha)
    clrMatrix.m(4, 4) = 1

    ' Use the existing buffer pCairoBuf that has (iRenderWidth * iRenderHeight * 4) bytes, 4 bytes per pixel for ARGB32 and zero row padding
    ' cairo_surface_t *pSurface = cairo_image_surface_create_for_data(pCairoBuf, CAIRO_FORMAT_ARGB32, iRenderWidth, iRenderHeight, (iRenderWidth * 4))

    ' set the image quality and smoothing mode
    If rDIconQuality = "0" Then
        imgQuality = &H1 '    ipmNearestNeighbor = &H5
        SmoothingMode = SmoothingModeNone
    End If
    If rDIconQuality = "1" Then
        imgQuality = &H6 '    ipmHighQualityBiLinear = &H6
        SmoothingMode = SmoothingModeHighSpeed
    End If
    If rDIconQuality = "2" Then
        imgQuality = &H7 '    ipmHighQualityBicubic = &H7
        SmoothingMode = SmoothingModeHighQuality
    End If
    
    'Creates a Bitmap object (surface) based on an array of bytes along with the destination size and format information img is the pointer to that bitmap object
    Call GdipCreateBitmapFromScan0(dxDst, dyDst, dxDst * 4, PixelFormat32bppPARGB, 0, img) ' Cairo.CreateSurface & Set_Device_Offset
    
    If img Then
        createScaledImg = img ' set the return value to the bitmap object
        'Creates a Graphics object context - ctx, that is associated with an Image bitmap object (surface) ie. the hw context of the image
        Call GdipGetImageGraphicsContext(img, Ctx)
    Else
        Err.Raise vbObjectError, , "unable to create scaled Img-Resource"
    End If
    
    If Ctx Then
        ' set the quality using three GDIP functions
        Call GdipSetPixelOffsetMode(Ctx, 3)            '     4=Half, 3=None
        Call GdipSetInterpolationMode(Ctx, imgQuality) ' three levels of quality
        Call GdipSetSmoothingMode(Ctx, SmoothingMode)  '          ditto
        
        ' Sets the compositing quality of this Graphics object when alpha blended. Speed vs quality. Used in conjunction with GdipSetCompositingMode
        'Call GdipSetCompositingQuality(Ctx, CompositingQualityHighQuality)  ' CompositingQualityHighSpeed
                                
        'Create storage for the image attributes struct used below
        Call GdipCreateImageAttributes(imgAttr)

        'Setup the image attributes using the transformation matrix and the enum values, ColorAdjustTypeBitmap and ColorMatrixFlagsDefault
        Call GdipSetImageAttributesColorMatrix(imgAttr, ColorAdjustTypeBitmap, 1, clrMatrix, graMatrix, ColorMatrixFlagsDefault)

        ' draw the loaded source image scaled onto a generated image to the desired scale with the above image quality and opacity attributes
        If SrcImg <> 0 Then
            GdipDrawImageRectRectI Ctx, SrcImg, 0, 0, dxDst, dyDst, 0, 0, dxSrc, dySrc, 2, imgAttr, 0, 0 ' Cairo.Cairo_Surface /  cairo_image_surface_create_for_data (BGRA = alpha)
            ' Set CC = Cairo.CreateSurface(Me.ScaleWidth, Me.ScaleHeight).CreateContext
        End If
        
        ' the image has now been drawn so we can now delete the now unwanted graphics context
        Call GdipDeleteGraphics(Ctx) ' cairo_destroy(cr) &  'cairo_surface_destroy(surface)
    End If

   On Error GoTo 0
   Exit Function

createScaledImg_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createScaledImg of module mdlMain.bas"
End Function



'---------------------------------------------------------------------------------------
' Procedure : resizeAndLoadImgToDict
' Author    : beededea
' Date      : 07/04/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'Uses an extracted function from Olaf Schmidt's code from gdiPlusCacheCls to read the file as a series of bytes that consumes memory 200K -800k approx each run.
'Creates a stream object stored in global memory using the location address of the variable where the data resides
'Creates a GDI+ Image object based on the stream, using GdipLoadImageFromStream that consumes memory 300k approx.
'Finally, uses a function GdipCreateBitmapFromScan0 to both create and resize the image that consumes memory 100k approx.
'
'This occurs for each image read into the collection but the memory is not being released.
'
'Tried releasing the memory by setting the variablles to erase or set to nothing or by assigning them to an empty object
'to no avail. Instead there is a workaround that combines two collections to form a new one, see removeImageFromDictionary directly below this routine


' .10 DAEB 01/05/2021 mdlMain.bas huge number of changes as I moved multiple declarations, subs and functions to mdlmain from frmMain.
Public Function resizeAndLoadImgToDict(ByRef thisDictionary As Dictionary, ByVal Key As String, ByVal strFilename As String, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal fullStringKey As String = "", Optional ByVal imageOpacity As Integer) As Long

    Dim thiskey As String
    Dim saveStatus As Boolean
    Dim encoderCLSID As CLSID
    Dim bytesFromFile() As Byte
    Dim Strm As stdole.IUnknown
    Dim img As Long
    Dim imgCrop As Long
    Dim imgCrop2 As Long
    Dim dx As Long
    Dim dy As Long
    Dim dockSkinWidth As Long
    Dim cropWidth As Long
    Dim action As String
    Dim lngPixelFormat As Long
    Dim iconBitmap As Long

    On Error GoTo resizeAndLoadImgToDict_Error
    
    cropWidth = 100 ' needs a param
    
    ' Get the CLSID of the PNG encoder
    Call GetEncoderClsid("image/png", encoderCLSID)
    
    ' uses an extracted function from Olaf Schmidt's code from gdiPlusCacheCls to read the file as a series of bytes
    bytesFromFile = ReadBytesFromFile(strFilename)  ' <consumes memory 200K -800k approx.

    ' creates a stream object stored in global memory using the location address of the variable where the data resides, Olaf Schmidt
    CreateStreamOnHGlobal VarPtr(bytesFromFile(0)), 0, Strm
    
    ' Creates a GDI+ Image object based on the stream, loads it into img - Olaf Schmidt
    Call GdipLoadImageFromStream(ObjPtr(Strm), img)       ' cairo_image_surface_create_from_png (const char *filename);
    If img = 0 Then Err.Raise vbObjectError, , "Could not load image " & strFilename & " with GDIPlus"

    'GDI+ API to determine image dimensions, Olaf Schmidt
    Call GdipGetImageWidth(img, dx) ' cairo_image_surface_get_width ' the width of the surface in pixels.
    If Width <= 0 Then Width = dx
    
    Call GdipGetImageHeight(img, dy) ' cairo_image_surface_get_height ()
    If Height <= 0 Then Height = dy
        
    ' a bit of a bodge but we need to handle the background image by cropping it
    ' Rocketdock has a background theme image in a single image, it is cropped left and right to extract the ends whilst the middle is both cropped and stretched.

    ' uses a function createScaledImg extracted from Olaf Schmidt's code in gdiPlusCacheCls to create and resize the image
'    If Key = "sDSkinMid" Then
'
'
'        ' Get the current image pixel format. The C++ SDK clone example used PixelFormatDontCare, but this can limit what you can do with the image.
'        Call GdipGetImagePixelFormat(img, lngPixelFormat)
'
'        ' when cloning a bitmap area beyond the original size of the image, it fails and creates nothing.
'        ' this ensures the cropped width is never greater than the original image.
'
'        ' Create a new Bitmap object (surface) by cropping a portion of the long 2000px bitmap to the calculated dock width - x, y, width, height
'        Call GdipCloneBitmapAreaI(0, 0, cropWidth, dy, lngPixelFormat, img, imgCrop) '
'        iconBitmap = createScaledImg(imgCrop, cropWidth, dy, cropWidth, Height, imageOpacity)
'    Else
        ' create a scaled version of the image surface
        iconBitmap = createScaledImg(img, dx, dy, Width, Height, imageOpacity) ' <consumes memory 100k approx.
'    End If
       
    'override the key
    If fullStringKey = "" Then
        ' create a unique key string
        thiskey = Key & "ResizedImg" & LTrim$(Str$(Width))
    Else
        thiskey = fullStringKey
    End If
    
    ' add the bitmap to the dictionary collection
    If thisDictionary.Exists(thiskey) Then
        thisDictionary.Remove thiskey
    End If
    thisDictionary.Add thiskey, iconBitmap
    
    GdipDisposeImage iconBitmap

   On Error GoTo 0
   Exit Function

resizeAndLoadImgToDict_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure resizeAndLoadImgToDict with this file - " & strFilename & " of module mdlMain.bas"
        
End Function
