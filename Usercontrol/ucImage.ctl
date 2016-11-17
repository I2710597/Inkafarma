VERSION 5.00
Begin VB.UserControl ucImage 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   1875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2325
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HasDC           =   0   'False
   PropertyPages   =   "ucImage.ctx":0000
   ScaleHeight     =   125
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   155
   ToolboxBitmap   =   "ucImage.ctx":000F
   Windowless      =   -1  'True
End
Attribute VB_Name = "ucImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module      : ucImage
' DateTime    : 04/03/2008 11:00
' Author      : Cobein
' Mail        : cobein27@hotmail.com
' Purpose     : Simple Image control replacement (Beta)
' Requirements: GDI Plus
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
'
' Credits     : LaVolpe, Paul Caton and http://www.activevb.de
'
' History     : 04/03/2008 Alpha realease
'               06/03/2008 Alpha 1
'               06/03/2008 Alpha 2
'               07/03/2008 Beta Release, added properties and methods
'               07/03/2008 Added bright and contrast
'               20/03/2008 Added 5 stretchig methods
'               22/03/2008 Major changes
'---------------------------------------------------------------------------------------
Option Explicit
Option Base 0

Private Const GWL_WNDPROC       As Long = -4
Private Const GW_OWNER          As Long = 4
Private Const WS_CHILD          As Long = &H40000000
Private Const UnitPixel         As Long = &H2&

Private Const InterpolationModeNearestNeighbor      As Long = &H5&
Private Const InterpolationModeHighQualityBicubic   As Long = &H7&
Private Const InterpolationModeHighQualityBilinear  As Long = &H6&

Private Enum ColorAdjustType
    ColorAdjustTypeDefault = 0
    ColorAdjustTypeBitmap = 1
    ColorAdjustTypeBrush = 2
    ColorAdjustTypePen = 3
    ColorAdjustTypeText = 4
    ColorAdjustTypeCount = 5
    ColorAdjustTypeAny = 6
End Enum

Private Enum ColorMatrixFlags
    ColorMatrixFlagsDefault = 0
    ColorMatrixFlagsSkipGrays = 1
    ColorMatrixFlagsAltGray = 2
End Enum

Enum eScaleMode
    eActualSize
    eStretch
    eScaleDown
    eScale
    eScaleUp
End Enum

Private Type RECTF
    nLeft                       As Single
    nTop                        As Single
    nWidth                      As Single
    nHeight                     As Single
End Type

Private Type GdiplusStartupInput
    GdiplusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type

Private Type POINTAPI
    X                           As Long
    y                           As Long
End Type

Private Type RECT
    Left                        As Long
    Top                         As Long
    Right                       As Long
    Bottom                      As Long
End Type

Private Type COLORMATRIX
    m(0 To 4, 0 To 4)           As Single
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, ByRef Image As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal bitmap As Long, ByRef hbmReturn As Long, ByVal Background As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As Long
Private Declare Function GdipGetImageBounds Lib "gdiplus.dll" (ByVal nImage As Long, srcRect As RECTF, srcUnit As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function CreateWindowExA Lib "user32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal y As Long) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal bitmap As Long, hbmReturn As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef imageattr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As Long, ByVal ColorAdjust As ColorAdjustType, ByVal EnableFlag As Boolean, ByRef MatrixColor As COLORMATRIX, ByRef MatrixGray As COLORMATRIX, ByVal flags As ColorMatrixFlags) As Long
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As Long, ByVal rfType As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Interpolation As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long

Private z_CbMem                 As Long    'Callback allocated memory address
Private z_Cb()                  As Long    'Callback thunk array

Public Event Click(ByVal Button As Integer)
Public Event DblClick(ByVal Button As Integer)
Public Event MouseExit()
Public Event MouseEnter()
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single, State As Integer)

Private c_lBtnClickTracker      As Long
Private c_lBitmap               As Long
Private c_lAttributes           As Long
Private c_lWidth                As Long
Private c_lHeight               As Long
Private c_bvData()              As Byte
Private c_sFilename             As String
Private c_eScale                As eScaleMode
Private c_bIn                   As Boolean
Private c_tPT                   As POINTAPI
Private c_lhWnd                 As Long
Private c_lContrast             As Long
Private c_lBrightness           As Long
Private c_lAlpha                As Long
Private c_bGrayScale            As Boolean
Private c_bFlipH                As Boolean
Private c_bFlipV                As Boolean
Private c_lAngle                As Long

Public Sub About()
Attribute About.VB_UserMemId = -552
    Call MsgBox("Cobein ucImage Control, Version 0.3" & _
       vbNewLine & vbNewLine & _
       "http://www.ClassicVisualBasic.com", , "About ucImage Control")
End Sub

'==================================================================================
'////////////////////////////         PROPERTIES         \\\\\\\\\\\\\\\\\\\\\\\\\\
'==================================================================================
Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

Public Property Get FlipHorizontal() As Boolean
    FlipHorizontal = c_bFlipH
End Property

Public Property Let FlipHorizontal(ByVal bFlipH As Boolean)
    c_bFlipH = bFlipH
    Call PropertyChanged("bFlipH")
    Call UserControl.Refresh
End Property

Public Property Get FlipVertical() As Boolean
    FlipVertical = c_bFlipV
End Property

Public Property Let FlipVertical(ByVal bFlipV As Boolean)
    c_bFlipV = bFlipV
    Call PropertyChanged("bFlipV")
    Call UserControl.Refresh
End Property

Public Property Get GrayScale() As Boolean
    GrayScale = c_bGrayScale
End Property

Public Property Let GrayScale(ByVal bGrayScale As Boolean)
    c_bGrayScale = bGrayScale
    Call PropertyChanged("bGrayScale")
    Call UserControl.Refresh
End Property

Public Property Get ScaleMode() As eScaleMode
    ScaleMode = c_eScale
End Property

Public Property Let ScaleMode(ByVal eScaleMode As eScaleMode)
    c_eScale = eScaleMode
    Call PropertyChanged("eScale")
    Call UserControl.Refresh
End Property

Public Property Get Brightness() As Long
    Brightness = c_lBrightness
End Property

Public Property Let Brightness(ByVal lBrightness As Long)
    c_lBrightness = lBrightness
    Call PropertyChanged("lBrightness")
    Call UserControl.Refresh
End Property

'Public Property Get Contrast() As Long
'    Contrast = c_lContrast
'End Property
'
'Public Property Let Contrast(ByVal lContrast As Long)
'    c_lContrast = lContrast
'    Call PropertyChanged("lContrast")
'    Call UserControl.Refresh
'End Property

Public Property Get Alpha() As Long
    Alpha = c_lAlpha
End Property

Public Property Let Alpha(ByVal lAlpha As Long)
    c_lAlpha = lAlpha
    Call PropertyChanged("lAlpha")
    Call UserControl.Refresh
End Property

Public Property Get Angle() As Long
    Angle = c_lAngle
End Property

Public Property Let Angle(ByVal lAngle As Long)
    c_lAngle = lAngle
    Call PropertyChanged("lAngle")
    Call UserControl.Refresh
End Property

Public Property Get PictureWidth() As Long
    PictureWidth = c_lWidth
End Property

Public Property Get PictureHeight() As Long
    PictureHeight = c_lHeight
End Property

'==================================================================================
'////////////////////////////          METHODS           \\\\\\\\\\\\\\\\\\\\\\\\\\
'==================================================================================
Public Sub Refresh()
    Call UserControl.Refresh
End Sub

Public Function PaintPicture( _
       ByVal lhDC As Long, _
       ByVal dstX As Long, _
       ByVal dstY As Long, _
       Optional ByVal dstWidth As Long, _
       Optional ByVal dstHeight As Long, _
       Optional ByVal SrcX As Long, _
       Optional ByVal SrcY As Long, _
       Optional ByVal srcWidth As Long, _
       Optional ByVal srcHeight As Long) As Boolean
       
    PaintPicture = RenderTo(lhDC, _
       dstX, dstY, dstWidth, dstHeight, _
       SrcX, SrcY, srcWidth, srcHeight)
       
End Function

Public Function IconHandle() As Long
    Call GdipCreateHICONFromBitmap(c_lBitmap, IconHandle)
End Function

Public Function GetStream() As Byte()
    GetStream = c_bvData
End Function

Public Sub LoadImageFromStream(ByRef bvStream() As Byte)
    c_bvData() = bvStream
    Call LoadFromStream(bvStream)
    Call UserControl.Refresh
End Sub

Public Function SaveToFile(ByVal sFile As String) As Boolean
    Dim iFile       As Integer
    
    On Local Error GoTo SaveToFile_Error

    iFile = FreeFile
    Open sFile For Binary Access Write As iFile
    Put iFile, , c_bvData
    Close iFile
    SaveToFile = True
    
    Exit Function
SaveToFile_Error:
End Function

Public Function GetFileName() As String
    GetFileName = c_sFilename
End Function

Public Function LoadImageFromFile(ByVal sFile As String) As Boolean
    LoadImageFromFile = ppgLoadStream(sFile)
End Function

Public Function LoadImageFromRes( _
       ByVal ResIndex As Variant, _
       ByVal ResSection As Variant, _
       Optional VBglobal As IUnknown) As Boolean
    
    Dim bvData()    As Byte
    Dim oVBglobal   As VB.Global
    
    On Local Error GoTo LoadImageFromCustomRes_Error

    If VBglobal Is Nothing Then
        Set oVBglobal = VB.Global
    ElseIf TypeOf VBglobal Is VB.Global Then
        Set oVBglobal = VBglobal
    ElseIf VBglobal Is Nothing Then
        Set oVBglobal = VB.Global
    End If
    
    bvData = oVBglobal.LoadResData(ResIndex, ResSection)
    
    LoadImageFromRes = LoadFromStream(bvData)

    Call UserControl.Cls
    Call LoadFromStream(bvData)
    Call UserControl_Paint

LoadImageFromCustomRes_Error:
End Function

'==================================================================================
'////////////////////////////       PROPERTY PAGE        \\\\\\\\\\\\\\\\\\\\\\\\\\
'==================================================================================
Friend Function ppgLoadStream(ByVal sFile As String) As Boolean
    Dim iFile       As Integer
    Dim bvData()    As Byte
    Dim svName()    As String
    
    On Local Error GoTo LoadStream_Error

    iFile = FreeFile
    Open sFile For Binary Access Read As iFile
    ReDim bvData(LOF(iFile) - 1)
    Get iFile, , bvData
    Close iFile
    
    svName = Split(sFile, "\")
    c_sFilename = svName(UBound(svName))
    c_bvData() = bvData
    
    Call PropertyChanged("bvData")
    Call PropertyChanged("Filename")
    
    Call UserControl.Cls
    Call LoadFromStream(bvData)
    Call UserControl_Paint

    ppgLoadStream = True
LoadStream_Error:
End Function

Friend Function ppgGetFilename() As String
    ppgGetFilename = c_sFilename
End Function

'==================================================================================
'////////////////////////////        USER CONTROL        \\\\\\\\\\\\\\\\\\\\\\\\\\
'==================================================================================
Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, y, State)
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click(c_lBtnClickTracker \ &H10)
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick(c_lBtnClickTracker \ &H10)
End Sub

Private Sub UserControl_HitTest(X As Single, y As Single, HitResult As Integer)
    HitResult = vbHitResultHit

    If Ambient.UserMode Then
        Dim PT  As POINTAPI
        Call GetCursorPos(c_tPT)
        Call ClientToScreen(c_lhWnd, PT)
        c_tPT.X = c_tPT.X - PT.X - X
        c_tPT.y = c_tPT.y - PT.y - y
    End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, X, y)

    If Not c_bIn Then
        c_bIn = True
        RaiseEvent MouseEnter
        SetTimer UserControl.hWnd, ObjPtr(Me) + 1, 10, zb_AddressOf(1, 4, 1)
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    c_lBtnClickTracker = (c_lBtnClickTracker Or Button)
    RaiseEvent MouseDown(Button, Shift, X, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    c_lBtnClickTracker = (c_lBtnClickTracker And &HF)
    If (c_lBtnClickTracker And Button) = Button Then
        c_lBtnClickTracker = (c_lBtnClickTracker Or Button * &H10)
    End If
    RaiseEvent MouseUp(Button, Shift, X, y)
End Sub

Private Sub UserControl_Paint()
    Dim lW As Long, lH As Long, lT As Long, lL As Long
    
    If Not c_lBitmap = 0 Then
        On Error Resume Next
        
        With UserControl
            If c_eScale = eActualSize Then
                .Height = c_lHeight * 15
                .Width = c_lWidth * 15
            End If
        
            ScalePicture c_eScale, c_lWidth, c_lHeight, _
               .Width / 15, .Height / 15, lW, lH, lL, lT
       
            Call RenderTo(.hdc, lL, lT, lW, lH)
        End With
        
    Else
        Call DrawFrame
    End If
End Sub

Private Sub UserControl_InitProperties()
    c_lAlpha = 100
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    c_lhWnd = UserControl.ContainerHwnd
    Call ManageGDIToken(c_lhWnd)
        
    With PropBag
        c_sFilename = .ReadProperty("Filename", vbNullString)
        c_eScale = .ReadProperty("eScale", 0)
        c_lContrast = .ReadProperty("lContrast", 100)
        c_lBrightness = .ReadProperty("lBrightness", 0)
        c_lAlpha = .ReadProperty("lAlpha", 100)
        c_bGrayScale = .ReadProperty("bGrayScale", False)
        c_lAngle = .ReadProperty("lAngle", 0)
        c_bFlipH = .ReadProperty("bFlipH", False)
        c_bFlipV = .ReadProperty("bFlipV", False)

        If CBool(.ReadProperty("bData", False)) Then
            c_bvData() = .ReadProperty("bvData")
            If c_lBitmap = 0 Then
                Dim bvData() As Byte
                bvData = c_bvData
                Call LoadFromStream(bvData)
            End If
        End If

    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        If IsArrayDim(VarPtrArray(c_bvData)) Then
            Call .WriteProperty("bvData", c_bvData)
            Call .WriteProperty("bData", True)
        Else
            Call .WriteProperty("bData", False)
        End If
        Call .WriteProperty("Filename", c_sFilename)
        Call .WriteProperty("eScale", c_eScale)
        Call .WriteProperty("lContrast", c_lContrast)
        Call .WriteProperty("lBrightness", c_lBrightness)
        Call .WriteProperty("lAlpha", c_lAlpha)
        Call .WriteProperty("bGrayScale", c_bGrayScale)
        Call .WriteProperty("lAngle", c_lAngle)
        Call .WriteProperty("bFlipH", c_bFlipH)
        Call .WriteProperty("bFlipV", c_bFlipV)
    End With
End Sub

Private Sub UserControl_Terminate()
    Call ClearUp
    Call KillTimer(UserControl.hWnd, ObjPtr(Me) + 1)
    Call zTerminate
End Sub

'==================================================================================
'////////////////////////////      HELPER FUNCTIONS      \\\\\\\\\\\\\\\\\\\\\\\\\\
'==================================================================================
Private Sub DrawFrame()
    Dim lhPen As Long
    On Error Resume Next
    If Not Ambient.UserMode Then
        With UserControl
            lhPen = CreatePen(2, 1, &HFF0000)
            Call SelectObject(.hdc, lhPen)
            Call Rectangle(.hdc, 0, 0, .Width / 15, .Height / 15)
            Call DeleteObject(lhPen)
        End With
    End If
End Sub

Private Function RenderTo( _
       ByVal lhDC As Long, _
       ByVal dstX As Long, _
       ByVal dstY As Long, _
       Optional ByVal dstWidth As Long, _
       Optional ByVal dstHeight As Long, _
       Optional ByVal SrcX As Long, _
       Optional ByVal SrcY As Long, _
       Optional ByVal srcWidth As Long, _
       Optional ByVal srcHeight As Long) As Boolean

    Dim hGraphics       As Long
    Dim hAttributes     As Long
    Dim bvData()        As Byte
        
    Dim dBrightness     As Double
    Dim dContrast       As Double
    Dim dAlpha          As Double
    Dim tMatrixColor    As COLORMATRIX
    Dim tMatrixGray     As COLORMATRIX
    
    bvData = c_bvData
    Call LoadFromStream(bvData)

    If c_lBitmap = 0 Then Exit Function
    
    If dstWidth = 0 Then dstWidth = c_lWidth
    If dstHeight = 0 Then dstHeight = c_lHeight
    If srcWidth = 0 Then srcWidth = c_lWidth
    If srcHeight = 0 Then srcHeight = c_lHeight
    
    dBrightness = ValidateValue(c_lBrightness)
    dContrast = ValidateValue(c_lContrast)
    dAlpha = ValidateValue(c_lAlpha)
    
    If GdipCreateFromHDC(lhDC, hGraphics) = 0 Then
        
        With tMatrixColor
            .m(0, 0) = 1
            .m(1, 1) = 1
            .m(2, 2) = 1
            .m(4, 4) = 1
            
            If Not dContrast = 0 Then
                .m(0, 0) = 1 + dContrast
                .m(1, 1) = .m(0, 0)
                .m(2, 2) = .m(0, 0)
            End If
            
            If Not dBrightness = 0 Then
                .m(0, 4) = dBrightness
                .m(1, 4) = .m(0, 4)
                .m(2, 4) = .m(0, 4)
            End If
     
            If Not dAlpha = 100 Then
                .m(3, 3) = dAlpha
            End If
            
            If c_bGrayScale Then
                .m(0, 0) = 0.299
                .m(1, 0) = .m(0, 0)
                .m(2, 0) = .m(0, 0)
                .m(0, 1) = 0.587
                .m(1, 1) = .m(0, 1)
                .m(2, 1) = .m(0, 1)
                .m(0, 2) = 0.114
                .m(1, 2) = .m(0, 2)
                .m(2, 2) = .m(0, 2)
            End If
        End With

        If c_bFlipH Then Call GdipImageRotateFlip(c_lBitmap, 4&)
        If c_bFlipV Then Call GdipImageRotateFlip(c_lBitmap, 6&)
                            
        If GdipCreateImageAttributes(hAttributes) = 0 Then
                
            If GdipSetImageAttributesColorMatrix( _
               hAttributes, ColorAdjustTypeDefault, True, _
               tMatrixColor, tMatrixGray, _
               ColorMatrixFlagsDefault) = 0 Then
           
                If c_lAngle = 0 Then
                    If GdipDrawImageRectRectI( _
                       hGraphics, _
                       c_lBitmap, _
                       dstX, dstY, dstWidth, dstHeight, _
                       SrcX, SrcY, srcWidth, srcHeight, _
                       UnitPixel, _
                       hAttributes) = 0 Then
                        RenderTo = True
                    End If
                Else
                    If GdipRotateWorldTransform(hGraphics, c_lAngle + 180, 0) = 0 Then
                        Call GdipTranslateWorldTransform( _
                           hGraphics, _
                           dstX + (dstWidth \ 2), dstY + (dstHeight \ 2), _
                           1)
                    End If
                    If GdipDrawImageRectRectI( _
                       hGraphics, _
                       c_lBitmap, _
                       dstWidth \ 2, dstHeight \ 2, -dstWidth, -dstHeight, _
                       SrcX, SrcY, srcWidth, srcHeight, _
                       UnitPixel, _
                       hAttributes) = 0 Then
                        RenderTo = True
                    End If
                End If
            End If
                
            Call GdipDisposeImageAttributes(hAttributes)
        End If
        
        Call GdipDeleteGraphics(hGraphics)
    End If
    
End Function

Private Function ValidateValue(ByVal dVal As Double) As Double
    If dVal < 0 Then
        ValidateValue = 0
        Exit Function
    ElseIf dVal > 100 Then
        dVal = 100
    End If
    ValidateValue = dVal / 100
End Function

Private Function LoadFromStream(ByRef bvData() As Byte) As Boolean
    Dim IStream     As IUnknown
    Dim lhBitmap    As Long
    Dim TR          As RECTF
    
    If Not IsArrayDim(VarPtrArray(bvData)) Then Exit Function
    
    Call ManageGDIToken(c_lhWnd)
    Call ClearUp
    Call CreateStreamOnHGlobal(bvData(0), False, IStream)
    
    If Not IStream Is Nothing Then
        If GdipLoadImageFromStream(IStream, c_lBitmap) = 0 Then
            If GdipCreateHBITMAPFromBitmap(c_lBitmap, lhBitmap, 0) = 0 Then
                Call GdipGetImageBounds(c_lBitmap, TR, UnitPixel)
                c_lWidth = TR.nWidth
                c_lHeight = TR.nHeight
                LoadFromStream = True
            End If
        End If
    End If

    Set IStream = Nothing
End Function

Private Sub ClearUp()
    If Not c_lBitmap = 0 Then
        Call GdipDisposeImage(c_lBitmap)
        c_lBitmap = 0: c_lWidth = 0: c_lHeight = 0
    End If
End Sub

Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress As Long
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function

Private Function ScalePicture( _
       ByVal eScaleMode As eScaleMode, _
       ByVal lSrcWidth As Long, _
       ByVal lSrcHeight As Long, _
       ByVal lDstWidth As Long, _
       ByVal lDstHeight As Long, _
       ByRef lNewWidth As Long, _
       ByRef lNewHeight As Long, _
       ByRef lNewLeft As Long, _
       ByRef lNewTop As Long)

    Dim dHRatio As Double
    Dim dVRatio As Double
    Dim dRatio  As Double
    
    dHRatio = lSrcWidth / lDstWidth
    dVRatio = lSrcHeight / lDstHeight
     
    Select Case eScaleMode
        Case eActualSize
            lNewWidth = lSrcWidth
            lNewHeight = lSrcHeight
        Case eStretch
            lNewWidth = lDstWidth
            lNewHeight = lDstHeight
        Case eScaleDown
            If dHRatio > 1 Or dVRatio > 1 Then
                If dHRatio > dVRatio Then
                    dRatio = dHRatio
                Else
                    dRatio = dVRatio
                End If
            Else
                lNewWidth = lSrcWidth
                lNewHeight = lSrcHeight
            End If
        Case eScale
            If dHRatio > dVRatio Then
                dRatio = dHRatio
            Else
                dRatio = dVRatio
            End If
        Case eScaleUp
            If dHRatio < dVRatio Then
                dRatio = dHRatio
            Else
                dRatio = dVRatio
            End If
    End Select
    
    If Not dRatio = 0 Then
        lNewWidth = lSrcWidth / dRatio
        lNewHeight = lSrcHeight / dRatio
    End If
    
    lNewLeft = (lDstWidth - lNewWidth) / 2
    lNewTop = (lDstHeight - lNewHeight) / 2
End Function

Private Function ManageGDIToken(ByVal projectHwnd As Long) As Long
    If projectHwnd = 0& Then Exit Function
    
    Dim hwndGDIsafe     As Long                 'API window to monitor IDE shutdown
    
    Do
        hwndGDIsafe = GetParent(projectHwnd)
        If Not hwndGDIsafe = 0& Then projectHwnd = hwndGDIsafe
    Loop Until hwndGDIsafe = 0&
    ' ok, got the highest level parent, now find highest level owner
    Do
        hwndGDIsafe = GetWindow(projectHwnd, GW_OWNER)
        If Not hwndGDIsafe = 0& Then projectHwnd = hwndGDIsafe
    Loop Until hwndGDIsafe = 0&
    
    hwndGDIsafe = FindWindowEx(projectHwnd, 0&, "Static", "GDI+Safe Patch")
    If hwndGDIsafe Then
        ManageGDIToken = hwndGDIsafe    ' we already have a manager running for this VB instance
        Exit Function                   ' can abort
    End If
    
    Dim gdiSI           As GdiplusStartupInput  'GDI+ startup info
    Dim gToken          As Long                 'GDI+ instance token
    
    On Error Resume Next
    gdiSI.GdiplusVersion = 1                    ' attempt to start GDI+
    GdiplusStartup gToken, gdiSI
    If gToken = 0& Then                         ' failed to start
        If Err Then Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    Dim z_ScMem         As Long                 'Thunk base address
    Dim z_Code()        As Long                 'Thunk machine-code initialised here
    Dim nAddr           As Long                 'hwndGDIsafe prev window procedure

    Const WNDPROC_OFF   As Long = &H30          'Offset where window proc starts from z_ScMem
    Const PAGE_RWX      As Long = &H40&         'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&       'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&       'Release allocated memory flag
    Const MEM_LEN       As Long = &HD4          'Byte length of thunk
        
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
    If z_ScMem <> 0 Then                                     'Ensure the allocation succeeded
        ' we make the api window a child so we can use FindWindowEx to locate it easily
        hwndGDIsafe = CreateWindowExA(0&, "Static", "GDI+Safe Patch", WS_CHILD, 0&, 0&, 0&, 0&, projectHwnd, 0&, App.hInstance, ByVal 0&)
        If hwndGDIsafe <> 0 Then
        
            ReDim z_Code(0 To MEM_LEN \ 4 - 1)
        
            z_Code(12) = &HD231C031: z_Code(13) = &HBBE58960: z_Code(14) = &H12345678: z_Code(15) = &H3FFF631: z_Code(16) = &H74247539: z_Code(17) = &H3075FF5B: z_Code(18) = &HFF2C75FF: z_Code(19) = &H75FF2875
            z_Code(20) = &H2C73FF24: z_Code(21) = &H890853FF: z_Code(22) = &HBFF1C45: z_Code(23) = &H2287D81: z_Code(24) = &H75000000: z_Code(25) = &H443C707: z_Code(26) = &H2&: z_Code(27) = &H2C753339: z_Code(28) = &H2047B81: z_Code(29) = &H75000000
            z_Code(30) = &H2C73FF23: z_Code(31) = &HFFFFFC68: z_Code(32) = &H2475FFFF: z_Code(33) = &H681C53FF: z_Code(34) = &H12345678: z_Code(35) = &H3268&: z_Code(36) = &HFF565600: z_Code(37) = &H43892053: z_Code(38) = &H90909020: z_Code(39) = &H10C261
            z_Code(40) = &H562073FF: z_Code(41) = &HFF2453FF: z_Code(42) = &H53FF1473: z_Code(43) = &H2873FF18: z_Code(44) = &H581053FF: z_Code(45) = &H89285D89: z_Code(46) = &H45C72C75: z_Code(47) = &H800030: z_Code(48) = &H20458B00: z_Code(49) = &H89145D89
            z_Code(50) = &H81612445: z_Code(51) = &H4C4&: z_Code(52) = &HC63FF00

            z_Code(1) = 0                                                   ' shutDown mode; used internally by ASM
            z_Code(2) = zFnAddr("user32", "CallWindowProcA")                ' function pointer CallWindowProc
            z_Code(3) = zFnAddr("kernel32", "VirtualFree")                  ' function pointer VirtualFree
            z_Code(4) = zFnAddr("kernel32", "FreeLibrary")                  ' function pointer FreeLibrary
            z_Code(5) = gToken                                              ' Gdi+ token
            z_Code(10) = LoadLibrary("gdiplus")                             ' library pointer (add reference)
            z_Code(6) = GetProcAddress(z_Code(10), "GdiplusShutdown")       ' function pointer GdiplusShutdown
            z_Code(7) = zFnAddr("user32", "SetWindowLongA")                 ' function pointer SetWindowLong
            z_Code(8) = zFnAddr("user32", "SetTimer")                       ' function pointer SetTimer
            z_Code(9) = zFnAddr("user32", "KillTimer")                      ' function pointer KillTimer
        
            z_Code(14) = z_ScMem                                            ' ASM ebx start point
            z_Code(34) = z_ScMem + WNDPROC_OFF                              ' subclass window procedure location
        
            RtlMoveMemory z_ScMem, VarPtr(z_Code(0)), MEM_LEN               'Copy the thunk code/data to the allocated memory
        
            nAddr = SetWindowLong(hwndGDIsafe, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Subclass our API window
            RtlMoveMemory z_ScMem + 44, VarPtr(nAddr), 4& ' Add prev window procedure to the thunk
            gToken = 0& ' zeroize so final check below does not release it
            
            ManageGDIToken = hwndGDIsafe    ' return handle of our GDI+ manager
        Else
            VirtualFree z_ScMem, 0, MEM_RELEASE     ' failure - release memory
            z_ScMem = 0&
        End If
    Else
        VirtualFree z_ScMem, 0, MEM_RELEASE           ' failure - release memory
        z_ScMem = 0&
    End If
    
    If gToken Then GdiplusShutdown gToken       ' release token if error occurred
    
End Function

Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)  'Get the specified procedure address
End Function

'============================================================================================
' /////////////////// C A L L B A C K   T H U N K I N G   R O U T I N E S \\\\\\\\\\\\\\\\\\\
'============================================================================================
'*************************************************************************************************
'* cCallback - Class generic callback template
'*
'* Note:
'*  The callback declarations and code are exactly the same for a Class, Form or UserControl.
'*  The callback declarations and code can co-exist with subclassing declarations and code.
'*    With both types of code in a single file,..
'*      delete the duplicated declarations and code, Ctrl+F5 will find them for you
'*      pay careful attention to the nOrdinal parameter to zAddressOf
'*
'* Paul_Caton@hotmail.com
'* Copyright free, use and abuse as you see fit.
'*
'* v1.0 The original..................................................................... 20060408
'* v1.1 Added multi-thunk support........................................................ 20060409
'* v1.2 Added optional IDE protection.................................................... 20060411
'* v1.3 Added an optional callback target object......................................... 20060413
'*************************************************************************************************

'-Callback code-----------------------------------------------------------------------------------
Private Function zb_AddressOf(ByVal nOrdinal As Long, _
       ByVal nParamCount As Long, _
       Optional ByVal nThunkNo As Long = 0, _
       Optional ByVal oCallback As Object = Nothing, _
       Optional ByVal bIdeSafety As Boolean = True) As Long   'Return the address of the specified callback thunk
    '*************************************************************************************************
    '* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
    '* nParamCount  - The number of parameters that will callback
    '* nThunkNo     - Optional, allows multiple simultaneous callbacks by referencing different thunks... adjust the MAX_THUNKS Const if you need to use more than two thunks simultaneously
    '* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety   - Optional, set to false to disable IDE protection.
    '*************************************************************************************************
    Const MAX_FUNKS   As Long = 2                                               'Number of simultaneous thunks, adjust to taste
    Const FUNK_LONGS  As Long = 22                                              'Number of Longs in the thunk
    Const FUNK_LEN    As Long = FUNK_LONGS * 4                                  'Bytes in a thunk
    Const MEM_LEN     As Long = MAX_FUNKS * FUNK_LEN                            'Memory bytes required for the callback thunk
    Const PAGE_RWX    As Long = &H40&                                           'Allocate executable memory
    Const MEM_COMMIT  As Long = &H1000&                                         'Commit allocated memory
    Dim nAddr       As Long
  
    If nThunkNo < 0 Or nThunkNo > (MAX_FUNKS - 1) Then
        MsgBox "nThunkNo doesn't exist.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
        Exit Function
    End If
  
    If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
        Set oCallback = Me                                                      'Then it is me
    End If
  
    nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the callback address of the specified ordinal
    If nAddr = 0 Then
        MsgBox "Callback address not found.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
        Exit Function
    End If
  
    If z_CbMem = 0 Then                                                       'If memory hasn't been allocated
        ReDim z_Cb(0 To FUNK_LONGS - 1, 0 To MAX_FUNKS - 1) As Long             'Create the machine-code array
        z_CbMem = VirtualAlloc(z_CbMem, MEM_LEN, MEM_COMMIT, PAGE_RWX)          'Allocate executable memory
    End If
  
    If z_Cb(0, nThunkNo) = 0 Then                                             'If this ThunkNo hasn't been initialized...
        z_Cb(3, nThunkNo) = _
           GetProcAddress(GetModuleHandleA("kernel32"), "IsBadCodePtr")
        z_Cb(4, nThunkNo) = &HBB60E089
        z_Cb(5, nThunkNo) = VarPtr(z_Cb(0, nThunkNo))                           'Set the data address
        z_Cb(6, nThunkNo) = &H73FFC589: z_Cb(7, nThunkNo) = &HC53FF04: z_Cb(8, nThunkNo) = &H7B831F75: z_Cb(9, nThunkNo) = &H20750008: z_Cb(10, nThunkNo) = &HE883E889: z_Cb(11, nThunkNo) = &HB9905004: z_Cb(13, nThunkNo) = &H74FF06E3: z_Cb(14, nThunkNo) = &HFAE2008D: z_Cb(15, nThunkNo) = &H53FF33FF: z_Cb(16, nThunkNo) = &HC2906104: z_Cb(18, nThunkNo) = &H830853FF: z_Cb(19, nThunkNo) = &HD87401F8: z_Cb(20, nThunkNo) = &H4589C031: z_Cb(21, nThunkNo) = &HEAEBFC
    End If
  
    z_Cb(0, nThunkNo) = ObjPtr(oCallback)                                     'Set the Owner
    z_Cb(1, nThunkNo) = nAddr                                                 'Set the callback address
  
    If bIdeSafety Then                                                        'If the user wants IDE protection
        z_Cb(2, nThunkNo) = GetProcAddress(GetModuleHandleA("vba6"), "EbMode")  'EbMode Address
    End If
    
    z_Cb(12, nThunkNo) = nParamCount                                          'Set the parameter count
    z_Cb(17, nThunkNo) = nParamCount * 4                                      'Set the number of stck bytes to release on thunk return
  
    nAddr = z_CbMem + (nThunkNo * FUNK_LEN)                                   'Calculate where in the allocated memory to copy the thunk
    RtlMoveMemory nAddr, VarPtr(z_Cb(0, nThunkNo)), FUNK_LEN                  'Copy thunk code to executable memory
    zb_AddressOf = nAddr + 16                                                 'Thunk code start address
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
    Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
    Dim bVal  As Byte
    Dim nAddr As Long                                                         'Address of the vTable
    Dim i     As Long                                                         'Loop index
    Dim j     As Long                                                         'Loop limit
  
    RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
    If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
        If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
            If Not zProbe(nAddr + &H7A4, i, bSub) Then                            'Probe for a UserControl method
                Exit Function                                                       'Bail...
            End If
        End If
    End If
  
    i = i + 4                                                                 'Bump to the next entry
    j = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
    Do While i < j
        RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry
    
        If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
            Exit Do                                                               'Bad method signature, quit loop
        End If

        RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
        If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
            Exit Do                                                               'Bad method signature, quit loop
        End If
    
        i = i + 4                                                             'Next vTable entry
    Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
    Dim bVal    As Byte
    Dim nAddr   As Long
    Dim nLimit  As Long
    Dim nEntry  As Long
  
    nAddr = nStart                                                            'Start address
    nLimit = nAddr + 32                                                       'Probe eight entries
    Do While nAddr < nLimit                                                   'While we've not reached our probe depth
        RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
        If nEntry <> 0 Then                                                     'If not an implemented interface
            RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
            If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
                nMethod = nAddr                                                     'Store the vTable entry
                bSub = bVal                                                         'Store the found method signature
                zProbe = True                                                       'Indicate success
                Exit Function                                                       'Return
            End If
        End If
    
        nAddr = nAddr + 4                                                       'Next vTable entry
    Loop
End Function

Private Sub zTerminate()
    
    Const MEM_RELEASE As Long = &H8000&                                'Release allocated memory flag
    If Not z_CbMem = 0 Then                                            'If memory allocated
        If Not VirtualFree(z_CbMem, 0, MEM_RELEASE) = 0 Then
            z_CbMem = 0  'Release; Indicate memory released
            Erase z_Cb()
        End If
    End If
End Sub

'*************************************************************************************************
'* Callbacks - the final private routine is ordinal #1, second last is ordinal #2 etc
'*************************************************************************************************
'Callback ordinal 1
Private Function Timer_MouseExit(ByVal hWnd As Long, ByVal tMsg As Long, ByVal TimerID As Long, ByVal tickCount As Long) As Long

    Dim PT As POINTAPI
    Dim CPT As POINTAPI
    Dim TR As RECT
    
    Call GetCursorPos(PT)
    Call ClientToScreen(c_lhWnd, CPT)
    Call SetRect(TR, 0, 0, UserControl.Width / 15, UserControl.Height / 15)
    
    CPT.X = PT.X - CPT.X - c_tPT.X
    CPT.y = PT.y - CPT.y - c_tPT.y
            
    If PtInRect(TR, CPT.X, CPT.y) = 0 Or _
       Not WindowFromPoint(PT.X, PT.y) = c_lhWnd Then
        Call KillTimer(hWnd, TimerID)
        c_bIn = False
        RaiseEvent MouseExit
    End If
    
    ' CAUTION: DO NOT ADD ANY ADDITIONAL CODE OR COMMENTS PAST THE "END FUNCTION"
    '          STATEMENT BELOW. Paul Caton's zProbe routine will read it as a start
    '          of a new function/sub and the class timer's will fail every time.
End Function
