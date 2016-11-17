VERSION 5.00
Begin VB.UserControl mButton 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1365
   ClipControls    =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   1365
   Begin VB.PictureBox Container 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   1020
      Left            =   0
      ScaleHeight     =   1020
      ScaleWidth      =   810
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Timer tEvents 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   15
   End
End
Attribute VB_Name = "mButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////
'///Información////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////
'///Nombre:         mButton
'///Autor:          SONIC88 - Marco A. Olivares Aracena
'///Mail:           sonic88@live.cl - maoa17@gmail.com
'///Decripción:     Botón simple de uso común mediante el API BitBlt.
'///Fecha:          22 de Abril de 2008 - 20:36
'///Versión:        1.31 (Nueva versión del BTNGRAFICO)
'///Comentario:     Todavía a pruebas, Alguna pifia favor de comunicar.
'///////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////

Option Explicit

'///Delaraciones///
Private Declare Function GetWindowRect Lib "user32.dll" ( _
ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Function GetCursorPos Lib "user32.dll" ( _
ByRef lpPoint As POINTAPI) As Long

Private Declare Function BitBlt Lib "gdi32.dll" ( _
ByVal hDestDC As Long, ByVal X As Long, _
ByVal Y As Long, ByVal nWidth As Long, _
ByVal nHeight As Long, ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal dwRop As Long) As Long

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long


'///Tipos///
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'///Enumeraciones///
Private Enum Events
mMouseHover = 1
mMouseDown = 2
mMouseup = 3
mDisabled = 4
End Enum

'///Variables///
Dim mPoint As POINTAPI
Dim mRect As RECT

Dim mHeight As Long, mWidth As Long, mSrc As Long, mDest As Long

Dim mX As Long, mY As Long, mLeft As Long, mTop As Long, mRight As Long, mBottom As Long
Dim mButtonX As Long, mDbClick As Boolean, Active As Boolean

'///Eventos///
Event Click()
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseEnter()
Event MouseOut()

Private Sub UserControl_DblClick()
'SetCapture UserControl.hWnd
tEvents_Timer
mDbClick = True
Active = True
End Sub

'///////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////
Private Sub UserControl_Resize()
On Error Resume Next
With UserControl
    If .Height < 1 Or .Width < 1 Then
        .Width = Container.Width: .Height = Container.Height / 4
    Else
    On Error Resume Next
        .Width = Container.Width: .Height = Container.Height / 4
    End If
End With

End Sub

Private Sub UserControl_Paint()
UserControl.Cls
tEvents_Timer
End Sub

Private Sub Container_Change()
tEvents_Timer
UserControl_Resize
End Sub

Private Sub tEvents_Timer()

'///Obtenemos la posición, ancho y alto del Usercontrol y la posicón del Mouse
'///luego las comparamos y si el Mouse esta pasando por encima del Usercontrol
'///creamos el evento MouseHover.

GetCursorPos mPoint
mX = mPoint.X: mY = mPoint.Y
GetWindowRect UserControl.hWnd, mRect
mLeft = mRect.Left: mTop = mRect.Top: mRight = mRect.Right: mBottom = mRect.Bottom


If Enabled = True Then

Select Case mButtonX
Case 0, 2, 4
    If (mX < mLeft Or mX > mRight) Or (mY < mTop Or mY > mBottom) Then
        '///Posición del Mouse fuera del control.
        RaiseEvent MouseOut
        mDrawButton mMouseup
        'tEvents = False
    Else
        '///Posición del Mouse dentro del control.
        'If mButtonX <> 1 Then
        RaiseEvent MouseEnter
        mDrawButton mMouseHover
        'End If
    End If
    
    
Case 3, 5, 6, 7

        RaiseEvent MouseOut
        mDrawButton mMouseHover
        tEvents = False
Case 1

    If (mX < mLeft Or mX > mRight) Or (mY < mTop Or mY > mBottom) Then
        '///Posición del Mouse fuera del control.
        RaiseEvent MouseOut
        mDrawButton mMouseup
        'tEvents = False

    Else
        '///Posición del Mouse dentro del control pero presionando el botón
        '///izquierdo del Mouse.
        RaiseEvent MouseEnter
        If Active = True Then
            ReleaseCapture
            mDrawButton mMouseDown
            SetCapture UserControl.hWnd
        Else
            mDrawButton mMouseup
            ReleaseCapture
        End If

    End If
End Select

Else
'///En caso de que el botón este desabilitado dibujamos otra imagen.
mDrawButton mDisabled
End If

End Sub

Private Sub UserControl_Click()
If mButtonX = 1 Then
    tEvents_Timer
    RaiseEvent Click
End If
mButtonX = 0
End Sub

Private Sub UserControl_Initialize()
tEvents_Timer
mButtonX = 0
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
mButtonX = Button

If tEvents = False Then
    tEvents = True
End If

If Button = 1 Then
    Active = True
Else
    Active = False
End If




End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mButtonX = Button
RaiseEvent MouseMove(Button, Shift, X, Y)

If tEvents = False Then
    tEvents = True
End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mButtonX = Button

If tEvents = False Then
    tEvents = True
End If

If (mX < mLeft Or mX > mRight) Or (mY < mTop Or mY > mBottom) Then
    Active = False
Else
    If mDbClick = True Then
        UserControl_Click
        mDbClick = False
    End If
End If

Active = False
tEvents = False
ReleaseCapture
RaiseEvent MouseUp(Button, Shift, X, Y)

If tEvents = False Then
    tEvents = True
End If
End Sub

'///Sub para dibujar las fases del botón en el hDC usando el API BitBlt.
Private Sub mDrawButton(mEvent As Events)

mDest = UserControl.hdc
mSrc = Container.hdc
mWidth = Container.Width / 15
mHeight = (Container.Height / 4) / 15

Select Case mEvent

Case 1
BitBlt mDest, 0, 0, mWidth, mHeight, mSrc, 0, mHeight, vbSrcCopy

Case 2
BitBlt mDest, 0, 0, mWidth, mHeight, mSrc, 0, mHeight * 2, vbSrcCopy

Case 3
BitBlt mDest, 0, 0, mWidth, mHeight, mSrc, 0, 0, vbSrcCopy

Case 4
BitBlt mDest, 0, 0, mWidth, mHeight, mSrc, 0, mHeight * 3, vbSrcCopy

End Select

End Sub
'///////////////////////////////////////////////////////////////////////////////

'///Propiedades///
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Devuelve o establece el gráfico que se mostrará en un control."
    Set Picture = Container.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set Container.Picture = New_Picture
    PropertyChanged "Picture"
    tEvents_Timer
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    tEvents_Timer
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    tEvents_Timer
End Property

