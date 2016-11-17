VERSION 5.00
Begin VB.UserControl BtnMenu 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1350
   ScaleWidth      =   2400
   Begin VB.Timer tRefresh 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   285
      Top             =   735
   End
   Begin VB.PictureBox CONT 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   2
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   13322
         SubFormatType   =   9
      EndProperty
      Enabled         =   0   'False
      Height          =   1185
      Left            =   765
      ScaleHeight     =   1185
      ScaleWidth      =   1530
      TabIndex        =   0
      Top             =   105
      Visible         =   0   'False
      Width           =   1530
   End
End
Attribute VB_Name = "BtnMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, _
ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'--------------------------------- MouseOver ------------------------------------'
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long

Private Enum Eventos
MouseHover = 1
MouseDown = 2
MouseUp = 3
Disabled = 4
End Enum

Private MouseEvent As Long
Private MouseTrack As Long
Private Activo As Boolean


'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Ocurre cuando el usuario presiona y libera un botón del mouse encima de un objeto."
'Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Ocurre cuando el usuario libera el botón del mouse mientras un objeto tiene el enfoque."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Ocurre cuando el usuario mueve el mouse."
Event Paint() 'MappingInfo=UserControl,UserControl,-1,Paint
Attribute Paint.VB_Description = "Ocurre cuando se mueve, se amplía o se expone cualquier parte de un formulario o un control PictureBox."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Ocurre cuando el usuario presiona el botón del mouse mientras un objeto tiene el enfoque."
'Default Property Values:
Const m_def_SubCheck = False
'Const m_def_WindowFocus = True
'Property Variables:
Dim m_SubCheck As Boolean
'Dim m_WindowFocus As Boolean


Private Sub tRefresh_Timer()
UserControl.Width = CONT.Width: UserControl.Height = CONT.Height / 4
UserControl_Resize
tRefresh = False
End Sub

Private Sub UserControl_Click()
If MouseTrack = 1 Then
    RaiseEvent Click
End If
    UserControl_Paint
UserControl.Refresh
MouseTrack = 0
End Sub

Public Property Let Enabled(ByVal New_Enabled As Boolean)
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    
    If UserControl.Enabled = False Then
    PBB Disabled
    Else
    PBB MouseUp
    End If
        UserControl_Paint
    UserControl.Refresh
End Property

Private Sub UserControl_Initialize()
MouseEvent = 1
UserControl_Paint
UserControl.Refresh
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseEvent = 1
    PBB MouseUp
Activo = False

RaiseEvent MouseUp(Button, Shift, X, Y)
ReleaseCapture
UserControl_Paint
UserControl.Refresh

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Activo = False Then
    If (X < 0 Or X > UserControl.Width) Or (Y < 0 Or Y > UserControl.Height) Then
       Call ReleaseCapture
        MouseEvent = 1
            PBB MouseUp
    ElseIf UserControl.hwnd <> GetCapture Then
        ReleaseCapture
       Call SetCapture(UserControl.hwnd)
        MouseEvent = 2
        PBB MouseHover
    End If
End If
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
UserControl_Paint
UserControl.Refresh
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    MouseTrack = 1
    Activo = True
    MouseEvent = 3
    PBB MouseDown
ElseIf Button = vbRightButton Then
    MouseTrack = 2
    Activo = False
    MouseEvent = 2
ElseIf Button = vbMiddleButton Then
    MouseTrack = 4
    Activo = False
    MouseEvent = 2
End If
    
RaiseEvent MouseDown(Button, Shift, X, Y)
UserControl_Paint
UserControl.Refresh
End Sub


Private Sub UserControl_Paint()

RaiseEvent Paint

UserControl.Cls

'If Not UserControl.Enabled = False Then

If SubCheck = False Then
    Select Case MouseEvent
    Case 1
        PBB MouseUp
    Case 2
        PBB MouseHover
    Case 3
        PBB MouseDown
    End Select
Else
    PBB MouseDown
End If
    
'End If


If UserControl.Enabled = False Then
    PBB Disabled
End If

UserControl.Refresh
End Sub


Private Sub UserControl_Resize()
On Error Resume Next
If (UserControl.Width < CONT.Width) Or (UserControl.Height < CONT.Height / 3) Or _
(UserControl.Width > CONT.Width) Or (UserControl.Height > CONT.Height / 3) Then
    With UserControl
        .Width = CONT.Width
        .Height = CONT.Height / 4
    End With
Else
    With UserControl
        .Width = CONT.Width
        .Height = CONT.Height / 4
    End With
End If
UserControl_Paint
UserControl.Refresh
End Sub

Private Function PBB(hEvento As Eventos)

Dim hHEIGHT As Long, hWIDTH As Long, hORIGEN As Long, hDESTINO As Long

hORIGEN = CONT.hDC
hDESTINO = UserControl.hDC
hWIDTH = CONT.Width / 15
hHEIGHT = (CONT.Height / 4) / 15

Select Case hEvento

Case 1
BitBlt hDESTINO, 0, 0, hWIDTH, hHEIGHT, hORIGEN, 0, hHEIGHT, vbSrcCopy
UserControl.Refresh

Case 2
BitBlt hDESTINO, 0, 0, hWIDTH, hHEIGHT, hORIGEN, 0, hHEIGHT * 2, vbSrcCopy
UserControl.Refresh

Case 3
BitBlt hDESTINO, 0, 0, hWIDTH, hHEIGHT, hORIGEN, 0, 0, vbSrcCopy
UserControl.Refresh

Case 4
BitBlt hDESTINO, 0, 0, hWIDTH, hHEIGHT, hORIGEN, 0, hHEIGHT * 3, vbSrcCopy
UserControl.Refresh
End Select

End Function

'//////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
    
    On Error Resume Next
    If UserControl.Enabled = False Then
    PBB Disabled
    Else
    PBB MouseUp
    End If
        UserControl_Paint
    UserControl.Refresh
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
    hDC = UserControl.hDC
        UserControl.Refresh
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
        UserControl_Paint
        UserControl.Refresh
End Property

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
m_SubCheck = m_def_SubCheck

UserControl_Paint
UserControl.Refresh
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("ToolTipText", CONT.ToolTipText, "")
    
    If Picture > 0 Then
            PBB MouseUp
    If (UserControl.Width < CONT.Width) Or (UserControl.Height < CONT.Height / 3) Or _
(UserControl.Width > CONT.Width) Or (UserControl.Height > CONT.Height / 3) Then
    With UserControl
        .Width = CONT.Width
        .Height = CONT.Height / 4
    End With
Else
    With UserControl
        .Width = CONT.Width
        .Height = CONT.Height / 4
    End With
End If
    Else
    UserControl.Cls
    End If
    UserControl_Paint
    UserControl.Refresh
    


    Call PropBag.WriteProperty("SubCheck", m_SubCheck, m_def_SubCheck)
End Sub




'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    CONT.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    
    If Picture > 0 Then
    PBB MouseUp
    Else
    UserControl.Cls
    End If
    
    UserControl_Paint
    UserControl.Refresh
    tRefresh = True

    m_SubCheck = PropBag.ReadProperty("SubCheck", m_def_SubCheck)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,HasDC
Public Property Get HasDC() As Boolean
Attribute HasDC.VB_Description = "Determina si hay asignado un contexto de presentación único para el control."
    HasDC = UserControl.HasDC
    UserControl_Paint
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,PaintPicture
Public Sub PaintPicture(ByVal Picture As Picture, ByVal X1 As Single, ByVal Y1 As Single, ByVal Width1 As Variant, ByVal Height1 As Variant, ByVal X2 As Variant, ByVal Y2 As Variant, ByVal Width2 As Variant, ByVal Height2 As Variant, ByVal Opcode As Variant)
Attribute PaintPicture.VB_Description = "Dibuja el contenido de un archivo de gráficos en un objeto Form, PictureBox o Printer."
    UserControl.PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, Opcode
    UserControl_Paint
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=CONT,CONT,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Devuelve o establece el texto mostrado cuando el mouse se sitúa sobre un control."
    ToolTipText = CONT.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    CONT.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
    UserControl.Refresh
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=CONT,CONT,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Devuelve o establece el gráfico que se mostrará en un control."
    Set Picture = CONT.Picture
        tRefresh = True
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set CONT.Picture = New_Picture
    PropertyChanged "Picture"
        UserControl.Refresh
        UserControl_Paint
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Obliga a volver a dibujar un objeto."
    UserControl.Refresh
    CONT.Refresh
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,False
Public Property Get SubCheck() As Boolean
    SubCheck = m_SubCheck
End Property

Public Property Let SubCheck(ByVal New_SubCheck As Boolean)
    m_SubCheck = New_SubCheck
    PropertyChanged "SubCheck"
    
    
    
        UserControl.Refresh
        UserControl_Paint
End Property

