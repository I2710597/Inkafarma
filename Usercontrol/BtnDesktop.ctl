VERSION 5.00
Begin VB.UserControl BtnDesktop 
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
Attribute VB_Name = "BtnDesktop"
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
'Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
'Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
'Event Paint() 'MappingInfo=UserControl,UserControl,-1,Paint
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
'Default Property Values:
Const m_def_SubCheck = False
'Const m_def_WindowFocus = True
'Property Variables:
Dim m_SubCheck As Boolean
'Dim m_WindowFocus As Boolean


Private Sub tRefresh_Timer()
UserControl.Width = CONT.Width: UserControl.Height = CONT.Height / 3
UserControl_Resize
tRefresh = False
End Sub

Private Sub UserControl_Initialize()
MouseEvent = 1
UserControl_Paint
UserControl.Refresh
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
If SubCheck = True Then
    If KeyAscii = vbKeyReturn Then
        RaiseEvent DblClick
    End If
End If
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
    
    'RaiseEvent MouseMove(Button, Shift, X, Y)
UserControl_Paint
UserControl.Refresh
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UnselectedOthers
If Button = vbLeftButton Then
    MouseTrack = 1
    SubCheck = True
    Activo = True
    MouseEvent = 3
    PBB MouseDown
ElseIf Button = vbRightButton Then
    MouseTrack = 2
    SubCheck = True
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

UserControl.Cls

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
    Select Case MouseEvent
    Case 1
        PBB MouseDown
    Case 2
        PBB MouseHover
    Case 3
        PBB MouseDown
    End Select
End If
UserControl.Refresh
End Sub


Private Sub UserControl_Resize()
On Error Resume Next
If (UserControl.Width < CONT.Width) Or (UserControl.Height < CONT.Height / 3) Or _
(UserControl.Width > CONT.Width) Or (UserControl.Height > CONT.Height / 3) Then
    With UserControl
        .Width = CONT.Width
        .Height = CONT.Height / 3
    End With
Else
    With UserControl
        .Width = CONT.Width
        .Height = CONT.Height / 3
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
hHEIGHT = (CONT.Height / 3) / 15

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

'Case 4
'BitBlt hDESTINO, 0, 0, hWIDTH, hHEIGHT, hORIGEN, 0, hHEIGHT * 3, vbSrcCopy
'UserControl.Refresh
End Select

End Function


Private Sub UserControl_DblClick()
If MouseTrack = 1 Then
    RaiseEvent DblClick
End If
    UserControl_Paint
UserControl.Refresh
MouseTrack = 0
End Sub

Private Sub UnselectedOthers()
UserControl.ParentControls.ParentControlsType = vbExtender
Dim Xu As Integer

For Xu = 0 To UserControl.ParentControls.Count - 1
    If TypeOf UserControl.ParentControls(Xu) Is BtnDesktop Then
        If UserControl.ParentControls(Xu).Name <> UserControl.Ambient.DisplayName Then
            UserControl.ParentControls(Xu).SubCheck = False
            
        End If
    End If
Next
        
        UserControl.Refresh
        UserControl_Paint
End Sub

'
''//////////////////////////////////////////////////////////////////////////////
''//////////////////////////////////////////////////////////////////////////////

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
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("SubCheck", m_SubCheck, m_def_SubCheck)
    
    If Picture > 0 Then
            PBB MouseUp
    If (UserControl.Width < CONT.Width) Or (UserControl.Height < CONT.Height / 3) Or _
(UserControl.Width > CONT.Width) Or (UserControl.Height > CONT.Height / 3) Then
    With UserControl
        .Width = CONT.Width
        .Height = CONT.Height / 3
    End With
Else
    With UserControl
        .Width = CONT.Width
        .Height = CONT.Height / 3
    End With
End If
    Else
    UserControl.Cls
    End If
    UserControl_Paint
    UserControl.Refresh
    

End Sub




'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_SubCheck = PropBag.ReadProperty("SubCheck", m_def_SubCheck)
     
    If Picture > 0 Then
    PBB MouseUp
    Else
    UserControl.Cls
    End If
    
    UserControl_Paint
    UserControl.Refresh
    tRefresh = True


End Sub

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

