VERSION 5.00
Begin VB.UserControl mAniControl 
   BackColor       =   &H000A0909&
   ClientHeight    =   2130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2565
   ClipBehavior    =   0  'None
   ScaleHeight     =   2130
   ScaleWidth      =   2565
   Begin VB.Timer tAniA 
      Enabled         =   0   'False
      Left            =   390
      Top             =   180
   End
   Begin VB.PictureBox ContAni 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   1800
      Left            =   1035
      Negotiate       =   -1  'True
      ScaleHeight     =   1800
      ScaleWidth      =   1350
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   1350
   End
End
Attribute VB_Name = "mAniControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////////////////////////
'///Información//////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////
'///Nombre:         mAniControl
'///Autor:          SONIC88 - Marco A. Olivares Aracena
'///Mail:           sonic88@live.cl
'///Descripción:    Control para mostrar una animación simple mediante el API BitBlt.
'///Distribución:   Distribución libre.
'/////////////////////////////////////////////////////////////////////////////////////


Option Explicit

'///Declaramos la API que hará la pega. jaja :p
Private Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal X As Long, _
    ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long


Dim ActiveAni As Boolean
Dim CurrentButton As Long

Dim xIndex As Long, tIndex As Long

'Event Declarations:

Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
'Default Property Values:
Const m_def_LayerHeight = 400
Const m_def_LockClick = 0
'Property Variables:
Dim m_LayerHeight As Long
Dim m_LockClick As Boolean


Private Sub ContAni_Change()
UserControl_Resize
End Sub

Private Sub UserControl_Initialize()
UserControl_Resize
End Sub

Private Sub UserControl_Paint()
UserControl.Cls

If ActiveAni = False Then
    DrawAni 0
Else
    DrawAni xIndex
End If

End Sub

Private Sub tAniA_Timer()
'tIndice contendrá el total de cuadros de la animación, se calcúla dividiendo
'el el tamaño de la imagen con el alto del cuadro (propiedad LayerHeight).
tIndex = ContAni.Height / LayerHeight

'Se le pasa el índice del cuadro a la Función DrawAni.
DrawAni xIndex

'En caso de que el índice coincida con el total de cuadros, el índice lo hacemos
'volver a 1 para que repita la secuencia y salimos del Sub.
'''If xIndex = tIndex - 1 Then xIndex = 1: Exit Sub

'*****Si deseas ocupar la secuencia completa reemplaza la anterior linea por la que sigue.
If xIndex = tIndex - 1 Then xIndex = 0: Exit Sub

'Mientras el índice no concida con el total de cuadros le seguiremos sumando 1
'al índice.
xIndex = xIndex + 1
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
'Sólo para saber que botón presionamos con el mouse.
CurrentButton = Button
End Sub

Private Sub UserControl_Click()

If LockClick = False Then
    If CurrentButton = 1 Then
        If ActiveAni = False Then
            PlayAni
        Else
            StopAni
        End If
        RaiseEvent Click
    Else
        CurrentButton = 0
    End If
        CurrentButton = 0
End If
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=ContAni,ContAni,-1,Picture
Public Property Get Picture() As StdPicture
Attribute Picture.VB_Description = "Devuelve o establece el gráfico que se mostrará en un control."
    Set Picture = ContAni.Picture
    UserControl_Resize
End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)
    Set ContAni.Picture = New_Picture
    PropertyChanged "Picture"
    
    UserControl_Resize
    UserControl.Refresh
End Property

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    tAniA.Interval = PropBag.ReadProperty("Interval", 0)
    m_LayerHeight = PropBag.ReadProperty("LayerHeight", m_def_LayerHeight)
    m_LockClick = PropBag.ReadProperty("LockClick", m_def_LockClick)
    
    UserControl_Resize
    UserControl.Refresh
    
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
UserControl.Width = ContAni.Width: UserControl.Height = LayerHeight
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Picture", ContAni.Picture, Nothing)
    Call PropBag.WriteProperty("Interval", tAniA.Interval, 0)
    Call PropBag.WriteProperty("LayerHeight", m_LayerHeight, m_def_LayerHeight)
    Call PropBag.WriteProperty("LockClick", m_LockClick, m_def_LockClick)
    
    UserControl_Resize
    UserControl.Refresh

End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=tAniA,tAniA,-1,Interval
Public Property Get Interval() As Long
Attribute Interval.VB_Description = "Devuelve o establece el número de milisegundos entre dos llamadas al evento Timer de un control Timer."
    Interval = tAniA.Interval
End Property

Public Property Let Interval(ByVal New_Interval As Long)
    tAniA.Interval() = New_Interval
    PropertyChanged "Interval"
End Property

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    m_LayerHeight = m_def_LayerHeight
    m_LockClick = m_def_LockClick
    UserControl_Resize
End Sub


'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,0
Public Property Get LockClick() As Boolean
    LockClick = m_LockClick
End Property

Public Property Let LockClick(ByVal New_LockClick As Boolean)
    m_LockClick = New_LockClick
    PropertyChanged "LockClick"
End Property


'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,400
Public Property Get LayerHeight() As Long
    LayerHeight = m_LayerHeight
End Property

Public Property Let LayerHeight(ByVal New_LayerHeight As Long)
    m_LayerHeight = New_LayerHeight
    PropertyChanged "LayerHeight"
End Property

'/////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////
'/// Funciones para la animación /////////////////////////////////////////////

'///Función para iniciar la secuencia.
Public Function PlayAni()
'Ponemos el indice en cero.
    xIndex = 0
'Obtenemos el Intevalo definido por usuario.
    tAniA.Interval = Interval
'Activamos el Timer(tAniA)
    tAniA.Enabled = True
'Volvemos a True a ActiveAni
    ActiveAni = True
'Refrescamos el control.
    UserControl.Refresh
End Function

'///Función para detener la secuencia.
Public Function StopAni()
'Ponemos el indice en cero.
    xIndex = 0
'Desactivamos el Timer.
    tAniA.Enabled = False
'Volvemos a False a ActiveAni
    ActiveAni = False
'Refrescamos el control.
    UserControl.Refresh
End Function

'///Sub que realiza el trabajo de recortar las imágenes de acuerdo
'///al índice que obtenga desde el Timer.
Private Sub DrawAni(nIndex As Long)
'API BitBlt para relizar la función de recortar las imágenes. Las coordenadas X, Y,
'nWidth, nHeight, xSrc, ySrc se deben dividir en 15 porque las API solo usan
'coordenadas expuestas en Píxeles, ya que este control usa la escala en "TWIP".

    BitBlt UserControl.hdc, 0, 0, ContAni.Width / 15, LayerHeight / 15, _
    ContAni.hdc, 0, (LayerHeight * nIndex) / 15, vbSrcCopy

End Sub

'///Función para saber si la secuencia animada esta activa o no.
Public Function GetActiveAni() As Boolean

If ActiveAni = True Then
    GetActiveAni = True
Else
    GetActiveAni = False
End If

End Function

Public Function CurrentFrame() As Long

If ContAni.Picture > 0 Then
    CurrentFrame = xIndex
Else
    CurrentFrame = 0
End If

End Function

Public Function Frames() As Long

If ContAni.Picture > 0 Then
    Frames = ContAni.Height / LayerHeight
Else
    Frames = 0
End If

End Function
