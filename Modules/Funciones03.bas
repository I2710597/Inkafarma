Attribute VB_Name = "Funciones03"
Option Explicit

Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
'/////////////////////////////////////////////////////////////////////////////////////////
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
'/////////////////////////////////////////////////////////////////////////////////////////
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'/////////////////////////////////////////////////////////////////////////////////////////
Private Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
'/////////////////////////////////////////////////////////////////////////////////////////
Private Declare Function CreateRoundRectRgn Lib "gdi32" ( _
    ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long, _
    ByVal X3 As Long, ByVal Y3 As Long) As Long

Private Declare Function SetWindowRgn Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hRgn As Long, _
    ByVal bRedraw As Boolean) As Long
'/////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////

Public Function Redondear(Objeto As Object, Radio As Long)

    Dim Region As Long
    Dim Ret As Long
    Dim Ancho As Long
    Dim Alto As Long

    'Obtenemos el ancho y alto de la region del Form
    Ancho = (Objeto.Width + 15) / Screen.TwipsPerPixelX
    Alto = (Objeto.Height + 15) / Screen.TwipsPerPixelY

    'Le pasamos el ancho alto del formualrio y el valor de _
    redondeo es decir el radio

    Region = CreateRoundRectRgn(0, 0, Ancho, Alto, Radio, Radio)

    'Aplica la región al formulario
    Ret = SetWindowRgn(Objeto.hwnd, Region, True)

End Function

Public Function MoverObjeto(Objeto As Object)
    Call ReleaseCapture
    Call SendMessage(Objeto.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Function

Public Function OpenEXE(ByVal Exefile As String) As Long

    OpenEXE = ShellExecute(wMain.hwnd, vbNullString, Exefile, vbNullString, "%SYSTEMROOT%", SW_SHOWNORMAL)
    'OpenEXE = WinExec(Exefile, SW_SHOWNORMAL)

End Function
