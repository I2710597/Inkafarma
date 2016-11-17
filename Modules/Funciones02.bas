Attribute VB_Name = "Funciones02"
Option Explicit

Public Enum StretchType
[Vertical] = 0
[Horizontal] = 1
[Both] = 2
End Enum

Public Sub StretchPicture(mPicture As PictureBox, mStretchType As StretchType, Optional Active As Boolean = True)

    If Active = True Then
        Select Case mStretchType
            Case 0
                mPicture.AutoRedraw = True
                mPicture.PaintPicture mPicture.Picture, 0, 0, , mPicture.ScaleHeight
            Case 1
                mPicture.AutoRedraw = True
                mPicture.PaintPicture mPicture.Picture, 0, 0, mPicture.Width
            Case 2
                mPicture.AutoRedraw = True
                mPicture.PaintPicture mPicture.Picture, 0, 0, mPicture.ScaleWidth, mPicture.ScaleHeight
            End Select
    Else
        mPicture.AutoRedraw = False
    End If

End Sub
