Attribute VB_Name = "modGUIPicture"
Option Explicit

Public Type guiPicture
    controlName As String
    controlRect As tRectangle

    textureIndex As Integer
    dragEnable As Boolean
End Type

Public Function guiCreatePicture(ByVal controlName As String, rectangle As tRectangle, textureIndex As Integer, Optional ByVal dragEnable As Boolean = False) As guiPicture
    guiCreatePicture.controlRect = rectangle
    guiCreatePicture.controlName = controlName
    guiCreatePicture.textureIndex = textureIndex
    guiCreatePicture.dragEnable = dragEnable
End Function

Public Sub pictureRender(picture As guiPicture, ByVal destX As Integer, ByVal destY As Integer)
    With picture
        guiTextureRender .textureIndex, destX + .controlRect.x, destY + .controlRect.y, .controlRect.width, .controlRect.height, D3DColorARGB(255, 255, 255, 255)
    End With
End Sub

Public Sub pictureWidthSet(picture As guiPicture, ByVal width As Integer)
    picture.controlRect.width = width
End Sub

Public Function pictureDrag(picture As guiPicture) As Boolean
    pictureDrag = picture.dragEnable
End Function

