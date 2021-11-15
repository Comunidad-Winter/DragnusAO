Attribute VB_Name = "modGUIButton"
Option Explicit

Private Const BTN_LEFT As Integer = 16009
Private Const BTN_MID As Integer = 16010
Private Const BTN_RIGHT As Integer = 16011

Public Type guiButton
    controlName As String
    controlRect As tRectangle
    normalTextureIndex As Integer
    rolloverTextureIndex As Integer
End Type


Public Function guiCreateButton(ByVal controlName As String, rectangle As tRectangle) As guiButton
    With guiCreateButton
        .controlName = controlName
        .controlRect = rectangle
    End With
End Function

Public Function buttonClick(Button As guiButton, ByVal mouseX As Integer, ByVal mouseY As Integer) As Boolean
    With Button
        buttonClick = rectMouseOver(.controlRect, mouseX, mouseY)
    End With
End Function

Public Sub buttonRender(Button As guiButton, ByVal destX As Integer, ByVal destY As Integer)
    With Button
        'dxDrawBox .controlRect.x + destX, .controlRect.y + destY, .controlRect.width, .controlRect.height, D3DColorXRGB(200, 150, 0)
        guiTextureRender .normalTextureIndex, destX + .controlRect.x, destY + .controlRect.y, .controlRect.width, .controlRect.height, D3DColorXRGB(255, 255, 255)
        'dxTextRender 1, "obj_button: " & .controlName, destX + .controlRect.x, destY + .controlRect.y, D3DColorXRGB(255, 255, 255)
        'guiTextureRender 16000, .controlRect.x + destX, .controlRect.y + destY, 120, 30, D3DColorXRGB(255, 255, 255)
    End With
End Sub
