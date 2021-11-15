Attribute VB_Name = "modGUILabel"
Public Type guiLabel
    controlName As String
    controlRect As tRectangle

    caption As String
End Type

Public Function guiCreateLabel(ByVal controlName As String, rectangle As tRectangle, Optional ByVal listType As Byte) As guiLabel
    guiCreateLabel.controlName = controlName
    guiCreateLabel.controlRect = rectangle
End Function


Public Function labelCaptionGet(label As guiLabel) As String
    labelCaptionGet = label.caption
End Function

Public Sub labelCaptionSet(label As guiLabel, caption As String)
    label.caption = caption
End Sub

Public Sub labelRender(label As guiLabel, ByVal destX As Integer, ByVal destY As Integer)
    With label
        dxDrawBox .controlRect.x + destX, .controlRect.y + destY, .controlRect.width, .controlRect.height, D3DColorXRGB(0, 150, 0)
        dxTextRender 1, "obj_label: " & .controlName, .controlRect.x + destX, .controlRect.y + destY, D3DColorXRGB(255, 255, 255)
    End With
End Sub
