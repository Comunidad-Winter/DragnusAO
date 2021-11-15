Attribute VB_Name = "modGUIAccess"
Public Function guiCreateAccess(ByVal controlName As String, rectangle As tRectangle) As guiAccess
    guiCreateAccess.controlName = controlName
    guiCreateAccess.controlRect = rectangle
End Function


Public Sub accessSet(access As guiAccess, accessType As eAccessType, ByVal sourceForm As Integer, ByVal sourceName As String, ByVal sourceSlot As Integer, ByVal grhIndex As Integer)
    access.used = True
    access.accessType = accessType
    access.sourceForm = sourceForm
    access.sourceName = sourceName
    access.sourceSlot = sourceSlot
    access.grhIndex = grhIndex
End Sub
Public Function accessGrhIndexGet(access As guiAccess) As Integer
    accessGrhIndexGet = access.grhIndex
End Function

Public Function accessUsedGet(access As guiAccess) As Boolean
    accessUsedGet = access.used
End Function

Public Sub accessClear(access As guiAccess)
    access.used = False
    access.sourceForm = 0
    access.sourceName = ""
    access.sourceSlot = 0
    access.grhIndex = 0
End Sub

Public Sub accessErase(access As guiAccess, sourceForm As Integer, sourceName As String, sourceSlot As Integer)
    With access
        If .used Then
            If .sourceForm = sourceForm And .sourceName = sourceName And .sourceSlot = sourceSlot Then
                accessClear access
            End If
        End If
    End With
End Sub

Public Function accessClick(access As guiAccess, ByVal mouseX As Integer, ByVal mouseY As Integer) As Boolean
    With access
        accessClick = rectMouseOver(.controlRect, mouseX, mouseY)
    End With
End Function

Public Sub accessRender(access As guiAccess, ByVal destX As Integer, ByVal destY As Integer)
    With access
        'guiTextureRender 16020, destX + .controlRect.x, destY + .controlRect.y, 48, 48, D3DColorARGB(200, 255, 255, 255)
    
        If .grhIndex > 0 Then
            GUI_Grh_Render .grhIndex, destX + .controlRect.x, destY + .controlRect.y, , , D3DColorARGB(255, 255, 255, 255)
        End If
    End With
End Sub

Public Function accessSlotGet(access As guiAccess) As Integer
    accessSlotGet = access.sourceSlot
End Function

Public Sub accessClicked(access As guiAccess)
    If access.used Then
        If access.accessType = eAccessType.Item Then
            Call EquiparItem(access.sourceSlot)
        ElseIf access.accessType = eAccessType.Spell Then
            actionCast access.sourceSlot
        End If
    End If
End Sub
