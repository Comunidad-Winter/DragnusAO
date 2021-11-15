Attribute VB_Name = "modGUIList"
Option Explicit

Private Const DegreeToRadian As Single = 0.0174532925

Private Const LISTARROWSIZE As Byte = 14
Private Const LISTITEMHEIGHT As Byte = 16
Private Const LISTARROWGRH As Long = 16002
Private Const LISTCHARWIDTH As Byte = 8

Public Type listItem
    item As String
    color As Long
End Type

Public Type guiList
    controlName As String
    controlRect As tRectangle
    
    List() As listItem
    listIndex As Integer
    listSize As Integer
    firstLine As Integer
    listType As Byte
    maxLines As Integer
End Type

Public Function guiCreateList(ByVal controlName As String, rectangle As tRectangle) As guiList
    guiCreateList.controlName = controlName
    guiCreateList.controlRect = rectangle
    guiCreateList.listIndex = 0
    guiCreateList.firstLine = 1
    guiCreateList.maxLines = 200
End Function

Public Function listRender(List As guiList, ByVal destX As Integer, ByVal destY As Integer)
    Dim i As Integer
    With List
        Dim boxRect As tRectangle
        boxRect.x = destX + .controlRect.x - 4
        boxRect.y = destY + .controlRect.y - 4
        
        boxRect.width = .controlRect.width + 7
        boxRect.height = (.controlRect.height \ LISTITEMHEIGHT) * 16 + 7
        
        guiDrawBox boxRect, 1
        
        'guiDrawBox boxRect, 0
        i = .firstLine
        Do While i < .firstLine + .controlRect.height \ LISTITEMHEIGHT And i <= .listSize
            If .listIndex > 0 And .listIndex = i Then
                guiTextureRender 16020, destX + .controlRect.x, destY + .controlRect.y + (i - .firstLine) * 16, .controlRect.width, 16, D3DColorXRGB(255, 255, 255)
            Else
                guiTextureRender 16017, destX + .controlRect.x, destY + .controlRect.y + (i - .firstLine) * 16, .controlRect.width, 16, D3DColorXRGB(255, 255, 255)
            End If
            dxTextRender 2, .List(i).item, destX + .controlRect.x, destY + .controlRect.y + (i - .firstLine) * 16, List.List(i).color
            i = i + 1
        Loop
    End With
    
End Function

Public Function listClick(List As guiList, ByVal mouseX As Integer, ByVal mouseY As Integer) As Boolean
    With List
        listClick = rectMouseOver(.controlRect, mouseX, mouseY)
    End With
End Function

Public Sub listAddItem(List As guiList, ByVal newItem As String, Optional ByVal color As Long)
    Dim i As Integer
    With List
        If .listSize < .maxLines Then
            .listSize = .listSize + 1
            ReDim Preserve .List(1 To .listSize)
            .List(.listSize).item = newItem
            .List(.listSize).color = color
        End If
    End With
End Sub

Private Function listArrowUp(List As guiList)
    With List
        If .listSize > .controlRect.height / LISTITEMHEIGHT Then
            If .firstLine > 1 Then
                .firstLine = .firstLine - 1
            End If
        End If
    End With
End Function

Private Function listArrowDown(List As guiList)
    With List
        If .listSize > .controlRect.height / LISTITEMHEIGHT Then
            If .firstLine + .controlRect.height / LISTITEMHEIGHT < .listSize Then
                .firstLine = .firstLine + 1
            End If
        End If
    End With
End Function

Private Function listItemSelect(List As guiList, listY As Integer)
    With List
        .listIndex = .firstLine + listY \ LISTITEMHEIGHT
    End With
End Function

Public Function listClicked(List As guiList, mouseX As Integer, mouseY As Integer)
    With List
        If mouseX > .controlRect.x + .controlRect.width - LISTARROWSIZE And mouseX < .controlRect.x + .controlRect.width Then
            If mouseY > .controlRect.y And mouseY < .controlRect.y + LISTARROWSIZE Then
                listArrowUp List
            ElseIf mouseY > .controlRect.y + .controlRect.height - LISTARROWSIZE And mouseY < .controlRect.y + .controlRect.height Then
                listArrowDown List
            End If
        Else
            listItemSelect List, mouseY - (.controlRect.y)
        End If
    End With
End Function

Public Function listIndexGet(List As guiList)
    listIndexGet = List.listIndex
End Function

Public Sub listClear(List As guiList)
    With List
        ReDim .List(1 To 1) As listItem
        .firstLine = 1
        .listIndex = 0
    End With
End Sub
