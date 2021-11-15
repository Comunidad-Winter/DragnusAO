Attribute VB_Name = "modGUIElement"
Option Explicit

Private Const TEXTURE_SIZE As Integer = 32

Private Const HORIZONTAL As Integer = 16004
Private Const VERTICAL As Integer = 16005
Private Const BACK As Integer = 16006

Private Const BTN_CLOSE As Integer = 16031

Public Enum eControlType
    guiList
    guiInventory
    guiTextbox
    guiScrollBar
    guiLabel
    guiButton
    guiPicture
    guiAccess
End Enum

Public Type guiElement
    windowRect As tRectangle
    
    windowElement As Boolean
    locked As Boolean
    renderBox As Boolean
    
    backGround As Integer
    
    listCount As Integer
    buttonCount As Integer
    pictureCount As Integer
    textboxCount As Integer
    inventoryCount As Integer
    labelCount As Integer
    
    focusedTextBox As Integer
    
    label() As guiLabel
    List() As guiList
    Button() As guiButton
    picture() As guiPicture
    textbox() As guiTextbox
    inventory() As guiInventory
End Type


Public Function formFocusGet(element As guiElement) As Boolean
    If element.focusedTextBox > 0 Then
        formFocusGet = True
    End If
End Function

Public Sub formFocusChange(element As guiElement)
    Dim i As Byte
    With element
        If .textboxCount Then
        
            i = .focusedTextBox + 1
            
            Do While i <= .textboxCount
                If .textbox(i).visible Then
                    .focusedTextBox = i
                    Exit Sub
                End If
                i = i + 1
            Loop
            
            
            i = 1
            Do While i < .focusedTextBox
                If .textbox(i).visible Then
                    .focusedTextBox = i
                    Exit Sub
                End If
                i = i + 1
            Loop
            
            .focusedTextBox = 0
            
        End If
    End With
End Sub

Public Sub formAddTextbox(element As guiElement, ByVal controlName As String, rectangle As tRectangle, Optional ByVal visible As Boolean = True)
    With element
        .textboxCount = .textboxCount + 1
        ReDim Preserve .textbox(1 To .textboxCount)
        
        .textbox(.textboxCount) = guiCreateTextbox(controlName, rectangle, visible)
    End With
End Sub

Public Sub formAddButton(element As guiElement, ByVal controlName As String, rectangle As tRectangle, ByVal normalTextureIndex As Integer, ByVal rolloverTextureIndex As Integer)
    With element
        .buttonCount = .buttonCount + 1
        ReDim Preserve .Button(1 To .buttonCount)
        
        .Button(.buttonCount) = guiCreateButton(controlName, rectangle)
        .Button(.buttonCount).normalTextureIndex = normalTextureIndex
        .Button(.buttonCount).rolloverTextureIndex = rolloverTextureIndex
    End With
End Sub

Public Sub formAddInventory(element As guiElement, ByVal controlName As String, rectangle As tRectangle, ByVal slots As Integer, Optional ByVal dragEnable As Boolean)
    With element
        .inventoryCount = .inventoryCount + 1
        ReDim Preserve .inventory(1 To .inventoryCount)
        
        .inventory(.inventoryCount) = guiCreateInventory(controlName, rectangle, slots, dragEnable)
    End With
End Sub

Public Sub formAddList(element As guiElement, ByVal controlName As String, rectangle As tRectangle)
    With element
        .listCount = .listCount + 1
        ReDim Preserve .List(1 To .listCount)
        .List(.listCount) = guiCreateList(controlName, rectangle)
    End With
End Sub

Public Sub formAddPicture(element As guiElement, ByVal controlName As String, rectangle As tRectangle, textureIndex As Integer, Optional ByVal dragEnable As Boolean)
    With element
        .pictureCount = .pictureCount + 1
        ReDim Preserve .picture(1 To .pictureCount)
        
        .picture(.pictureCount) = guiCreatePicture(controlName, rectangle, textureIndex, dragEnable)
    End With
End Sub

Public Sub formAddLabel(element As guiElement, ByVal controlName As String, rectangle As tRectangle)
    With element
        .labelCount = .labelCount + 1
        ReDim Preserve .label(1 To .labelCount)
        
        .label(.labelCount) = guiCreateLabel(controlName, rectangle)
    End With
End Sub

Public Sub formBoxRender(element As guiElement)
    With element
        guiDrawBox .windowRect, 0
        
        If Not element.locked Then
            'guiTextureRender BTN_CLOSE, .windowRect.x + .windowRect.width - 20, .windowRect.y + 3, 32, 32, D3DColorARGB(255, 255, 255, 255)
        End If
    End With
End Sub

Public Function formCloseClick(element As guiElement, ByVal mouseX As Integer, ByVal mouseY As Integer) As Boolean
    With element
        formMousePosGet element, mouseX, mouseY
    
        If mouseX > .windowRect.width - 20 And mouseX < .windowRect.width Then
            If mouseY > 0 And mouseY < 20 Then
                formCloseClick = True
            End If
        End If
    End With
End Function

Public Sub formRender(element As guiElement)
    Dim i As Integer
    
'    dxDrawBox element.windowRect.x, element.windowRect.y, element.windowRect.width, element.windowRect.height, D3DColorARGB(150, 100, 100, 100)
    If element.renderBox = False Then
        formBoxRender element
        guiTextureRender 16018, element.windowRect.x + 3, element.windowRect.y + 3, 128, 128, D3DColorXRGB(255, 255, 255)
    End If
    
    
    
    If element.backGround > 0 Then _
        guiTextureRender element.backGround, element.windowRect.x, element.windowRect.y, element.windowRect.width, element.windowRect.height, D3DColorXRGB(255, 255, 255)

    
    If element.pictureCount > 0 Then
        For i = 1 To element.pictureCount
            pictureRender element.picture(i), element.windowRect.x, element.windowRect.y
        Next i
    End If
    
    If element.buttonCount > 0 Then
        For i = 1 To element.buttonCount
            buttonRender element.Button(i), element.windowRect.x, element.windowRect.y
        Next i
    End If
    
    If element.listCount > 0 Then
        For i = 1 To element.listCount
            listRender element.List(i), element.windowRect.x, element.windowRect.y
        Next i
    End If
    
    If element.textboxCount > 0 Then
        For i = 1 To element.textboxCount
            textboxRender element.textbox(i), element.windowRect.x, element.windowRect.y
        Next i
    End If
    
    If element.inventoryCount > 0 Then
        For i = 1 To element.inventoryCount
            inventoryRender element.inventory(i), element.windowRect.x, element.windowRect.y
        Next i
    End If

End Sub

Public Function getControlID(element As guiElement, ByVal controlName As String, controlType As eControlType) As Integer
    Dim i As Integer
    With element
        Select Case controlType
            Case eControlType.guiButton
                Do While i < .buttonCount And getControlID < 1
                    i = i + 1
                    If .Button(i).controlName = controlName Then
                        getControlID = i
                    End If
                Loop
            Case eControlType.guiInventory
                Do While i < .inventoryCount And getControlID < 1
                    i = i + 1
                    If .inventory(i).controlName = controlName Then
                        getControlID = i
                    End If
                Loop
            Case eControlType.guiList
                Do While i < .listCount And getControlID < 1
                    i = i + 1
                    If .List(i).controlName = controlName Then
                        getControlID = i
                    End If
                Loop
            Case eControlType.guiTextbox
                Do While i < .textboxCount And getControlID < 1
                    i = i + 1
                    If .textbox(i).controlName = controlName Then
                        getControlID = i
                    End If
                Loop
            Case eControlType.guiPicture
                Do While i < .pictureCount And getControlID < 1
                    i = i + 1
                    If .picture(i).controlName = controlName Then
                        getControlID = i
                    End If
                Loop
            Case eControlType.guiLabel
                Do While i < .labelCount And getControlID < 1
                    i = i + 1
                    If .label(i).controlName = controlName Then
                        getControlID = i
                    End If
                Loop
        End Select
    End With
End Function

Public Function formClick(element As guiElement, ByVal mouseX As Integer, ByVal mouseY As Integer) As Boolean
    With element
        formClick = rectMouseOver(.windowRect, mouseX, mouseY)
    End With
End Function

Public Function formBarClick(element As guiElement, ByVal mouseX As Integer, ByVal mouseY As Integer)
    With element
        If mouseX + .windowRect.x > .windowRect.x And mouseX + .windowRect.x < .windowRect.x + .windowRect.width Then
            If mouseY + .windowRect.y > .windowRect.y And mouseY + .windowRect.y < .windowRect.y + 18 Then
                formBarClick = True
            End If
        End If
    End With
End Function

Public Function formButtonClick(element As guiElement, ByVal mouseX As Integer, ByVal mouseY As Integer) As String
    Dim i As Integer
    
    formButtonClick = ""
    
    Do While i < element.buttonCount And formButtonClick = ""
        i = i + 1
        With element
            If buttonClick(.Button(i), mouseX, mouseY) Then
                formButtonClick = .Button(i).controlName
            End If
        End With
    Loop
End Function

Public Function formListClick(element As guiElement, ByVal mouseX As Integer, ByVal mouseY As Integer) As String
    Dim i As Integer
    
    formListClick = ""
    
    Do While i < element.listCount And formListClick = ""
        i = i + 1
        With element
            If listClick(.List(i), mouseX, mouseY) Then
                formListClick = .List(i).controlName
            End If
        End With
    Loop
End Function

Public Function formTextboxClick(element As guiElement, ByVal mouseX As Integer, ByVal mouseY As Integer) As String
    Dim i As Integer
    
    formTextboxClick = ""
    
    Do While i < element.textboxCount And formTextboxClick = ""
        i = i + 1
        With element
            If textboxClick(.textbox(i), mouseX, mouseY) Then
                formTextboxClick = .textbox(i).controlName
            End If
        End With
    Loop
End Function

Public Function formInventoryClick(element As guiElement, ByVal mouseX As Integer, ByVal mouseY As Integer) As String
    Dim i As Integer
    
    formInventoryClick = ""
    
    Do While i < element.inventoryCount And formInventoryClick = ""
        i = i + 1
        With element
            If inventoryClick(.inventory(i), mouseX, mouseY) Then
                formInventoryClick = .inventory(i).controlName
            End If
        End With
    Loop
End Function

Public Function formKeyPress(element As guiElement, ByVal KeyAscii As Byte) As Boolean
    Select Case KeyAscii
        Case 9 'TAB
            formFocusChange element
            formKeyPress = True
        Case Else
            If element.focusedTextBox > 0 Then _
                formKeyPress = textboxKeyPress(element.textbox(element.focusedTextBox), KeyAscii)
    End Select
End Function

Public Sub formMousePosGet(element As guiElement, mouseX As Integer, mouseY As Integer)
    mouseX = mouseX - element.windowRect.x
    mouseY = mouseY - element.windowRect.y
End Sub

Public Function formControlFocusGet(element As guiElement) As Integer
    formControlFocusGet = element.focusedTextBox
End Function

Public Function formControlFocusSet(element As guiElement, ByVal controlName As String) As Integer
    element.focusedTextBox = getControlID(element, controlName, eControlType.guiTextbox)
End Function


'OBJECTS PROPERTIES
'INVENTORY
Public Function formInventoryItemGet(element As guiElement, ByVal controlName As String) As Integer
    formInventoryItemGet = inventoryItemGet(element.inventory(getControlID(element, controlName, eControlType.guiInventory)))
End Function

Public Function formInventoryInvGet(element As guiElement, ByVal controlName) As inventory
    formInventoryInvGet = inventoryInvGet(element.inventory(getControlID(element, controlName, eControlType.guiInventory)))
End Function

Public Function formInventoryItemSet(element As guiElement, ByVal controlName As String, ByVal slot As Byte, ByVal objIndex As Integer, ByVal amount As Integer, ByVal equipped As Byte, ByVal grhIndex As Integer, ByVal objType As Integer, ByVal maxHit As Integer, ByVal minHit As Integer, ByVal def As Integer, ByVal valor As Long, ByVal name As String) As Integer
    inventoryItemSet element.inventory(getControlID(element, controlName, eControlType.guiInventory)), slot, objIndex, amount, equipped, grhIndex, objType, maxHit, minHit, def, valor, name
End Function

Public Sub formInventoryGLDSelect(element As guiElement, ByVal controlName As String)
    inventoryGLDSelect element.inventory(getControlID(element, controlName, eControlType.guiInventory))
End Sub

Public Sub formInventoryScroll(element As guiElement, ByVal controlName As String, ByVal up As Boolean)
    inventoryScroll element.inventory(getControlID(element, controlName, eControlType.guiInventory)), up
End Sub

Public Function formInventoryAmountGet(element As guiElement, ByVal controlName As String, ByVal slot As Byte) As Integer
    formInventoryAmountGet = inventoryAmountGet(element.inventory(getControlID(element, controlName, eControlType.guiInventory)), slot)
End Function

Public Function formInventoryGrhIndexGet(element As guiElement, ByVal controlName As String, ByVal slot As Byte) As Integer
    formInventoryGrhIndexGet = inventoryGrhIndexGet(element.inventory(getControlID(element, controlName, eControlType.guiInventory)), slot)
End Function

Public Function formInventoryMaxHitGet(element As guiElement, ByVal controlName As String, ByVal slot As Byte) As Integer
    formInventoryMaxHitGet = inventoryMaxHitGet(element.inventory(getControlID(element, controlName, eControlType.guiInventory)), slot)
End Function

Public Function formInventoryMinHitGet(element As guiElement, ByVal controlName As String, ByVal slot As Byte) As Integer
    formInventoryMinHitGet = inventoryMinHitGet(element.inventory(getControlID(element, controlName, eControlType.guiInventory)), slot)
End Function

Public Function formInventoryValueGet(element As guiElement, ByVal controlName As String, ByVal slot As Byte) As Long
    formInventoryValueGet = inventoryValueGet(element.inventory(getControlID(element, controlName, eControlType.guiInventory)), slot)
End Function

Public Function formInventoryNameGet(element As guiElement, ByVal controlName As String, ByVal slot As Byte) As String
    formInventoryNameGet = inventoryNameGet(element.inventory(getControlID(element, controlName, eControlType.guiInventory)), slot)
End Function

Public Function formInventoryEquipGet(element As guiElement, ByVal controlName As String, ByVal slot As Byte) As Byte
    formInventoryEquipGet = inventoryEquipGet(element.inventory(getControlID(element, controlName, eControlType.guiInventory)), slot)
End Function

Public Function formInventoryObjIndexGet(element As guiElement, ByVal controlName As String, ByVal slot As Byte) As Integer
    formInventoryObjIndexGet = inventoryObjIndexGet(element.inventory(getControlID(element, controlName, eControlType.guiInventory)), slot)
End Function

Public Function formInventoryObjTypeGet(element As guiElement, ByVal controlName As String, ByVal slot As Byte) As Integer
    formInventoryObjTypeGet = inventoryObjTypeGet(element.inventory(getControlID(element, controlName, eControlType.guiInventory)), slot)
End Function

Public Sub formTextboxTextSet(element As guiElement, controlName As String, ByVal Text As String)
    textboxSet element.textbox(getControlID(element, controlName, eControlType.guiTextbox)), Text
End Sub
'TEXTBOX
Public Sub formTextboxShow(element As guiElement, ByVal controlName As String)
    textboxShow element.textbox(getControlID(element, controlName, eControlType.guiTextbox))
    formControlFocusSet element, controlName
End Sub

Public Sub formTextboxHide(element As guiElement, ByVal controlName As String)
    textboxHide element.textbox(getControlID(element, controlName, eControlType.guiTextbox))
    formFocusChange element
End Sub

Public Sub formTextboxClear(element As guiElement, ByVal controlName As String)
    textboxClear element.textbox(getControlID(element, controlName, eControlType.guiTextbox))
End Sub

Public Function formTextboxTextGet(element As guiElement, ByVal controlName As String) As String
    formTextboxTextGet = textboxTextGet(element.textbox(getControlID(element, controlName, eControlType.guiTextbox)))
End Function

Public Function formTextboxVisibleGet(element As guiElement, ByVal controlName As String) As Boolean
    formTextboxVisibleGet = textboxVisibleGet(element.textbox(getControlID(element, controlName, eControlType.guiTextbox)))
End Function

Public Function formLabelCaptionGet(element As guiElement, ByVal controlName As String) As String
    formLabelCaptionGet = labelCaptionGet(element.label(getControlID(element, controlName, eControlType.guiLabel)))
End Function

Public Sub formLabelCaptionSet(element As guiElement, ByVal controlName As String, caption As String)
    labelCaptionSet element.label(getControlID(element, controlName, eControlType.guiLabel)), caption
End Sub

'LIST
Public Sub formListAddItem(element As guiElement, ByVal controlName As String, ByVal newItem As String, Optional ByVal color As Long)
    listAddItem element.List(getControlID(element, controlName, eControlType.guiList)), newItem, color
End Sub
Public Sub formListClear(element As guiElement, ByVal controlName As String)
    listClear element.List(getControlID(element, controlName, eControlType.guiList))
End Sub

Public Function formListIndexGet(element As guiElement, ByVal controlName) As Integer
    formListIndexGet = listIndexGet(element.List(getControlID(element, controlName, eControlType.guiList)))
End Function

Public Sub formPictureWidthSet(element As guiElement, controlName As String, ByVal width As Integer)
    pictureWidthSet element.picture(getControlID(element, controlName, eControlType.guiPicture)), width
End Sub

Public Function formPictureDrag(element As guiElement, controlName As String) As Boolean
    formPictureDrag = pictureDrag(element.picture(getControlID(element, controlName, eControlType.guiPicture)))
End Function

Public Function formInventoryDrag(element As guiElement, controlName As String) As Boolean
    formInventoryDrag = inventoryDrag(element.inventory(getControlID(element, controlName, eControlType.guiInventory)))
End Function

Public Sub formPosSet(element As guiElement, ByVal x As Integer, ByVal y As Integer)
    element.windowRect.x = x
    element.windowRect.y = y
End Sub

Public Function formControlClick(element As guiElement, ByVal mouseX As Integer, ByVal mouseY As Integer, controlName As String, controlType As eControlType) As Boolean
    With element
        formMousePosGet element, mouseX, mouseY
    
        'Check buttons
        controlName = formButtonClick(element, mouseX, mouseY)
        If controlName <> "" Then
            controlType = eControlType.guiButton
            formControlClick = True
            Exit Function
        End If
        
        'Check lists
        controlName = formListClick(element, mouseX, mouseY)
        If controlName <> "" Then
            controlType = eControlType.guiList
            formControlClick = True
            Exit Function
        End If
            
        'Check textbox
        controlName = formTextboxClick(element, mouseX, mouseY)
        If controlName <> "" Then
            controlType = eControlType.guiTextbox
            formControlClick = True
            Exit Function
        End If
        
        'Check Inventory
        controlName = formInventoryClick(element, mouseX, mouseY)
        If controlName <> "" Then
            controlType = eControlType.guiInventory
            formControlClick = True
            Exit Function
        End If
    End With
End Function

Public Function formClicked(element As guiElement, ByVal mouseX As Integer, ByVal mouseY As Integer, controlName As String, controlType As eControlType) As Boolean

    formControlFocusSet element, ""
    formMousePosGet element, mouseX, mouseY
    
    With element
        If controlName <> "" Then
            Select Case controlType
                Case eControlType.guiInventory
                    inventoryClicked .inventory(getControlID(element, controlName, eControlType.guiInventory)), mouseX, mouseY
                Case eControlType.guiList
                    listClicked .List(getControlID(element, controlName, eControlType.guiList)), mouseX, mouseY
                Case eControlType.guiTextbox
                    formControlFocusSet element, controlName
            End Select
        End If
    End With
End Function

