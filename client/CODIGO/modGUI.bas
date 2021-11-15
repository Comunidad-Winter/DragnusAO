Attribute VB_Name = "modGUI"
'Module modGUI
Option Explicit

Public Enum guiElements
    frmCargando = 1
    frmPres
    
    'main
    frmMain
    console
    statMenu
    inventory
    macros
    spells
    
    frmConectar
    frmCrearPJ
    frmCuenta
    frmBancoObj
    frmCambiaMotd
    frmCantidad
    frmCarp
    frmCharInfo
    frmComerciar
    frmComerciarUsu
    frmCommet
    frmEntrenador
    frmEstadisticas
    frmForo
    frmGuildAdm
    frmGuildBrief
    frmGuildDetails
    frmGuildFoundation
    frmGuildLeader
    frmGuildNews
    frmGuildSol
    frmGuildURL
    frmHerrero
    frmKeypad
    frmMensaje
    frmMSG
    frmOpciones
    frmPanelGm
    frmPasswd
    frmPeaceProp
    frmPremios
    frmScreenshots
    frmSkills3
    frmSpawnList
    frmSubasta
    frmtip
    frmTorneo
    frmUserRequest
End Enum

Public Enum eDragObjType
    Form
    picture
    item
End Enum

Private Type tDragObj
    dragObjType As eDragObjType
    
    sourceForm As Integer
    sourceName As String
    sourceIndex As Integer
    
    x As Integer
    y As Integer
End Type

Public Type tGuiBox
    topLeftCorner As Integer
    topMid As Integer
    topRightCorner As Integer
    
    midLeft As Integer
    midMid As Integer
    midRight As Integer
    
    bottomLeftCorner As Integer
    bottomMid As Integer
    bottomRightCorner As Integer
End Type

Private dragObj As tDragObj
Private bDrag As Boolean

Private zOrder() As Integer
Private zOrderCount As Integer

Private menuElements() As guiElement
Private menuElementsCount As Integer
Private focusedMenu As Integer
Private focusLocked As Integer

Private box() As tGuiBox

Private Const GUI_PATH As String = "\resources\init\gui\"

Public Sub guiInit()
    focusLocked = -1
    guiLoadBoxes
    createGuiElements
End Sub
Public Sub lockFocus(elementID As Integer)
    focusLocked = elementID
End Sub

Public Sub unlockFocus()
    focusLocked = -1
End Sub

Private Function getZOrder(elementID As Integer) As Integer
    Dim i As Integer
    Do While i < zOrderCount And getZOrder < 1
        i = i + 1
        If zOrder(i) = elementID Then _
            getZOrder = i
    Loop
End Function

Public Function GUIVisibleGet(elementID As Integer) As Boolean
    If getZOrder(elementID) > 0 Then _
        GUIVisibleGet = True
End Function

Public Sub GUIHideForm(elementID As Integer)
    Dim i As Integer
    If GUIVisibleGet(elementID) Then
        If elementID = focusLocked Then
            unlockFocus
        End If
        i = getZOrder(elementID)
    
        If i > 0 Then
            Do While i < zOrderCount
                zOrder(i) = zOrder(i + 1)
                i = i + 1
            Loop
        End If
    
        zOrderCount = zOrderCount - 1
    
        If zOrderCount > 0 Then _
            ReDim Preserve zOrder(1 To zOrderCount)
    End If
End Sub
Public Sub GUIShowForm(elementID As Integer, Optional ByVal x As Integer, Optional ByVal y As Integer)
    Dim i As Integer
    If Not GUIVisibleGet(elementID) Then
        zOrderCount = zOrderCount + 1
        ReDim Preserve zOrder(1 To zOrderCount)
        If menuElements(elementID).windowElement Then
            i = zOrderCount
            Do While i > 1
                zOrder(i) = zOrder(i - 1)
                i = i - 1
            Loop
            zOrder(1) = elementID
        Else
            zOrder(zOrderCount) = elementID
        End If
    End If
End Sub

Private Function focusSet(elementID As Integer)
    Dim i As Integer
    i = getZOrder(elementID)
    If i > 0 Then
        If Not menuElements(elementID).windowElement Then
            Do While i < zOrderCount
                zOrder(i) = zOrder(i + 1)
                i = i + 1
            Loop
            zOrder(zOrderCount) = elementID
        End If
        focusedMenu = elementID
    End If
End Function


'*****************************
'PROPERTIES
'*****************************

'INVENTORY
Public Sub guiPictureWidthSet(ByVal elementID As Integer, ByVal controlName As String, ByVal width As Integer)
    formPictureWidthSet menuElements(elementID), controlName, width
End Sub

Public Function guiInventoryInvGet(ByVal elementID As Integer, ByVal controlName As String) As inventory
    guiInventoryInvGet = formInventoryInvGet(menuElements(elementID), controlName)
End Function

Public Sub guiInventoryInvSet(ByVal elementID As Integer, ByVal controlName As String)

End Sub

Public Function guiInventoryItemGet(ByVal elementID As Integer, ByVal controlName As String) As Integer
    guiInventoryItemGet = formInventoryItemGet(menuElements(elementID), controlName)
End Function

Public Function guiInventoryItemSet(ByVal elementID As Integer, ByVal controlName As String, ByVal slot As Byte, ByVal objIndex As Integer, ByVal amount As Integer, ByVal equipped As Byte, ByVal grhIndex As Integer, ByVal objType As Integer, ByVal maxHit As Integer, ByVal minHit As Integer, ByVal def As Integer, ByVal valor As Long, ByVal name As String) As Integer
    formInventoryItemSet menuElements(elementID), controlName, slot, objIndex, amount, equipped, grhIndex, objType, maxHit, minHit, def, valor, name
End Function

Public Sub guiInventoryGLDSelect(ByVal elementID As Integer, ByVal controlName As String)
    formInventoryGLDSelect menuElements(elementID), controlName
End Sub

Public Sub guiInventoryScroll(ByVal elementID As Integer, ByVal controlName As String, ByVal up As Boolean)
    formInventoryScroll menuElements(elementID), controlName, up
End Sub

Public Function guiInventoryAmountGet(ByVal elementID As Integer, ByVal controlName As String, ByVal slot As Byte) As Integer
    guiInventoryAmountGet = formInventoryAmountGet(menuElements(elementID), controlName, slot)
End Function

Public Function guiInventoryGrhIndexGet(ByVal elementID As Integer, ByVal controlName As String, ByVal slot As Byte) As Integer
    guiInventoryGrhIndexGet = formInventoryGrhIndexGet(menuElements(elementID), controlName, slot)
End Function

Public Function guiInventoryMaxHitGet(ByVal elementID As Integer, ByVal controlName As String, ByVal slot As Byte) As Integer
    guiInventoryMaxHitGet = formInventoryMaxHitGet(menuElements(elementID), controlName, slot)
End Function

Public Function guiInventoryMinHitGet(ByVal elementID As Integer, ByVal controlName As String, ByVal slot As Byte) As Integer
    guiInventoryMinHitGet = formInventoryMinHitGet(menuElements(elementID), controlName, slot)
End Function

Public Function guiInventoryValueGet(ByVal elementID As Integer, ByVal controlName As String, ByVal slot As Byte) As Long
    guiInventoryValueGet = formInventoryValueGet(menuElements(elementID), controlName, slot)
End Function

Public Function guiInventoryNameGet(ByVal elementID As Integer, ByVal controlName As String, ByVal slot As Byte) As String
    guiInventoryNameGet = formInventoryNameGet(menuElements(elementID), controlName, slot)
End Function

Public Function guiInventoryEquipGet(ByVal elementID As Integer, ByVal controlName As String, ByVal slot As Byte) As Byte
    guiInventoryEquipGet = formInventoryEquipGet(menuElements(elementID), controlName, slot)
End Function

Public Function guiInventoryObjIndexGet(ByVal elementID As Integer, ByVal controlName As String, ByVal slot As Byte) As Integer
    guiInventoryObjIndexGet = formInventoryObjIndexGet(menuElements(elementID), controlName, slot)
End Function

Public Function guiInventoryObjTypeGet(ByVal elementID As Integer, ByVal controlName As String, ByVal slot As Byte) As Integer
    guiInventoryObjTypeGet = formInventoryObjTypeGet(menuElements(elementID), controlName, slot)
End Function

Public Function guiTextboxFocusGet() As Boolean
    If focusedMenu > 0 Then _
        guiTextboxFocusGet = formFocusGet(menuElements(focusedMenu))
End Function


'TEXTBOX
Public Sub guiTextboxShow(ByVal elementID As Integer, ByVal controlName As String)
    focusSet elementID
    textboxShow menuElements(elementID).textbox(getControlID(menuElements(elementID), controlName, eControlType.guiTextbox))
    formControlFocusSet menuElements(elementID), controlName
End Sub

Public Sub guiTextboxHide(ByVal elementID As Integer, ByVal controlName As String)
    formTextboxHide menuElements(elementID), controlName
End Sub

Public Sub guiTextboxClear(ByVal elementID As Integer, ByVal controlName As String)
    formTextboxClear menuElements(elementID), controlName
End Sub

Public Function guiTextboxTextGet(ByVal elementID As Integer, ByVal controlName As String) As String
    guiTextboxTextGet = formTextboxTextGet(menuElements(elementID), controlName)
End Function

Public Function guiTextboxVisibleGet(ByVal elementID As Integer, ByVal controlName As String) As Boolean
    guiTextboxVisibleGet = formTextboxVisibleGet(menuElements(elementID), controlName)
End Function
Public Function guiLabelCaptionGet(ByVal elementID As Integer, ByVal controlName As String) As String
    guiLabelCaptionGet = formLabelCaptionGet(menuElements(elementID), controlName)
End Function

Public Sub guiLabelCaptionSet(ByVal elementID As Integer, ByVal controlName As String, ByVal caption As String)
    formLabelCaptionSet menuElements(elementID), controlName, caption
End Sub

'LIST
Public Sub guiListAddItem(elementID As Integer, ByVal controlName As String, ByVal newItem As String, Optional ByVal color As Long)
    If color = 0 Then _
        color = D3DColorXRGB(255, 255, 255)
    formListAddItem menuElements(elementID), controlName, newItem, color
End Sub
Public Sub guiListClear(elementID As Integer, ByVal controlName As String)
    formListClear menuElements(elementID), controlName
End Sub
Public Function guiListIndexGet(elementID As Integer, ByVal controlName As String) As Integer
    guiListIndexGet = formListIndexGet(menuElements(elementID), controlName)
End Function

Public Sub guiTextboxTextSet(elementID As Integer, ByVal controlName As String, ByVal Text As String)
    formTextboxTextSet menuElements(elementID), controlName, Text
End Sub

Public Sub guiFormPosSet(elementID As Integer, ByVal x As Integer, ByVal y As Integer)
    formPosSet menuElements(elementID), x, y
End Sub
'*********************
'INPUT
'*********************
Public Sub GUIStartDrag(ByVal mouseX As Integer, ByVal mouseY As Integer)
    Dim elementID As Integer
    Dim controlName As String
    Dim controlType As eControlType
    
    elementID = formClickGet(mouseX, mouseY)
    If elementID > 0 Then
        If formControlClick(menuElements(elementID), mouseX, mouseY, controlName, controlType) Then
            Select Case controlType
                Case eControlType.guiPicture
                    If formPictureDrag(menuElements(elementID), controlName) Then
                        dragObj.dragObjType = eDragObjType.picture
                        dragObj.sourceForm = elementID
                        dragObj.sourceName = controlName
                        bDrag = True
                    End If
                Case eControlType.guiInventory
                    If formInventoryDrag(menuElements(elementID), controlName) Then
                        formClicked menuElements(elementID), mouseX, mouseY, controlName, controlType
                        dragObj.dragObjType = eDragObjType.item
                        dragObj.sourceForm = elementID
                        dragObj.sourceName = controlName
                        dragObj.sourceIndex = guiInventoryItemGet(elementID, controlName)
                        If dragObj.sourceIndex > 0 Then _
                            bDrag = True
                    End If
            End Select
            
            dragObj.x = mouseX
            dragObj.y = mouseY
        Else
            formMousePosGet menuElements(elementID), mouseX, mouseY
            If formBarClick(menuElements(elementID), mouseX, mouseY) Then
                If Not menuElements(elementID).locked Then
                    focusSet elementID
                    dragObj.dragObjType = eDragObjType.Form
                    dragObj.sourceForm = elementID
                    dragObj.x = mouseX
                    dragObj.y = mouseY
                    bDrag = True
                End If
            End If
        End If
    End If
End Sub
Public Sub GUIDrag(ByVal mouseX As Integer, ByVal mouseY As Integer)
    If bDrag Then
        Select Case dragObj.dragObjType
            Case eDragObjType.Form
                If mouseX - dragObj.x < 800 - menuElements(dragObj.sourceForm).windowRect.width / 3 And mouseY - dragObj.y < 600 - menuElements(dragObj.sourceForm).windowRect.height / 3 Then
                    menuElements(dragObj.sourceForm).windowRect.x = mouseX - dragObj.x
                    menuElements(dragObj.sourceForm).windowRect.y = mouseY - dragObj.y
                End If
            Case Else
                dragObj.x = mouseX
                dragObj.y = mouseY
        End Select
    End If
End Sub
Public Sub GUIEndDrag(ByVal mouseX As Integer, ByVal mouseY As Integer)
    
    Dim targetElementID As Integer
    Dim targetControlName As String
    Dim targetControlType As eControlType
    
    targetElementID = formClickGet(mouseX, mouseY)
    
    If targetElementID Then _
        formControlClick menuElements(targetElementID), mouseX, mouseY, targetControlName, targetControlType
    
    Select Case dragObj.dragObjType
        Case eDragObjType.picture
            eventPictureDrop targetElementID, targetControlName, targetControlType, dragObj.sourceForm, dragObj.sourceName
        Case eDragObjType.item
            eventItemDrop targetElementID, targetControlName, targetControlType, dragObj.sourceForm, dragObj.sourceName, dragObj.sourceIndex
    End Select
    
    bDrag = False
    dragObj.sourceForm = 0
    dragObj.sourceName = ""
    dragObj.sourceIndex = 0
End Sub

Public Function guiDragGet() As Boolean
    guiDragGet = bDrag
End Function

Private Function formClickGet(ByVal mouseX As Integer, ByVal mouseY As Integer) As Integer
    Dim i As Integer
    i = zOrderCount
    
    Do While i > 0 And formClickGet < 1
        If formClick(menuElements(zOrder(i)), mouseX, mouseY) Then
            formClickGet = zOrder(i)
        End If
        i = i - 1
    Loop
End Function

Public Function GUIInputClick(ByVal mouseX As Integer, ByVal mouseY As Integer) As Boolean
    Dim elementID As Integer
    
    Dim controlName As String
    Dim controlType As eControlType
    
    elementID = formClickGet(mouseX, mouseY)
    
    If elementID > 0 And ((elementID = focusLocked) Or focusLocked = -1) Then
        GUIInputClick = True
        focusSet elementID
        If formControlClick(menuElements(elementID), mouseX, mouseY, controlName, controlType) Then
            formClicked menuElements(elementID), mouseX, mouseY, controlName, controlType
            
            'EVENT Manager
            Select Case controlType
                Case eControlType.guiButton
                    eventButtonClick elementID, controlName
            End Select
        ElseIf formCloseClick(menuElements(elementID), mouseX, mouseY) Then
            If Not menuElements(elementID).windowElement Then
                GUIHideForm elementID
                eventFormClose elementID
            End If
        End If
    Else
        focusedMenu = 0
    End If
End Function

Public Function guiInputKey(ByVal KeyAscii As Byte) As Boolean
    If focusedMenu > 0 Then
        guiInputKey = formKeyPress(menuElements(focusedMenu), KeyAscii)
    End If
End Function


'*****************************
'RENDERS
'*****************************
Public Function guiRender()
    Dim i As Integer
    Dim mouseX As Integer
    Dim mouseY As Integer
    
    For i = 1 To zOrderCount
        formRender menuElements(zOrder(i))
    Next i
    
    dragDropRender
    inputMouseGet mouseX, mouseY
    'guiTextureRender 16021, mouseX - 4, mouseY - 4, 32, 32, D3DColorARGB(255, 255, 255, 255)
    dxTextRender 1, "(" & mouseX & "," & mouseY & ")", mouseX, mouseY, D3DColorXRGB(255, 255, 255)
End Function

Private Sub dragDropRender()
    Dim grhIndex As Integer

    If bDrag Then
        If dragObj.dragObjType = eDragObjType.item Then
            grhIndex = guiInventoryGrhIndexGet(dragObj.sourceForm, dragObj.sourceName, dragObj.sourceIndex)
            GUI_Grh_Render grhIndex, dragObj.x - 16, dragObj.y - 16, , , D3DColorARGB(255, 255, 255, 255)
        ElseIf dragObj.dragObjType = eDragObjType.picture Then
            
        End If
    End If
End Sub

Public Sub guiTextureRender(ByVal textureIndex As Integer, ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal color As Long, Optional ByVal alphaBlend As Boolean, Optional ByVal angle As Single)
    Dim rgbList(3) As Long
    
    rgbList(0) = color
    rgbList(1) = color
    rgbList(2) = color
    rgbList(3) = color
    
    dxTextureRender textureIndex, x, y, width, height, rgbList, 1, 1, width, height, alphaBlend, angle, 1
End Sub


Public Function guiCreateRectangle(x As Integer, y As Integer, width As Integer, height As Integer) As tRectangle
    With guiCreateRectangle
        .x = x
        .y = y
        .width = width
        .height = height
    End With
End Function

'*****************************
'ELEMENT CREATORS
'*****************************
Private Function createGuiElement(ByVal elementID As Integer)
    Select Case elementID
        Case guiElements.spells
            loadMenu menuElements(guiElements.spells), App.Path & GUI_PATH & "frmSpells.ini"
        Case guiElements.inventory
            loadMenu menuElements(guiElements.inventory), App.Path & GUI_PATH & "frmInventory.ini"
        Case guiElements.frmMain
            loadMenu menuElements(guiElements.frmMain), App.Path & GUI_PATH & "frmMain.ini"
        Case guiElements.console
            formAddTextbox menuElements(elementID), "txtChat", guiCreateRectangle(0, 0, 10, 10)
            guiTextboxHide elementID, "txtChat"
        Case guiElements.frmComerciar
            loadMenu menuElements(guiElements.frmComerciar), App.Path & GUI_PATH & "frmComerciar.ini"

        Case guiElements.frmBancoObj
            loadMenu menuElements(guiElements.frmBancoObj), App.Path & GUI_PATH & "frmBanco.ini"
            
            
        Case guiElements.frmCambiaMotd
        
        Case guiElements.frmCantidad
            loadMenu menuElements(guiElements.frmCantidad), App.Path & GUI_PATH & "frmCantidad.ini"
        
        Case guiElements.frmCargando
        
        Case guiElements.frmCarp
            menuElements(elementID).windowRect = guiCreateRectangle(250, 200, 300, 200)
            'formAddButton menuElements(elementID), "btnConstruir", guiCreateRectangle(25, 145, 250, 30)
            formAddList menuElements(elementID), "listCarpintero", guiCreateRectangle(15, 15, 270, 120)
        
        Case guiElements.frmCharInfo
        
        Case guiElements.frmComerciarUsu
        
        Case guiElements.frmCommet
        
        Case guiElements.frmConectar
            loadMenu menuElements(guiElements.frmConectar), App.Path & GUI_PATH & "frmConectar.ini"
            
        Case guiElements.frmCrearPJ
        
        Case guiElements.frmCuenta
        
        Case guiElements.frmEntrenador
            loadMenu menuElements(guiElements.frmEntrenador), App.Path & GUI_PATH & "frmEstadisticas.ini"
            
        Case guiElements.frmEstadisticas
            loadMenu menuElements(guiElements.frmEstadisticas), App.Path & GUI_PATH & "frmEstadisticas.ini"
            
        Case guiElements.frmForo
        
        Case guiElements.frmGuildAdm
            loadMenu menuElements(guiElements.frmGuildAdm), App.Path & GUI_PATH & "frmGuildAdm.ini"
            
        Case guiElements.frmGuildBrief
            
        Case guiElements.frmGuildDetails
        
        Case guiElements.frmGuildFoundation
        
        Case guiElements.frmGuildLeader
        
        Case guiElements.frmGuildNews
        
        Case guiElements.frmGuildSol
        
        Case guiElements.frmGuildURL
        
        Case guiElements.frmHerrero
        
        Case guiElements.frmMensaje
        
        Case guiElements.frmMSG
        
        Case guiElements.frmOpciones
        
        Case guiElements.frmPanelGm
        
        Case guiElements.frmPeaceProp
        
    End Select
End Function

Private Function createGuiElements()
    Dim i As Integer
    
    menuElementsCount = guiElements.frmUserRequest
    ReDim menuElements(1 To menuElementsCount) As guiElement
    Do While i < menuElementsCount
        i = i + 1
        createGuiElement i
    Loop
End Function

Private Sub loadMenu(element As guiElement, file As String)
    On Error GoTo errhandler
    Dim controls As Integer
    Dim i As Integer
    
    element.windowRect = guiCreateRectangle(General_Var_Get(file, "INIT", "rectX"), General_Var_Get(file, "INIT", "rectY"), General_Var_Get(file, "INIT", "rectWidth"), General_Var_Get(file, "INIT", "rectHeight"))
    element.locked = IIf(General_Var_Get(file, "INIT", "locked") <> "", CBool(General_Var_Get(file, "INIT", "locked")), False)
    element.windowElement = IIf(General_Var_Get(file, "INIT", "windowElement") <> "", CBool(General_Var_Get(file, "INIT", "windowElement")), False)
    element.backGround = General_Var_Get(file, "INIT", "background")
    
    controls = General_Var_Get(file, "CONTROLS", "inventoryCount")
    For i = 1 To controls
        formAddInventory element, General_Var_Get(file, "INVENTORY" & i, "controlName"), guiCreateRectangle(General_Var_Get(file, "INVENTORY" & i, "rectX"), General_Var_Get(file, "INVENTORY" & i, "rectY"), General_Var_Get(file, "INVENTORY" & i, "rectWidth"), General_Var_Get(file, "INVENTORY" & i, "rectHeight")), General_Var_Get(file, "INVENTORY" & i, "slots"), General_Var_Get(file, "INVENTORY" & i, "dragEnable")
    Next i
    
    controls = General_Var_Get(file, "CONTROLS", "listCount")
    For i = 1 To controls
        formAddList element, General_Var_Get(file, "LIST" & i, "controlName"), guiCreateRectangle(General_Var_Get(file, "LIST" & i, "rectX"), General_Var_Get(file, "LIST" & i, "rectY"), General_Var_Get(file, "LIST" & i, "rectWidth"), General_Var_Get(file, "LIST" & i, "rectHeight"))
    Next i
    
    controls = General_Var_Get(file, "CONTROLS", "pictureCount")
    For i = 1 To controls
        formAddPicture element, General_Var_Get(file, "PICTURE" & i, "controlName"), guiCreateRectangle(General_Var_Get(file, "PICTURE" & i, "rectX"), General_Var_Get(file, "PICTURE" & i, "rectY"), General_Var_Get(file, "PICTURE" & i, "rectWidth"), General_Var_Get(file, "PICTURE" & i, "rectHeight")), General_Var_Get(file, "PICTURE" & i, "textureIndex"), General_Var_Get(file, "PICTURE" & i, "dragEnable")
    Next i
    
    controls = General_Var_Get(file, "CONTROLS", "buttonCount")
    For i = 1 To controls
        formAddButton element, General_Var_Get(file, "BUTTON" & i, "controlName"), guiCreateRectangle(General_Var_Get(file, "BUTTON" & i, "rectX"), General_Var_Get(file, "BUTTON" & i, "rectY"), General_Var_Get(file, "BUTTON" & i, "rectWidth"), General_Var_Get(file, "BUTTON" & i, "rectHeight")), Val(General_Var_Get(file, "BUTTON" & i, "normalTextureIndex")), Val(General_Var_Get(file, "BUTTON" & i, "rolloverTextureIndex"))
    Next i
    
    controls = General_Var_Get(file, "CONTROLS", "labelCount")
    For i = 1 To controls
        formAddLabel element, General_Var_Get(file, "LABEL" & i, "controlName"), guiCreateRectangle(General_Var_Get(file, "LABEL" & i, "rectX"), General_Var_Get(file, "LABEL" & i, "rectY"), General_Var_Get(file, "LABEL" & i, "rectWidth"), General_Var_Get(file, "LABEL" & i, "rectHeight"))
    Next i
    
    controls = General_Var_Get(file, "CONTROLS", "textboxCount")
    For i = 1 To controls
        formAddTextbox element, General_Var_Get(file, "TEXTBOX" & i, "controlName"), guiCreateRectangle(General_Var_Get(file, "TEXTBOX" & i, "rectX"), General_Var_Get(file, "TEXTBOX" & i, "rectY"), General_Var_Get(file, "TEXTBOX" & i, "rectWidth"), General_Var_Get(file, "TEXTBOX" & i, "rectHeight"))
    Next i
    Exit Sub
errhandler:
    MsgBox "failed loading: " & file & " - Error:" & Err.Description
End Sub

Public Sub guiLoadBoxes()
    Dim boxCount As Integer
    Dim i As Integer
    Dim file As String
    
    file = App.Path & GUI_PATH & "boxes.ini"
    
    boxCount = Val(General_Var_Get(file, "INIT", "cantidad"))
    
    ReDim box(boxCount)
    For i = 0 To boxCount - 1
        box(i).topLeftCorner = Val(General_Var_Get(file, "BOX" & i, "topLeftCorner"))
        box(i).topMid = Val(General_Var_Get(file, "BOX" & i, "topMid"))
        box(i).topRightCorner = Val(General_Var_Get(file, "BOX" & i, "topRightCorner"))
        
        box(i).bottomLeftCorner = Val(General_Var_Get(file, "BOX" & i, "bottomLeftCorner"))
        box(i).bottomMid = Val(General_Var_Get(file, "BOX" & i, "bottomMid"))
        box(i).bottomRightCorner = Val(General_Var_Get(file, "BOX" & i, "bottomRight"))
        
        box(i).midLeft = Val(General_Var_Get(file, "BOX" & i, "midLeft"))
        box(i).midMid = Val(General_Var_Get(file, "BOX" & i, "midMid"))
        box(i).midRight = Val(General_Var_Get(file, "BOX" & i, "midRight"))
    Next i
End Sub

Public Sub guiDrawBox(RECT As tRectangle, boxIndex As Integer)
    
    'dxDrawBox RECT.x, RECT.y, RECT.width, RECT.height, D3DColorARGB(50, 10, 10, 10), D3DColorARGB(50, 10, 10, 10), 1
    If boxIndex > -1 Then
        With box(boxIndex)
            If .midMid > 0 Then _
                guiTextureRender .midMid, RECT.x + 10, RECT.y + 10, RECT.width - 10, RECT.height - 10, D3DColorXRGB(255, 255, 255)
                
            If .midLeft > 0 Then _
                guiTextureRender .midLeft, RECT.x, RECT.y + 64, 64, RECT.height - 128, D3DColorXRGB(255, 255, 255)
                
            If .midRight > 0 Then _
                guiTextureRender .midRight, RECT.x + RECT.width - 64, RECT.y + 64, 64, RECT.height - 128, D3DColorXRGB(255, 255, 255)
                
            If .topMid > 0 Then _
                guiTextureRender .topMid, RECT.x + 64, RECT.y, RECT.width - 128, 64, D3DColorXRGB(255, 255, 255)
                
            If .bottomMid > 0 Then _
                guiTextureRender .bottomMid, RECT.x + 64, RECT.y + RECT.height - 64, RECT.width - 128, 64, D3DColorXRGB(255, 255, 255)
            
            If .topRightCorner > 0 Then _
                guiTextureRender .topRightCorner, RECT.x + RECT.width - 64, RECT.y, 64, 64, D3DColorXRGB(255, 255, 255)
            
            If .topLeftCorner > 0 Then _
                guiTextureRender .topLeftCorner, RECT.x, RECT.y, 64, 64, D3DColorXRGB(255, 255, 255)
            
            If .bottomLeftCorner > 0 Then _
                guiTextureRender .bottomLeftCorner, RECT.x, RECT.y + RECT.height - 64, 64, 64, D3DColorXRGB(255, 255, 255)
            
            If .bottomRightCorner > 0 Then _
                guiTextureRender .bottomRightCorner, RECT.x + RECT.width - 64, RECT.y + RECT.height - 64, 64, 64, D3DColorXRGB(255, 255, 255)
        End With
    End If
End Sub
