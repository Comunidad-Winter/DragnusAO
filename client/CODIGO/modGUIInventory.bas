Attribute VB_Name = "modGUIInventory"
Option Explicit

Public Type guiInventory
    controlName As String
    controlRect As tRectangle
    invent As clsGrapchicalInventory
    dragEnable As Boolean
End Type


Public Function guiCreateInventory(ByVal controlName As String, rectangle As tRectangle, ByVal slots As Integer, Optional ByVal dragEnable As Boolean = False) As guiInventory
        guiCreateInventory.controlName = controlName
        guiCreateInventory.controlRect = rectangle
        guiCreateInventory.dragEnable = dragEnable
        
        Set guiCreateInventory.invent = New clsGrapchicalInventory
        
        guiCreateInventory.invent.init guiCreateInventory.controlRect.x, guiCreateInventory.controlRect.y, guiCreateInventory.controlRect.width, guiCreateInventory.controlRect.height, slots
End Function

Public Function inventoryClick(inventory As guiInventory, ByVal mouseX As Integer, ByVal mouseY As Integer) As Boolean
    With inventory
        inventoryClick = rectMouseOver(.controlRect, mouseX, mouseY)
    End With
End Function

Public Sub inventoryClicked(inventory As guiInventory, ByVal mouseX As Integer, ByVal mouseY As Integer)
    inventory.invent.inventoryMouseUp mouseX, mouseY
End Sub

Public Sub inventoryDrop(inventory As guiInventory, ByVal mouseX As Integer, ByVal mouseY As Integer, ByVal slot As Integer)
    
End Sub

'*********************
'PROPERTIES
'*********************

Public Sub inventoryItemSet(inventory As guiInventory, ByVal slot As Byte, ByVal objIndex As Integer, ByVal amount As Integer, ByVal equipped As Byte, ByVal grhIndex As Integer, ByVal objType As Integer, ByVal maxHit As Integer, ByVal minHit As Integer, ByVal def As Integer, ByVal valor As Long, ByVal name As String)
    With inventory
        .invent.SetItem slot, objIndex, amount, equipped, grhIndex, objType, maxHit, minHit, def, valor, name
    End With
End Sub

Public Function inventoryInvGet(inventory As guiInventory) As inventory
    'inventoryInvGet = inventory.invent.invGet
End Function

Public Sub inventoryGLDSelect(inventory As guiInventory)
    With inventory
        .invent.SelectGold
    End With
End Sub

Public Function inventoryItemGet(inventory As guiInventory) As Integer
    With inventory
        inventoryItemGet = .invent.SelectedItem
    End With
End Function

Public Function inventoryScroll(inventory As guiInventory, up As Boolean) As Integer
    With inventory
        .invent.ScrollInventory up
    End With
End Function

Public Function inventoryAmountGet(inventory As guiInventory, ByVal slot As Byte) As Integer
    With inventory
        inventoryAmountGet = .invent.amount(slot)
    End With
End Function

Public Function inventoryGrhIndexGet(inventory As guiInventory, ByVal slot As Byte) As Integer
    With inventory
        inventoryGrhIndexGet = .invent.grhIndex(slot)
    End With
End Function

Public Function inventoryMaxHitGet(inventory As guiInventory, ByVal slot As Byte) As Integer
    With inventory
        inventoryMaxHitGet = .invent.maxHit(slot)
    End With
End Function

Public Function inventoryMinHitGet(inventory As guiInventory, ByVal slot As Byte) As Integer
    With inventory
        inventoryMinHitGet = .invent.minHit(slot)
    End With
End Function

Public Function inventoryValueGet(inventory As guiInventory, ByVal slot As Byte) As Long
    With inventory
        inventoryValueGet = .invent.valor(slot)
    End With
End Function

Public Function inventoryNameGet(inventory As guiInventory, ByVal slot As Byte) As String
    With inventory
        inventoryNameGet = .invent.ItemName(slot)
    End With
End Function


Public Function inventoryEquipGet(inventory As guiInventory, ByVal slot As Byte) As Byte
    With inventory
        inventoryEquipGet = .invent.equipped(slot)
    End With
End Function

Public Function inventoryObjIndexGet(inventory As guiInventory, ByVal slot As Byte) As Integer
    With inventory
        inventoryObjIndexGet = .invent.objIndex(slot)
    End With
End Function

Public Function inventoryObjTypeGet(inventory As guiInventory, ByVal slot As Byte) As Integer
    With inventory
        inventoryObjTypeGet = .invent.objType(slot)
    End With
End Function


Public Sub inventoryRender(inventory As guiInventory, ByVal destX As Integer, ByVal destY As Integer)
    
    With inventory
        '(CAJITA)
        guiDrawBox guiCreateRectangle(destX + .controlRect.x - 3, destY + .controlRect.y - 3, .controlRect.width, .controlRect.height), 1
        .invent.DrawInventory destX, destY
    End With
End Sub

Public Function inventoryDrag(inventory As guiInventory) As Boolean
    inventoryDrag = inventory.dragEnable
End Function
