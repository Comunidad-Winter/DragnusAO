VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Graphical inventory for ORE 0.4.1
'Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
'juansotuyo@hotmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Dim InventoryOffset As Long             'Number of lines we scrolled down from topmost
Dim InvSelectedItem As Long             'Currently selected item
Dim base_tile_size As Byte

Dim slots As Byte

'Dim InvSurface As DirectDrawSurface7            'DD Surface used to render everything

Dim UserInventory() As inventory    'User's inventory

Private Type tInventoryRect
    x As Integer
    y As Integer
    width As Integer
    height As Integer
End Type

Dim inventoryRect As tInventoryRect

#If ConMenuesConextuales = 1 Then
    Dim ItemMenu As Menu    'Menu to be shown as pop up
#End If

Public Sub setInventoryRect(ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer)
    inventoryRect.x = x
    inventoryRect.y = y
    inventoryRect.width = width
    inventoryRect.height = height
End Sub


Public Sub ScrollInventory(ByVal up As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Scrolls the graphical inventory up / down
'***************************************************
    'Only move if there is something to show up /down
    If up Then
        If (InventoryOffset + (inventoryRect.height \ base_tile_size)) * (inventoryRect.width \ base_tile_size) < slots Then _
            InventoryOffset = InventoryOffset + 1
    Else
        If InventoryOffset > 0 Then _
            InventoryOffset = InventoryOffset - 1
    End If
    
    Render_Inventory = True
    
End Sub

Public Sub SelectGold()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Sets the gold as the currently selected item
'***************************************************
    Dim prevSelectedItem As Long
    
    'Store preivously selected item
    prevSelectedItem = InvSelectedItem
    
    'Select the gold
    InvSelectedItem = FLAGORO
    
End Sub

Public Property Get SelectedItem() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item index
'***************************************************
    SelectedItem = InvSelectedItem
End Property

Public Property Get maxHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the max hit of the selected item
'***************************************************
    maxHit = UserInventory(slot).maxHit
End Property

Public Property Get minHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the min hit of the selected item
'***************************************************
    minHit = UserInventory(slot).minHit
End Property

Public Property Get def(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the defense of the selected item
'***************************************************
    def = UserInventory(slot).def
End Property

Public Property Get grhIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the grh index of the selected item
'***************************************************
    grhIndex = UserInventory(slot).grhIndex
End Property

Public Property Get valor(ByVal slot As Byte) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the value of the selected item
'***************************************************
    valor = UserInventory(slot).valor
End Property

Public Property Get amount(ByVal slot As Byte) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's amount
'***************************************************
    If slot = FLAGORO Then
        amount = UserGLD
    ElseIf slot >= LBound(UserInventory) And slot <= UBound(UserInventory) Then
        amount = UserInventory(slot).amount
    End If
End Property

Public Property Get objIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object index
'***************************************************
    objIndex = UserInventory(slot).objIndex
End Property

Public Property Get objType(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object type
'***************************************************
    objType = UserInventory(slot).objType
End Property

Public Property Get ItemName(ByVal slot As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's name
'***************************************************
    ItemName = UserInventory(slot).name
End Property

Public Property Get equipped(ByVal slot As Byte) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves True if the item at the given pos is eqiupped
'***************************************************
    equipped = UserInventory(slot).equipped
End Property

Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Integer, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Integer, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal eDef As Integer, ByVal eValor As Long, ByVal eName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/04/06
'Sets all data for a given inventory slot
'***************************************************
    If slot < 1 Or slot > slots Then Exit Sub
    
    With UserInventory(slot)
        .amount = eAmount
        .def = eDef
        .equipped = eEquipped
        .grhIndex = eGrhIndex
        .maxHit = eMaxHit
        .minHit = eMinHit
        .name = eName
        .objIndex = eOBJIndex
        .objType = eObjType
        .valor = eValor
    End With
    
    Render_Inventory = True
    
End Sub

Private Function ClickItem(ByVal x As Long, ByVal y As Long) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Selects the item clicked if it's valid and return's it's index
'***************************************************
    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
    Dim AuxInv As inventory
    
    temp_x = x \ 32
    temp_y = y \ 32
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (inventoryRect.width \ base_tile_size) + 1
    
    'Make sure it's within limits
    If TempItem <= slots Then
        'Make sure slot isn't empty
        If UserInventory(TempItem).grhIndex Then
            ClickItem = TempItem
        Else
            ClickItem = 0
        End If
    End If
    
    Render_Inventory = True
End Function

Public Sub DrawInventory(ByVal destX As Integer, ByVal destY As Integer)
    Dim loopc As Byte
    Dim filas As Integer
    Dim x As Long
    Dim y As Long
    Dim h_slots As Byte
    
    h_slots = Fix(inventoryRect.width / base_tile_size)
    
    For loopc = InventoryOffset * h_slots + 1 To UBound(UserInventory)
        If loopc - filas * h_slots > h_slots Then
            filas = filas + 1
        End If
        
        x = destX + inventoryRect.x + ((loopc - 1) - (filas * h_slots)) * 32
        y = destY + inventoryRect.y + filas * 32
        
        
        'Call guiTextureRender(16016, x, y, 40, 40, D3DColorARGB(180, 255, 255, 255))
        
        If loopc = InvSelectedItem Then
            'dxDrawBox x, y, 32, 32, D3DColorARGB(120, 50, 50, 50)
            Call GUI_Grh_Render(UserInventory(loopc).grhIndex, x, y, , False, D3DColorXRGB(255, 255, 255))
        Else
            'dxDrawBox x, y, 32, 32, D3DColorARGB(120, 100, 100, 100)
            Call GUI_Grh_Render(UserInventory(loopc).grhIndex, x, y, , False, D3DColorXRGB(150, 150, 150))
        End If
        
        If UserInventory(loopc).grhIndex Then
            Call dxTextRender(2, UserInventory(loopc).amount, x + 2, y + 1, D3DColorXRGB(255, 255, 255))
            If UserInventory(loopc).equipped Then
                Call dxTextRender(2, "+", x + 20, y + 20, D3DColorXRGB(255, 255, 0))
            End If
        End If
        
    Next loopc
End Sub


Public Sub init(ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal NumSlots As Byte)

    'Cambiamos la cantidad de items del inventario.
    slots = NumSlots
    
    ReDim UserInventory(1 To slots) As inventory
    
    setInventoryRect x, y, width, height
    base_tile_size = 32
    
    InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item
    
End Sub


Private Sub inventoryMouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse move event of the inventory picture box
'Displays a ToolTip of the item under the mouse
'***************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer
    
    'Exit if it got outside the control's area
    If x < 0 Or y < 0 Or x > inventoryRect.width Or y > inventoryRect.height Then _
        Exit Sub
    
    temp_x = x \ base_tile_size
    temp_y = y \ base_tile_size
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (inventoryRect.width \ base_tile_size) + 1
    
    If TempItem <= slots Then
        'InventoryWindow.ToolTipText = UserInventory(TempItem).Name
    End If
End Sub

Public Sub inventoryMouseUp(ByVal x As Integer, ByVal y As Integer)
    'Store previously selected item
    Dim prevSelItem As Long
    'Exit if it got outside the control's area
    If x < 0 Or y < 0 Or x > inventoryRect.x + inventoryRect.width Or y > inventoryRect.y + inventoryRect.height Then _
        Exit Sub
    
    'Inventory x,y position.
    x = x - inventoryRect.x
    y = y - inventoryRect.y
    
    prevSelItem = InvSelectedItem

    'Get the currently clickced item
    InvSelectedItem = ClickItem(CInt(x), CInt(y))
End Sub

Public Sub dropItem(ByVal mouseX As Integer, ByVal mouseY As Integer)
    
End Sub

Public Sub Reset()
    '***************************************************
    'Author: Nicolás Serafini (ElNiCk)
    'Last Modify Date: 18/7/06
    'Reset Inventory items.
    '***************************************************
    Dim slot As Integer
    For slot = LBound(UserInventory) To UBound(UserInventory)
        Call SetItem(slot, 0, 0, 0, 0, 0, 0, 0, 0, 0, "")
    Next slot
End Sub


