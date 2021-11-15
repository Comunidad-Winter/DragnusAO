Attribute VB_Name = "ModAreas"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

'LAS GUARDAMOS PARA PROCESAR LOS MPs y sabes si borrar personajes
Public MinLimiteX As Integer
Public MaxLimiteX As Integer
Public MinLimiteY As Integer
Public MaxLimiteY As Integer

Public Sub CambioDeArea(ByVal x As Byte, ByVal y As Byte)
    Dim loopX As Long, loopY As Long
    Dim char_index As Integer
    
    MinLimiteX = (x \ 13 - 1) * 13
    MaxLimiteX = MinLimiteX + 38
    
    MinLimiteY = (y \ 13 - 1) * 13
    MaxLimiteY = MinLimiteY + 38
    
    For loopX = 1 To 100
        For loopY = 1 To 100
            
            If (loopY < MinLimiteY) Or (loopY > MaxLimiteY) Or (loopX < MinLimiteX) Or (loopX > MaxLimiteX) Then
                'Erase NPCs
                char_index = Engine.Map_Char_Get(loopX, loopY)
                If char_index > 0 Then
                    If char_index <> Engine.User_Char_Index_Get Then
                        Call Engine.Char_Remove(char_index)
                    End If
                End If
                
                'Erase OBJs
                Call Engine.Map_Item_Remove(loopX, loopY)
            End If
        Next
    Next
    
    Call Engine.Char_Refresh_All
End Sub

Public Function isPcArea(ByVal tX As Integer, ByVal tY As Integer) As Boolean
    If Not ((tY < MinLimiteY) Or (tY > MaxLimiteY) Or (tX < MinLimiteX) Or (tX > MaxLimiteX)) Then
        isPcArea = True
    End If
End Function
