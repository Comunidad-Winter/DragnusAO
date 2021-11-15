Attribute VB_Name = "InvUsuario"
'Argentum Online 0.11.6
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Function TieneObjetosRobables(ByVal userIndex As Integer) As Boolean

'17/09/02
'Agregue que la función se asegure que el objeto no es un barco

On Error Resume Next

Dim i As Integer
Dim ObjIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
    ObjIndex = UserList(userIndex).Invent.Object(i).ObjIndex
    If ObjIndex > 0 Then
            If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And _
                ObjData(ObjIndex).OBJType <> eOBJType.otBarcos) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
    
    End If
Next i


End Function

Function ClasePuedeUsarItem(ByVal userIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo manejador

'Call LogTarea("ClasePuedeUsarItem")

Dim flag As Boolean

'Admins can use ANYTHING!
If UserList(userIndex).flags.Privilegios And PlayerType.User Then
    If ObjData(ObjIndex).ExclusivoClase = 0 Then
        If ObjData(ObjIndex).ClaseProhibida(1) <> 0 Then
            Dim i As Integer
            For i = 1 To NUMClaseS
                If ObjData(ObjIndex).ClaseProhibida(i) = UserList(userIndex).Clase Then
                    ClasePuedeUsarItem = False
                    Exit Function
                End If
            Next i
        End If
    Else
        If ObjData(ObjIndex).ExclusivoClase <> UserList(userIndex).Clase Then
            ClasePuedeUsarItem = False
            Exit Function
        End If
    End If
End If

ClasePuedeUsarItem = True

Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Sub QuitarNewbieObj(ByVal userIndex As Integer)
Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(userIndex).Invent.Object(j).ObjIndex > 0 Then
             
             If ObjData(UserList(userIndex).Invent.Object(j).ObjIndex).Newbie = 1 Then _
                    Call QuitarUserInvItem(userIndex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, userIndex, j)
        
        End If
Next j

'[Barrin 17-12-03] Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon
'es transportado a Ulla
If UCase$(MapInfo(UserList(userIndex).pos.Map).Restringir) = "NEWBIE" Then
    Call WarpUserChar(userIndex, Ullathorpe.Map, Ullathorpe.X, Ullathorpe.Y, True)
End If
'[/Barrin]

End Sub

Sub LimpiarInventario(ByVal userIndex As Integer)


Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        UserList(userIndex).Invent.Object(j).ObjIndex = 0
        UserList(userIndex).Invent.Object(j).amount = 0
        UserList(userIndex).Invent.Object(j).Equipped = 0
        
Next

UserList(userIndex).Invent.NroItems = 0

UserList(userIndex).Invent.ArmourEqpObjIndex = 0
UserList(userIndex).Invent.ArmourEqpSlot = 0

UserList(userIndex).Invent.WeaponEqpObjIndex = 0
UserList(userIndex).Invent.WeaponEqpSlot = 0

UserList(userIndex).Invent.CascoEqpObjIndex = 0
UserList(userIndex).Invent.CascoEqpSlot = 0

UserList(userIndex).Invent.EscudoEqpObjIndex = 0
UserList(userIndex).Invent.EscudoEqpSlot = 0

UserList(userIndex).Invent.AnilloEqpObjIndex = 0
UserList(userIndex).Invent.AnilloEqpSlot = 0

UserList(userIndex).Invent.MunicionEqpObjIndex = 0
UserList(userIndex).Invent.MunicionEqpSlot = 0

UserList(userIndex).Invent.BarcoObjIndex = 0
UserList(userIndex).Invent.BarcoSlot = 0

UserList(userIndex).Invent.MonturaObjIndex = 0
UserList(userIndex).Invent.MonturaSlot = 0

End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal userIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
'***************************************************
On Error GoTo errhandler

'If Cantidad > 100000 Then Exit Sub

'SI EL Pjta TIENE ORO LO TIRAMOS
If (Cantidad > 0) And (Cantidad <= UserList(userIndex).Stats.GLD) Then
        Dim i As Byte
        Dim MiObj As Obj
        'info debug
        Dim loops As Integer
        
        'Seguridad Alkon
        If Cantidad > 39999 Then
            Dim j As Integer
            Dim k As Integer
            Dim M As Integer
            Dim Cercanos As String
            M = UserList(userIndex).pos.Map
            For j = UserList(userIndex).pos.X - 10 To UserList(userIndex).pos.X + 10
                For k = UserList(userIndex).pos.Y - 10 To UserList(userIndex).pos.Y + 10
                    If InMapBounds(M, j, k) Then
                        If MapData(M, j, k).userIndex > 0 Then
                            Cercanos = Cercanos & UserList(MapData(M, j, k).userIndex).name & ","
                        End If
                    End If
                Next k
            Next j
            Call LogDesarrollo(UserList(userIndex).name & " tira oro. Cercanos: " & Cercanos)
        End If
        '/Seguridad
        Dim Extra As Long
        Dim TeniaOro As Long
        TeniaOro = UserList(userIndex).Stats.GLD
        If Cantidad > 500000 Then 'Para evitar explotar demasiado
            Extra = Cantidad - 500000
            Cantidad = 500000
        End If
        
        Do While (Cantidad > 0)
            
            If Cantidad > MAX_INVENTORY_OBJS And UserList(userIndex).Stats.GLD > MAX_INVENTORY_OBJS Then
                MiObj.amount = MAX_INVENTORY_OBJS
                Cantidad = Cantidad - MiObj.amount
            Else
                MiObj.amount = Cantidad
                Cantidad = Cantidad - MiObj.amount
            End If

            MiObj.ObjIndex = iORO
            
            If EsGM(userIndex) Then Call LogGM(UserList(userIndex).name, "Tiro cantidad:" & MiObj.amount & " Objeto:" & ObjData(MiObj.ObjIndex).name)
            Dim AuxPos As WorldPos
            
            If UserList(userIndex).Clase = eClass.Pirat And UserList(userIndex).Invent.BarcoObjIndex = 476 Then
                AuxPos = TirarItemAlPiso(UserList(userIndex).pos, MiObj, False)
                If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                    UserList(userIndex).Stats.GLD = UserList(userIndex).Stats.GLD - MiObj.amount
                End If
            Else
                AuxPos = TirarItemAlPiso(UserList(userIndex).pos, MiObj, True)
                If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                    UserList(userIndex).Stats.GLD = UserList(userIndex).Stats.GLD - MiObj.amount
                End If
            End If
            
            'info debug
            loops = loops + 1
            If loops > 100 Then
                LogError ("Error en tiraroro")
                Exit Sub
            End If
            
        Loop
        If TeniaOro = UserList(userIndex).Stats.GLD Then Extra = 0
        If Extra > 0 Then
            UserList(userIndex).Stats.GLD = UserList(userIndex).Stats.GLD - Extra
        End If
    
End If

Exit Sub

errhandler:

End Sub

Sub QuitarUserInvItem(ByVal userIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)

Dim MiObj As Obj
'Desequipa
If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub

If UserList(userIndex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(userIndex, Slot)

If UserList(userIndex).flags.Montado = 1 And Slot = UserList(userIndex).Invent.MonturaSlot Then
    Call DoMontar(userIndex, ObjData(UserList(userIndex).Invent.MonturaObjIndex), UserList(userIndex).Invent.MonturaSlot)
End If

'Quita un objeto
UserList(userIndex).Invent.Object(Slot).amount = UserList(userIndex).Invent.Object(Slot).amount - Cantidad
'¿Quedan mas?
If UserList(userIndex).Invent.Object(Slot).amount <= 0 Then
    UserList(userIndex).Invent.NroItems = UserList(userIndex).Invent.NroItems - 1
    UserList(userIndex).Invent.Object(Slot).ObjIndex = 0
    UserList(userIndex).Invent.Object(Slot).amount = 0
End If
    
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal userIndex As Integer, ByVal Slot As Byte)

Dim NullObj As UserOBJ
Dim LoopC As Long

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(userIndex).Invent.Object(Slot).ObjIndex > 0 Then
        Call ChangeUserInv(userIndex, Slot, UserList(userIndex).Invent.Object(Slot))
    Else
        Call ChangeUserInv(userIndex, Slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        'Actualiza el inventario
        If UserList(userIndex).Invent.Object(LoopC).ObjIndex > 0 Then
            Call ChangeUserInv(userIndex, LoopC, UserList(userIndex).Invent.Object(LoopC))
        Else
            Call ChangeUserInv(userIndex, LoopC, NullObj)
        End If
    Next LoopC
End If

End Sub

Sub DropObj(ByVal userIndex As Integer, ByVal Slot As Byte, ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

Dim Obj As Obj

If num > 0 Then
  
  If num > UserList(userIndex).Invent.Object(Slot).amount Then num = UserList(userIndex).Invent.Object(Slot).amount
  
  'Check objeto en el suelo
  If MapData(UserList(userIndex).pos.Map, X, Y).ObjInfo.ObjIndex = 0 Or MapData(UserList(userIndex).pos.Map, X, Y).ObjInfo.ObjIndex = UserList(userIndex).Invent.Object(Slot).ObjIndex Then
        If UserList(userIndex).Invent.Object(Slot).Equipped = 1 Then Call Desequipar(userIndex, Slot)
        Obj.ObjIndex = UserList(userIndex).Invent.Object(Slot).ObjIndex
        
        If num + MapData(UserList(userIndex).pos.Map, X, Y).ObjInfo.amount > MAX_INVENTORY_OBJS Then
            num = MAX_INVENTORY_OBJS - MapData(UserList(userIndex).pos.Map, X, Y).ObjInfo.amount
        End If
        
        Obj.amount = num
        
        Call MakeObj(Map, Obj, Map, X, Y)
        Call QuitarUserInvItem(userIndex, Slot, num)
        Call UpdateUserInv(False, userIndex, Slot)
        
        If ObjData(Obj.ObjIndex).OBJType = eOBJType.otBarcos Then
            Call WriteConsoleMsg(userIndex, "¡¡ATENCION!! ¡ACABAS DE TIRAR TU BARCA!", FontTypeNames.FONTTYPE_TALK)
        End If
        
        If ObjData(Obj.ObjIndex).OBJType = eOBJType.otMontura Then
             Call WriteConsoleMsg(userIndex, "¡¡ATENCION!! ¡ACABAS DE TIRAR TU MONTURA!", FontTypeNames.FONTTYPE_TALK)
        End If

        
        If Not UserList(userIndex).flags.Privilegios And PlayerType.User Then Call LogGM(UserList(userIndex).name, "Tiro cantidad:" & num & " Objeto:" & ObjData(Obj.ObjIndex).name)
  Else
    Call WriteConsoleMsg(userIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
  End If
    
End If

End Sub

Sub EraseObj(ByVal sndIndex As Integer, ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

MapData(Map, X, Y).ObjInfo.amount = MapData(Map, X, Y).ObjInfo.amount - num

If MapData(Map, X, Y).ObjInfo.amount <= 0 Then
    MapData(Map, X, Y).ObjInfo.ObjIndex = 0
    MapData(Map, X, Y).ObjInfo.amount = 0
    
    Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectDelete(X, Y))
End If

End Sub

Sub MakeObj(ByVal sndIndex As Integer, ByRef Obj As Obj, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

If Obj.ObjIndex > 0 And Obj.ObjIndex <= UBound(ObjData) Then

    If MapData(Map, X, Y).ObjInfo.ObjIndex = Obj.ObjIndex Then
        MapData(Map, X, Y).ObjInfo.amount = MapData(Map, X, Y).ObjInfo.amount + Obj.amount
    Else
        MapData(Map, X, Y).ObjInfo = Obj
        
        Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(ObjData(Obj.ObjIndex).GrhIndex, X, Y))
    End If
End If

End Sub

Function MeterItemEnInventario(ByVal userIndex As Integer, ByRef MiObj As Obj) As Boolean
On Error GoTo errhandler

'Call LogTarea("MeterItemEnInventario")
 
Dim X As Integer
Dim Y As Integer
Dim Slot As Byte

'¿el user ya tiene un objeto del mismo tipo?
Slot = 1
Do Until UserList(userIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And _
         UserList(userIndex).Invent.Object(Slot).amount + MiObj.amount <= MAX_INVENTORY_OBJS
   Slot = Slot + 1
   If Slot > MAX_INVENTORY_SLOTS Then
         Exit Do
   End If
Loop
    
'Sino busca un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
   Slot = 1
   Do Until UserList(userIndex).Invent.Object(Slot).ObjIndex = 0
       Slot = Slot + 1
       If Slot > MAX_INVENTORY_SLOTS Then
           Call WriteConsoleMsg(userIndex, "No podes cargar mas objetos.", FontTypeNames.FONTTYPE_FIGHT)
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(userIndex).Invent.NroItems = UserList(userIndex).Invent.NroItems + 1
End If
    
'Mete el objeto
If UserList(userIndex).Invent.Object(Slot).amount + MiObj.amount <= MAX_INVENTORY_OBJS Then
   'Menor que MAX_INV_OBJS
   UserList(userIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
   UserList(userIndex).Invent.Object(Slot).amount = UserList(userIndex).Invent.Object(Slot).amount + MiObj.amount
Else
   UserList(userIndex).Invent.Object(Slot).amount = MAX_INVENTORY_OBJS
End If
    
MeterItemEnInventario = True
       
Call UpdateUserInv(False, userIndex, Slot)


Exit Function
errhandler:

End Function


Sub GetObj(ByVal userIndex As Integer)

Dim Obj As ObjData
Dim MiObj As Obj

'¿Hay algun obj?
If MapData(UserList(userIndex).pos.Map, UserList(userIndex).pos.X, UserList(userIndex).pos.Y).ObjInfo.ObjIndex > 0 Then
    '¿Esta permitido agarrar este obj?
    If ObjData(MapData(UserList(userIndex).pos.Map, UserList(userIndex).pos.X, UserList(userIndex).pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then
        Dim X As Integer
        Dim Y As Integer
        Dim Slot As Byte
        
        X = UserList(userIndex).pos.X
        Y = UserList(userIndex).pos.Y
        Obj = ObjData(MapData(UserList(userIndex).pos.Map, UserList(userIndex).pos.X, UserList(userIndex).pos.Y).ObjInfo.ObjIndex)
        MiObj.amount = MapData(UserList(userIndex).pos.Map, X, Y).ObjInfo.amount
        MiObj.ObjIndex = MapData(UserList(userIndex).pos.Map, X, Y).ObjInfo.ObjIndex
        
        If Not MeterItemEnInventario(userIndex, MiObj) Then
            'Call WriteConsoleMsg(UserIndex, "No puedo cargar mas objetos.", FontTypeNames.FONTTYPE_INFO)
        Else
            'Quitamos el objeto
            Call EraseObj(UserList(userIndex).pos.Map, MapData(UserList(userIndex).pos.Map, X, Y).ObjInfo.amount, UserList(userIndex).pos.Map, UserList(userIndex).pos.X, UserList(userIndex).pos.Y)
            If Not UserList(userIndex).flags.Privilegios And PlayerType.User Then Call LogGM(UserList(userIndex).name, "Agarro:" & MiObj.amount & " Objeto:" & ObjData(MiObj.ObjIndex).name)
        End If
        
    End If
Else
    Call WriteConsoleMsg(userIndex, "No hay nada aqui.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Sub Desequipar(ByVal userIndex As Integer, ByVal Slot As Byte)
'Desequipa el item slot del inventario
Dim Obj As ObjData


If (Slot < LBound(UserList(userIndex).Invent.Object)) Or (Slot > UBound(UserList(userIndex).Invent.Object)) Then
    Exit Sub
ElseIf UserList(userIndex).Invent.Object(Slot).ObjIndex = 0 Then
    Exit Sub
End If

Obj = ObjData(UserList(userIndex).Invent.Object(Slot).ObjIndex)

Select Case Obj.OBJType
    Case eOBJType.otWeapon
        UserList(userIndex).Invent.Object(Slot).Equipped = 0
        UserList(userIndex).Invent.WeaponEqpObjIndex = 0
        UserList(userIndex).Invent.WeaponEqpSlot = 0
        If Not UserList(userIndex).flags.Mimetizado = 1 Then
            UserList(userIndex).Char.Aura = 0
            UserList(userIndex).Char.WeaponAnim = NingunArma
            Call ChangeUserChar(userIndex, UserList(userIndex).Char.body, UserList(userIndex).Char.head, UserList(userIndex).Char.Heading, UserList(userIndex).Char.WeaponAnim, UserList(userIndex).Char.ShieldAnim, UserList(userIndex).Char.CascoAnim, UserList(userIndex).Char.Aura)
        End If
        Call WriteUpdateHit(userIndex)
        
    Case eOBJType.otFlechas
        UserList(userIndex).Invent.Object(Slot).Equipped = 0
        UserList(userIndex).Invent.MunicionEqpObjIndex = 0
        UserList(userIndex).Invent.MunicionEqpSlot = 0
    
    Case eOBJType.otAnillo
        UserList(userIndex).Invent.Object(Slot).Equipped = 0
        UserList(userIndex).Invent.AnilloEqpObjIndex = 0
        UserList(userIndex).Invent.AnilloEqpSlot = 0
    
    Case eOBJType.otArmadura
        UserList(userIndex).Invent.Object(Slot).Equipped = 0
        UserList(userIndex).Invent.ArmourEqpObjIndex = 0
        UserList(userIndex).Invent.ArmourEqpSlot = 0
        Call DarCuerpoDesnudo(userIndex, UserList(userIndex).flags.Mimetizado = 1)
        Call ChangeUserChar(userIndex, UserList(userIndex).Char.body, UserList(userIndex).Char.head, UserList(userIndex).Char.Heading, UserList(userIndex).Char.WeaponAnim, UserList(userIndex).Char.ShieldAnim, UserList(userIndex).Char.CascoAnim)
        Call WriteUpdateArmor(userIndex)
        
    Case eOBJType.otCASCO
        UserList(userIndex).Invent.Object(Slot).Equipped = 0
        UserList(userIndex).Invent.CascoEqpObjIndex = 0
        UserList(userIndex).Invent.CascoEqpSlot = 0
        If Not UserList(userIndex).flags.Mimetizado = 1 Then
            UserList(userIndex).Char.CascoAnim = NingunCasco
            Call ChangeUserChar(userIndex, UserList(userIndex).Char.body, UserList(userIndex).Char.head, UserList(userIndex).Char.Heading, UserList(userIndex).Char.WeaponAnim, UserList(userIndex).Char.ShieldAnim, UserList(userIndex).Char.CascoAnim)
        End If
        Call WriteUpdateCasco(userIndex)
        
    Case eOBJType.otESCUDO
        UserList(userIndex).Invent.Object(Slot).Equipped = 0
        UserList(userIndex).Invent.EscudoEqpObjIndex = 0
        UserList(userIndex).Invent.EscudoEqpSlot = 0
        If Not UserList(userIndex).flags.Mimetizado = 1 Then
            UserList(userIndex).Char.ShieldAnim = NingunEscudo
            Call ChangeUserChar(userIndex, UserList(userIndex).Char.body, UserList(userIndex).Char.head, UserList(userIndex).Char.Heading, UserList(userIndex).Char.WeaponAnim, UserList(userIndex).Char.ShieldAnim, UserList(userIndex).Char.CascoAnim)
        End If
        Call WriteUpdateEscu(userIndex)
    
End Select

Call WriteUpdateUserStats(userIndex)
Call UpdateUserInv(False, userIndex, Slot)


End Sub

Function SexoPuedeUsarItem(ByVal userIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo errhandler

If ObjData(ObjIndex).Mujer = 1 Then
    SexoPuedeUsarItem = UserList(userIndex).Genero <> eGenero.Hombre
ElseIf ObjData(ObjIndex).Hombre = 1 Then
    SexoPuedeUsarItem = UserList(userIndex).Genero <> eGenero.Mujer
Else
    SexoPuedeUsarItem = True
End If

Exit Function
errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal userIndex As Integer, ByVal ObjIndex As Integer) As Boolean

If ObjData(ObjIndex).Real = 1 Then
    FaccionPuedeUsarItem = (UserList(userIndex).Faccion.Alineacion = e_Alineacion.Real)
ElseIf ObjData(ObjIndex).Caos = 1 Then
    FaccionPuedeUsarItem = (UserList(userIndex).Faccion.Alineacion = e_Alineacion.Caos)
Else
    FaccionPuedeUsarItem = True
End If

End Function

Sub EquiparInvItem(ByVal userIndex As Integer, ByVal Slot As Byte)
On Error GoTo errhandler

'Equipa un item del inventario
Dim Obj As ObjData
Dim ObjIndex As Integer

ObjIndex = UserList(userIndex).Invent.Object(Slot).ObjIndex
Obj = ObjData(ObjIndex)

If Obj.Newbie = 1 And Not EsNewbie(userIndex) Then
     Call WriteConsoleMsg(userIndex, "Solo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)
     Exit Sub
End If
        
Select Case Obj.OBJType
    Case eOBJType.otWeapon
       If ClasePuedeUsarItem(userIndex, ObjIndex) And _
          FaccionPuedeUsarItem(userIndex, ObjIndex) Then
            'Si esta equipado lo quita
            If UserList(userIndex).Invent.Object(Slot).Equipped Then
                'Quitamos del inv el item
                Call Desequipar(userIndex, Slot)
                'Animacion por defecto
                If UserList(userIndex).flags.Mimetizado = 1 Then
                    UserList(userIndex).CharMimetizado.WeaponAnim = NingunArma
                Else
                    UserList(userIndex).Char.WeaponAnim = NingunArma
                    Call ChangeUserChar(userIndex, UserList(userIndex).Char.body, UserList(userIndex).Char.head, UserList(userIndex).Char.Heading, UserList(userIndex).Char.WeaponAnim, UserList(userIndex).Char.ShieldAnim, UserList(userIndex).Char.CascoAnim)
                End If
                Exit Sub
            End If
            
            'Quitamos el elemento anterior
            If UserList(userIndex).Invent.WeaponEqpObjIndex > 0 Then
                Call Desequipar(userIndex, UserList(userIndex).Invent.WeaponEqpSlot)
            End If
            
            UserList(userIndex).Invent.Object(Slot).Equipped = 1
            UserList(userIndex).Invent.WeaponEqpObjIndex = UserList(userIndex).Invent.Object(Slot).ObjIndex
            UserList(userIndex).Invent.WeaponEqpSlot = Slot

            If ObjData(UserList(userIndex).Invent.WeaponEqpObjIndex).Aura Then
                UserList(userIndex).Char.Aura = ObjData(UserList(userIndex).Invent.WeaponEqpObjIndex).Aura
            Else
                UserList(userIndex).Char.Aura = 0
            End If
            
            'Call writechangeCharaura(UserIndex)
            
            'Sonido
            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_SACARARMA))
            
            If UserList(userIndex).flags.Mimetizado = 1 Then
                UserList(userIndex).CharMimetizado.WeaponAnim = Obj.WeaponAnim
            Else
                UserList(userIndex).Char.WeaponAnim = Obj.WeaponAnim
                
                If UserList(userIndex).flags.Montado = 0 Then
                    Call ChangeUserChar(userIndex, UserList(userIndex).Char.body, UserList(userIndex).Char.head, UserList(userIndex).Char.Heading, UserList(userIndex).Char.WeaponAnim, UserList(userIndex).Char.ShieldAnim, UserList(userIndex).Char.CascoAnim, UserList(userIndex).Char.Aura)
                End If
                
            End If
            
            'Si lo equipa, actualizamos labels.
            Call WriteUpdateHit(userIndex)
       Else
            Call WriteConsoleMsg(userIndex, "Tu Clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
    
    Case eOBJType.otAnillo
       If ClasePuedeUsarItem(userIndex, ObjIndex) And _
          FaccionPuedeUsarItem(userIndex, ObjIndex) Then
                'Si esta equipado lo quita
                If UserList(userIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(userIndex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(userIndex).Invent.AnilloEqpObjIndex > 0 Then
                    Call Desequipar(userIndex, UserList(userIndex).Invent.AnilloEqpSlot)
                End If
        
                UserList(userIndex).Invent.Object(Slot).Equipped = 1
                UserList(userIndex).Invent.AnilloEqpObjIndex = ObjIndex
                UserList(userIndex).Invent.AnilloEqpSlot = Slot
                
       Else
            Call WriteConsoleMsg(userIndex, "Tu Clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
    
    Case eOBJType.otFlechas
       If ClasePuedeUsarItem(userIndex, UserList(userIndex).Invent.Object(Slot).ObjIndex) And _
          FaccionPuedeUsarItem(userIndex, UserList(userIndex).Invent.Object(Slot).ObjIndex) Then
                
                'Si esta equipado lo quita
                If UserList(userIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(userIndex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(userIndex).Invent.MunicionEqpObjIndex > 0 Then
                    Call Desequipar(userIndex, UserList(userIndex).Invent.MunicionEqpSlot)
                End If
        
                UserList(userIndex).Invent.Object(Slot).Equipped = 1
                UserList(userIndex).Invent.MunicionEqpObjIndex = UserList(userIndex).Invent.Object(Slot).ObjIndex
                UserList(userIndex).Invent.MunicionEqpSlot = Slot
                
       Else
            Call WriteConsoleMsg(userIndex, "Tu Clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
    
    Case eOBJType.otArmadura
        If UserList(userIndex).flags.Navegando = 1 Or UserList(userIndex).flags.Montado = 1 Then Exit Sub
        'Nos aseguramos que puede usarla
        If ClasePuedeUsarItem(userIndex, UserList(userIndex).Invent.Object(Slot).ObjIndex) And _
           SexoPuedeUsarItem(userIndex, UserList(userIndex).Invent.Object(Slot).ObjIndex) And _
           CheckRazaUsaRopa(userIndex, UserList(userIndex).Invent.Object(Slot).ObjIndex) And _
           FaccionPuedeUsarItem(userIndex, UserList(userIndex).Invent.Object(Slot).ObjIndex) Then
           
           'Si esta equipado lo quita
            If UserList(userIndex).Invent.Object(Slot).Equipped Then
                Call Desequipar(userIndex, Slot)
                Call DarCuerpoDesnudo(userIndex, UserList(userIndex).flags.Mimetizado = 1)
                If Not UserList(userIndex).flags.Mimetizado = 1 Then
                    Call ChangeUserChar(userIndex, UserList(userIndex).Char.body, UserList(userIndex).Char.head, UserList(userIndex).Char.Heading, UserList(userIndex).Char.WeaponAnim, UserList(userIndex).Char.ShieldAnim, UserList(userIndex).Char.CascoAnim)
                End If
                Exit Sub
            End If
    
            'Quita el anterior
            If UserList(userIndex).Invent.ArmourEqpObjIndex > 0 Then
                Call Desequipar(userIndex, UserList(userIndex).Invent.ArmourEqpSlot)
            End If
    
            'Lo equipa
            UserList(userIndex).Invent.Object(Slot).Equipped = 1
            UserList(userIndex).Invent.ArmourEqpObjIndex = UserList(userIndex).Invent.Object(Slot).ObjIndex
            UserList(userIndex).Invent.ArmourEqpSlot = Slot
                
            If UserList(userIndex).flags.Mimetizado = 1 Then
                UserList(userIndex).CharMimetizado.body = Obj.Ropaje
            Else
                UserList(userIndex).Char.body = Obj.Ropaje
                Call ChangeUserChar(userIndex, UserList(userIndex).Char.body, UserList(userIndex).Char.head, UserList(userIndex).Char.Heading, UserList(userIndex).Char.WeaponAnim, UserList(userIndex).Char.ShieldAnim, UserList(userIndex).Char.CascoAnim)
            End If
            UserList(userIndex).flags.Desnudo = 0
            
            'Si lo equipa, actualizamos labels.
            Call WriteUpdateArmor(userIndex)
        Else
            Call WriteConsoleMsg(userIndex, "Tu Clase,Genero o Raza no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
        End If
    
    Case eOBJType.otCASCO
        If UserList(userIndex).flags.Navegando = 1 Or UserList(userIndex).flags.Montado = 1 Then Exit Sub
        If ClasePuedeUsarItem(userIndex, UserList(userIndex).Invent.Object(Slot).ObjIndex) Then
            'Si esta equipado lo quita
            If UserList(userIndex).Invent.Object(Slot).Equipped Then
                Call Desequipar(userIndex, Slot)
                If UserList(userIndex).flags.Mimetizado = 1 Then
                    UserList(userIndex).CharMimetizado.CascoAnim = NingunCasco
                Else
                    UserList(userIndex).Char.CascoAnim = NingunCasco
                    Call ChangeUserChar(userIndex, UserList(userIndex).Char.body, UserList(userIndex).Char.head, UserList(userIndex).Char.Heading, UserList(userIndex).Char.WeaponAnim, UserList(userIndex).Char.ShieldAnim, UserList(userIndex).Char.CascoAnim)
                End If
                Exit Sub
            End If
    
            'Quita el anterior
            If UserList(userIndex).Invent.CascoEqpObjIndex > 0 Then
                Call Desequipar(userIndex, UserList(userIndex).Invent.CascoEqpSlot)
            End If
    
            'Lo equipa
            
            UserList(userIndex).Invent.Object(Slot).Equipped = 1
            UserList(userIndex).Invent.CascoEqpObjIndex = UserList(userIndex).Invent.Object(Slot).ObjIndex
            UserList(userIndex).Invent.CascoEqpSlot = Slot
            If UserList(userIndex).flags.Mimetizado = 1 Then
                UserList(userIndex).CharMimetizado.CascoAnim = Obj.CascoAnim
            Else
                UserList(userIndex).Char.CascoAnim = Obj.CascoAnim
                Call ChangeUserChar(userIndex, UserList(userIndex).Char.body, UserList(userIndex).Char.head, UserList(userIndex).Char.Heading, UserList(userIndex).Char.WeaponAnim, UserList(userIndex).Char.ShieldAnim, UserList(userIndex).Char.CascoAnim)
            End If
            
            'Si lo equipa, actualizamos labels.
            Call WriteUpdateCasco(userIndex)
        Else
            Call WriteConsoleMsg(userIndex, "Tu Clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
        End If
    
    Case eOBJType.otESCUDO
        If UserList(userIndex).flags.Navegando = 1 Or UserList(userIndex).flags.Montado = 1 Then Exit Sub
         If ClasePuedeUsarItem(userIndex, UserList(userIndex).Invent.Object(Slot).ObjIndex) And _
             FaccionPuedeUsarItem(userIndex, UserList(userIndex).Invent.Object(Slot).ObjIndex) Then

             'Si esta equipado lo quita
             If UserList(userIndex).Invent.Object(Slot).Equipped Then
                 Call Desequipar(userIndex, Slot)
                 If UserList(userIndex).flags.Mimetizado = 1 Then
                     UserList(userIndex).CharMimetizado.ShieldAnim = NingunEscudo
                 Else
                     UserList(userIndex).Char.ShieldAnim = NingunEscudo
                     Call ChangeUserChar(userIndex, UserList(userIndex).Char.body, UserList(userIndex).Char.head, UserList(userIndex).Char.Heading, UserList(userIndex).Char.WeaponAnim, UserList(userIndex).Char.ShieldAnim, UserList(userIndex).Char.CascoAnim)
                 End If
                 Exit Sub
             End If
     
             'Quita el anterior
             If UserList(userIndex).Invent.EscudoEqpObjIndex > 0 Then
                 Call Desequipar(userIndex, UserList(userIndex).Invent.EscudoEqpSlot)
             End If
     
             'Lo equipa
             
             UserList(userIndex).Invent.Object(Slot).Equipped = 1
             UserList(userIndex).Invent.EscudoEqpObjIndex = UserList(userIndex).Invent.Object(Slot).ObjIndex
             UserList(userIndex).Invent.EscudoEqpSlot = Slot
             
             If UserList(userIndex).flags.Mimetizado = 1 Then
                 UserList(userIndex).CharMimetizado.ShieldAnim = Obj.ShieldAnim
             Else
                 UserList(userIndex).Char.ShieldAnim = Obj.ShieldAnim
                 If UserList(userIndex).flags.Montado = 0 Then
                    Call ChangeUserChar(userIndex, UserList(userIndex).Char.body, UserList(userIndex).Char.head, UserList(userIndex).Char.Heading, UserList(userIndex).Char.WeaponAnim, UserList(userIndex).Char.ShieldAnim, UserList(userIndex).Char.CascoAnim)
                 End If
             End If
             
             'Si lo equipa, actualizamos labels.
             Call WriteUpdateEscu(userIndex)
         Else
             Call WriteConsoleMsg(userIndex, "Tu Clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
         End If
End Select

'Actualiza
Call UpdateUserInv(False, userIndex, Slot)

Exit Sub
errhandler:
Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.Number & " - Error Description : " & Err.description)
End Sub

Private Function CheckRazaUsaRopa(ByVal userIndex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo errhandler

'Verifica si la Raza puede usar la ropa
If UserList(userIndex).Raza = eRaza.Humano Or _
   UserList(userIndex).Raza = eRaza.Elfo Or _
   UserList(userIndex).Raza = eRaza.ElfoOscuro Or _
   UserList(userIndex).Raza = eRaza.Orco Then
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
End If

'Solo se habilita la ropa exclusiva para Drows por ahora. Pablo (ToxicWaste)
If (UserList(userIndex).Raza <> eRaza.ElfoOscuro) And ObjData(ItemIndex).RazaDrow Then
    CheckRazaUsaRopa = False
End If

Exit Function
errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal userIndex As Integer, ByVal Slot As Byte)
'*************************************************
'Author: Unknown
'Last modified: 24/01/2007
'Handels the usage of items from inventory box.
'24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legión.
'24/01/2007 Pablo (ToxicWaste) - Utilización nueva de Barco en lvl 20 por Clase Pirata y Pescador.
'*************************************************

Dim Obj As ObjData
Dim ObjIndex As Integer
Dim TargObj As ObjData
Dim MiObj As Obj

If UserList(userIndex).Invent.Object(Slot).amount = 0 Then Exit Sub

Obj = ObjData(UserList(userIndex).Invent.Object(Slot).ObjIndex)

If Obj.Newbie = 1 And Not EsNewbie(userIndex) Then
    Call WriteConsoleMsg(userIndex, "Solo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If Obj.OBJType = eOBJType.otWeapon Then
    
    If Obj.proyectil = 1 Then
        
    If Obj.Boat = 1 And UserList(userIndex).flags.Navegando = 0 Then
        Call WriteConsoleMsg(userIndex, "No puedes utilizar este tipo de armas si no estas en un bote.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
        'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
        If Not IntervaloPermiteUsar(userIndex, False) Then Exit Sub
    Else
        'dagas
        If Not IntervaloPermiteUsar(userIndex) Then Exit Sub
    End If
Else
    If Not IntervaloPermiteUsar(userIndex) Then Exit Sub
End If

ObjIndex = UserList(userIndex).Invent.Object(Slot).ObjIndex
UserList(userIndex).flags.TargetObjInvIndex = ObjIndex
UserList(userIndex).flags.TargetObjInvSlot = Slot

Select Case Obj.OBJType
    Case eOBJType.otUseOnce
        If UserList(userIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'Usa el item
        UserList(userIndex).Stats.MinHam = UserList(userIndex).Stats.MinHam + Obj.MinHam
        If UserList(userIndex).Stats.MinHam > UserList(userIndex).Stats.MaxHam Then _
            UserList(userIndex).Stats.MinHam = UserList(userIndex).Stats.MaxHam
        UserList(userIndex).flags.Hambre = 0
        Call WriteUpdateHungerAndThirst(userIndex)
        'Sonido
        
        If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userIndex, e_SoundIndex.MORFAR_MANZANA)
        Else
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userIndex, e_SoundIndex.SOUND_COMIDA)
        End If
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(userIndex, Slot, 1)
        
        Call UpdateUserInv(False, userIndex, Slot)

    Case eOBJType.otGuita
        If UserList(userIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        UserList(userIndex).Stats.GLD = UserList(userIndex).Stats.GLD + UserList(userIndex).Invent.Object(Slot).amount
        UserList(userIndex).Invent.Object(Slot).amount = 0
        UserList(userIndex).Invent.Object(Slot).ObjIndex = 0
        UserList(userIndex).Invent.NroItems = UserList(userIndex).Invent.NroItems - 1
        
        Call UpdateUserInv(False, userIndex, Slot)
        Call WriteUpdateGold(userIndex)
        
    Case eOBJType.otWeapon
        If UserList(userIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not UserList(userIndex).Stats.MinSta > 0 Then
            Call WriteConsoleMsg(userIndex, "Estas muy cansado", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        
        If ObjData(ObjIndex).proyectil = 1 Then
            'liquid: muevo esto aca adentro, para que solo pida modo combate si estamos por usar el arco
            Call WriteWorkRequestTarget(userIndex, Proyectiles)
        Else
            If UserList(userIndex).flags.TargetObj = Leña Then
                If UserList(userIndex).Invent.Object(Slot).ObjIndex = DAGA Then
                    Call TratarDeHacerFogata(UserList(userIndex).flags.TargetObjMap, _
                         UserList(userIndex).flags.TargetObjX, UserList(userIndex).flags.TargetObjY, userIndex)
                End If
            End If
        End If
        
        'Solo si es herramienta ;) (en realidad si no es ni proyectil ni daga)
        If UserList(userIndex).Invent.Object(Slot).Equipped = 0 Then
            Call WriteConsoleMsg(userIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Select Case ObjIndex
            Case CAÑA_PESCA, RED_PESCA
                Call WriteWorkRequestTarget(userIndex, eSkill.Pesca)
            Case HACHA_LEÑADOR
                Call WriteWorkRequestTarget(userIndex, eSkill.Talar)
            Case PIQUETE_MINERO
                Call WriteWorkRequestTarget(userIndex, eSkill.Mineria)
            Case MARTILLO_HERRERO
                Call WriteWorkRequestTarget(userIndex, eSkill.Herreria)
            Case SERRUCHO_CARPINTERO
                Call EnivarObjConstruibles(userIndex)
                Call WriteShowCarpenterForm(userIndex)
        End Select
        
    
    Case eOBJType.otPociones
        If UserList(userIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'If Not IntervaloPermiteAtacar(UserIndex, False) Then
            'Para evitar el lag...
            'Call WriteConsoleMsg(UserIndex, "¡¡Debes esperar unos momentos para tomar otra pocion!!", FontTypeNames.FONTTYPE_INFO)
        '    Exit Sub
        'End If
        
        UserList(userIndex).flags.TomoPocion = True
        UserList(userIndex).flags.TipoPocion = Obj.TipoPocion
                
        Select Case UserList(userIndex).flags.TipoPocion
        
            Case 1 'Modif la agilidad
                UserList(userIndex).flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then _
                    UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
                If UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad) > 2 * UserList(userIndex).Stats.UserAtributosBackUP(Agilidad) Then UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad) = 2 * UserList(userIndex).Stats.UserAtributosBackUP(Agilidad)
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_BEBER))
        
            Case 2 'Modif la fuerza
                UserList(userIndex).flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then _
                    UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
                If UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) > 2 * UserList(userIndex).Stats.UserAtributosBackUP(Fuerza) Then UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) = 2 * UserList(userIndex).Stats.UserAtributosBackUP(Fuerza)
                
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_BEBER))
                
            Case 3 'Pocion roja, restaura HP
                'Usa el item
                UserList(userIndex).Stats.MinHP = UserList(userIndex).Stats.MinHP + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If UserList(userIndex).Stats.MinHP > UserList(userIndex).Stats.MaxHP Then _
                    UserList(userIndex).Stats.MinHP = UserList(userIndex).Stats.MaxHP
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_BEBER))
            
            Case 4 'Pocion azul, restaura MANA
                'Usa el item
                UserList(userIndex).Stats.MinMAN = UserList(userIndex).Stats.MinMAN + Porcentaje(UserList(userIndex).Stats.MaxMAN, 5)
                If UserList(userIndex).Stats.MinMAN > UserList(userIndex).Stats.MaxMAN Then _
                    UserList(userIndex).Stats.MinMAN = UserList(userIndex).Stats.MaxMAN
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(userIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_BEBER))
                
            Case 5 ' Pocion violeta
                If UserList(userIndex).flags.Envenenado = 1 Then
                    UserList(userIndex).flags.Envenenado = 0
                    Call WriteConsoleMsg(userIndex, "Te has curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
                End If
                'Quitamos del inv el item
                Call QuitarUserInvItem(userIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_BEBER))
            Case 6  ' Pocion Negra
                If UserList(userIndex).flags.Privilegios And PlayerType.User Then
                    Call QuitarUserInvItem(userIndex, Slot, 1)
                    Call UserDie(userIndex)
                    Call WriteConsoleMsg(userIndex, "Sientes un gran mareo y pierdes el conocimiento.", FontTypeNames.FONTTYPE_FIGHT)
                End If
       End Select
       
       Call WriteUpdateUserStats(userIndex)
       Call UpdateUserInv(False, userIndex, Slot)
       Call WriteUpdateStrengthAgility(userIndex)
       
     Case eOBJType.otBebidas
        If UserList(userIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        UserList(userIndex).Stats.MinAGU = UserList(userIndex).Stats.MinAGU + Obj.MinSed
        If UserList(userIndex).Stats.MinAGU > UserList(userIndex).Stats.MaxAGU Then _
            UserList(userIndex).Stats.MinAGU = UserList(userIndex).Stats.MaxAGU
        UserList(userIndex).flags.Sed = 0
        Call WriteUpdateHungerAndThirst(userIndex)
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(userIndex, Slot, 1)
        
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_BEBER))
        
        Call UpdateUserInv(False, userIndex, Slot)
    
    Case eOBJType.otLlaves
        If UserList(userIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(userIndex).flags.TargetObj = 0 Then Exit Sub
        TargObj = ObjData(UserList(userIndex).flags.TargetObj)
        '¿El objeto clickeado es una puerta?
        If TargObj.OBJType = eOBJType.otPuertas Then
            '¿Esta cerrada?
            If TargObj.Cerrada = 1 Then
                  '¿Cerrada con llave?
                  If TargObj.Llave > 0 Then
                     If TargObj.clave = Obj.clave Then
         
                        MapData(UserList(userIndex).flags.TargetObjMap, UserList(userIndex).flags.TargetObjX, UserList(userIndex).flags.TargetObjY).ObjInfo.ObjIndex _
                        = ObjData(MapData(UserList(userIndex).flags.TargetObjMap, UserList(userIndex).flags.TargetObjX, UserList(userIndex).flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
                        UserList(userIndex).flags.TargetObj = MapData(UserList(userIndex).flags.TargetObjMap, UserList(userIndex).flags.TargetObjX, UserList(userIndex).flags.TargetObjY).ObjInfo.ObjIndex
                        Call WriteConsoleMsg(userIndex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     Else
                        Call WriteConsoleMsg(userIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     End If
                  Else
                     If TargObj.clave = Obj.clave Then
                        MapData(UserList(userIndex).flags.TargetObjMap, UserList(userIndex).flags.TargetObjX, UserList(userIndex).flags.TargetObjY).ObjInfo.ObjIndex _
                        = ObjData(MapData(UserList(userIndex).flags.TargetObjMap, UserList(userIndex).flags.TargetObjX, UserList(userIndex).flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
                        Call WriteConsoleMsg(userIndex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                        UserList(userIndex).flags.TargetObj = MapData(UserList(userIndex).flags.TargetObjMap, UserList(userIndex).flags.TargetObjX, UserList(userIndex).flags.TargetObjY).ObjInfo.ObjIndex
                        Exit Sub
                     Else
                        Call WriteConsoleMsg(userIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     End If
                  End If
            Else
                  Call WriteConsoleMsg(userIndex, "No esta cerrada.", FontTypeNames.FONTTYPE_INFO)
                  Exit Sub
            End If
        End If
    
    Case eOBJType.otBotellaVacia
        If UserList(userIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Not HayAgua(UserList(userIndex).pos.Map, UserList(userIndex).flags.targetX, UserList(userIndex).flags.targetY) Then
            Call WriteConsoleMsg(userIndex, "No hay agua allí.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        MiObj.amount = 1
        MiObj.ObjIndex = ObjData(UserList(userIndex).Invent.Object(Slot).ObjIndex).IndexAbierta
        Call QuitarUserInvItem(userIndex, Slot, 1)
        If Not MeterItemEnInventario(userIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(userIndex).pos, MiObj)
        End If
        
        Call UpdateUserInv(False, userIndex, Slot)
    
    Case eOBJType.otBotellaLlena
        If UserList(userIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        UserList(userIndex).Stats.MinAGU = UserList(userIndex).Stats.MinAGU + Obj.MinSed
        If UserList(userIndex).Stats.MinAGU > UserList(userIndex).Stats.MaxAGU Then _
            UserList(userIndex).Stats.MinAGU = UserList(userIndex).Stats.MaxAGU
        UserList(userIndex).flags.Sed = 0
        Call WriteUpdateHungerAndThirst(userIndex)
        MiObj.amount = 1
        MiObj.ObjIndex = ObjData(UserList(userIndex).Invent.Object(Slot).ObjIndex).IndexCerrada
        Call QuitarUserInvItem(userIndex, Slot, 1)
        If Not MeterItemEnInventario(userIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(userIndex).pos, MiObj)
        End If
        
        Call UpdateUserInv(False, userIndex, Slot)
    
    Case eOBJType.otPergaminos
        If UserList(userIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(userIndex).Stats.MaxMAN > 0 Then
            If UserList(userIndex).flags.Hambre = 0 And _
                UserList(userIndex).flags.Sed = 0 Then
                Call AgregarHechizo(userIndex, Slot)
                Call UpdateUserInv(False, userIndex, Slot)
            Else
                Call WriteConsoleMsg(userIndex, "Estas demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(userIndex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_INFO)
        End If
    Case eOBJType.otMinerales
        If UserList(userIndex).flags.Muerto = 1 Then
             Call WriteConsoleMsg(userIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        Call WriteWorkRequestTarget(userIndex, FundirMetal)
       
    Case eOBJType.otInstrumentos
        If UserList(userIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Obj.Real Then '¿Es el Cuerno Real?
            If FaccionPuedeUsarItem(userIndex, ObjIndex) Then
                If MapInfo(UserList(userIndex).pos.Map).Pk = False Then
                    Call WriteConsoleMsg(userIndex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call SendData(SendTarget.toMap, UserList(userIndex).pos.Map, PrepareMessagePlayWave(Obj.Snd1))
                Exit Sub
            End If
        ElseIf Obj.Caos Then '¿Es el Cuerno Legión?
            If FaccionPuedeUsarItem(userIndex, ObjIndex) Then
                If MapInfo(UserList(userIndex).pos.Map).Pk = False Then
                    Call WriteConsoleMsg(userIndex, "No hay Peligro aquí. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call SendData(SendTarget.toMap, UserList(userIndex).pos.Map, PrepareMessagePlayWave(Obj.Snd1))
                Exit Sub
            End If
        End If
        'Si llega aca es porque es o Laud o Tambor o Flauta
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(Obj.Snd1))
       
    Case eOBJType.otBarcos
        'Verifica si esta aproximado al agua antes de permitirle navegar
        If UserList(userIndex).Stats.ELV < 25 Then
            If UserList(userIndex).Clase <> eClass.Fisher And UserList(userIndex).Clase <> eClass.Pirat Then
                Call WriteConsoleMsg(userIndex, "Para recorrer los mares debes ser nivel 25 o superior.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
                If UserList(userIndex).Stats.ELV < 20 Then
                    Call WriteConsoleMsg(userIndex, "Para recorrer los mares debes ser nivel 20 o superior.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
        End If
        

        
        If ((LegalPos(UserList(userIndex).pos.Map, UserList(userIndex).pos.X - 1, UserList(userIndex).pos.Y, True, False) _
                Or LegalPos(UserList(userIndex).pos.Map, UserList(userIndex).pos.X, UserList(userIndex).pos.Y - 1, True, False) _
                Or LegalPos(UserList(userIndex).pos.Map, UserList(userIndex).pos.X + 1, UserList(userIndex).pos.Y, True, False) _
                Or LegalPos(UserList(userIndex).pos.Map, UserList(userIndex).pos.X, UserList(userIndex).pos.Y + 1, True, False)) _
                And UserList(userIndex).flags.Navegando = 0) _
                Or UserList(userIndex).flags.Navegando = 1 Then
            Call DoNavega(userIndex, Obj, Slot)
        Else
            Call WriteConsoleMsg(userIndex, "¡Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)
        End If
        
    Case eOBJType.otMontura
    'Verifica todo lo que requiere la montura
        If UserList(userIndex).Stats.ELV < 15 Then
                Call WriteConsoleMsg(userIndex, "Para montar debes ser nivel 15 o superior.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
        End If
        Call DoMontar(userIndex, Obj, Slot)
        
    Case eOBJType.otPasajes
        'Por si se traba el muy gil.
        'If UserList(UserIndex).flags.Muerto = 1 Then
        '    Call WriteConsoleMsg(UserIndex, "No puedes usar este item estando muerto.", FontTypeNames.FONTTYPE_INFO)
        '    Exit Sub
        'End If
        If UserList(userIndex).pos.Map <> ObjData(ObjIndex).OrigMap Then
            Call WriteConsoleMsg(userIndex, "Debes estar en " & MapInfo(ObjData(ObjIndex).OrigMap).name & " para utilizar este pasaje.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Call WarpUserChar(userIndex, ObjData(ObjIndex).pos.Map, ObjData(ObjIndex).pos.X, ObjData(ObjIndex).pos.Y, True)
        Call QuitarUserInvItem(userIndex, Slot, 1)
        Call UpdateUserInv(False, userIndex, Slot)
End Select

End Sub

Sub EnivarArmasConstruibles(ByVal userIndex As Integer)

Call WriteBlacksmithWeapons(userIndex)

End Sub
Sub EnivarCascosConstruibles(ByVal userIndex As Integer)

Call WriteBlacksmithHelmets(userIndex)

End Sub
Sub EnivarEscudosConstruibles(ByVal userIndex As Integer)

Call WriteBlacksmithShields(userIndex)

End Sub
 
Sub EnivarObjConstruibles(ByVal userIndex As Integer)

Call WriteCarpenterObjects(userIndex)

End Sub

Sub EnivarArmadurasConstruibles(ByVal userIndex As Integer)

Call WriteBlacksmithArmors(userIndex)

End Sub

Sub TirarTodo(ByVal userIndex As Integer)
On Error Resume Next

If MapData(UserList(userIndex).pos.Map, UserList(userIndex).pos.X, UserList(userIndex).pos.Y).trigger = 6 Then Exit Sub

Call TirarTodosLosItems(userIndex)

Dim Cantidad As Long
Cantidad = UserList(userIndex).Stats.GLD

If Cantidad < 100000 Then _
    Call TirarOro(Cantidad, userIndex)

End Sub

Public Function ItemSeCae(ByVal index As Integer) As Boolean

ItemSeCae = (ObjData(index).Real <> 1 Or ObjData(index).NoSeCae = 0) And _
            (ObjData(index).Caos <> 1 Or ObjData(index).NoSeCae = 0) And _
            ObjData(index).OBJType <> eOBJType.otLlaves And _
            ObjData(index).OBJType <> eOBJType.otBarcos And _
            ObjData(index).NoSeCae = 0 And _
            ObjData(index).OBJType <> eOBJType.otMontura


End Function

Sub TirarTodosLosItems(ByVal userIndex As Integer)
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    
    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(userIndex).Invent.Object(i).ObjIndex
        If ItemIndex > 0 Then
             If ItemSeCae(ItemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                
                'Creo el Obj
                MiObj.amount = UserList(userIndex).Invent.Object(i).amount
                MiObj.ObjIndex = ItemIndex
                'Pablo (ToxicWaste) 24/01/2007
                'Si es pirata y usa un Galeón entonces no explota los items. (en el agua)
                If UserList(userIndex).Clase = eClass.Pirat And UserList(userIndex).Invent.BarcoObjIndex = 476 Then
                    Tilelibre UserList(userIndex).pos, NuevaPos, MiObj, False, True
                Else
                    Tilelibre UserList(userIndex).pos, NuevaPos, MiObj, True, True
                End If
                
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    Call DropObj(userIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                End If
             End If
        End If
    Next i
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean

ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function

Sub TirarTodosLosItemsNoNewbies(ByVal userIndex As Integer)
Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As Obj
Dim ItemIndex As Integer

If MapData(UserList(userIndex).pos.Map, UserList(userIndex).pos.X, UserList(userIndex).pos.Y).trigger = 6 Then Exit Sub

For i = 1 To MAX_INVENTORY_SLOTS
    ItemIndex = UserList(userIndex).Invent.Object(i).ObjIndex
    If ItemIndex > 0 Then
        If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
            NuevaPos.X = 0
            NuevaPos.Y = 0
            
            'Creo MiObj
            MiObj.amount = UserList(userIndex).Invent.Object(i).ObjIndex
            MiObj.ObjIndex = ItemIndex
            'Pablo (ToxicWaste) 24/01/2007
            'Tira los Items no newbies en todos lados.
            Tilelibre UserList(userIndex).pos, NuevaPos, MiObj, True, True
            If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                If MapData(NuevaPos.Map, NuevaPos.X, NuevaPos.Y).ObjInfo.ObjIndex = 0 Then Call DropObj(userIndex, i, MAX_INVENTORY_OBJS, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
            End If
        End If
    End If
Next i

End Sub



Public Function userCheckAmmo(ByVal userIndex As Integer) As Boolean
    Dim dummyInt As Integer

    'Make sure the item is valid and there is ammo equipped.
    With UserList(userIndex).Invent
        If .WeaponEqpObjIndex = 0 Then
            dummyInt = 1
        ElseIf .WeaponEqpSlot < 1 Or .WeaponEqpSlot > MAX_INVENTORY_SLOTS Then
            dummyInt = 1
        ElseIf .MunicionEqpSlot < 1 Or .MunicionEqpSlot > MAX_INVENTORY_SLOTS Then
            dummyInt = 1
        ElseIf .MunicionEqpObjIndex = 0 Then
            dummyInt = 1
        ElseIf ObjData(.WeaponEqpObjIndex).proyectil <> 1 Then
            dummyInt = 2
        ElseIf ObjData(.MunicionEqpObjIndex).OBJType <> eOBJType.otFlechas Then
            dummyInt = 1
        ElseIf .Object(.MunicionEqpSlot).amount < 1 Then
            dummyInt = 1
        ElseIf ObjData(.WeaponEqpObjIndex).Boat <> ObjData(.MunicionEqpObjIndex).Boat Then
            dummyInt = 1
        End If
            
        If dummyInt <> 0 Then
            If dummyInt = 1 Then
                Call WriteConsoleMsg(userIndex, "No tenés municiones.", FontTypeNames.FONTTYPE_INFO)
                Call Desequipar(userIndex, .WeaponEqpSlot)
            End If
            Call Desequipar(userIndex, .MunicionEqpSlot)
            userCheckAmmo = False
        Else
            userCheckAmmo = True
        End If
    End With
End Function

