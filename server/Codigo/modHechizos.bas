Attribute VB_Name = "modHechizos"
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

Public Const HELEMENTAL_FUEGO As Integer = 26
Public Const HELEMENTAL_TIERRA As Integer = 28
Public Const SUPERANILLO As Integer = 700



Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal userIndex As Integer, ByVal Spell As Integer)

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
If UserList(userIndex).flags.invisible = 1 Or UserList(userIndex).flags.Oculto = 1 Then Exit Sub

Npclist(NpcIndex).CanAttack = 0
Dim daño As Integer

If Hechizos(Spell).SubeHP = 1 Then

    daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV))
    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCreateFX(UserList(userIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))

    UserList(userIndex).Stats.MinHP = UserList(userIndex).Stats.MinHP + daño
    If UserList(userIndex).Stats.MinHP > UserList(userIndex).Stats.MaxHP Then UserList(userIndex).Stats.MinHP = UserList(userIndex).Stats.MaxHP
    
    Call WriteConsoleMsg(userIndex, Npclist(NpcIndex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    Call WriteUpdateUserStats(userIndex)

ElseIf Hechizos(Spell).SubeHP = 2 Then
    
    If UserList(userIndex).flags.Privilegios And PlayerType.User Then
    
        daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        
        If UserList(userIndex).Invent.CascoEqpObjIndex > 0 Then
            daño = daño - RandomNumber(ObjData(UserList(userIndex).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(userIndex).Invent.CascoEqpObjIndex).DefensaMagicaMax)
        End If
        
        If UserList(userIndex).Invent.AnilloEqpObjIndex > 0 Then
            daño = daño - RandomNumber(ObjData(UserList(userIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(UserList(userIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMax)
        End If
        
        If daño < 0 Then daño = 0
        
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV))
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCreateFX(UserList(userIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
    
        UserList(userIndex).Stats.MinHP = UserList(userIndex).Stats.MinHP - daño
        
        Call WriteConsoleMsg(userIndex, Npclist(NpcIndex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteUpdateUserStats(userIndex)
        
        'Muere
        If UserList(userIndex).Stats.MinHP < 1 Then
            UserList(userIndex).Stats.MinHP = 0
            Call UserDie(userIndex)
            '[Barrin 1-12-03]
            If Npclist(NpcIndex).MaestroUser > 0 Then
                'Store it!
                Call Statistics.StoreFrag(Npclist(NpcIndex).MaestroUser, userIndex)
                Call ContarMuerte(userIndex, Npclist(NpcIndex).MaestroUser)
                Call ActStats(userIndex, Npclist(NpcIndex).MaestroUser)
            End If
            '[/Barrin]
        End If
    
    End If
    
End If

If Hechizos(Spell).Paraliza = 1 Then
     If UserList(userIndex).flags.Paralizado = 0 Then
          Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(Hechizos(Spell).WAV))
          Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCreateFX(UserList(userIndex).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
          
            If UserList(userIndex).Invent.AnilloEqpObjIndex = SUPERANILLO Then
                Call WriteConsoleMsg(userIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
          UserList(userIndex).flags.Paralizado = 1
          UserList(userIndex).Counters.Paralisis = IntervaloParalizado
          
          Call WriteParalizeOK(userIndex)
     End If
     
     
End If


End Sub


Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal Spell As Integer)
'solo hechizos ofensivos!

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
Npclist(NpcIndex).CanAttack = 0

Dim daño As Integer

If Hechizos(Spell).SubeHP = 2 Then
    
        daño = RandomNumber(Hechizos(Spell).MinHP, Hechizos(Spell).MaxHP)
        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(Spell).WAV))
        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(Npclist(TargetNPC).Char.CharIndex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
        
        Npclist(TargetNPC).Stats.MinHP = Npclist(TargetNPC).Stats.MinHP - daño
        
        'Muere
        If Npclist(TargetNPC).Stats.MinHP < 1 Then
            Npclist(TargetNPC).Stats.MinHP = 0
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
            Else
                Call MuereNpc(TargetNPC, 0)
            End If
        End If
    
End If
    
End Sub



Function TieneHechizo(ByVal i As Integer, ByVal userIndex As Integer) As Boolean

On Error GoTo errhandler
    
    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(userIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

Exit Function
errhandler:

End Function

Sub AgregarHechizo(ByVal userIndex As Integer, ByVal Slot As Integer)
Dim hIndex As Integer
Dim j As Integer
hIndex = ObjData(UserList(userIndex).Invent.Object(Slot).ObjIndex).HechizoIndex

If Not TieneHechizo(hIndex, userIndex) Then
    'Buscamos un slot vacio
    For j = 1 To MAXUSERHECHIZOS
        If UserList(userIndex).Stats.UserHechizos(j) = 0 Then Exit For
    Next j
        
    If UserList(userIndex).Stats.UserHechizos(j) <> 0 Then
        Call WriteConsoleMsg(userIndex, "No tenes espacio para mas hechizos.", FontTypeNames.FONTTYPE_INFO)
    Else
        UserList(userIndex).Stats.UserHechizos(j) = hIndex
        Call UpdateUserHechizos(False, userIndex, CByte(j))
        'Quitamos del inv el item
        Call QuitarUserInvItem(userIndex, CByte(Slot), 1)
    End If
Else
    Call WriteConsoleMsg(userIndex, "Ya tenes ese hechizo.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub
            
Sub DecirPalabrasMagicas(ByVal S As String, ByVal userIndex As Integer)
On Error Resume Next
    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageChatOverHead(S, UserList(userIndex).Char.CharIndex, vbCyan))
    Exit Sub
End Sub

Function PuedeLanzar(ByVal userIndex As Integer, ByVal HechizoIndex As Integer) As Boolean

If UserList(userIndex).flags.Muerto = 0 Then
    Dim wp2 As WorldPos
    wp2.Map = UserList(userIndex).flags.TargetMap
    wp2.X = UserList(userIndex).flags.targetX
    wp2.Y = UserList(userIndex).flags.targetY
    
    If Hechizos(HechizoIndex).NeedStaff > 0 Then
        If UserList(userIndex).Clase = eClass.Mage Then
            If UserList(userIndex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(userIndex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                    Call WriteConsoleMsg(userIndex, "Tu Báculo no es lo suficientemente poderoso para que puedas lanzar el conjuro.", FontTypeNames.FONTTYPE_INFO)
                    PuedeLanzar = False
                    Exit Function
                End If
            Else
                Call WriteConsoleMsg(userIndex, "No puedes lanzar este conjuro sin la ayuda de un báculo.", FontTypeNames.FONTTYPE_INFO)
                PuedeLanzar = False
                Exit Function
            End If
        End If
    End If
        
    If UserList(userIndex).Stats.MinMAN >= Hechizos(HechizoIndex).ManaRequerido Then
        If UserList(userIndex).Stats.UserSkills(eSkill.Magia) >= Hechizos(HechizoIndex).MinSkill Then
            If UserList(userIndex).Stats.MinSta >= Hechizos(HechizoIndex).StaRequerido Then
                PuedeLanzar = True
            Else
                Call WriteConsoleMsg(userIndex, "Estás muy cansado para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
                PuedeLanzar = False
            End If
                
        Else
            Call WriteConsoleMsg(userIndex, "No tenes suficientes puntos de magia para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
            PuedeLanzar = False
        End If
    Else
            Call WriteConsoleMsg(userIndex, "No tenes suficiente mana.", FontTypeNames.FONTTYPE_INFO)
            PuedeLanzar = False
    End If
Else
   Call WriteConsoleMsg(userIndex, "No podes lanzar hechizos porque estas muerto.", FontTypeNames.FONTTYPE_INFO)
   PuedeLanzar = False
End If

End Function

Sub HechizoTerrenoEstado(ByVal userIndex As Integer, ByRef b As Boolean)
Dim PosCasteadaX As Integer
Dim PosCasteadaY As Integer
Dim PosCasteadaM As Integer
Dim H As Integer
Dim TempX As Integer
Dim TempY As Integer


    PosCasteadaX = UserList(userIndex).flags.targetX
    PosCasteadaY = UserList(userIndex).flags.targetY
    PosCasteadaM = UserList(userIndex).flags.TargetMap
    
    H = UserList(userIndex).Stats.UserHechizos(UserList(userIndex).flags.Hechizo)
    
    If Hechizos(H).RemueveInvisibilidadParcial = 1 Then
        b = True
        For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
            For TempY = PosCasteadaY - 8 To PosCasteadaY + 8
                If InMapBounds(PosCasteadaM, TempX, TempY) Then
                    If MapData(PosCasteadaM, TempX, TempY).userIndex > 0 Then
                        'hay un user
                        If UserList(MapData(PosCasteadaM, TempX, TempY).userIndex).flags.invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).userIndex).flags.AdminInvisible = 0 Then
                            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCreateFX(UserList(MapData(PosCasteadaM, TempX, TempY).userIndex).Char.CharIndex, Hechizos(H).FXgrh, Hechizos(H).loops))
                        End If
                    End If
                End If
            Next TempY
        Next TempX
    
        Call InfoHechizo(userIndex)
    End If

End Sub

Sub HechizoInvocacion(ByVal userIndex As Integer, ByRef b As Boolean)

If UserList(userIndex).NroMacotas >= MAXMASCOTAS Then Exit Sub

'No permitimos se invoquen criaturas en zonas seguras
If MapInfo(UserList(userIndex).Pos.Map).Pk = False Or MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
    Call WriteConsoleMsg(userIndex, "En zona segura no puedes invocar criaturas.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

Dim H As Integer, j As Integer, ind As Integer, Index As Integer
Dim TargetPos As WorldPos


TargetPos.Map = UserList(userIndex).flags.TargetMap
TargetPos.X = UserList(userIndex).flags.targetX
TargetPos.Y = UserList(userIndex).flags.targetY

H = UserList(userIndex).Stats.UserHechizos(UserList(userIndex).flags.Hechizo)
    
    
For j = 1 To Hechizos(H).cant
    
    If UserList(userIndex).NroMacotas < MAXMASCOTAS Then
        ind = SpawnNpc(Hechizos(H).NumNpc, TargetPos, True, False)
        If ind > 0 Then
            UserList(userIndex).NroMacotas = UserList(userIndex).NroMacotas + 1
            
            Index = FreeMascotaIndex(userIndex)
            
            UserList(userIndex).MascotasIndex(Index) = ind
            UserList(userIndex).MascotasType(Index) = Npclist(ind).Numero
            
            Npclist(ind).MaestroUser = userIndex
            Npclist(ind).Contadores.TiempoExistencia = IntervaloInvocacion
            Npclist(ind).GiveGlD = 0
            
            Call FollowAmo(ind)
        End If
            
    Else
        Exit For
    End If
    
Next j


Call InfoHechizo(userIndex)
b = True


End Sub

Sub HandleHechizoTerreno(ByVal userIndex As Integer, ByVal uh As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 01/10/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'Antes de procesar cualquier hechizo chequea de que este en modo de combate el
'usuario
'***************************************************

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uInvocacion '
        Call HechizoInvocacion(userIndex, b)
    Case TipoHechizo.uEstado
        Call HechizoTerrenoEstado(userIndex, b)
    
End Select

If b Then
    Call SubirSkill(userIndex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(userIndex).Stats.MinMAN = UserList(userIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(userIndex).Stats.MinMAN < 0 Then UserList(userIndex).Stats.MinMAN = 0
    UserList(userIndex).Stats.MinSta = UserList(userIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(userIndex).Stats.MinSta < 0 Then UserList(userIndex).Stats.MinSta = 0
    Call WriteUpdateUserStats(userIndex)
End If


End Sub

Sub HandleHechizoUsuario(ByVal userIndex As Integer, ByVal uh As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 01/10/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'Antes de procesar cualquier hechizo chequea de que este en modo de combate el
'usuario
'***************************************************

Dim b As Boolean
Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
       Call HechizoEstadoUsuario(userIndex, b)
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
       Call HechizoPropUsuario(userIndex, b)
End Select
If b Then
    Call SubirSkill(userIndex, Magia)
    'If Hechizos(uh).Resis = 1 Then Call SubirSkill(UserList(UserIndex).Flags.TargetUser, Resis)
    UserList(userIndex).Stats.MinMAN = UserList(userIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(userIndex).Stats.MinMAN < 0 Then UserList(userIndex).Stats.MinMAN = 0
    UserList(userIndex).Stats.MinSta = UserList(userIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(userIndex).Stats.MinSta < 0 Then UserList(userIndex).Stats.MinSta = 0
    Call WriteUpdateUserStats(userIndex)
    Call WriteUpdateUserStats(UserList(userIndex).flags.TargetUser)
    UserList(userIndex).flags.TargetUser = 0
End If

End Sub

Sub HandleHechizoNPC(ByVal userIndex As Integer, ByVal uh As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 01/10/07
'Last Modified By: Agustin Andreucci (Blizzard)
'Antes de procesar cualquier hechizo chequea de que este en modo de combate el
'usuario
'Antes de procesar hechizo se fija si puede atacar al npc. (Fuertes)
'***************************************************

Dim b As Boolean

Select Case Hechizos(uh).Tipo
    Case TipoHechizo.uEstado ' Afectan estados (por ejem : Envenenamiento)
        Call HechizoEstadoNPC(UserList(userIndex).flags.TargetNPC, uh, b, userIndex)
    Case TipoHechizo.uPropiedades ' Afectan HP,MANA,STAMINA,ETC
        Call HechizoPropNPC(uh, UserList(userIndex).flags.TargetNPC, userIndex, b)
End Select

If b Then
    Call SubirSkill(userIndex, Magia)
    UserList(userIndex).flags.TargetNPC = 0
    UserList(userIndex).Stats.MinMAN = UserList(userIndex).Stats.MinMAN - Hechizos(uh).ManaRequerido
    If UserList(userIndex).Stats.MinMAN < 0 Then UserList(userIndex).Stats.MinMAN = 0
    UserList(userIndex).Stats.MinSta = UserList(userIndex).Stats.MinSta - Hechizos(uh).StaRequerido
    If UserList(userIndex).Stats.MinSta < 0 Then UserList(userIndex).Stats.MinSta = 0
    Call WriteUpdateUserStats(userIndex)
End If

End Sub


Sub userCastSpell(Index As Integer, userIndex As Integer)

Dim uh As Integer
Dim exito As Boolean

uh = UserList(userIndex).Stats.UserHechizos(Index)
If PuedeLanzar(userIndex, uh) Then
    Select Case Hechizos(uh).target
        Case spellTargetType.uUsuarios
            If UserList(userIndex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(userIndex).flags.TargetUser).Pos.Y - UserList(userIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(userIndex, uh)
                Else
                    Call WriteConsoleMsg(userIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                End If
            Else
                Call WriteConsoleMsg(userIndex, "Este hechizo actua solo sobre usuarios.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        Case spellTargetType.uNPC
            If UserList(userIndex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(userIndex).flags.TargetNPC).Pos.Y - UserList(userIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(userIndex, uh)
                Else
                    Call WriteConsoleMsg(userIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                End If
            Else
                Call WriteConsoleMsg(userIndex, "Este hechizo solo afecta a los npcs.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        Case spellTargetType.uUsuariosYnpc
            If UserList(userIndex).flags.TargetUser > 0 Then
                If Abs(UserList(UserList(userIndex).flags.TargetUser).Pos.Y - UserList(userIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoUsuario(userIndex, uh)
                Else
                    Call WriteConsoleMsg(userIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                End If
            ElseIf UserList(userIndex).flags.TargetNPC > 0 Then
                If Abs(Npclist(UserList(userIndex).flags.TargetNPC).Pos.Y - UserList(userIndex).Pos.Y) <= RANGO_VISION_Y Then
                    Call HandleHechizoNPC(userIndex, uh)
                Else
                    Call WriteConsoleMsg(userIndex, "Estas demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                End If
            Else
                Call WriteConsoleMsg(userIndex, "Target invalido.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        Case spellTargetType.uTerreno
            Call HandleHechizoTerreno(userIndex, uh)
    End Select
    
End If

If UserList(userIndex).Counters.Trabajando Then _
    UserList(userIndex).Counters.Trabajando = UserList(userIndex).Counters.Trabajando - 1

If UserList(userIndex).Counters.Ocultando Then _
    UserList(userIndex).Counters.Ocultando = UserList(userIndex).Counters.Ocultando - 1
    
End Sub

Sub HechizoEstadoUsuario(ByVal userIndex As Integer, ByRef b As Boolean)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 24/01/2007
'Handles the Spells that afect the Stats of an User
'24/01/2007 Pablo (ToxicWaste) - Invisibilidad no permitida en Mapas con InviSinEfecto
'26/01/2007 Pablo (ToxicWaste) - Cambios que permiten mejor manejo de ataques en los rings.
'26/01/2007 Pablo (ToxicWaste) - Revivir no permitido en Mapas con ResuSinEfecto
'***************************************************


Dim H As Integer, tU As Integer

H = UserList(userIndex).Stats.UserHechizos(UserList(userIndex).flags.Hechizo)

tU = UserList(userIndex).flags.TargetUser

If Hechizos(H).Invisibilidad = 1 Then
   
    If UserList(tU).flags.Muerto = 1 Then
        Call WriteConsoleMsg(userIndex, "¡Está muerto!", FontTypeNames.FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
    
    'No usar invi mapas InviSinEfecto
    If MapInfo(UserList(tU).Pos.Map).InviSinEfecto > 0 Then
        Call WriteConsoleMsg(userIndex, "¡La invisibilidad no funciona aquí!", FontTypeNames.FONTTYPE_INFO)
        b = False
        Exit Sub
    End If
    
    'Para poder tirar invi a un pk en el ring
    If (TriggerZonaPelea(userIndex, tU) <> TRIGGER6_PERMITE) Then
        If Faccion(tU) = e_Alineacion.Caos And Faccion(userIndex) = e_Alineacion.Real Then
            Call WriteConsoleMsg(userIndex, "No puedes ayudar a miembros de la faccion del Caos.", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        Else
            b = True
        End If
    End If
    
   
    UserList(tU).flags.invisible = 1
    Call SendData(SendTarget.ToPCArea, tU, PrepareMessageSetInvisible(UserList(tU).Char.CharIndex, True))

    Call InfoHechizo(userIndex)
    b = True
End If

If Hechizos(H).Mimetiza = 1 Then
    If UserList(tU).flags.Muerto = 1 Then
        Exit Sub
    End If
    
    If UserList(tU).flags.Navegando = 1 Then
        Exit Sub
    End If
    
    If UserList(tU).flags.Montado = 1 Then
        Exit Sub
    End If
    
    If UserList(userIndex).flags.Montado = 1 Then
        Exit Sub
    End If
    
    If UserList(userIndex).flags.Navegando = 1 Then
        Exit Sub
    End If
    
    If Not UserList(tU).flags.Privilegios And PlayerType.User Then
        Exit Sub
    End If
    
    If UserList(userIndex).flags.Mimetizado = 1 Then
        Call WriteConsoleMsg(userIndex, "Ya te encuentras transformado. El hechizo no ha tenido efecto", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(userIndex).flags.AdminInvisible = 1 Then Exit Sub
    
    'copio el char original al mimetizado
    
    With UserList(userIndex)
        .CharMimetizado.body = .Char.body
        .CharMimetizado.head = .Char.head
        .CharMimetizado.CascoAnim = .Char.CascoAnim
        .CharMimetizado.ShieldAnim = .Char.ShieldAnim
        .CharMimetizado.WeaponAnim = .Char.WeaponAnim
        
        .flags.Mimetizado = 1
        
        'ahora pongo local el del enemigo
        .Char.body = UserList(tU).Char.body
        .Char.head = UserList(tU).Char.head
        .Char.CascoAnim = UserList(tU).Char.CascoAnim
        .Char.ShieldAnim = UserList(tU).Char.ShieldAnim
        .Char.WeaponAnim = UserList(tU).Char.WeaponAnim
    
        Call ChangeUserChar(userIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
    End With
   
   Call InfoHechizo(userIndex)
   b = True
End If


If Hechizos(H).Envenena = 1 Then
        If Not userCanAttackUser(userIndex, tU) Then Exit Sub
        If userIndex <> tU Then
            Call userAttackedUser(userIndex, tU)
        End If
        UserList(tU).flags.Envenenado = 1
        Call InfoHechizo(userIndex)
        b = True
End If

If Hechizos(H).CuraVeneno = 1 Then
    'Para poder tirar curar veneno a un pk en el ring
    If (TriggerZonaPelea(userIndex, tU) <> TRIGGER6_PERMITE) Then
        If userCanAttackUser(userIndex, tU) Then
            b = True
        End If
    End If
        
    UserList(tU).flags.Envenenado = 0
    Call InfoHechizo(userIndex)
    b = True
End If

If Hechizos(H).Maldicion = 1 Then
        If Not userCanAttackUser(userIndex, tU) Then Exit Sub
        If userIndex <> tU Then
            Call userAttackedUser(userIndex, tU)
        End If
        UserList(tU).flags.Maldicion = 1
        Call InfoHechizo(userIndex)
        b = True
End If

If Hechizos(H).RemoverMaldicion = 1 Then
        UserList(tU).flags.Maldicion = 0
        Call InfoHechizo(userIndex)
        b = True
End If

If Hechizos(H).Bendicion = 1 Then
        UserList(tU).flags.Bendicion = 1
        Call InfoHechizo(userIndex)
        b = True
End If

If Hechizos(H).Paraliza = 1 Or Hechizos(H).Inmoviliza = 1 Then
     If UserList(tU).flags.Paralizado = 0 Then
            If Not userCanAttackUser(userIndex, tU) Then Exit Sub
            
            If userIndex <> tU Then
                Call userAttackedUser(userIndex, tU)
            End If
            
            Call InfoHechizo(userIndex)
            b = True
            If UserList(tU).Invent.AnilloEqpObjIndex = SUPERANILLO Then
                Call WriteConsoleMsg(tU, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(userIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
                Call FlushBuffer(tU)
                Exit Sub
            End If
            
            UserList(tU).flags.Paralizado = 1
            UserList(tU).Counters.Paralisis = IntervaloParalizado
            
            Call WriteParalizeOK(tU)
            Call FlushBuffer(tU)

            
    End If
End If


If Hechizos(H).RemoverParalisis = 1 Then
    If UserList(tU).flags.Paralizado = 1 Then
        'Para poder tirar remo a un pk en el ring
        If (TriggerZonaPelea(userIndex, tU) <> TRIGGER6_PERMITE) Then
            If Faccion(tU) = e_Alineacion.Caos And Faccion(userIndex) = e_Alineacion.Real Then
                Call WriteConsoleMsg(userIndex, "No puedes ayudar a miembros de la faccion del Caos.", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
        End If
        
        UserList(tU).flags.Paralizado = 0
        'no need to crypt this
        Call WriteParalizeOK(tU)
        Call InfoHechizo(userIndex)
        b = True
    End If
End If

If Hechizos(H).RemoverEstupidez = 1 Then
    If UserList(tU).flags.Estupidez = 1 Then
        'Para poder tirar remo estu a un pk en el ring
        If (TriggerZonaPelea(userIndex, tU) <> TRIGGER6_PERMITE) Then
            If Not userCanAttackUser(userIndex, tU) Then
                b = True
            End If
        End If
    
        UserList(tU).flags.Estupidez = 0
        'no need to crypt this
        Call WriteDumbNoMore(tU)
        Call FlushBuffer(tU)
        Call InfoHechizo(userIndex)
        b = True
    End If
End If


If Hechizos(H).Revivir = 1 Then
    If UserList(tU).flags.Muerto = 1 Then
    
        'No usar resu en mapas con ResuSinEfecto
        If MapInfo(UserList(tU).Pos.Map).ResuSinEfecto > 0 Then
            Call WriteConsoleMsg(userIndex, "¡Revivir no está permitido aqui! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        
        'No podemos resucitar si nuestra barra de energía no está llena. (GD: 29/04/07)
        If UserList(userIndex).Stats.MaxSta <> UserList(userIndex).Stats.MinSta Then
            Call WriteConsoleMsg(userIndex, "No puedes resucitar si no tienes tu barra de energía llena.", FontTypeNames.FONTTYPE_INFO)
            b = False
            Exit Sub
        End If
        
        'revisamos si necesita vara
        If UserList(userIndex).Clase = eClass.Mage Then
            If UserList(userIndex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(userIndex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(H).NeedStaff Then
                    Call WriteConsoleMsg(userIndex, "Necesitas un mejor báculo para este hechizo", FontTypeNames.FONTTYPE_INFO)
                    b = False
                    Exit Sub
                End If
            End If
        ElseIf UserList(userIndex).Clase = eClass.Bard Then
            If UserList(userIndex).Invent.AnilloEqpObjIndex <> LAUDMAGICO Then
                Call WriteConsoleMsg(userIndex, "Necesitas un instrumento mágico para devolver la vida", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
        End If
        
        'Para poder tirar revivir a un pk en el ring
        If (TriggerZonaPelea(userIndex, tU) <> TRIGGER6_PERMITE) Then
            If Faccion(tU) = e_Alineacion.Caos And Faccion(userIndex) = e_Alineacion.Real Then
                Call WriteConsoleMsg(userIndex, "No puedes ayudar a miembros de la faccion del Caos.", FontTypeNames.FONTTYPE_INFO)
                b = False
                Exit Sub
            End If
        End If
        
        
        'Pablo Toxic Waste (GD: 29/04/07)
        UserList(tU).Stats.MinAGU = 0
        UserList(tU).flags.Sed = 1
        UserList(tU).Stats.MinHam = 0
        UserList(tU).flags.Hambre = 1
        Call WriteUpdateHungerAndThirst(tU)
        Call InfoHechizo(userIndex)
        UserList(tU).Stats.MinMAN = 0
        UserList(tU).Stats.MinSta = 0
        Dim aux As Double
        aux = UserList(tU).Stats.ELV / 100
        aux = UserList(userIndex).Stats.MaxHP * aux
        'Solo saco vida si es User. no quiero que exploten GMs por ahi.
        If UserList(userIndex).flags.Privilegios And PlayerType.User Then
            UserList(userIndex).Stats.MinHP = UserList(userIndex).Stats.MinHP - aux
        End If
        If (UserList(userIndex).Stats.MinHP <= 0) Then
            Call UserDie(userIndex)
            Call WriteConsoleMsg(userIndex, "El esfuerzo de Resucitar fue demasiado grande", FontTypeNames.FONTTYPE_INFO)
            b = False
        Else
            Call WriteConsoleMsg(userIndex, "El esfuerzo de resucitar te ha debilitado", FontTypeNames.FONTTYPE_INFO)
            b = True
        End If
        
        Call RevivirUsuario(tU)
    Else
        b = False
    End If

End If

If Hechizos(H).Ceguera = 1 Then
        If Not userCanAttackUser(userIndex, tU) Then Exit Sub
        If userIndex <> tU Then
            Call userAttackedUser(userIndex, tU)
        End If
        UserList(tU).flags.Ceguera = 1
        UserList(tU).Counters.Ceguera = IntervaloParalizado / 3

        Call WriteBlind(tU)
        Call FlushBuffer(tU)
        Call InfoHechizo(userIndex)
        b = True
End If

If Hechizos(H).Estupidez = 1 Then
        If Not userCanAttackUser(userIndex, tU) Then Exit Sub
        If userIndex <> tU Then
            Call userAttackedUser(userIndex, tU)
        End If
        If UserList(tU).flags.Estupidez = 0 Then
            UserList(tU).flags.Estupidez = 1
            UserList(tU).Counters.Ceguera = IntervaloParalizado
        End If
        Call WriteDumb(tU)
        Call FlushBuffer(tU)

        Call InfoHechizo(userIndex)
        b = True
End If

End Sub

Sub RevisoAtaqueNPC(ByVal NpcIndex As Integer, ByVal userIndex As Integer, ByRef b As Boolean, ByRef ExitSub As Boolean)
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 26/01/2007
'Finds out if the UserIndex can attack the NpcIndex
'***************************************************
    
    'Es guardia caos y lo quiere atacar un caos?
    If Npclist(NpcIndex).NPCtype = eNPCType.Guardiascaos & UserList(userIndex).Faccion.Alineacion = e_Alineacion.Caos Then
        Call WriteConsoleMsg(userIndex, "No puedes atacar Guardias del Caos siendo Legionario", FontTypeNames.FONTTYPE_WARNING)
        b = False
        ExitSub = True
        Exit Sub
    End If
    'Es guardia Real?
    If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
        If Faccion(userIndex) = e_Alineacion.Real Then
            Call WriteConsoleMsg(userIndex, "No atacar Guardias Reales siendo un Soldado Real.", FontTypeNames.FONTTYPE_INFO)
            b = False
            ExitSub = True
            Exit Sub
        End If
    End If
    If Npclist(NpcIndex).MaestroUser > 0 Then 'Es mascota?
        'Puede atacar mascota?
        If Not userCanAttackUser(userIndex, Npclist(NpcIndex).MaestroUser) Then
            Call WriteConsoleMsg(userIndex, "No puedes atacar mascotas de miembros de tu faccion.", FontTypeNames.FONTTYPE_WARNING)
            b = False
            ExitSub = True
            Exit Sub
        End If
    End If

    Call npcAttacked(NpcIndex, userIndex)

End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal hIndex As Integer, ByRef b As Boolean, ByVal userIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 26/01/2007
'Handles the Spells that afect the Stats of an NPC
'26/01/2007 Pablo (ToxicWaste) - Modificaciones por funcionamiento en los Rings y ataque a guardias
'***************************************************
Dim ExitSub As Boolean

If Hechizos(hIndex).Invisibilidad = 1 Then
   Call InfoHechizo(userIndex)
   Npclist(NpcIndex).flags.invisible = 1
   b = True
End If

If Hechizos(hIndex).Envenena = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call WriteConsoleMsg(userIndex, "No podes atacar a ese npc.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
   End If
   
   ExitSub = False
   Call RevisoAtaqueNPC(NpcIndex, userIndex, b, ExitSub)
   If ExitSub = True Then Exit Sub
        
   Call InfoHechizo(userIndex)
   Npclist(NpcIndex).flags.Envenenado = 1
   b = True
End If

If Hechizos(hIndex).CuraVeneno = 1 Then
   Call InfoHechizo(userIndex)
   Npclist(NpcIndex).flags.Envenenado = 0
   b = True
End If

If Hechizos(hIndex).Maldicion = 1 Then
   If Npclist(NpcIndex).Attackable = 0 Then
        Call WriteConsoleMsg(userIndex, "No podes atacar a ese npc.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
   End If
    
    Call InfoHechizo(userIndex)
    Npclist(NpcIndex).flags.Maldicion = 1
    b = True
End If

If Hechizos(hIndex).RemoverMaldicion = 1 Then
   Call InfoHechizo(userIndex)
   Npclist(NpcIndex).flags.Maldicion = 0
   b = True
End If

If Hechizos(hIndex).Bendicion = 1 Then
   Call InfoHechizo(userIndex)
   Npclist(NpcIndex).flags.Bendicion = 1
   b = True
End If

If Hechizos(hIndex).Paraliza = 1 Then
    If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
        
        ExitSub = False
        Call RevisoAtaqueNPC(NpcIndex, userIndex, b, ExitSub)
        If ExitSub = True Then Exit Sub
        
        Call InfoHechizo(userIndex)
        Npclist(NpcIndex).flags.Paralizado = 1
        Npclist(NpcIndex).flags.Inmovilizado = 0
        Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        b = True
    Else
        Call WriteConsoleMsg(userIndex, "El npc es inmune a este hechizo.", FontTypeNames.FONTTYPE_FIGHT)
    End If
End If

'[Barrin 16-2-04]
If Hechizos(hIndex).RemoverParalisis = 1 Then
   If Npclist(NpcIndex).flags.Paralizado = 1 And Npclist(NpcIndex).MaestroUser = userIndex Then
            Call InfoHechizo(userIndex)
            Npclist(NpcIndex).flags.Paralizado = 0
            Npclist(NpcIndex).Contadores.Paralisis = 0
            b = True
   Else
      Call WriteConsoleMsg(userIndex, "Este hechizo solo afecta NPCs que tengan amo.", FontTypeNames.FONTTYPE_WARNING)
   End If
End If
'[/Barrin]
 
If Hechizos(hIndex).Inmoviliza = 1 Then
    If Npclist(NpcIndex).flags.AfectaParalisis = 0 Then
        
        ExitSub = False
        Call RevisoAtaqueNPC(NpcIndex, userIndex, b, ExitSub)
        If ExitSub = True Then Exit Sub
        
        Npclist(NpcIndex).flags.Inmovilizado = 1
        Npclist(NpcIndex).flags.Paralizado = 0
        Npclist(NpcIndex).Contadores.Paralisis = IntervaloParalizado
        Call InfoHechizo(userIndex)
        b = True
    Else
        Call WriteConsoleMsg(userIndex, "El npc es inmune a este hechizo.", FontTypeNames.FONTTYPE_FIGHT)
    End If
End If

End Sub

Sub HechizoPropNPC(ByVal hIndex As Integer, ByVal NpcIndex As Integer, ByVal userIndex As Integer, ByRef b As Boolean)
Dim daño As Long

'Salud
If Hechizos(hIndex).SubeHP = 1 Then
    daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(userIndex).Stats.ELV)
    
    Call InfoHechizo(userIndex)
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP + daño
    If Npclist(NpcIndex).Stats.MinHP > Npclist(NpcIndex).Stats.MaxHP Then _
        Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP
    Call WriteConsoleMsg(userIndex, "Has curado " & daño & " puntos de salud a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
    b = True
ElseIf Hechizos(hIndex).SubeHP = 2 Then
    If Not userCanAttackNPC(userIndex, NpcIndex) Then
        b = False
        Exit Sub
    End If
    
    daño = RandomNumber(Hechizos(hIndex).MinHP, Hechizos(hIndex).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(userIndex).Stats.ELV)

    If Hechizos(hIndex).StaffAffected Then
        If UserList(userIndex).Clase = eClass.Mage Then
            If UserList(userIndex).Invent.WeaponEqpObjIndex > 0 Then
                daño = (daño * (ObjData(UserList(userIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                'Aumenta daño segun el staff-
                'Daño = (Daño* (80 + BonifBáculo)) / 100
            Else
                daño = daño * 0.7 'Baja daño a 70% del original
            End If
        End If
    End If
    If UserList(userIndex).Invent.AnilloEqpObjIndex = LAUDMAGICO Then
        daño = daño * 1.04  'laud magico de los bardos
    End If

    If UserList(userIndex).Invent.WeaponEqpObjIndex = VaraMataDragonesIndex Then
        If Npclist(NpcIndex).NPCtype = eNPCType.DRAGON Then
            daño = Npclist(NpcIndex).Stats.MinHP + Npclist(NpcIndex).Stats.def
            Call QuitarObjetos(VaraMataDragonesIndex, 1, userIndex)
        Else
            daño = 1
        End If
    End If
    
    Call InfoHechizo(userIndex)
    b = True
    Call npcAttacked(NpcIndex, userIndex)
    If Npclist(NpcIndex).flags.Snd2 > 0 Then
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2))
    End If
    
    Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - daño
    Call WriteConsoleMsg(userIndex, "Le has causado " & daño & " puntos de daño a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
    Call CalcularDarExp(userIndex, NpcIndex, daño)

    If Npclist(NpcIndex).Stats.MinHP < 1 Then
        Npclist(NpcIndex).Stats.MinHP = 0
        Call MuereNpc(NpcIndex, userIndex)
    Else
        If Npclist(UserList(userIndex).flags.TargetNPC).EsRey Then
            CastleUnderAttack Npclist(UserList(userIndex).flags.TargetNPC).EsRey
        End If
        Call CheckPets(UserList(userIndex).flags.TargetNPC, userIndex)
    End If
End If

End Sub

Sub InfoHechizo(ByVal userIndex As Integer)


    Dim H As Integer
    
    H = UserList(userIndex).Stats.UserHechizos(UserList(userIndex).flags.Hechizo)
    
    Call DecirPalabrasMagicas(Hechizos(H).PalabrasMagicas, userIndex)
    
    If UserList(userIndex).flags.TargetUser > 0 Then
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCreateFX(UserList(UserList(userIndex).flags.TargetUser).Char.CharIndex, Hechizos(H).FXgrh, Hechizos(H).loops))
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(Hechizos(H).WAV)) 'Esta linea faltaba. Pablo (ToxicWaste)
    ElseIf UserList(userIndex).flags.TargetNPC > 0 Then
        Call SendData(SendTarget.ToNPCArea, UserList(userIndex).flags.TargetNPC, PrepareMessageCreateFX(Npclist(UserList(userIndex).flags.TargetNPC).Char.CharIndex, Hechizos(H).FXgrh, Hechizos(H).loops))
        Call SendData(SendTarget.ToNPCArea, UserList(userIndex).flags.TargetNPC, PrepareMessagePlayWave(Hechizos(H).WAV))
    End If
    
    If UserList(userIndex).flags.TargetUser > 0 Then
        If userIndex <> UserList(userIndex).flags.TargetUser Then
            If UserList(userIndex).showName Then
                Call WriteConsoleMsg(userIndex, Hechizos(H).HechizeroMsg & " " & UserList(UserList(userIndex).flags.TargetUser).name, FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(userIndex, Hechizos(H).HechizeroMsg & " alguien.", FontTypeNames.FONTTYPE_FIGHT)
            End If
            Call WriteConsoleMsg(UserList(userIndex).flags.TargetUser, UserList(userIndex).name & " " & Hechizos(H).TargetMsg, FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(userIndex, Hechizos(H).PropioMsg, FontTypeNames.FONTTYPE_FIGHT)
        End If
    ElseIf UserList(userIndex).flags.TargetNPC > 0 Then
        Call WriteConsoleMsg(userIndex, Hechizos(H).HechizeroMsg & " " & "la criatura.", FontTypeNames.FONTTYPE_FIGHT)
    End If

End Sub

Sub HechizoPropUsuario(ByVal userIndex As Integer, ByRef b As Boolean)

Dim H As Integer
Dim daño As Integer
Dim tempChr As Integer
    
    

H = UserList(userIndex).Stats.UserHechizos(UserList(userIndex).flags.Hechizo)


tempChr = UserList(userIndex).flags.TargetUser
      
      
'Hambre
If Hechizos(H).SubeHam = 1 Then
    
    Call InfoHechizo(userIndex)
    
    daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam + daño
    If UserList(tempChr).Stats.MinHam > UserList(tempChr).Stats.MaxHam Then _
        UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MaxHam
    
    If userIndex <> tempChr Then
        Call WriteConsoleMsg(userIndex, "Le has restaurado " & daño & " puntos de hambre a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(userIndex).name & " te ha restaurado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(userIndex, "Te has restaurado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    Call WriteUpdateHungerAndThirst(tempChr)
    b = True
    
ElseIf Hechizos(H).SubeHam = 2 Then
    If Not userCanAttackUser(userIndex, tempChr) Then Exit Sub
    
    If userIndex <> tempChr Then
        Call userAttackedUser(userIndex, tempChr)
    Else
        Exit Sub
    End If
    
    Call InfoHechizo(userIndex)
    
    daño = RandomNumber(Hechizos(H).MinHam, Hechizos(H).MaxHam)
    
    UserList(tempChr).Stats.MinHam = UserList(tempChr).Stats.MinHam - daño
    
    If UserList(tempChr).Stats.MinHam < 0 Then UserList(tempChr).Stats.MinHam = 0
    
    If userIndex <> tempChr Then
        Call WriteConsoleMsg(userIndex, "Le has quitado " & daño & " puntos de hambre a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(userIndex).name & " te ha quitado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(userIndex, "Te has quitado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    Call WriteUpdateHungerAndThirst(tempChr)
    
    b = True
    
    If UserList(tempChr).Stats.MinHam < 1 Then
        UserList(tempChr).Stats.MinHam = 0
        UserList(tempChr).flags.Hambre = 1
    End If
    
End If

'Sed
If Hechizos(H).SubeSed = 1 Then
    
    Call InfoHechizo(userIndex)
    
    daño = RandomNumber(Hechizos(H).MinSed, Hechizos(H).MaxSed)
    
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU + daño
    If UserList(tempChr).Stats.MinAGU > UserList(tempChr).Stats.MaxAGU Then _
        UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MaxAGU
         
    If userIndex <> tempChr Then
      Call WriteConsoleMsg(userIndex, "Le has restaurado " & daño & " puntos de sed a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
      Call WriteConsoleMsg(tempChr, UserList(userIndex).name & " te ha restaurado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
    Else
      Call WriteConsoleMsg(userIndex, "Te has restaurado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(H).SubeSed = 2 Then
    
    If Not userCanAttackUser(userIndex, tempChr) Then Exit Sub
    
    If userIndex <> tempChr Then
        Call userAttackedUser(userIndex, tempChr)
    End If
    
    Call InfoHechizo(userIndex)
    
    daño = RandomNumber(Hechizos(H).MinSed, Hechizos(H).MaxSed)
    
    UserList(tempChr).Stats.MinAGU = UserList(tempChr).Stats.MinAGU - daño
    
    If userIndex <> tempChr Then
        Call WriteConsoleMsg(userIndex, "Le has quitado " & daño & " puntos de sed a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(userIndex).name & " te ha quitado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(userIndex, "Te has quitado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    If UserList(tempChr).Stats.MinAGU < 1 Then
            UserList(tempChr).Stats.MinAGU = 0
            UserList(tempChr).flags.Sed = 1
    End If
    
    b = True
End If

' <-------- Agilidad ---------->
If Hechizos(H).SubeAgilidad = 1 Then
    
    Call InfoHechizo(userIndex)
    daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    
    UserList(tempChr).flags.DuracionEfecto = 1200
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) + daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2) Then _
        UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Agilidad) * 2)
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
ElseIf Hechizos(H).SubeAgilidad = 2 Then
    
    If Not userCanAttackUser(userIndex, tempChr) Then Exit Sub
    
    If userIndex <> tempChr Then
        Call userAttackedUser(userIndex, tempChr)
    End If
    
    Call InfoHechizo(userIndex)
    
    UserList(tempChr).flags.TomoPocion = True
    daño = RandomNumber(Hechizos(H).MinAgilidad, Hechizos(H).MaxAgilidad)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) - daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
    b = True
    
    Call WriteUpdateStrengthAgility(tempChr)
    
End If

' <-------- Fuerza ---------->
If Hechizos(H).SubeFuerza = 1 Then
    
    Call InfoHechizo(userIndex)
    daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    
    UserList(tempChr).flags.DuracionEfecto = 1200

    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) + daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) * 2) Then _
        UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, UserList(tempChr).Stats.UserAtributosBackUP(Fuerza) * 2)
    
    UserList(tempChr).flags.TomoPocion = True
    b = True
    
    Call WriteUpdateStrengthAgility(tempChr)
    
ElseIf Hechizos(H).SubeFuerza = 2 Then

    If Not userCanAttackUser(userIndex, tempChr) Then Exit Sub
    
    If userIndex <> tempChr Then
        Call userAttackedUser(userIndex, tempChr)
    End If
    
    Call InfoHechizo(userIndex)
    
    UserList(tempChr).flags.TomoPocion = True
    
    daño = RandomNumber(Hechizos(H).MinFuerza, Hechizos(H).MaxFuerza)
    UserList(tempChr).flags.DuracionEfecto = 700
    UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) - daño
    If UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then UserList(tempChr).Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
    b = True
    
End If

'Salud
If Hechizos(H).SubeHP = 1 Then
    daño = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    daño = daño + Porcentaje(daño, 3 * UserList(userIndex).Stats.ELV)
    
    Call InfoHechizo(userIndex)

    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP + daño
    If UserList(tempChr).Stats.MinHP > UserList(tempChr).Stats.MaxHP Then _
        UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MaxHP
    
    If userIndex <> tempChr Then
        Call WriteConsoleMsg(userIndex, "Le has restaurado " & daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(userIndex).name & " te ha restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(userIndex, "Te has restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    b = True
ElseIf Hechizos(H).SubeHP = 2 Then
    
    If userIndex = tempChr Then
        Call WriteConsoleMsg(userIndex, "No podes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    End If
    
    daño = RandomNumber(Hechizos(H).MinHP, Hechizos(H).MaxHP)
    
    daño = daño + Porcentaje(daño, 3 * UserList(userIndex).Stats.ELV)
    
    If Hechizos(H).StaffAffected Then
        If UserList(userIndex).Clase = eClass.Mage Then
            If UserList(userIndex).Invent.WeaponEqpObjIndex > 0 Then
                daño = (daño * (ObjData(UserList(userIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
            Else
                daño = daño * 0.7 'Baja daño a 70% del original
            End If
        End If
    End If
    
    If UserList(userIndex).Invent.AnilloEqpObjIndex = LAUDMAGICO Then
        daño = daño * 1.04  'laud magico de los bardos
    End If
    
    'cascos antimagia
    If (UserList(tempChr).Invent.CascoEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.CascoEqpObjIndex).DefensaMagicaMax)
    End If
    
    'anillos
    If (UserList(tempChr).Invent.AnilloEqpObjIndex > 0) Then
        daño = daño - RandomNumber(ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(UserList(tempChr).Invent.AnilloEqpObjIndex).DefensaMagicaMax)
    End If
    
    If daño < 0 Then daño = 0
    
    If Not userCanAttackUser(userIndex, tempChr) Then Exit Sub
    
    If userIndex <> tempChr Then
        Call userAttackedUser(userIndex, tempChr)
    End If
    
    Call InfoHechizo(userIndex)
    
    UserList(tempChr).Stats.MinHP = UserList(tempChr).Stats.MinHP - daño
    
    Call WriteConsoleMsg(userIndex, "Le has quitado " & daño & " puntos de vida a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
    Call WriteConsoleMsg(tempChr, UserList(userIndex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
    
    'Si Lanza hechizo a otro usuario y le quita vida, entonces ya no es invisible!
    UserList(userIndex).flags.invisible = 0
    UserList(userIndex).Counters.Invisibilidad = 0
    If UserList(userIndex).flags.Oculto = 0 Then
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageSetInvisible(UserList(userIndex).Char.CharIndex, False))
        Call WriteConsoleMsg(userIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
    End If
    
    'Muere
    If UserList(tempChr).Stats.MinHP < 1 Then
        'Store it!
        Call Statistics.StoreFrag(userIndex, tempChr)
        
        Call ContarMuerte(tempChr, userIndex)
        UserList(tempChr).Stats.MinHP = 0
        Call ActStats(tempChr, userIndex)
        'GRAVE ERROR?
        'Call UserDie(tempChr)
    End If
    
    b = True
End If

'Mana
If Hechizos(H).SubeMana = 1 Then
    
    Call InfoHechizo(userIndex)
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN + daño
    If UserList(tempChr).Stats.MinMAN > UserList(tempChr).Stats.MaxMAN Then _
        UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MaxMAN
    
    If userIndex <> tempChr Then
        Call WriteConsoleMsg(userIndex, "Le has restaurado " & daño & " puntos de mana a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(userIndex).name & " te ha restaurado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(userIndex, "Te has restaurado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    b = True
    
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not userCanAttackUser(userIndex, tempChr) Then Exit Sub
    
    If userIndex <> tempChr Then
        Call userAttackedUser(userIndex, tempChr)
    End If
    
    Call InfoHechizo(userIndex)
    
    If userIndex <> tempChr Then
        Call WriteConsoleMsg(userIndex, "Le has quitado " & daño & " puntos de mana a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(userIndex).name & " te ha quitado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(userIndex, "Te has quitado " & daño & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinMAN = UserList(tempChr).Stats.MinMAN - daño
    If UserList(tempChr).Stats.MinMAN < 1 Then UserList(tempChr).Stats.MinMAN = 0
    b = True
    
End If

'Stamina
If Hechizos(H).SubeSta = 1 Then
    Call InfoHechizo(userIndex)
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta + daño
    If UserList(tempChr).Stats.MinSta > UserList(tempChr).Stats.MaxSta Then _
        UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MaxSta
    If userIndex <> tempChr Then
        Call WriteConsoleMsg(userIndex, "Le has restaurado " & daño & " puntos de vitalidad a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(userIndex).name & " te ha restaurado " & daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(userIndex, "Te has restaurado " & daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    b = True
ElseIf Hechizos(H).SubeMana = 2 Then
    If Not userCanAttackUser(userIndex, tempChr) Then Exit Sub
    
    If userIndex <> tempChr Then
        Call userAttackedUser(userIndex, tempChr)
    End If
    
    Call InfoHechizo(userIndex)
    
    If userIndex <> tempChr Then
        Call WriteConsoleMsg(userIndex, "Le has quitado " & daño & " puntos de vitalidad a " & UserList(tempChr).name, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(tempChr, UserList(userIndex).name & " te ha quitado " & daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(userIndex, "Te has quitado " & daño & " puntos de vitalidad.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    UserList(tempChr).Stats.MinSta = UserList(tempChr).Stats.MinSta - daño
    
    If UserList(tempChr).Stats.MinSta < 1 Then UserList(tempChr).Stats.MinSta = 0
    b = True
End If

Call FlushBuffer(tempChr)

End Sub

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal userIndex As Integer, ByVal Slot As Byte)

'Call LogTarea("Sub UpdateUserHechizos")

Dim LoopC As Byte

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(userIndex).Stats.UserHechizos(Slot) > 0 Then
        Call ChangeUserHechizo(userIndex, Slot, UserList(userIndex).Stats.UserHechizos(Slot))
    Else
        Call ChangeUserHechizo(userIndex, Slot, 0)
    End If

Else

'Actualiza todos los slots
For LoopC = 1 To MAXUSERHECHIZOS

        'Actualiza el inventario
        If UserList(userIndex).Stats.UserHechizos(LoopC) > 0 Then
            Call ChangeUserHechizo(userIndex, LoopC, UserList(userIndex).Stats.UserHechizos(LoopC))
        Else
            Call ChangeUserHechizo(userIndex, LoopC, 0)
        End If

Next LoopC

End If

End Sub

Sub ChangeUserHechizo(ByVal userIndex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)

'Call LogTarea("ChangeUserHechizo")

UserList(userIndex).Stats.UserHechizos(Slot) = Hechizo


If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
    
    Call WriteChangeSpellSlot(userIndex, Slot)

Else

    Call WriteChangeSpellSlot(userIndex, Slot)

End If


End Sub


Public Sub DesplazarHechizo(ByVal userIndex As Integer, ByVal Dire As Integer, ByVal CualHechizo As Integer)

If (Dire <> 1 And Dire <> -1) Then Exit Sub
If Not (CualHechizo >= 1 And CualHechizo <= MAXUSERHECHIZOS) Then Exit Sub

Dim TempHechizo As Integer

If Dire = 1 Then 'Mover arriba
    If CualHechizo = 1 Then
        Call WriteConsoleMsg(userIndex, "No puedes mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(userIndex).Stats.UserHechizos(CualHechizo)
        UserList(userIndex).Stats.UserHechizos(CualHechizo) = UserList(userIndex).Stats.UserHechizos(CualHechizo - 1)
        UserList(userIndex).Stats.UserHechizos(CualHechizo - 1) = TempHechizo

        'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
        If UserList(userIndex).flags.Hechizo > 0 Then
            UserList(userIndex).flags.Hechizo = UserList(userIndex).flags.Hechizo - 1
        End If
    End If
Else 'mover abajo
    If CualHechizo = MAXUSERHECHIZOS Then
        Call WriteConsoleMsg(userIndex, "No puedes mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    Else
        TempHechizo = UserList(userIndex).Stats.UserHechizos(CualHechizo)
        UserList(userIndex).Stats.UserHechizos(CualHechizo) = UserList(userIndex).Stats.UserHechizos(CualHechizo + 1)
        UserList(userIndex).Stats.UserHechizos(CualHechizo + 1) = TempHechizo

        'Prevent the user from casting other spells than the one he had selected when he hitted "cast".
        If UserList(userIndex).flags.Hechizo > 0 Then
            UserList(userIndex).flags.Hechizo = UserList(userIndex).flags.Hechizo + 1
        End If
    End If
End If
End Sub

