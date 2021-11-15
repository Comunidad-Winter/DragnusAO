Attribute VB_Name = "modNPCCombat"
'Argentum Online 0.11.6
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'
'Dise�o y correcci�n del modulo de combate por
'Gerardo Saiz, gerardosaiz@yahoo.com
'

Option Explicit

Public Const MAXDISTANCIAARCO As Byte = 18
Public Const MAXDISTANCIAMAGIA As Byte = 18

Function ModificadorEvasion(ByVal Clase As eClass) As Single

Select Case Clase
    Case eClass.Warrior
        ModificadorEvasion = 1
    Case eClass.Hunter
        ModificadorEvasion = 0.9
    Case eClass.Paladin
        ModificadorEvasion = 0.9
    Case eClass.Bandit
        ModificadorEvasion = 0.9
    Case eClass.Assasin
        ModificadorEvasion = 1.1
    Case eClass.Pirat
        ModificadorEvasion = 0.9
    Case eClass.Thief
        ModificadorEvasion = 1.1
    Case eClass.Bard
        ModificadorEvasion = 1.1
    Case eClass.Mage
        ModificadorEvasion = 0.4
    Case eClass.Druid
        ModificadorEvasion = 0.75
    Case Else
        ModificadorEvasion = 0.8
End Select
End Function

Function ModificadorPoderAtaqueArmas(ByVal Clase As eClass) As Single
Select Case UCase$(Clase)
    Case eClass.Warrior
        ModificadorPoderAtaqueArmas = 1
    Case eClass.Paladin
        ModificadorPoderAtaqueArmas = 0.9
    Case eClass.Hunter
        ModificadorPoderAtaqueArmas = 0.8
    Case eClass.Assasin
        ModificadorPoderAtaqueArmas = 0.85
    Case eClass.Pirat
        ModificadorPoderAtaqueArmas = 0.8
    Case eClass.Thief
        ModificadorPoderAtaqueArmas = 0.75
    Case eClass.Bandit
        ModificadorPoderAtaqueArmas = 0.7
    Case eClass.Cleric
        ModificadorPoderAtaqueArmas = 0.75
    Case eClass.Bard
        ModificadorPoderAtaqueArmas = 0.7
    Case eClass.Druid
        ModificadorPoderAtaqueArmas = 0.65
    Case eClass.Fisher
        ModificadorPoderAtaqueArmas = 0.6
    Case eClass.Lumberjack
        ModificadorPoderAtaqueArmas = 0.6
    Case eClass.Miner
        ModificadorPoderAtaqueArmas = 0.6
    Case eClass.Blacksmith
        ModificadorPoderAtaqueArmas = 0.6
    Case eClass.Carpenter
        ModificadorPoderAtaqueArmas = 0.6
    Case Else
        ModificadorPoderAtaqueArmas = 0.5
End Select
End Function

Function ModificadorPoderAtaqueProyectiles(ByVal Clase As eClass) As Single
Select Case UCase$(Clase)
    Case eClass.Warrior
        ModificadorPoderAtaqueProyectiles = 0.8
    Case eClass.Hunter
        ModificadorPoderAtaqueProyectiles = 1
    Case eClass.Paladin
        ModificadorPoderAtaqueProyectiles = 0.75
    Case eClass.Assasin
        ModificadorPoderAtaqueProyectiles = 0.75
    Case eClass.Pirat
        ModificadorPoderAtaqueProyectiles = 0.75
    Case eClass.Thief
        ModificadorPoderAtaqueProyectiles = 0.8
    Case eClass.Bandit
        ModificadorPoderAtaqueProyectiles = 0.75
    Case eClass.Cleric
        ModificadorPoderAtaqueProyectiles = 0.7
    Case eClass.Bard
        ModificadorPoderAtaqueProyectiles = 0.7
    Case eClass.Druid
        ModificadorPoderAtaqueProyectiles = 0.75
    Case eClass.Fisher
        ModificadorPoderAtaqueProyectiles = 0.65
    Case eClass.Lumberjack
        ModificadorPoderAtaqueProyectiles = 0.7
    Case eClass.Miner
        ModificadorPoderAtaqueProyectiles = 0.65
    Case eClass.Blacksmith
        ModificadorPoderAtaqueProyectiles = 0.65
    Case eClass.Carpenter
        ModificadorPoderAtaqueProyectiles = 0.7
    Case Else
        ModificadorPoderAtaqueProyectiles = 0.5
End Select
End Function

Function ModicadorDa�oClaseArmas(ByVal Clase As eClass) As Single
Select Case UCase$(Clase)
    Case eClass.Warrior
        ModicadorDa�oClaseArmas = 1.1
    Case eClass.Paladin
        ModicadorDa�oClaseArmas = 0.95
    Case eClass.Hunter
        ModicadorDa�oClaseArmas = 0.9
    Case eClass.Assasin
        ModicadorDa�oClaseArmas = 0.9
    Case eClass.Thief
        ModicadorDa�oClaseArmas = 0.8
    Case eClass.Pirat
        ModicadorDa�oClaseArmas = 0.8
    Case eClass.Bandit
        ModicadorDa�oClaseArmas = 1
    Case eClass.Cleric
        ModicadorDa�oClaseArmas = 0.8
    Case eClass.Bard
        ModicadorDa�oClaseArmas = 0.75
    Case eClass.Druid
        ModicadorDa�oClaseArmas = 0.7
    Case eClass.Fisher
        ModicadorDa�oClaseArmas = 0.6
    Case eClass.Lumberjack
        ModicadorDa�oClaseArmas = 0.7
    Case eClass.Miner
        ModicadorDa�oClaseArmas = 0.75
    Case eClass.Blacksmith
        ModicadorDa�oClaseArmas = 0.75
    Case eClass.Carpenter
        ModicadorDa�oClaseArmas = 0.7
    Case Else
        ModicadorDa�oClaseArmas = 0.5
End Select
End Function

Function ModicadorDa�oClaseWrestling(ByVal Clase As eClass) As Single
'Pablo (ToxicWaste): Esto en proxima versi�n habr� que balancearlo para cada Clase
'Hoy por hoy est� solo hecho para el bandido.
Select Case UCase$(Clase)
    Case eClass.Warrior
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Paladin
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Hunter
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Assasin
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Thief
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Pirat
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Bandit
        ModicadorDa�oClaseWrestling = 1.1
    Case eClass.Cleric
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Bard
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Druid
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Fisher
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Lumberjack
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Miner
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Blacksmith
        ModicadorDa�oClaseWrestling = 0.4
    Case eClass.Carpenter
        ModicadorDa�oClaseWrestling = 0.4
    Case Else
        ModicadorDa�oClaseWrestling = 0.4
End Select
End Function


Function ModicadorDa�oClaseProyectiles(ByVal Clase As eClass) As Single
Select Case Clase
    Case eClass.Hunter
        ModicadorDa�oClaseProyectiles = 1.1
    Case eClass.Warrior
        ModicadorDa�oClaseProyectiles = 0.9
    Case eClass.Paladin
        ModicadorDa�oClaseProyectiles = 0.8
    Case eClass.Assasin
        ModicadorDa�oClaseProyectiles = 0.8
    Case eClass.Thief
        ModicadorDa�oClaseProyectiles = 0.75
    Case eClass.Pirat
        ModicadorDa�oClaseProyectiles = 0.75
    Case eClass.Bandit
        ModicadorDa�oClaseProyectiles = 0.8
    Case eClass.Cleric
        ModicadorDa�oClaseProyectiles = 0.7
    Case eClass.Bard
        ModicadorDa�oClaseProyectiles = 0.7
    Case eClass.Druid
        ModicadorDa�oClaseProyectiles = 0.75
    Case eClass.Fisher
        ModicadorDa�oClaseProyectiles = 0.6
    Case eClass.Lumberjack
        ModicadorDa�oClaseProyectiles = 0.7
    Case eClass.Miner
        ModicadorDa�oClaseProyectiles = 0.6
    Case eClass.Blacksmith
        ModicadorDa�oClaseProyectiles = 0.6
    Case eClass.Carpenter
        ModicadorDa�oClaseProyectiles = 0.7
    Case Else
        ModicadorDa�oClaseProyectiles = 0.5
End Select
End Function

Function ModEvasionDeEscudoClase(ByVal Clase As eClass) As Single

Select Case Clase
    Case eClass.Warrior
        ModEvasionDeEscudoClase = 1
    Case eClass.Hunter
        ModEvasionDeEscudoClase = 0.8
    Case eClass.Paladin
        ModEvasionDeEscudoClase = 1
    Case eClass.Assasin
        ModEvasionDeEscudoClase = 0.8
    Case eClass.Thief
        ModEvasionDeEscudoClase = 0.7
    Case eClass.Bandit
        ModEvasionDeEscudoClase = 2
    Case eClass.Pirat
        ModEvasionDeEscudoClase = 0.75
    Case eClass.Cleric
        ModEvasionDeEscudoClase = 0.85
    Case eClass.Bard
        ModEvasionDeEscudoClase = 0.8
    Case eClass.Druid
        ModEvasionDeEscudoClase = 0.75
    Case eClass.Fisher
        ModEvasionDeEscudoClase = 0.7
    Case eClass.Lumberjack
        ModEvasionDeEscudoClase = 0.7
    Case eClass.Miner
        ModEvasionDeEscudoClase = 0.7
    Case eClass.Blacksmith
        ModEvasionDeEscudoClase = 0.7
    Case eClass.Carpenter
        ModEvasionDeEscudoClase = 0.7
    Case Else
        ModEvasionDeEscudoClase = 0.6
End Select

End Function
Function Minimo(ByVal a As Single, ByVal b As Single) As Single
If a > b Then
    Minimo = b
    Else: Minimo = a
End If
End Function

Function MinimoInt(ByVal a As Integer, ByVal b As Integer) As Integer
If a > b Then
    MinimoInt = b
    Else: MinimoInt = a
End If
End Function

Function Maximo(ByVal a As Single, ByVal b As Single) As Single
If a > b Then
    Maximo = a
    Else: Maximo = b
End If
End Function

Function MaximoInt(ByVal a As Integer, ByVal b As Integer) As Integer
If a > b Then
    MaximoInt = a
    Else: MaximoInt = b
End If
End Function


Function PoderEvasionEscudo(ByVal userIndex As Integer) As Long

PoderEvasionEscudo = (UserList(userIndex).Stats.UserSkills(eSkill.Defensa) * _
ModEvasionDeEscudoClase(UserList(userIndex).Clase)) / 2

End Function

Function PoderEvasion(ByVal userIndex As Integer) As Long
    Dim lTemp As Long
     With UserList(userIndex)
       lTemp = (.Stats.UserSkills(eSkill.Tacticas) + _
          .Stats.UserSkills(eSkill.Tacticas) / 33 * .Stats.UserAtributos(eAtributos.Agilidad)) * _
          ModificadorEvasion(.Clase)
       
        PoderEvasion = (lTemp + (2.5 * Maximo(CInt(.Stats.ELV) - 12, 0)))
    End With
End Function

Function PoderAtaqueArma(ByVal userIndex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(userIndex).Stats.UserSkills(eSkill.Armas) < 31 Then
    PoderAtaqueTemp = (UserList(userIndex).Stats.UserSkills(eSkill.Armas) * _
    ModificadorPoderAtaqueArmas(UserList(userIndex).Clase))
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Armas) < 61 Then
    PoderAtaqueTemp = ((UserList(userIndex).Stats.UserSkills(eSkill.Armas) + _
    UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad)) * _
    ModificadorPoderAtaqueArmas(UserList(userIndex).Clase))
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Armas) < 91 Then
    PoderAtaqueTemp = ((UserList(userIndex).Stats.UserSkills(eSkill.Armas) + _
    (2 * UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
    ModificadorPoderAtaqueArmas(UserList(userIndex).Clase))
Else
   PoderAtaqueTemp = ((UserList(userIndex).Stats.UserSkills(eSkill.Armas) + _
   (3 * UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
   ModificadorPoderAtaqueArmas(UserList(userIndex).Clase))
End If

PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * Maximo(CInt(UserList(userIndex).Stats.ELV) - 12, 0)))
End Function

Function PoderAtaqueProyectil(ByVal userIndex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(userIndex).Stats.UserSkills(eSkill.Proyectiles) < 31 Then
    PoderAtaqueTemp = (UserList(userIndex).Stats.UserSkills(eSkill.Proyectiles) * _
    ModificadorPoderAtaqueProyectiles(UserList(userIndex).Clase))
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Proyectiles) < 61 Then
        PoderAtaqueTemp = ((UserList(userIndex).Stats.UserSkills(eSkill.Proyectiles) + _
        UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad)) * _
        ModificadorPoderAtaqueProyectiles(UserList(userIndex).Clase))
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Proyectiles) < 91 Then
        PoderAtaqueTemp = ((UserList(userIndex).Stats.UserSkills(eSkill.Proyectiles) + _
        (2 * UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
        ModificadorPoderAtaqueProyectiles(UserList(userIndex).Clase))
Else
       PoderAtaqueTemp = ((UserList(userIndex).Stats.UserSkills(eSkill.Proyectiles) + _
      (3 * UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
      ModificadorPoderAtaqueProyectiles(UserList(userIndex).Clase))
End If

PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * Maximo(CInt(UserList(userIndex).Stats.ELV) - 12, 0)))

End Function

Function PoderAtaqueWrestling(ByVal userIndex As Integer) As Long
Dim PoderAtaqueTemp As Long

If UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) < 31 Then
    PoderAtaqueTemp = (UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) * _
    ModificadorPoderAtaqueArmas(UserList(userIndex).Clase))
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) < 61 Then
        PoderAtaqueTemp = ((UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) + _
        UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad)) * _
        ModificadorPoderAtaqueArmas(UserList(userIndex).Clase))
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) < 91 Then
        PoderAtaqueTemp = ((UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) + _
        (2 * UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
        ModificadorPoderAtaqueArmas(UserList(userIndex).Clase))
Else
       PoderAtaqueTemp = ((UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) + _
       (3 * UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad))) * _
       ModificadorPoderAtaqueArmas(UserList(userIndex).Clase))
End If

PoderAtaqueWrestling = (PoderAtaqueTemp + (2.5 * Maximo(CInt(UserList(userIndex).Stats.ELV) - 12, 0)))

End Function


Public Function UserImpactoNpc(ByVal userIndex As Integer, ByVal NpcIndex As Integer) As Boolean
Dim PoderAtaque As Long
Dim Arma As Integer
Dim proyectil As Boolean
Dim ProbExito As Long

Arma = UserList(userIndex).Invent.WeaponEqpObjIndex
If Arma = 0 Then proyectil = False Else proyectil = ObjData(Arma).proyectil = 1

If Arma > 0 Then 'Usando un arma
    If proyectil Then
        PoderAtaque = PoderAtaqueProyectil(userIndex)
    Else
        PoderAtaque = PoderAtaqueArma(userIndex)
    End If
Else 'Peleando con pu�os
    PoderAtaque = PoderAtaqueWrestling(userIndex)
End If


ProbExito = Maximo(10, Minimo(90, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))

UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)

If UserImpactoNpc Then
    If Arma <> 0 Then
       If proyectil Then
            Call SubirSkill(userIndex, Proyectiles)
       Else
            Call SubirSkill(userIndex, Armas)
       End If
    Else
        Call SubirSkill(userIndex, Wrestling)
    End If
End If


End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal userIndex As Integer) As Boolean
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Revisa si un NPC logra impactar a un user o no
'03/15/2006 Maraxus - Evit� una divisi�n por cero que eliminaba NPCs
'*************************************************
Dim Rechazo As Boolean
Dim ProbRechazo As Long
Dim ProbExito As Long
Dim UserEvasion As Long
Dim NpcPoderAtaque As Long
Dim PoderEvasioEscudo As Long
Dim SkillTacticas As Long
Dim SkillDefensa As Long

UserEvasion = PoderEvasion(userIndex)
NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
PoderEvasioEscudo = PoderEvasionEscudo(userIndex)

SkillTacticas = UserList(userIndex).Stats.UserSkills(eSkill.Tacticas)
SkillDefensa = UserList(userIndex).Stats.UserSkills(eSkill.Defensa)

'Esta usando un escudo ???
If UserList(userIndex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo

ProbExito = Maximo(10, Minimo(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))

NpcImpacto = (RandomNumber(1, 100) <= ProbExito)

' el usuario esta usando un escudo ???
If UserList(userIndex).Invent.EscudoEqpObjIndex > 0 Then
    If Not NpcImpacto Then
        If SkillDefensa + SkillTacticas > 0 Then  'Evitamos divisi�n por cero
            ProbRechazo = Maximo(10, Minimo(90, 100 * (SkillDefensa / (SkillDefensa + SkillTacticas))))
            Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
            If Rechazo = True Then
                'Se rechazo el ataque con el escudo
                Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_ESCUDO))
                Call WriteBlockedWithShieldUser(userIndex)
                Call SubirSkill(userIndex, Defensa)
            End If
        End If
    End If
End If
End Function

Public Function CalcularDa�o(ByVal userIndex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long
Dim Da�oArma As Long, Da�oUsuario As Long, Arma As ObjData, ModifClase As Single
Dim proyectil As ObjData
Dim Da�oMaxArma As Long

''sacar esto si no queremos q la matadracos mate el Dragon si o si
Dim matoDragon As Boolean
matoDragon = False


If UserList(userIndex).Invent.WeaponEqpObjIndex > 0 Then
    Arma = ObjData(UserList(userIndex).Invent.WeaponEqpObjIndex)
    
    
    ' Ataca a un npc?
    If NpcIndex > 0 Then
        
        'Usa la mata Dragones?
        If UserList(userIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then ' Usa la mataDragones?
            ModifClase = ModicadorDa�oClaseArmas(UserList(userIndex).Clase)
            
            If Npclist(NpcIndex).NPCtype = DRAGON Then 'Ataca Dragon?
                Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                Da�oMaxArma = Arma.MaxHIT
                matoDragon = True ''sacar esto si no queremos q la matadracos mate el Dragon si o si
            Else ' Sino es Dragon da�o es 1
                Da�oArma = 1
                Da�oMaxArma = 1
            End If
        Else ' da�o comun
           If Arma.proyectil = 1 Then
                ModifClase = ModicadorDa�oClaseProyectiles(UserList(userIndex).Clase)
                Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                Da�oMaxArma = Arma.MaxHIT
                If Arma.Municion = 1 Then
                    proyectil = ObjData(UserList(userIndex).Invent.MunicionEqpObjIndex)
                    Da�oArma = Da�oArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                    Da�oMaxArma = Arma.MaxHIT
                End If
           Else
                ModifClase = ModicadorDa�oClaseArmas(UserList(userIndex).Clase)
                Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                Da�oMaxArma = Arma.MaxHIT
           End If
        End If
    
    Else ' Ataca usuario
        If UserList(userIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
            ModifClase = ModicadorDa�oClaseArmas(UserList(userIndex).Clase)
            Da�oArma = 1 ' Si usa la espada mataDragones da�o es 1
            Da�oMaxArma = 1
        Else
           If Arma.proyectil = 1 Then
                ModifClase = ModicadorDa�oClaseProyectiles(UserList(userIndex).Clase)
                Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                Da�oMaxArma = Arma.MaxHIT
                
                If Arma.Municion = 1 Then
                    proyectil = ObjData(UserList(userIndex).Invent.MunicionEqpObjIndex)
                    Da�oArma = Da�oArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                    Da�oMaxArma = Arma.MaxHIT
                End If
           Else
                ModifClase = ModicadorDa�oClaseArmas(UserList(userIndex).Clase)
                Da�oArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
                Da�oMaxArma = Arma.MaxHIT
           End If
        End If
    End If
Else
    'Pablo (ToxicWaste)
    ModifClase = ModicadorDa�oClaseWrestling(UserList(userIndex).Clase)
    Da�oArma = RandomNumber(1, 3) 'Hacemos que sea "tipo" una daga el ataque de Wrestling
    Da�oMaxArma = 3
End If

Da�oUsuario = RandomNumber(UserList(userIndex).Stats.MinHIT, UserList(userIndex).Stats.MaxHIT)

''sacar esto si no queremos q la matadracos mate el Dragon si o si
If matoDragon Then
    CalcularDa�o = Npclist(NpcIndex).Stats.MinHP + Npclist(NpcIndex).Stats.def
Else
    CalcularDa�o = (((3 * Da�oArma) + ((Da�oMaxArma / 5) * Maximo(0, (UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) - 15))) + Da�oUsuario) * ModifClase)
End If

End Function

Public Sub UserDa�oNpc(ByVal userIndex As Integer, ByVal NpcIndex As Integer)

Dim da�o As Long

da�o = CalcularDa�o(userIndex, NpcIndex)

'esta navegando? si es asi le sumamos el da�o del barco
If UserList(userIndex).flags.Navegando = 1 And UserList(userIndex).Invent.BarcoObjIndex > 0 Then _
        da�o = da�o + RandomNumber(ObjData(UserList(userIndex).Invent.BarcoObjIndex).MinHIT, ObjData(UserList(userIndex).Invent.BarcoObjIndex).MaxHIT)

If UserList(userIndex).flags.Montado = 1 And UserList(userIndex).Invent.MonturaObjIndex > 0 Then _
        da�o = da�o + RandomNumber(ObjData(UserList(userIndex).Invent.MonturaObjIndex).MinHIT, ObjData(UserList(userIndex).Invent.MonturaObjIndex).MaxHIT)

da�o = da�o - Npclist(NpcIndex).Stats.def

If da�o < 0 Then da�o = 0

'[KEVIN]
Call WriteUserHitNPC(userIndex, da�o)
Call CalcularDarExp(userIndex, NpcIndex, da�o)
Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MinHP - da�o
'[/KEVIN]

If Npclist(NpcIndex).Stats.MinHP > 0 Then
    'Trata de apu�alar por la espalda al enemigo
    If PuedeApu�alar(userIndex) Then
       Call DoApu�alar(userIndex, NpcIndex, 0, da�o)
       Call SubirSkill(userIndex, Apu�alar)
    End If
    'trata de dar golpe cr�tico
    Call DoGolpeCritico(userIndex, NpcIndex, 0, da�o)
    
End If

 
If Npclist(NpcIndex).Stats.MinHP <= 0 Then
        
        ' Si era un Dragon perdemos la espada mataDragones
        If Npclist(NpcIndex).NPCtype = DRAGON Then
            'Si tiene equipada la matadracos se la sacamos
            If UserList(userIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                Call QuitarObjetos(EspadaMataDragonesIndex, 1, userIndex)
            End If
            If Npclist(NpcIndex).Stats.MaxHP > 100000 Then Call LogDesarrollo(UserList(userIndex).name & " mat� un drag�n")
        End If
        
        
        ' Para que las mascotas no sigan intentando luchar y
        ' comiencen a seguir al amo
        
        Dim j As Integer
        For j = 1 To MAXMASCOTAS
            If UserList(userIndex).MascotasIndex(j) > 0 Then
                If Npclist(UserList(userIndex).MascotasIndex(j)).TargetNPC = NpcIndex Then Npclist(UserList(userIndex).MascotasIndex(j)).TargetNPC = 0
                Npclist(UserList(userIndex).MascotasIndex(j)).Movement = TipoAI.SigueAmo
            End If
        Next j
        
        Call MuereNpc(NpcIndex, userIndex)
End If

End Sub


Public Sub NpcDa�o(ByVal NpcIndex As Integer, ByVal userIndex As Integer)

Dim da�o As Integer, Lugar As Integer, absorbido As Integer, npcfile As String
Dim antda�o As Integer, defbarco As Integer
Dim Obj As ObjData



da�o = RandomNumber(Npclist(NpcIndex).Stats.MinHIT, Npclist(NpcIndex).Stats.MaxHIT)
antda�o = da�o

If UserList(userIndex).flags.Navegando = 1 And UserList(userIndex).Invent.BarcoObjIndex > 0 Then
    Obj = ObjData(UserList(userIndex).Invent.BarcoObjIndex)
    defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If

Dim defMontura As Integer

If UserList(userIndex).flags.Montado = 1 And UserList(userIndex).Invent.MonturaObjIndex > 0 Then
    Obj = ObjData(UserList(userIndex).Invent.MonturaObjIndex)
    defMontura = RandomNumber(Obj.MinDef, Obj.MaxDef)
End If

Lugar = RandomNumber(1, 6)


Select Case Lugar
  Case PartesCuerpo.bCabeza
        'Si tiene casco absorbe el golpe
        If UserList(userIndex).Invent.CascoEqpObjIndex > 0 Then
           Obj = ObjData(UserList(userIndex).Invent.CascoEqpObjIndex)
           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           absorbido = absorbido + defbarco
           da�o = da�o - absorbido
           If da�o < 1 Then da�o = 1
        End If
  Case Else
        'Si tiene armadura absorbe el golpe
        If UserList(userIndex).Invent.ArmourEqpObjIndex > 0 Then
           Dim Obj2 As ObjData
           Obj = ObjData(UserList(userIndex).Invent.ArmourEqpObjIndex)
           If UserList(userIndex).Invent.EscudoEqpObjIndex Then
                Obj2 = ObjData(UserList(userIndex).Invent.EscudoEqpObjIndex)
                absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
           Else
                absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
           End If
           absorbido = absorbido + defbarco
           da�o = da�o - absorbido
           If da�o < 1 Then da�o = 1
        End If
End Select

Call WriteNPCHitUser(userIndex, Lugar, da�o)

If UserList(userIndex).flags.Privilegios And PlayerType.User Then UserList(userIndex).Stats.MinHP = UserList(userIndex).Stats.MinHP - da�o

If UserList(userIndex).flags.Meditando Then
    If da�o > Fix(UserList(userIndex).Stats.MinHP / 100 * UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia) * UserList(userIndex).Stats.UserSkills(eSkill.Meditar) / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
        UserList(userIndex).flags.Meditando = False
        Call WriteMeditateToggle(userIndex)
        Call WriteConsoleMsg(userIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
        UserList(userIndex).Char.FX = 0
        UserList(userIndex).Char.loops = 0
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCreateFX(UserList(userIndex).Char.CharIndex, 0, 0))
    End If
End If

'Muere el usuario
If UserList(userIndex).Stats.MinHP <= 0 Then

    Call WriteNPCKillUser(userIndex) ' Le informamos que ha muerto ;)
    
    If Npclist(NpcIndex).MaestroUser > 0 Then
        Call AllFollowAmo(Npclist(NpcIndex).MaestroUser)
    Else
        'Al matarlo no lo sigue mas
        If Npclist(NpcIndex).Stats.Alineacion = 0 Then
                    Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
                    Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
                    Npclist(NpcIndex).flags.AttackedBy = vbNullString
        End If
    End If
    
    
    Call UserDie(userIndex)

End If

End Sub



Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal userIndex As Integer, Optional ByVal CheckElementales As Boolean = True)

Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(userIndex).MascotasIndex(j) > 0 Then
       If UserList(userIndex).MascotasIndex(j) <> NpcIndex Then
        If CheckElementales Or (Npclist(UserList(userIndex).MascotasIndex(j)).Numero <> ELEMENTALFUEGO And Npclist(UserList(userIndex).MascotasIndex(j)).Numero <> ELEMENTALTIERRA) Then
            If Npclist(UserList(userIndex).MascotasIndex(j)).TargetNPC = 0 Then Npclist(UserList(userIndex).MascotasIndex(j)).TargetNPC = NpcIndex
            'Npclist(UserList(UserIndex).MascotasIndex(j)).Flags.OldMovement = Npclist(UserList(UserIndex).MascotasIndex(j)).Movement
            Npclist(UserList(userIndex).MascotasIndex(j)).Movement = TipoAI.NpcAtacaNpc
        End If
       End If
    End If
Next j

End Sub
Public Sub AllFollowAmo(ByVal userIndex As Integer)
Dim j As Integer
For j = 1 To MAXMASCOTAS
    If UserList(userIndex).MascotasIndex(j) > 0 Then
        Call FollowAmo(UserList(userIndex).MascotasIndex(j))
    End If
Next j
End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal userIndex As Integer) As Boolean

If UserList(userIndex).flags.AdminInvisible = 1 Then Exit Function
If (Not UserList(userIndex).flags.Privilegios And PlayerType.User) <> 0 And Not UserList(userIndex).flags.AdminPerseguible Then Exit Function

' El npc puede atacar ???
If Npclist(NpcIndex).CanAttack = 1 Then
    NpcAtacaUser = True
    Call CheckPets(NpcIndex, userIndex, False)

    If Npclist(NpcIndex).Target = 0 Then Npclist(NpcIndex).Target = userIndex

    If UserList(userIndex).flags.AtacadoPorNpc = 0 And _
       UserList(userIndex).flags.AtacadoPorUser = 0 Then UserList(userIndex).flags.AtacadoPorNpc = NpcIndex
Else
    NpcAtacaUser = False
    Exit Function
End If

Npclist(NpcIndex).CanAttack = 0

If Npclist(NpcIndex).flags.Snd1 > 0 Then
    Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd1))
End If

If NpcImpacto(NpcIndex, userIndex) Then
    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_IMPACTO))
    
    If UserList(userIndex).flags.Meditando = False Then
        If UserList(userIndex).flags.Navegando = 0 Or UserList(userIndex).flags.Montado = 0 Then
            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCreateFX(UserList(userIndex).Char.CharIndex, FXSANGRE, 0))
        End If
    End If
    
    Call NpcDa�o(NpcIndex, userIndex)
    Call WriteUpdateHP(userIndex)
    '�Puede envenenar?
    If Npclist(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(userIndex)
Else
    Call WriteNPCSwing(userIndex)
End If



'-----Tal vez suba los skills------
Call SubirSkill(userIndex, Tacticas)

'Controla el nivel del usuario
Call CheckUserLevel(userIndex)

End Function

Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
Dim PoderAtt As Long, PoderEva As Long, dif As Long
Dim ProbExito As Long

PoderAtt = Npclist(Atacante).PoderAtaque
PoderEva = Npclist(Victima).PoderEvasion
ProbExito = Maximo(10, Minimo(90, 50 + _
            ((PoderAtt - PoderEva) * 0.4)))
NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)


End Function

Public Sub NpcDa�oNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
Dim da�o As Integer
Dim ANpc As npc, DNpc As npc
ANpc = Npclist(Atacante)

da�o = RandomNumber(ANpc.Stats.MinHIT, ANpc.Stats.MaxHIT)
Npclist(Victima).Stats.MinHP = Npclist(Victima).Stats.MinHP - da�o

If Npclist(Victima).Stats.MinHP < 1 Then
        
        If LenB(Npclist(Atacante).flags.AttackedBy) <> 0 Then
            Npclist(Atacante).Movement = Npclist(Atacante).flags.OldMovement
            Npclist(Atacante).Hostile = Npclist(Atacante).flags.OldHostil
        Else
            Npclist(Atacante).Movement = Npclist(Atacante).flags.OldMovement
        End If
        
        Call FollowAmo(Atacante)
        
        Call MuereNpc(Victima, Npclist(Atacante).MaestroUser)
End If

End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMOvimiento As Boolean = True)

' El npc puede atacar ???
If Npclist(Atacante).CanAttack = 1 Then
       Npclist(Atacante).CanAttack = 0
        If cambiarMOvimiento Then
            Npclist(Victima).TargetNPC = Atacante
            Npclist(Victima).Movement = TipoAI.NpcAtacaNpc
        End If
Else
    Exit Sub
End If

If Npclist(Atacante).flags.Snd1 > 0 Then
    Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(Npclist(Atacante).flags.Snd1))
End If

If NpcImpactoNpc(Atacante, Victima) Then
    
    If Npclist(Victima).flags.Snd2 > 0 Then
        Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(Npclist(Victima).flags.Snd2))
    Else
        Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO2))
    End If

    If Npclist(Atacante).MaestroUser > 0 Then
        Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_IMPACTO))
    Else
        Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO))
    End If
    Call NpcDa�oNpc(Atacante, Victima)
    
Else
    If Npclist(Atacante).MaestroUser > 0 Then
        Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_SWING))
    Else
        Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_SWING))
    End If
End If

End Sub

Public Sub UsuarioAtacaNpc(ByVal userIndex As Integer, ByVal NpcIndex As Integer)

If UserList(userIndex).flags.Privilegios And PlayerType.Consejero Then Exit Sub

If Distancia(UserList(userIndex).Pos, Npclist(NpcIndex).Pos) > MAXDISTANCIAARCO Then
   Call WriteConsoleMsg(userIndex, "Est�s muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
   Exit Sub
End If

If Npclist(NpcIndex).MaestroUser <> 0 Then
    If UserList(Npclist(NpcIndex).MaestroUser).Faccion.Alineacion = UserList(userIndex).Faccion.Alineacion And UserList(userIndex).Faccion.Alineacion <> e_Alineacion.Neutro Then
        Call WriteConsoleMsg(userIndex, "No puedes atacar a usuarios de tu faccion.", FontTypeNames.FONTTYPE_WARNING)
        Exit Sub
    End If
End If

If Npclist(NpcIndex).EsRey Then
    If UserList(userIndex).Faccion.Alineacion = e_Alineacion.Neutro Then
        Call WriteConsoleMsg(userIndex, "Debes pertenecer a una faccion para atacar a este npc.", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    ElseIf UserList(userIndex).Faccion.Alineacion = Castillo(Npclist(NpcIndex).EsRey).LeaderFaccion Then
        Call WriteConsoleMsg(userIndex, "No puedes atacar al rey de tu castillo.", FontTypeNames.FONTTYPE_FIGHT)
        Exit Sub
    Else
        CastleUnderAttack Npclist(NpcIndex).EsRey
    End If
End If

'Revisa que el Rey pretoriano no est� solo.
If esPretoriano(NpcIndex) = 4 Then
    If Npclist(NpcIndex).Pos.X < 50 Then
        If pretorianosVivos(1) > 0 Then
            Call WriteConsoleMsg(userIndex, "Debes matar al resto del ejercito antes de atacar al rey!", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
    Else
        If pretorianosVivos(2) > 0 Then
            Call WriteConsoleMsg(userIndex, "Debes matar al resto del ejercito antes de atacar al rey!", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
    End If
End If

Call NPCAtacado(NpcIndex, userIndex)
Call CheckPets(NpcIndex, userIndex)

If UserImpactoNpc(userIndex, NpcIndex) Then
    
    If Npclist(NpcIndex).flags.Snd2 > 0 Then
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2))
    Else
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_IMPACTO2))
    End If
    
    Call UserDa�oNpc(userIndex, NpcIndex)
   
Else
    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_SWING))
    Call WriteUserSwing(userIndex)
End If

End Sub

Sub UsuarioAtacadoPorUsuario(ByVal attackerIndex As Integer, ByVal VictimIndex As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 03/09/06 Nacho
'Usuario deja de meditar
'***************************************************
    If TriggerZonaPelea(attackerIndex, VictimIndex) = TRIGGER6_PERMITE Or UserList(attackerIndex).flags.EnDuelo Or UserList(attackerIndex).Pos.Map = MAPATORNEO Or IsInCastle(attackerIndex) Then Exit Sub
    
    Dim EraCriminal As Boolean
    
    If UserList(VictimIndex).flags.Meditando Then
        UserList(VictimIndex).flags.Meditando = False
        Call WriteMeditateToggle(VictimIndex)
        Call WriteConsoleMsg(VictimIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
        UserList(VictimIndex).Char.FX = 0
        UserList(VictimIndex).Char.loops = 0
        Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageCreateFX(UserList(VictimIndex).Char.CharIndex, 0, 0))
    End If
    
    Call AllMascotasAtacanUser(attackerIndex, VictimIndex)
    Call AllMascotasAtacanUser(VictimIndex, attackerIndex)
    
    Call FlushBuffer(VictimIndex)
End Sub

Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)
'Reaccion de las mascotas
Dim iCount As Integer

For iCount = 1 To MAXMASCOTAS
    If UserList(Maestro).MascotasIndex(iCount) > 0 Then
            Npclist(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = UserList(victim).name
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Movement = TipoAI.NPCDEFENSA
            Npclist(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1
    End If
Next iCount

End Sub



Public Function PuedeAtacarNPC(ByVal attackerIndex As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Autor: Unknown Author (Original version)
'Returns True if AttackerIndex can attack the NpcIndex
'Last Modification: 24/01/2007
'24/01/2007 Pablo (ToxicWaste) - Orden y correcci�n de ataque sobre una mascota y guardias
'***************************************************

'Estas muerto?
If UserList(attackerIndex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(attackerIndex, "No pod�s atacar porque estas muerto", FontTypeNames.FONTTYPE_INFO)
    PuedeAtacarNPC = False
    Exit Function
End If

'Es el NPC mascota de alguien?
If Npclist(NpcIndex).MaestroUser > 0 Then
    'De un cudadanos y sos Armada?
    If Not PuedeAtacar(attackerIndex, Npclist(NpcIndex).MaestroUser) Then
        PuedeAtacarNPC = False
        Exit Function
    End If
End If

'Sos consejero? no podes atacar nunca.
If UserList(attackerIndex).flags.Privilegios And PlayerType.Consejero Then
    PuedeAtacarNPC = False
    Exit Function
End If

'Es el Rey Preatoriano?
If esPretoriano(NpcIndex) = 4 Then
    If Npclist(NpcIndex).Pos.X < 50 Then
        If pretorianosVivos(1) > 0 Then
            Call WriteConsoleMsg(attackerIndex, "Debes matar al resto del ejercito antes de atacar al rey!", FontTypeNames.FONTTYPE_FIGHT)
            PuedeAtacarNPC = False
            Exit Function
        End If
    Else
        If pretorianosVivos(2) > 0 Then
            Call WriteConsoleMsg(attackerIndex, "Debes matar al resto del ejercito antes de atacar al rey!", FontTypeNames.FONTTYPE_FIGHT)
            PuedeAtacarNPC = False
            Exit Function
        End If
    End If
End If
Debug.Print "0 -- 0"

If Npclist(NpcIndex).EsRey Then
    If Not UserList(attackerIndex).GuildIndex > 0 Then
        Debug.Print "3"
        Call WriteConsoleMsg(attackerIndex, "Debes pertenecer a un clan para atacar a este npc.", FontTypeNames.FONTTYPE_FIGHT)
        PuedeAtacarNPC = False
        Exit Function
    End If
End If
If Npclist(NpcIndex).EsRey Then
    If UserList(attackerIndex).Faccion.Alineacion = Castillo(Npclist(NpcIndex).EsRey).LeaderFaccion Then
        Call WriteConsoleMsg(attackerIndex, "No podes atacar al rey de tu castillo.", FontTypeNames.FONTTYPE_FIGHT)
        PuedeAtacarNPC = False
        Exit Function
    ElseIf UserList(attackerIndex).Faccion.Alineacion = e_Alineacion.Neutro Then
        Call WriteConsoleMsg(attackerIndex, "Debes pertenecer a una faccion para atacar a este npc.", FontTypeNames.FONTTYPE_FIGHT)
        PuedeAtacarNPC = False
        Exit Function
    End If
End If

PuedeAtacarNPC = True
End Function

Sub CalcularDarExp(ByVal userIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDa�o As Long)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/09/06 Nacho
'Reescribi gran parte del Sub
'Ahora, da toda la experiencia del npc mientras este vivo.
'***************************************************
Dim ExpaDar As Long

'[Nacho] Chekeamos que las variables sean validas para las operaciones
If ElDa�o <= 0 Then ElDa�o = 0
If Npclist(NpcIndex).Stats.MaxHP <= 0 Then Exit Sub
If ElDa�o > Npclist(NpcIndex).Stats.MinHP Then ElDa�o = Npclist(NpcIndex).Stats.MinHP

If ElDa�o = Npclist(NpcIndex).Stats.MinHP Then
    ExpaDar = Npclist(NpcIndex).GiveEXP
Else
    ExpaDar = CLng((ElDa�o) * (Npclist(NpcIndex).GiveEXP / Npclist(NpcIndex).Stats.MaxHP))
End If

'[Nacho] Le damos la exp al user
If ExpaDar > 0 Then
    If ClanPoseeMapa(UserList(userIndex).GuildIndex, Npclist(NpcIndex).Pos.Map) Then
        UserList(userIndex).Stats.Exp = UserList(userIndex).Stats.Exp + ExpaDar * 1.1
    Else
        UserList(userIndex).Stats.Exp = UserList(userIndex).Stats.Exp + ExpaDar
    End If
    If UserList(userIndex).Stats.Exp > MAXEXP Then _
        UserList(userIndex).Stats.Exp = MAXEXP
    Call WriteConsoleMsg(userIndex, "Has ganado " & ExpaDar & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)
    
    Call CheckUserLevel(userIndex)
End If

End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6
'TODO: Pero que rebuscado!!
'Nigo:  Te lo redise�e, pero no te borro el TODO para que lo revises.
On Error GoTo errhandler
    Dim tOrg As eTrigger
    Dim tDst As eTrigger
    
    tOrg = MapData(UserList(Origen).Pos.Map, UserList(Origen).Pos.X, UserList(Origen).Pos.Y).trigger
    tDst = MapData(UserList(Destino).Pos.Map, UserList(Destino).Pos.X, UserList(Destino).Pos.Y).trigger
    
    If tOrg = eTrigger.ZONAPELEA Or tDst = eTrigger.ZONAPELEA Then
        If tOrg = tDst Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE
        End If
    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE
    End If

Exit Function
errhandler:
    TriggerZonaPelea = TRIGGER6_AUSENTE
    LogError ("Error en TriggerZonaPelea - " & Err.description)
End Function

Sub UserEnvenena(ByVal atacanteindex As Integer, ByVal victimaindex As Integer)
Dim ArmaObjInd As Integer, ObjInd As Integer
Dim num As Long

ArmaObjInd = UserList(atacanteindex).Invent.WeaponEqpObjIndex
ObjInd = 0

If ArmaObjInd > 0 Then
    If ObjData(ArmaObjInd).proyectil = 0 Then
        ObjInd = ArmaObjInd
    Else
        ObjInd = UserList(atacanteindex).Invent.MunicionEqpObjIndex
    End If
    
    If ObjInd > 0 Then
        If (ObjData(ObjInd).Envenena = 1) Then
            num = RandomNumber(1, 100)
            
            If num < 60 Then
                UserList(victimaindex).flags.Envenenado = 1
                Call WriteConsoleMsg(victimaindex, UserList(atacanteindex).name & " te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(atacanteindex, "Has envenenado a " & UserList(victimaindex).name & "!!", FontTypeNames.FONTTYPE_FIGHT)
            End If
        End If
    End If
End If

Call FlushBuffer(victimaindex)
End Sub
