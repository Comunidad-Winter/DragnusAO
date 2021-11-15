Attribute VB_Name = "AI"
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

Option Explicit

Public Enum TipoAI
    ESTATICO = 1
    MueveAlAzar = 2
    NpcMaloAtacaUsersBuenos = 3
    NPCDEFENSA = 4
    GuardiasAtacanCriminales = 5
    SigueAmo = 8
    npcAttackNPC = 9
    NpcPathfinding = 10
    NpcProtegeCastillo = 11
End Enum

Public Const ELEMENTALFUEGO As Integer = 93
Public Const ELEMENTALTIERRA As Integer = 94
Public Const ELEMENTALAGUA As Integer = 92

'Damos a los NPCs el mismo rango de visi�n que un PJ
Public Const RANGO_VISION_X As Byte = 12
Public Const RANGO_VISION_Y As Byte = 9

Public Enum e_Alineacion
    Neutro = 0
    Real = 1
    Caos = 2
End Enum

Public Enum e_Personalidad
''Inerte: no tiene objetivos de ningun tipo (npcs vendedores, curas, etc)
''Agresivo no magico: Su objetivo es acercarse a las victimas para atacarlas
''Agresivo magico: Su objetivo es mantenerse lo mas lejos posible de sus victimas y atacarlas con magia
''Mascota: Solo ataca a quien ataque a su amo.
''Pacifico: No ataca.
    ninguna = 0
    Inerte = 1
    AgresivoNoMagico = 2
    AgresivoMagico = 3
    Macota = 4
    Pacifico = 5
End Enum

'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'                        Modulo AI_NPC
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'AI de los NPC
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�
'?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�?�

Private Sub GuardiasAI(ByVal NpcIndex As Integer, Optional ByVal DelCaos As Boolean = False)
Dim nPos As WorldPos
Dim Headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim UI As Integer

For Headingloop = eHeading.NORTH To eHeading.WEST
    nPos = Npclist(NpcIndex).Pos
    If Npclist(NpcIndex).flags.Inmovilizado = 0 Or Headingloop = Npclist(NpcIndex).Char.Heading Then
        Call HeadtoPos(Headingloop, nPos)
        If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
            UI = MapData(nPos.Map, nPos.X, nPos.Y).userIndex
            If UI > 0 Then
                  If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible Then
                         '�ES CRIMINAL?
                        If Not DelCaos Then
                            If UserList(UI).Faccion.Alineacion = e_Alineacion.Caos Then
                                   If npcAttackUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(NpcIndex, Npclist(NpcIndex).Char.body, Npclist(NpcIndex).Char.head, Headingloop)
                                   End If
                                   Exit Sub
                            ElseIf Npclist(NpcIndex).flags.AttackedBy = UserList(UI).name _
                                      And Not Npclist(NpcIndex).flags.Follow Then
                                  
                                  If npcAttackUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(NpcIndex, Npclist(NpcIndex).Char.body, Npclist(NpcIndex).Char.head, Headingloop)
                                  End If
                                  Exit Sub
                            End If
                        Else
                            If UserList(UI).Faccion.Alineacion = e_Alineacion.Real Then
                                   
                                   If npcAttackUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(NpcIndex, Npclist(NpcIndex).Char.body, Npclist(NpcIndex).Char.head, Headingloop)
                                   End If
                                   Exit Sub
                            ElseIf Npclist(NpcIndex).flags.AttackedBy = UserList(UI).name _
                                      And Not Npclist(NpcIndex).flags.Follow Then
                                  
                                  If npcAttackUser(NpcIndex, UI) Then
                                        Call ChangeNPCChar(NpcIndex, Npclist(NpcIndex).Char.body, Npclist(NpcIndex).Char.head, Headingloop)
                                  End If
                                  Exit Sub
                            End If
                        End If
                  End If
            End If
        End If
    End If  'not inmovil
Next Headingloop

Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub HostilMalvadoAI(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim Headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim UI As Integer
Dim NPCI As Integer
Dim atacoPJ As Boolean

atacoPJ = False

For Headingloop = eHeading.NORTH To eHeading.WEST
    nPos = Npclist(NpcIndex).Pos
    If Npclist(NpcIndex).flags.Inmovilizado = 0 Or Npclist(NpcIndex).Char.Heading = Headingloop Then
        Call HeadtoPos(Headingloop, nPos)
        If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
            UI = MapData(nPos.Map, nPos.X, nPos.Y).userIndex
            NPCI = MapData(nPos.Map, nPos.X, nPos.Y).NpcIndex
            If UI > 0 And Not atacoPJ Then
                If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible Then
                    atacoPJ = True
                    If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then
                        Call NpcLanzaUnSpell(NpcIndex, UI)
                    End If
                    If npcAttackUser(NpcIndex, MapData(nPos.Map, nPos.X, nPos.Y).userIndex) Then
                        Call ChangeNPCChar(NpcIndex, Npclist(NpcIndex).Char.body, Npclist(NpcIndex).Char.head, Headingloop)
                    End If
                    Exit Sub
                End If
            ElseIf NPCI > 0 Then
                    If Npclist(NPCI).MaestroUser > 0 And Npclist(NPCI).flags.Paralizado = 0 Then
                        Call ChangeNPCChar(NpcIndex, Npclist(NpcIndex).Char.body, Npclist(NpcIndex).Char.head, Headingloop)
                        Call modCombat.npcAttackNPC(NpcIndex, NPCI, False)
                        Exit Sub
                    End If
            End If
        End If
    End If  'inmo
Next Headingloop

Call RestoreOldMovement(NpcIndex)

End Sub


Private Sub HostilBuenoAI(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim Headingloop As eHeading
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim UI As Integer
For Headingloop = eHeading.NORTH To eHeading.WEST
    nPos = Npclist(NpcIndex).Pos
    If Npclist(NpcIndex).flags.Inmovilizado = 0 Or Npclist(NpcIndex).Char.Heading = Headingloop Then
        Call HeadtoPos(Headingloop, nPos)
        If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
            UI = MapData(nPos.Map, nPos.X, nPos.Y).userIndex
            If UI > 0 Then
                If UserList(UI).name = Npclist(NpcIndex).flags.AttackedBy Then
                    If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible Then
                            If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                              Call NpcLanzaUnSpell(NpcIndex, UI)
                            End If
                            If npcAttackUser(NpcIndex, UI) Then
                                Call ChangeNPCChar(NpcIndex, Npclist(NpcIndex).Char.body, Npclist(NpcIndex).Char.head, Headingloop)
                            End If
                            Exit Sub
                    End If
                End If
            End If
        End If
    End If
Next Headingloop

Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub IrUsuarioCercano(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim Headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim UI As Integer
Dim SignoNS As Integer
Dim SignoEO As Integer

If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
    Select Case Npclist(NpcIndex).Char.Heading
        Case eHeading.NORTH
            SignoNS = -1
            SignoEO = 0
        Case eHeading.EAST
            SignoNS = 0
            SignoEO = 1
        Case eHeading.SOUTH
            SignoNS = 1
            SignoEO = 0
        Case eHeading.WEST
            SignoEO = -1
            SignoNS = 0
    End Select
    
    For Y = Npclist(NpcIndex).Pos.Y To Npclist(NpcIndex).Pos.Y + SignoNS * RANGO_VISION_Y Step IIf(SignoNS = 0, 1, SignoNS)
        For X = Npclist(NpcIndex).Pos.X To Npclist(NpcIndex).Pos.X + SignoEO * RANGO_VISION_X Step IIf(SignoEO = 0, 1, SignoEO)
            
            If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                   UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).userIndex
                   If UI > 0 Then
                      If UserList(UI).flags.Muerto = 0 Then
                            If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                            Exit Sub
                      End If
                   End If
            End If
            
        Next X
    Next Y
    
Else
    For Y = Npclist(NpcIndex).Pos.Y - RANGO_VISION_Y To Npclist(NpcIndex).Pos.Y + RANGO_VISION_Y
        For X = Npclist(NpcIndex).Pos.X - RANGO_VISION_X To Npclist(NpcIndex).Pos.X + RANGO_VISION_X
            If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).userIndex
                If UI > 0 Then
                     If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Then
                         If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                         tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).userIndex).Pos)
                         Call MoveNPCChar(NpcIndex, tHeading)
                         Exit Sub
                     End If
                End If
            End If
        Next X
    Next Y
End If

Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub SeguirAgresor(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim Headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim UI As Integer

Dim SignoNS As Integer
Dim SignoEO As Integer

If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
    Select Case Npclist(NpcIndex).Char.Heading
        Case eHeading.NORTH
            SignoNS = -1
            SignoEO = 0
        Case eHeading.EAST
            SignoNS = 0
            SignoEO = 1
        Case eHeading.SOUTH
            SignoNS = 1
            SignoEO = 0
        Case eHeading.WEST
            SignoEO = -1
            SignoNS = 0
    End Select
    
    
    For Y = Npclist(NpcIndex).Pos.Y To Npclist(NpcIndex).Pos.Y + SignoNS * RANGO_VISION_Y Step IIf(SignoNS = 0, 1, SignoNS)
        For X = Npclist(NpcIndex).Pos.X To Npclist(NpcIndex).Pos.X + SignoEO * RANGO_VISION_X Step IIf(SignoEO = 0, 1, SignoEO)

            If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).userIndex
                If UI > 0 Then
                    If UserList(UI).name = Npclist(NpcIndex).flags.AttackedBy Then
                        If Npclist(NpcIndex).MaestroUser > 0 Then
                            If UserList(Npclist(NpcIndex).MaestroUser).Faccion.Alineacion = UserList(UI).Faccion.Alineacion And Not (UserList(UI).Faccion.Alineacion = e_Alineacion.Neutro Or UserList(Npclist(NpcIndex).MaestroUser).Faccion.Alineacion = e_Alineacion.Neutro) Then
                                Call WriteConsoleMsg(Npclist(NpcIndex).MaestroUser, "La mascota no atacar� a miembros de tu faccion.", FontTypeNames.FONTTYPE_INFO)
                                Call FlushBuffer(Npclist(NpcIndex).MaestroUser) 'CHECK
                                Npclist(NpcIndex).flags.AttackedBy = vbNullString
                                Exit Sub
                            End If
                        End If
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                             If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                                  Call NpcLanzaUnSpell(NpcIndex, UI)
                             End If
                             Exit Sub
                        End If
                    End If
                End If
            End If
        Next X
    Next Y
Else
    For Y = Npclist(NpcIndex).Pos.Y - RANGO_VISION_Y To Npclist(NpcIndex).Pos.Y + RANGO_VISION_Y
        For X = Npclist(NpcIndex).Pos.X - RANGO_VISION_X To Npclist(NpcIndex).Pos.X + RANGO_VISION_X
            If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).userIndex
                If UI > 0 Then
                    If UserList(UI).name = Npclist(NpcIndex).flags.AttackedBy Then
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                             If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                                  Call NpcLanzaUnSpell(NpcIndex, UI)
                             End If
                             tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).userIndex).Pos)
                             Call MoveNPCChar(NpcIndex, tHeading)
                             Exit Sub
                        End If
                    End If
                End If
            End If
        Next X
    Next Y
End If
Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)

If Npclist(NpcIndex).MaestroUser = 0 Then
    Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
    Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
    Npclist(NpcIndex).flags.AttackedBy = vbNullString
End If

End Sub

Private Sub PersigueCiudadano(ByVal NpcIndex As Integer)
Dim UI As Integer
Dim nPos As WorldPos
Dim Headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
For Y = Npclist(NpcIndex).Pos.Y - RANGO_VISION_Y To Npclist(NpcIndex).Pos.Y + RANGO_VISION_Y
    For X = Npclist(NpcIndex).Pos.X - RANGO_VISION_X To Npclist(NpcIndex).Pos.X + RANGO_VISION_X
        If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
           UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).userIndex
           If UI > 0 Then
                If Not UserList(UI).Faccion.Alineacion = e_Alineacion.Caos Then
                   If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                        If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                              Call NpcLanzaUnSpell(NpcIndex, UI)
                        End If
                        tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).userIndex).Pos)
                        Call MoveNPCChar(NpcIndex, tHeading)
                        Exit Sub
                   End If
                End If
           End If
        End If
    Next X
Next Y

Call RestoreOldMovement(NpcIndex)

End Sub


Private Sub PersigueCriminal(ByVal NpcIndex As Integer)
Dim UI As Integer
Dim nPos As WorldPos
Dim Headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim SignoNS As Integer
Dim SignoEO As Integer

If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
    Select Case Npclist(NpcIndex).Char.Heading
        Case eHeading.NORTH
            SignoNS = -1
            SignoEO = 0
        Case eHeading.EAST
            SignoNS = 0
            SignoEO = 1
        Case eHeading.SOUTH
            SignoNS = 1
            SignoEO = 0
        Case eHeading.WEST
            SignoEO = -1
            SignoNS = 0
    End Select
    
    For Y = Npclist(NpcIndex).Pos.Y To Npclist(NpcIndex).Pos.Y + SignoNS * RANGO_VISION_Y Step IIf(SignoNS = 0, 1, SignoNS)
        For X = Npclist(NpcIndex).Pos.X To Npclist(NpcIndex).Pos.X + SignoEO * RANGO_VISION_X Step IIf(SignoEO = 0, 1, SignoEO)


            If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
               UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).userIndex
               If UI > 0 Then
                    If Not UserList(UI).Faccion.Alineacion = e_Alineacion.Real And Not UserList(UI).Faccion.Alineacion = e_Alineacion.Neutro Then
                       If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Then
                            If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                                  Call NpcLanzaUnSpell(NpcIndex, UI)
                            End If
                            Exit Sub
                       End If
                    End If
               End If
            End If


        Next X
    Next Y
Else
    For Y = Npclist(NpcIndex).Pos.Y - RANGO_VISION_Y To Npclist(NpcIndex).Pos.Y + RANGO_VISION_Y
        For X = Npclist(NpcIndex).Pos.X - RANGO_VISION_X To Npclist(NpcIndex).Pos.X + RANGO_VISION_X
            If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
               UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).userIndex
               If UI > 0 Then
                    If Not UserList(UI).Faccion.Alineacion = e_Alineacion.Real And Not UserList(UI).Faccion.Alineacion = e_Alineacion.Neutro Then
                       If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Then
                            If Npclist(NpcIndex).flags.LanzaSpells > 0 Then
                                  Call NpcLanzaUnSpell(NpcIndex, UI)
                            End If
                            If Npclist(NpcIndex).flags.Inmovilizado = 1 Then Exit Sub
                            tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).userIndex).Pos)
                            Call MoveNPCChar(NpcIndex, tHeading)
                            Exit Sub
                       End If
                    End If
               End If
            End If
        Next X
    Next Y
End If
Call RestoreOldMovement(NpcIndex)

End Sub

Public Sub SeguirAmo(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim Headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim UI As Integer
For Y = Npclist(NpcIndex).Pos.Y - 10 To Npclist(NpcIndex).Pos.Y + 10
    For X = Npclist(NpcIndex).Pos.X - 10 To Npclist(NpcIndex).Pos.X + 10
        If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
            If Npclist(NpcIndex).target = 0 And Npclist(NpcIndex).TargetNPC = 0 Then
                UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).userIndex
                If UI > 0 Then
                   If UserList(UI).flags.Muerto = 0 _
                   And UserList(UI).flags.invisible = 0 _
                   And UserList(UI).flags.Oculto = 0 _
                   And UI = Npclist(NpcIndex).MaestroUser _
                   And Distancia(Npclist(NpcIndex).Pos, UserList(UI).Pos) > 3 Then
                        tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).userIndex).Pos)
                        Call MoveNPCChar(NpcIndex, tHeading)
                        Exit Sub
                   End If
                End If
            End If
        End If
    Next X
Next Y

Call RestoreOldMovement(NpcIndex)

End Sub

Private Sub AinpcAttackNPC(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim Headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim NI As Integer
Dim bNoEsta As Boolean

Dim SignoNS As Integer
Dim SignoEO As Integer

If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
    Select Case Npclist(NpcIndex).Char.Heading
        Case eHeading.NORTH
            SignoNS = -1
            SignoEO = 0
        Case eHeading.EAST
            SignoNS = 0
            SignoEO = 1
        Case eHeading.SOUTH
            SignoNS = 1
            SignoEO = 0
        Case eHeading.WEST
            SignoEO = -1
            SignoNS = 0
    End Select
    
    For Y = Npclist(NpcIndex).Pos.Y To Npclist(NpcIndex).Pos.Y + SignoNS * RANGO_VISION_Y Step IIf(SignoNS = 0, 1, SignoNS)
        For X = Npclist(NpcIndex).Pos.X To Npclist(NpcIndex).Pos.X + SignoEO * RANGO_VISION_X Step IIf(SignoEO = 0, 1, SignoEO)
            If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
               NI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).NpcIndex
               If NI > 0 Then
                    If Npclist(NpcIndex).TargetNPC = NI Then
                         bNoEsta = True
                         If Npclist(NpcIndex).Numero = ELEMENTALFUEGO Then
                             Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)
                             If Npclist(NI).NPCtype = DRAGON Then
                                Npclist(NI).CanAttack = 1
                                Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)
                             End If
                         Else
                            'aca verificamosss la distancia de ataque
                            If Distancia(Npclist(NpcIndex).Pos, Npclist(NI).Pos) <= 1 Then
                                Call modCombat.npcAttackNPC(NpcIndex, NI)
                            End If
                         End If
                         Exit Sub
                    End If
               End If
            End If
        Next X
    Next Y
Else
    For Y = Npclist(NpcIndex).Pos.Y - RANGO_VISION_Y To Npclist(NpcIndex).Pos.Y + RANGO_VISION_Y
        For X = Npclist(NpcIndex).Pos.X - RANGO_VISION_Y To Npclist(NpcIndex).Pos.X + RANGO_VISION_Y
            If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
               NI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).NpcIndex
               If NI > 0 Then
                    If Npclist(NpcIndex).TargetNPC = NI Then
                         bNoEsta = True
                         If Npclist(NpcIndex).Numero = ELEMENTALFUEGO Then
                             Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)
                             If Npclist(NI).NPCtype = DRAGON Then
                                Npclist(NI).CanAttack = 1
                                Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)
                             End If
                         Else
                            'aca verificamosss la distancia de ataque
                            If Distancia(Npclist(NpcIndex).Pos, Npclist(NI).Pos) <= 1 Then
                                Call modCombat.npcAttackNPC(NpcIndex, NI)
                            End If
                         End If
                         If Npclist(NpcIndex).flags.Inmovilizado = 1 Then Exit Sub
                         If Npclist(NpcIndex).TargetNPC = 0 Then Exit Sub
                         tHeading = FindDirection(Npclist(NpcIndex).Pos, Npclist(MapData(Npclist(NpcIndex).Pos.Map, X, Y).NpcIndex).Pos)
                         Call MoveNPCChar(NpcIndex, tHeading)
                         Exit Sub
                    End If
               End If
            End If
        Next X
    Next Y
End If

If Not bNoEsta Then
    If Npclist(NpcIndex).MaestroUser > 0 Then
        Call FollowAmo(NpcIndex)
    Else
        Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
        Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
    End If
End If
    
End Sub

Function NPCAI(ByVal NpcIndex As Integer)
On Error GoTo ErrorHandler
        '<<<<<<<<<<< Ataques >>>>>>>>>>>>>>>>
        If Npclist(NpcIndex).MaestroUser = 0 Then
            'Busca a alguien para atacar
            '�Es un guardia?
            If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                    Call GuardiasAI(NpcIndex)
            ElseIf Npclist(NpcIndex).NPCtype = eNPCType.Guardiascaos Then
                    Call GuardiasAI(NpcIndex, True)
            ElseIf Npclist(NpcIndex).Hostile And Npclist(NpcIndex).Stats.Alineacion <> 0 Then
                    Call HostilMalvadoAI(NpcIndex)
            ElseIf Npclist(NpcIndex).Hostile And Npclist(NpcIndex).Stats.Alineacion = 0 Then
                    Call HostilBuenoAI(NpcIndex)
            End If
        Else
            If False Then Exit Function
            'Evitamos que ataque a su amo, a menos
            'que el amo lo ataque.
            'Call HostilBuenoAI(NpcIndex)
        End If
        
        '<<<<<<<<<<<Movimiento>>>>>>>>>>>>>>>>
        Select Case Npclist(NpcIndex).Movement
            Case TipoAI.MueveAlAzar
                If Npclist(NpcIndex).flags.Inmovilizado = 1 Then Exit Function
                If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                    End If
                    Call PersigueCriminal(NpcIndex)
                ElseIf Npclist(NpcIndex).NPCtype = eNPCType.Guardiascaos Then
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                    End If
                    Call PersigueCiudadano(NpcIndex)
                Else
                    If RandomNumber(1, 12) = 3 Then
                        Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                    End If
                End If
            'Protege al castillo
            Case TipoAI.NpcProtegeCastillo
                Call ProtegerCastillo(NpcIndex)
            'Va hacia el usuario cercano
            Case TipoAI.NpcMaloAtacaUsersBuenos
                Call IrUsuarioCercano(NpcIndex)
            'Va hacia el usuario que lo ataco(FOLLOW)
            Case TipoAI.NPCDEFENSA
                Call SeguirAgresor(NpcIndex)
            'Persigue criminales
            Case TipoAI.GuardiasAtacanCriminales
                Call PersigueCriminal(NpcIndex)
            Case TipoAI.SigueAmo
                If Npclist(NpcIndex).flags.Inmovilizado = 1 Then Exit Function
                Call SeguirAmo(NpcIndex)
                If RandomNumber(1, 12) = 3 Then
                    Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                End If
            Case TipoAI.npcAttackNPC
                Call AinpcAttackNPC(NpcIndex)
            Case TipoAI.NpcPathfinding
                If Npclist(NpcIndex).flags.Inmovilizado = 1 Then Exit Function
                If ReCalculatePath(NpcIndex) Then
                    Call PathFindingAI(NpcIndex)
                    'Existe el camino?
                    If Npclist(NpcIndex).PFINFO.NoPath Then 'Si no existe nos movemos al azar
                        'Move randomly
                        Call MoveNPCChar(NpcIndex, RandomNumber(eHeading.NORTH, eHeading.WEST))
                    End If
                Else
                    If Not PathEnd(NpcIndex) Then
                        Call FollowPath(NpcIndex)
                    Else
                        Npclist(NpcIndex).PFINFO.PathLenght = 0
                    End If
                End If

        End Select


Exit Function


ErrorHandler:
    Call LogError("NPCAI " & Npclist(NpcIndex).name & " " & Npclist(NpcIndex).MaestroUser & " " & Npclist(NpcIndex).MaestroNpc & " mapa:" & Npclist(NpcIndex).Pos.Map & " x:" & Npclist(NpcIndex).Pos.X & " y:" & Npclist(NpcIndex).Pos.Y & " Mov:" & Npclist(NpcIndex).Movement & " TargU:" & Npclist(NpcIndex).target & " TargN:" & Npclist(NpcIndex).TargetNPC)
    Dim MiNPC As npc
    MiNPC = Npclist(NpcIndex)
    Call QuitarNPC(NpcIndex)
    Call ReSpawnNpc(MiNPC)
    
End Function


Function UserNear(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Returns True if there is an user adjacent to the npc position.
'#################################################################
UserNear = Not Int(Distance(Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, UserList(Npclist(NpcIndex).PFINFO.TargetUser).Pos.X, UserList(Npclist(NpcIndex).PFINFO.TargetUser).Pos.Y)) > 1
End Function

Function ReCalculatePath(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Returns true if we have to seek a new path
'#################################################################
If Npclist(NpcIndex).PFINFO.PathLenght = 0 Then
    ReCalculatePath = True
ElseIf Not UserNear(NpcIndex) And Npclist(NpcIndex).PFINFO.PathLenght = Npclist(NpcIndex).PFINFO.CurPos - 1 Then
    ReCalculatePath = True
End If
End Function

Function SimpleAI(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Old Ore4 AI function
'#################################################################
Dim nPos As WorldPos
Dim Headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer

For Y = Npclist(NpcIndex).Pos.Y - 5 To Npclist(NpcIndex).Pos.Y + 5    'Makes a loop that looks at
    For X = Npclist(NpcIndex).Pos.X - 5 To Npclist(NpcIndex).Pos.X + 5   '5 tiles in every direction
           'Make sure tile is legal
            If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then
                'look for a user
                If MapData(Npclist(NpcIndex).Pos.Map, X, Y).userIndex > 0 Then
                    'Move towards user
                    tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).userIndex).Pos)
                    MoveNPCChar NpcIndex, tHeading
                    'Leave
                    Exit Function
                End If
            End If
    Next X
Next Y

End Function

Function PathEnd(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock
'Returns if the npc has arrived to the end of its path
'#################################################################
PathEnd = Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.PathLenght
End Function

Function FollowPath(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock
'Moves the npc.
'#################################################################

Dim tmpPos As WorldPos
Dim tHeading As Byte

tmpPos.Map = Npclist(NpcIndex).Pos.Map
tmpPos.X = Npclist(NpcIndex).PFINFO.Path(Npclist(NpcIndex).PFINFO.CurPos).Y ' invert� las coordenadas
tmpPos.Y = Npclist(NpcIndex).PFINFO.Path(Npclist(NpcIndex).PFINFO.CurPos).X

'Debug.Print "(" & tmpPos.X & "," & tmpPos.Y & ")"

tHeading = FindDirection(Npclist(NpcIndex).Pos, tmpPos)

MoveNPCChar NpcIndex, tHeading

Npclist(NpcIndex).PFINFO.CurPos = Npclist(NpcIndex).PFINFO.CurPos + 1

End Function

Function PathFindingAI(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock / 11-07-02
'www.geocities.com/gmorgolock
'morgolock@speedy.com.ar
'This function seeks the shortest path from the Npc
'to the user's location.
'#################################################################
Dim nPos As WorldPos
Dim Headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer

For Y = Npclist(NpcIndex).Pos.Y - 10 To Npclist(NpcIndex).Pos.Y + 10    'Makes a loop that looks at
     For X = Npclist(NpcIndex).Pos.X - 10 To Npclist(NpcIndex).Pos.X + 10   '5 tiles in every direction

         'Make sure tile is legal
         If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then
         
             'look for a user
             If MapData(Npclist(NpcIndex).Pos.Map, X, Y).userIndex > 0 Then
                 'Move towards user
                  Dim tmpUserIndex As Integer
                  tmpUserIndex = MapData(Npclist(NpcIndex).Pos.Map, X, Y).userIndex
                  With UserList(tmpUserIndex)
                    If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And .flags.AdminPerseguible Then
                        'We have to invert the coordinates, this is because
                        'ORE refers to maps in converse way of my pathfinding
                        'routines.
                        Npclist(NpcIndex).PFINFO.target.X = UserList(tmpUserIndex).Pos.Y
                        Npclist(NpcIndex).PFINFO.target.Y = UserList(tmpUserIndex).Pos.X 'ops!
                        Npclist(NpcIndex).PFINFO.TargetUser = tmpUserIndex
                        Call SeekPath(NpcIndex)
                        Exit Function
                    End If
                  End With
             End If
             
         End If
              
     Next X
 Next Y
End Function


Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer, ByVal userIndex As Integer)

If UserList(userIndex).flags.invisible = 1 Or UserList(userIndex).flags.Oculto = 1 Then Exit Sub

Dim k As Integer
k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
Call NpcLanzaSpellSobreUser(NpcIndex, userIndex, Npclist(NpcIndex).Spells(k))

End Sub


Sub NpcLanzaUnSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer)

Dim k As Integer
    k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNPC, Npclist(NpcIndex).Spells(k))

End Sub

Private Sub ProtegerCastillo(ByVal NpcIndex As Integer)
Dim nPos As WorldPos
Dim Headingloop As Byte
Dim tHeading As Byte
Dim Y As Integer
Dim X As Integer
Dim UI As Integer
Dim SignoNS As Integer
Dim SignoEO As Integer

If Npclist(NpcIndex).flags.Inmovilizado = 1 Then
    Select Case Npclist(NpcIndex).Char.Heading
        Case eHeading.NORTH
            SignoNS = -1
            SignoEO = 0
        Case eHeading.EAST
            SignoNS = 0
            SignoEO = 1
        Case eHeading.SOUTH
            SignoNS = 1
            SignoEO = 0
        Case eHeading.WEST
            SignoEO = -1
            SignoNS = 0
    End Select
    
    For Y = Npclist(NpcIndex).Pos.Y To Npclist(NpcIndex).Pos.Y + SignoNS * RANGO_VISION_Y Step IIf(SignoNS = 0, 1, SignoNS)
        For X = Npclist(NpcIndex).Pos.X To Npclist(NpcIndex).Pos.X + SignoEO * RANGO_VISION_X Step IIf(SignoEO = 0, 1, SignoEO)
            
            If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                   UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).userIndex
                   If UI > 0 Then
                      If UserList(UI).flags.Muerto = 0 And Not EsCastillo(UI) Then
                            If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                            Exit Sub
                      End If
                   End If
            End If
            
        Next X
    Next Y
    
Else
    For Y = Npclist(NpcIndex).Pos.Y - RANGO_VISION_Y To Npclist(NpcIndex).Pos.Y + RANGO_VISION_Y
        For X = Npclist(NpcIndex).Pos.X - RANGO_VISION_X To Npclist(NpcIndex).Pos.X + RANGO_VISION_X
            If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                UI = MapData(Npclist(NpcIndex).Pos.Map, X, Y).userIndex
                If UI > 0 Then
                     If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible And Not EsCastillo(UI) Then
                         If Npclist(NpcIndex).flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                         tHeading = FindDirection(Npclist(NpcIndex).Pos, UserList(MapData(Npclist(NpcIndex).Pos.Map, X, Y).userIndex).Pos)
                         Call MoveNPCChar(NpcIndex, tHeading)
                         Exit Sub
                     End If
                End If
            End If
        Next X
    Next Y
End If

Call RestoreOldMovement(NpcIndex)

End Sub
