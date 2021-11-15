Attribute VB_Name = "modFacciones"
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

Public ArmaduraImperial1 As Integer 'Primer jerarquia
Public ArmaduraImperial2 As Integer 'Segunda jerarquía
Public ArmaduraImperial3 As Integer 'Enanos
Public TunicaMagoImperial As Integer 'Magos
Public TunicaMagoImperialEnanos As Integer 'Magos

Public VestimentaImperialHumano As Integer
Public VestimentaImperialEnano As Integer
Public TunicaConspicuaHumano As Integer
Public TunicaConspicuaEnano As Integer
Public ArmaduraNobilisimaHumano As Integer
Public ArmaduraNobilisimaEnano As Integer
Public ArmaduraGranSacerdote As Integer

Public VestimentaLegionHumano As Integer
Public VestimentaLegionEnano As Integer
Public TunicaLobregaHumano As Integer
Public TunicaLobregaEnano As Integer
Public TunicaEgregiaHumano As Integer
Public TunicaEgregiaEnano As Integer
Public SacerdoteDemoniaco As Integer

Public ArmaduraCaos1 As Integer
Public TunicaMagoCaos As Integer
Public TunicaMagoCaosEnanos As Integer
Public ArmaduraCaos2 As Integer
Public ArmaduraCaos3 As Integer

Public Const ExpAlUnirse As Long = 50000
Public Const ExpX100 As Integer = 5000


Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Handles the entrance of users to the "Armada Real"
'***************************************************
If UserList(UserIndex).Faccion.Alineacion = e_Alineacion.Real Then
    Call WriteChatOverHead(UserIndex, "¡¡¡Ya perteneces a las tropas reales!!! Ve a combatir criminales", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.Alineacion = e_Alineacion.Caos Then
    Call WriteChatOverHead(UserIndex, "¡¡¡Maldito insolente!!! vete de aqui seguidor de las sombras", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.SalioFaccion Then
    Call WriteChatOverHead(UserIndex, "Acabas de salir de la Legion Oscura, no confio en ti, vuelve en " & UserList(UserIndex).Faccion.SalioFaccionCounter & " horas y veremos.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 25 Then
    Call WriteChatOverHead(UserIndex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

UserList(UserIndex).Faccion.Alineacion = e_Alineacion.Real
UserList(UserIndex).Faccion.RangoFaccionario = 0

Call WriteVar(CharPath, "FACCIONES", "Alineacion", CStr(UserList(UserIndex).Faccion.Alineacion))
Call WriteVar(CharPath & UserList(UserIndex).name & ".chr", "FACCIONES", "RangoFaccionario", CStr(UserList(UserIndex).Faccion.RangoFaccionario))

If UserList(UserIndex).GuildIndex Then _
        Call m_ValidarPermanencia(UserIndex)

Call LogEjercitoReal(UserList(UserIndex).name & " ingresó el " & Date & " cuando era nivel " & UserList(UserIndex).Stats.ELV)

End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)


End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer)

    UserList(UserIndex).Faccion.Alineacion = e_Alineacion.Neutro
    'Call PerderItemsFaccionarios(UserIndex)
    Call WriteConsoleMsg(UserIndex, "Has sido expulsado de las tropas reales!!!.", FontTypeNames.FONTTYPE_FIGHT)
    'Desequipamos la armadura real si está equipada
    If UserList(UserIndex).Invent.ArmourEqpObjIndex Then _
        If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    
    Call WriteVar(CharPath & UserList(UserIndex).name & ".chr", "FACCIONES", "Alineacion", CStr(UserList(UserIndex).Faccion.Alineacion))
    Call WriteVar(CharPath & UserList(UserIndex).name & ".chr", "FACCIONES", "RangoFaccionario", CStr(UserList(UserIndex).Faccion.RangoFaccionario))
    
    UserList(UserIndex).Faccion.SalioFaccion = 1
    UserList(UserIndex).Faccion.SalioFaccionCounter = 48 'No volves a hacerte faccion por 2 dias.
    
    If UserList(UserIndex).GuildIndex Then _
        Call m_ValidarPermanencia(UserIndex)
End Sub

Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer)

    UserList(UserIndex).Faccion.Alineacion = e_Alineacion.Neutro
    'Call PerderItemsFaccionarios(UserIndex)
    Call WriteConsoleMsg(UserIndex, "Has sido expulsado de la legión oscura!!!.", FontTypeNames.FONTTYPE_FIGHT)
    'Desequipamos la armadura caos si está equipada
    If UserList(UserIndex).Invent.ArmourEqpObjIndex Then _
        If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    
    Call WriteVar(CharPath & UserList(UserIndex).name & ".chr", "FACCIONES", "Alineacion", CStr(UserList(UserIndex).Faccion.Alineacion))
    Call WriteVar(CharPath & UserList(UserIndex).name & ".chr", "FACCIONES", "RangoFaccionario", CStr(UserList(UserIndex).Faccion.RangoFaccionario))
    
    UserList(UserIndex).Faccion.SalioFaccion = 1
    UserList(UserIndex).Faccion.SalioFaccionCounter = 48 'No volves a hacerte faccion por 2 dias.
    
    If UserList(UserIndex).GuildIndex Then _
        Call m_ValidarPermanencia(UserIndex)
End Sub

Public Function TituloReal(ByVal UserIndex As Integer) As String
'***************************************************
'Autor: Unknown
'Last Modification: 23/01/2007 Pablo (ToxicWaste)
'Handles the titles of the members of the "Armada Real"
'***************************************************
Select Case UserList(UserIndex).Faccion.RangoFaccionario
'Rango 1: Aprendiz (30 Criminales)
'Rango 2: Escudero (70 Criminales)
'Rango 3: Soldado (130 Criminales)
'Rango 4: Sargento (210 Criminales)
'Rango 5: Caballero (320 Criminales)
'Rango 6: Comandante (460 Criminales)
'Rango 7: Capitán (640 Criminales + > lvl 27)
'Rango 8: Senescal (870 Criminales)
'Rango 9: Mariscal (1160 Criminales)
'Rango 10: Condestable (2000 Criminales + > lvl 30)
'Rangos de Honor de la Armada Real: (Consejo de Bander)
'Rango 11: Ejecutor Imperial (2500 Criminales + 2.000.000 Nobleza)
'Rango 12: Protector del Reino (3000 Criminales + 3.000.000 Nobleza)
'Rango 13: Avatar de la Justicia (3500 Criminales + 4.000.000 Nobleza + > lvl 35)
'Rango 14: Guardián del Bien (4000 Criminales + 5.000.000 Nobleza + > lvl 36)
'Rango 15: Campeón de la Luz (5000 Criminales + 6.000.000 Nobleza + > lvl 37)
    
    Case 0
        TituloReal = "Aprendiz"
    Case 1
        TituloReal = "Escudero"
    Case 2
        TituloReal = "Soldado"
    Case 3
        TituloReal = "Sargento"
    Case 4
        TituloReal = "Caballero"
    Case 5
        TituloReal = "Comandante"
    Case 6
        TituloReal = "Capitán"
    Case 7
        TituloReal = "Senescal"
    Case 8
        TituloReal = "Mariscal"
    Case 9
        TituloReal = "Condestable"
    Case 10
        TituloReal = "Ejecutor Imperial"
    Case 11
        TituloReal = "Protector del Reino"
    Case 12
        TituloReal = "Avatar de la Justicia"
    Case 13
        TituloReal = "Guardián del Bien"
    Case Else
        TituloReal = "Campeón de la Luz"
End Select


End Function

Public Sub EnlistarCaos(ByVal UserIndex As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Handles the entrance of users to the "Legión Oscura"
'***************************************************

If UserList(UserIndex).Faccion.Alineacion = e_Alineacion.Caos Then
    Call WriteChatOverHead(UserIndex, "¡¡¡Ya perteneces a la legión oscura!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.Alineacion = e_Alineacion.Real Then
    Call WriteChatOverHead(UserIndex, "Las sombras reinarán en Argentum. ¡¡¡Fuera de aqui insecto Real!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.SalioFaccion Then
    Call WriteChatOverHead(UserIndex, "Acabas de salir de la Armada Real, no confio en ti, vuelve en " & UserList(UserIndex).Faccion.SalioFaccionCounter & " horas y veremos.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 25 Then
    Call WriteChatOverHead(UserIndex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

UserList(UserIndex).Faccion.Alineacion = e_Alineacion.Caos
UserList(UserIndex).Faccion.RangoFaccionario = 0

Call WriteVar(CharPath & UserList(UserIndex).name & ".chr", "FACCIONES", "Alineacion", CStr(UserList(UserIndex).Faccion.Alineacion))
Call WriteVar(CharPath & UserList(UserIndex).name & ".chr", "FACCIONES", "RangoFaccionario", CStr(UserList(UserIndex).Faccion.RangoFaccionario))

If UserList(UserIndex).GuildIndex Then _
        Call m_ValidarPermanencia(UserIndex)

Call LogEjercitoCaos(UserList(UserIndex).name & " ingresó el " & Date & " cuando era nivel " & UserList(UserIndex).Stats.ELV)

End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Handles the way of gaining new ranks in the "Legión Oscura"
'***************************************************


End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As String
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 23/01/2007 Pablo (ToxicWaste)
'Handles the titles of the members of the "Legión Oscura"
'***************************************************
'Rango 1: Acólito (70)
'Rango 2: Alma Corrupta (160)
'Rango 3: Paria (300)
'Rango 4: Condenado (490)
'Rango 5: Esbirro (740)
'Rango 6: Sanguinario (1100)
'Rango 7: Corruptor (1500 + lvl 27)
'Rango 8: Heraldo Impio (2010)
'Rango 9: Caballero de la Oscuridad (2700)
'Rango 10: Señor del Miedo (4600 + lvl 30)
'Rango 11: Ejecutor Infernal (5800 + lvl 31)
'Rango 12: Protector del Averno (6990 + lvl 33)
'Rango 13: Avatar de la Destrucción (8100 + lvl 35)
'Rango 14: Guardián del Mal (9300 + lvl 36)
'Rango 15: Campeón de la Oscuridad (11500 + lvl 37)

Select Case UserList(UserIndex).Faccion.RangoFaccionario
    Case 0
        TituloCaos = "Acólito"
    Case 1
        TituloCaos = "Alma Corrupta"
    Case 2
        TituloCaos = "Paria"
    Case 3
        TituloCaos = "Condenado"
    Case 4
        TituloCaos = "Esbirro"
    Case 5
        TituloCaos = "Sanguinario"
    Case 6
        TituloCaos = "Corruptor"
    Case 7
        TituloCaos = "Heraldo Impío"
    Case 8
        TituloCaos = "Caballero de la Oscuridad"
    Case 9
        TituloCaos = "Señor del Miedo"
    Case 10
        TituloCaos = "Ejecutor Infernal"
    Case 11
        TituloCaos = "Protector del Averno"
    Case 12
        TituloCaos = "Avatar de la Destrucción"
    Case 13
        TituloCaos = "Guardián del Mal"
    Case Else
        TituloCaos = "Campeón de la Oscuridad"
End Select

End Function
