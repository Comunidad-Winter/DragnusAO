Attribute VB_Name = "Trabajo"
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

Public Sub DoPermanecerOculto(ByVal userIndex As Integer)
'********************************************************
'Autor: Nacho (Integer)
'Last Modif: 28/01/2007
'Chequea si ya debe mostrarse
'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
'********************************************************

UserList(userIndex).Counters.TiempoOculto = UserList(userIndex).Counters.TiempoOculto - 1
If UserList(userIndex).Counters.TiempoOculto <= 0 Then
    
    UserList(userIndex).Counters.TiempoOculto = IntervaloOculto
    If UserList(userIndex).Clase = eClass.Hunter And UserList(userIndex).Stats.UserSkills(eSkill.Ocultarse) > 90 Then
        If UserList(userIndex).Invent.ArmourEqpObjIndex = 648 Or UserList(userIndex).Invent.ArmourEqpObjIndex = 360 Then
            Exit Sub
        End If
    End If
    UserList(userIndex).Counters.TiempoOculto = 0
    UserList(userIndex).flags.Oculto = 0
    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageSetInvisible(UserList(userIndex).Char.CharIndex, False))
    Call WriteConsoleMsg(userIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
End If



Exit Sub

errhandler:
    Call LogError("Error en Sub DoPermanecerOculto")


End Sub

Public Sub DoOcultarse(ByVal userIndex As Integer)
'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
'Modifique la fórmula y ahora anda bien.
On Error GoTo errhandler

Dim Suerte As Double
Dim res As Integer
Dim Skill As Integer
Skill = UserList(userIndex).Stats.UserSkills(eSkill.Ocultarse)

Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100

res = RandomNumber(1, 100)

If res <= Suerte Then

    UserList(userIndex).flags.Oculto = 1
    Suerte = (-0.000001 * (100 - Skill) ^ 3)
    Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
    Suerte = Suerte + (-0.0088 * (100 - Skill))
    Suerte = Suerte + (0.9571)
    Suerte = Suerte * IntervaloOculto
    UserList(userIndex).Counters.TiempoOculto = Suerte
  
    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageSetInvisible(UserList(userIndex).Char.CharIndex, True))

    Call WriteConsoleMsg(userIndex, "¡Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
    Call SubirSkill(userIndex, Ocultarse)
Else
    '[CDT 17-02-2004]
    If Not UserList(userIndex).flags.UltimoMensaje = 4 Then
        Call WriteConsoleMsg(userIndex, "¡No has logrado esconderte!", FontTypeNames.FONTTYPE_INFO)
        UserList(userIndex).flags.UltimoMensaje = 4
    End If
    '[/CDT]
End If

UserList(userIndex).Counters.Ocultando = UserList(userIndex).Counters.Ocultando + 1

Exit Sub

errhandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub


Public Sub DoNavega(ByVal userIndex As Integer, ByRef Barco As ObjData, ByVal Slot As Integer)

If UserList(userIndex).flags.Montado = 1 Then
Call WriteConsoleMsg(userIndex, "No puedes navegar estando montado.", FontTypeNames.FONTTYPE_INFO)
Exit Sub
End If

Dim ModNave As Long
ModNave = ModNavegacion(UserList(userIndex).Clase)

If UserList(userIndex).Stats.UserSkills(eSkill.Navegacion) / ModNave < Barco.MinSkill Then
    Call WriteConsoleMsg(userIndex, "No tenes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(userIndex, "Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

UserList(userIndex).Invent.BarcoObjIndex = UserList(userIndex).Invent.Object(Slot).ObjIndex
UserList(userIndex).Invent.BarcoSlot = Slot

If UserList(userIndex).flags.Navegando = 0 Then
    
    UserList(userIndex).Char.head = 0
    
    If UserList(userIndex).flags.Muerto = 0 Then
        '(Nacho)
        If UserList(userIndex).Faccion.Alineacion = e_Alineacion.Real Then
            UserList(userIndex).Char.body = iFragataReal
        ElseIf UserList(userIndex).Faccion.Alineacion = e_Alineacion.Caos Then
            UserList(userIndex).Char.body = iFragataCaos
        Else
            UserList(userIndex).Char.body = Barco.Ropaje
        End If
    Else
        UserList(userIndex).Char.body = iFragataFantasmal
    End If
    
    UserList(userIndex).Char.ShieldAnim = NingunEscudo
    UserList(userIndex).Char.WeaponAnim = NingunArma
    UserList(userIndex).Char.CascoAnim = NingunCasco
    UserList(userIndex).flags.Navegando = 1
    
Else
    
    UserList(userIndex).flags.Navegando = 0
    
    If UserList(userIndex).flags.Muerto = 0 Then
        UserList(userIndex).Char.head = UserList(userIndex).OrigChar.head
        
        If UserList(userIndex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(userIndex).Char.body = ObjData(UserList(userIndex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(userIndex)
        End If
        
        If UserList(userIndex).Invent.EscudoEqpObjIndex > 0 Then _
            UserList(userIndex).Char.ShieldAnim = ObjData(UserList(userIndex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(userIndex).Invent.WeaponEqpObjIndex > 0 Then _
            UserList(userIndex).Char.WeaponAnim = ObjData(UserList(userIndex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(userIndex).Invent.CascoEqpObjIndex > 0 Then _
            UserList(userIndex).Char.CascoAnim = ObjData(UserList(userIndex).Invent.CascoEqpObjIndex).CascoAnim
    Else
        UserList(userIndex).Char.body = iCuerpoMuerto
        UserList(userIndex).Char.head = iCabezaMuerto
        UserList(userIndex).Char.ShieldAnim = NingunEscudo
        UserList(userIndex).Char.WeaponAnim = NingunArma
        UserList(userIndex).Char.CascoAnim = NingunCasco
    End If
End If

Call ChangeUserChar(userIndex, UserList(userIndex).Char.body, UserList(userIndex).Char.head, UserList(userIndex).Char.Heading, UserList(userIndex).Char.WeaponAnim, UserList(userIndex).Char.ShieldAnim, UserList(userIndex).Char.CascoAnim)
Call WriteNavigateToggle(userIndex)

End Sub

Public Sub FundirMineral(ByVal userIndex As Integer)
'Call LogTarea("Sub FundirMineral")

If UserList(userIndex).flags.TargetObjInvIndex > 0 Then
   
   If ObjData(UserList(userIndex).flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And ObjData(UserList(userIndex).flags.TargetObjInvIndex).MinSkill <= UserList(userIndex).Stats.UserSkills(eSkill.Mineria) / ModFundicion(UserList(userIndex).Clase) Then
        Call DoLingotes(userIndex)
   Else
        Call WriteConsoleMsg(userIndex, "No tenes conocimientos de mineria suficientes para trabajar este mineral.", FontTypeNames.FONTTYPE_INFO)
   End If

End If

End Sub
Function TieneObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal userIndex As Integer) As Boolean
'Call LogTarea("Sub TieneObjetos")

Dim i As Integer
Dim Total As Long
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(userIndex).Invent.Object(i).ObjIndex = ItemIndex Then
        Total = Total + UserList(userIndex).Invent.Object(i).amount
    End If
Next i

If cant <= Total Then
    TieneObjetos = True
    Exit Function
End If
        
End Function

Function QuitarObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal userIndex As Integer) As Boolean
'Call LogTarea("Sub QuitarObjetos")

Dim i As Integer
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(userIndex).Invent.Object(i).ObjIndex = ItemIndex Then
        
        Call Desequipar(userIndex, i)
        
        UserList(userIndex).Invent.Object(i).amount = UserList(userIndex).Invent.Object(i).amount - cant
        If (UserList(userIndex).Invent.Object(i).amount <= 0) Then
            cant = Abs(UserList(userIndex).Invent.Object(i).amount)
            UserList(userIndex).Invent.Object(i).amount = 0
            UserList(userIndex).Invent.Object(i).ObjIndex = 0
        Else
            cant = 0
        End If
        
        Call UpdateUserInv(False, userIndex, i)
        
        If (cant = 0) Then
            QuitarObjetos = True
            Exit Function
        End If
    End If
Next i

End Function

Sub HerreroQuitarMateriales(ByVal userIndex As Integer, ByVal ItemIndex As Integer)
    If ObjData(ItemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(ItemIndex).LingH, userIndex)
    If ObjData(ItemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(ItemIndex).LingP, userIndex)
    If ObjData(ItemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(ItemIndex).LingO, userIndex)
End Sub

Sub CarpinteroQuitarMateriales(ByVal userIndex As Integer, ByVal ItemIndex As Integer)
    If ObjData(ItemIndex).Madera > 0 Then Call QuitarObjetos(Leña, ObjData(ItemIndex).Madera, userIndex)
End Sub

Function CarpinteroTieneMateriales(ByVal userIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    
    If ObjData(ItemIndex).Madera > 0 Then
            If Not TieneObjetos(Leña, ObjData(ItemIndex).Madera, userIndex) Then
                    Call WriteConsoleMsg(userIndex, "No tenes suficientes madera.", FontTypeNames.FONTTYPE_INFO)
                    CarpinteroTieneMateriales = False
                    Exit Function
            End If
    End If
    
    CarpinteroTieneMateriales = True

End Function
 
Function HerreroTieneMateriales(ByVal userIndex As Integer, ByVal ItemIndex As Integer) As Boolean
    If ObjData(ItemIndex).LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, ObjData(ItemIndex).LingH, userIndex) Then
                    Call WriteConsoleMsg(userIndex, "No tenes suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingP > 0 Then
            If Not TieneObjetos(LingotePlata, ObjData(ItemIndex).LingP, userIndex) Then
                    Call WriteConsoleMsg(userIndex, "No tenes suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    If ObjData(ItemIndex).LingO > 0 Then
            If Not TieneObjetos(LingoteOro, ObjData(ItemIndex).LingO, userIndex) Then
                    Call WriteConsoleMsg(userIndex, "No tenes suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
                    HerreroTieneMateriales = False
                    Exit Function
            End If
    End If
    HerreroTieneMateriales = True
End Function

Public Function PuedeConstruir(ByVal userIndex As Integer, ByVal ItemIndex As Integer) As Boolean
PuedeConstruir = HerreroTieneMateriales(userIndex, ItemIndex) And UserList(userIndex).Stats.UserSkills(eSkill.Herreria) >= _
 ObjData(ItemIndex).SkHerreria
End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ArmasHerrero)
    If ArmasHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i
For i = 1 To UBound(ArmadurasHerrero)
    If ArmadurasHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i
PuedeConstruirHerreria = False
End Function


Public Sub HerreroConstruirItem(ByVal userIndex As Integer, ByVal ItemIndex As Integer)
'Call LogTarea("Sub HerreroConstruirItem")
If PuedeConstruir(userIndex, ItemIndex) And PuedeConstruirHerreria(ItemIndex) Then
    Call HerreroQuitarMateriales(userIndex, ItemIndex)
    ' AGREGAR FX
    If ObjData(ItemIndex).OBJType = eOBJType.otWeapon Then
        Call WriteConsoleMsg(userIndex, "Has construido el arma!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otESCUDO Then
        Call WriteConsoleMsg(userIndex, "Has construido el escudo!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otCASCO Then
        Call WriteConsoleMsg(userIndex, "Has construido el casco!.", FontTypeNames.FONTTYPE_INFO)
    ElseIf ObjData(ItemIndex).OBJType = eOBJType.otArmadura Then
        Call WriteConsoleMsg(userIndex, "Has construido la armadura!.", FontTypeNames.FONTTYPE_INFO)
    End If
    Dim MiObj As Obj
    MiObj.amount = 1
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(userIndex, MiObj) Then
                    Call TirarItemAlPiso(UserList(userIndex).pos, MiObj)
    End If
    Call SubirSkill(userIndex, Herreria)
    Call UpdateUserInv(True, userIndex, 0)
    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(MARTILLOHERRERO))
    
End If

UserList(userIndex).Counters.Trabajando = UserList(userIndex).Counters.Trabajando + 1

End Sub

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ObjCarpintero)
    If ObjCarpintero(i) = ItemIndex Then
        PuedeConstruirCarpintero = True
        Exit Function
    End If
Next i
PuedeConstruirCarpintero = False

End Function

Public Sub CarpinteroConstruirItem(ByVal userIndex As Integer, ByVal ItemIndex As Integer)

If CarpinteroTieneMateriales(userIndex, ItemIndex) And _
   UserList(userIndex).Stats.UserSkills(eSkill.Carpinteria) >= _
   ObjData(ItemIndex).SkCarpinteria And _
   PuedeConstruirCarpintero(ItemIndex) And _
   UserList(userIndex).Invent.WeaponEqpObjIndex = SERRUCHO_CARPINTERO Then

    Call CarpinteroQuitarMateriales(userIndex, ItemIndex)
    Call WriteConsoleMsg(userIndex, "Has construido el objeto!.", FontTypeNames.FONTTYPE_INFO)
    
    Dim MiObj As Obj
    MiObj.amount = 1
    MiObj.ObjIndex = ItemIndex
    If Not MeterItemEnInventario(userIndex, MiObj) Then
                    Call TirarItemAlPiso(UserList(userIndex).pos, MiObj)
    End If
    
    Call SubirSkill(userIndex, Carpinteria)
    Call UpdateUserInv(True, userIndex, 0)
    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(LABUROCARPINTERO))
End If

UserList(userIndex).Counters.Trabajando = UserList(userIndex).Counters.Trabajando + 1

End Sub

Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer
    Select Case Lingote
        Case iMinerales.HierroCrudo
            MineralesParaLingote = 14
        Case iMinerales.PlataCruda
            MineralesParaLingote = 20
        Case iMinerales.OroCrudo
            MineralesParaLingote = 35
        Case Else
            MineralesParaLingote = 10000
    End Select
End Function


Public Sub DoLingotes(ByVal userIndex As Integer)
'    Call LogTarea("Sub DoLingotes")
Dim Slot As Integer
Dim obji As Integer
Dim i As Integer 'Imposible que haga mas de 32k...

    Slot = UserList(userIndex).flags.TargetObjInvSlot
    obji = UserList(userIndex).Invent.Object(Slot).ObjIndex
    
    If UserList(userIndex).Invent.Object(Slot).amount < MineralesParaLingote(obji) Or _
        ObjData(obji).OBJType <> eOBJType.otMinerales Then
            Call WriteConsoleMsg(userIndex, "No tienes suficientes minerales para hacer un lingote.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
    End If
    
    If Not UserList(userIndex).Invent.Object(Slot).amount / MineralesParaLingote(obji) > 30000 Then
        i = UserList(userIndex).Invent.Object(Slot).amount / MineralesParaLingote(obji) 'cantidad de lingotes que hizo.
    Else
        LogError ("Mas de 30k de lingotes " & UserList(userIndex).name)
    End If
    
    'Le quitamos los minerales
    UserList(userIndex).Invent.Object(Slot).amount = UserList(userIndex).Invent.Object(Slot).amount - MineralesParaLingote(obji) * i
    
    If UserList(userIndex).Invent.Object(Slot).amount < 1 Then
        UserList(userIndex).Invent.Object(Slot).amount = 0
        UserList(userIndex).Invent.Object(Slot).ObjIndex = 0
    End If
    
    Dim nPos As WorldPos
    Dim MiObj As Obj
    MiObj.amount = i
    MiObj.ObjIndex = ObjData(UserList(userIndex).flags.TargetObjInvIndex).LingoteIndex
    If Not MeterItemEnInventario(userIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(userIndex).pos, MiObj)
    End If
    Call UpdateUserInv(False, userIndex, Slot)
    Call WriteConsoleMsg(userIndex, "¡Has obtenido un lingote!", FontTypeNames.FONTTYPE_INFO)

    UserList(userIndex).Counters.Trabajando = UserList(userIndex).Counters.Trabajando + 1
End Sub

Function ModNavegacion(ByVal Clase As eClass) As Integer

Select Case Clase
    Case eClass.Pirat
        ModNavegacion = 1
    Case eClass.Fisher
        ModNavegacion = 1.2
    Case Else
        ModNavegacion = 2.3
End Select

End Function


Function ModFundicion(ByVal Clase As eClass) As Integer

Select Case Clase
    Case eClass.Miner
        ModFundicion = 1
    Case eClass.Blacksmith
        ModFundicion = 1.2
    Case Else
        ModFundicion = 3
End Select

End Function

Function ModCarpinteria(ByVal Clase As eClass) As Integer

Select Case Clase
    Case eClass.Carpenter
        ModCarpinteria = 1
    Case Else
        ModCarpinteria = 3
End Select

End Function

Function ModHerreriA(ByVal Clase As eClass) As Integer

Select Case Clase
    Case eClass.Blacksmith
        ModHerreriA = 1
    Case eClass.Miner
        ModHerreriA = 1.2
    Case Else
        ModHerreriA = 4
End Select

End Function

Function ModDomar(ByVal Clase As eClass) As Integer
    Select Case Clase
        Case eClass.Druid
            ModDomar = 6
        Case eClass.Hunter
            ModDomar = 6
        Case eClass.Cleric
            ModDomar = 7
        Case Else
            ModDomar = 10
    End Select
End Function

Function CalcularPoderDomador(ByVal userIndex As Integer) As Long
    With UserList(userIndex).Stats
        CalcularPoderDomador = .UserAtributos(eAtributos.Carisma) _
            * (.UserSkills(eSkill.Domar) / ModDomar(UserList(userIndex).Clase)) _
            + RandomNumber(1, .UserAtributos(eAtributos.Carisma) / 3) _
            + RandomNumber(1, .UserAtributos(eAtributos.Carisma) / 3) _
            + RandomNumber(1, .UserAtributos(eAtributos.Carisma) / 3)
    End With
End Function

Function FreeMascotaIndex(ByVal userIndex As Integer) As Integer
    Dim j As Integer
    For j = 1 To MAXMASCOTAS
        If UserList(userIndex).MascotasIndex(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function
        End If
    Next j
End Function

Sub DoDomar(ByVal userIndex As Integer, ByVal NpcIndex As Integer)
'Call LogTarea("Sub DoDomar")

If UserList(userIndex).NroMacotas < MAXMASCOTAS Then
    
    If Npclist(NpcIndex).MaestroUser = userIndex Then
        Call WriteConsoleMsg(userIndex, "La criatura ya te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
        Call WriteConsoleMsg(userIndex, "La criatura ya tiene amo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Npclist(NpcIndex).flags.Domable <= CalcularPoderDomador(userIndex) Then
        Dim index As Integer
        UserList(userIndex).NroMacotas = UserList(userIndex).NroMacotas + 1
        index = FreeMascotaIndex(userIndex)
        UserList(userIndex).MascotasIndex(index) = NpcIndex
        UserList(userIndex).MascotasType(index) = Npclist(NpcIndex).Numero
        
        Npclist(NpcIndex).MaestroUser = userIndex
        
        Call FollowAmo(NpcIndex)
        
        Call WriteConsoleMsg(userIndex, "La criatura te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)
        Call SubirSkill(userIndex, Domar)
    Else
        If Not UserList(userIndex).flags.UltimoMensaje = 5 Then
            Call WriteConsoleMsg(userIndex, "No has logrado domar la criatura.", FontTypeNames.FONTTYPE_INFO)
            UserList(userIndex).flags.UltimoMensaje = 5
        End If
    End If
Else
    Call WriteConsoleMsg(userIndex, "No podes controlar mas criaturas.", FontTypeNames.FONTTYPE_INFO)
End If
End Sub

Sub DoAdminInvisible(ByVal userIndex As Integer)
    
    If UserList(userIndex).flags.AdminInvisible = 0 Then
        
        ' Sacamos el mimetizmo
        If UserList(userIndex).flags.Mimetizado = 1 Then
            UserList(userIndex).Char.body = UserList(userIndex).CharMimetizado.body
            UserList(userIndex).Char.head = UserList(userIndex).CharMimetizado.head
            UserList(userIndex).Char.CascoAnim = UserList(userIndex).CharMimetizado.CascoAnim
            UserList(userIndex).Char.ShieldAnim = UserList(userIndex).CharMimetizado.ShieldAnim
            UserList(userIndex).Char.WeaponAnim = UserList(userIndex).CharMimetizado.WeaponAnim
            UserList(userIndex).Counters.Mimetismo = 0
            UserList(userIndex).flags.Mimetizado = 0
        End If
        
        UserList(userIndex).flags.AdminInvisible = 1
        UserList(userIndex).flags.invisible = 1
        UserList(userIndex).flags.Oculto = 1
        UserList(userIndex).flags.OldBody = UserList(userIndex).Char.body
        UserList(userIndex).flags.OldHead = UserList(userIndex).Char.head
        UserList(userIndex).Char.body = 0
        UserList(userIndex).Char.head = 0
        
    Else
        
        UserList(userIndex).flags.AdminInvisible = 0
        UserList(userIndex).flags.invisible = 0
        UserList(userIndex).flags.Oculto = 0
        UserList(userIndex).Counters.TiempoOculto = 0
        UserList(userIndex).Char.body = UserList(userIndex).flags.OldBody
        UserList(userIndex).Char.head = UserList(userIndex).flags.OldHead
        
    End If
    
    'vuelve a ser visible por la fuerza
    Call ChangeUserChar(userIndex, UserList(userIndex).Char.body, UserList(userIndex).Char.head, UserList(userIndex).Char.Heading, UserList(userIndex).Char.WeaponAnim, UserList(userIndex).Char.ShieldAnim, UserList(userIndex).Char.CascoAnim)
    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageSetInvisible(UserList(userIndex).Char.CharIndex, False))
End Sub

Sub TratarDeHacerFogata(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal userIndex As Integer)

Dim Suerte As Byte
Dim exito As Byte
Dim raise As Byte
Dim Obj As Obj
Dim posMadera As WorldPos

If Not LegalPos(Map, X, Y) Then Exit Sub

With posMadera
    .Map = Map
    .X = X
    .Y = Y
End With

If MapData(Map, X, Y).ObjInfo.ObjIndex <> 58 Then
    Call WriteConsoleMsg(userIndex, "Necesitas clickear sobre Leña para hacer ramitas", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If Distancia(posMadera, UserList(userIndex).pos) > 2 Then
    Call WriteConsoleMsg(userIndex, "Estás demasiado lejos para prender la fogata.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(userIndex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(userIndex, "No puedes hacer fogatas estando muerto.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If MapData(Map, X, Y).ObjInfo.amount < 3 Then
    Call WriteConsoleMsg(userIndex, "Necesitas por lo menos tres troncos para hacer una fogata.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If


If UserList(userIndex).Stats.UserSkills(eSkill.supervivencia) >= 0 And UserList(userIndex).Stats.UserSkills(eSkill.supervivencia) < 6 Then
    Suerte = 3
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.supervivencia) >= 6 And UserList(userIndex).Stats.UserSkills(eSkill.supervivencia) <= 34 Then
    Suerte = 2
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.supervivencia) >= 35 Then
    Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    Obj.ObjIndex = FOGATA_APAG
    Obj.amount = MapData(Map, X, Y).ObjInfo.amount \ 3
    
    Call WriteConsoleMsg(userIndex, "Has hecho " & Obj.amount & " fogatas.", FontTypeNames.FONTTYPE_INFO)
    
    Call MakeObj(Map, Obj, Map, X, Y)
    
    'Seteamos la fogata como el nuevo TargetObj del user
    UserList(userIndex).flags.TargetObj = FOGATA_APAG
Else
    '[CDT 17-02-2004]
    If Not UserList(userIndex).flags.UltimoMensaje = 10 Then
        Call WriteConsoleMsg(userIndex, "No has podido hacer la fogata.", FontTypeNames.FONTTYPE_INFO)
        UserList(userIndex).flags.UltimoMensaje = 10
    End If
    '[/CDT]
End If

Call SubirSkill(userIndex, supervivencia)


End Sub

Public Sub DoPescar(ByVal userIndex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer


If UserList(userIndex).Clase = eClass.Fisher Then
    Call QuitarSta(userIndex, EsfuerzoPescarPescador)
Else
    Call QuitarSta(userIndex, EsfuerzoPescarGeneral)
End If

Dim Skill As Integer
Skill = UserList(userIndex).Stats.UserSkills(eSkill.Pesca)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
    If UserList(userIndex).Clase = eClass.Fisher Then
        MiObj.amount = RandomNumber(150, 300)
    Else
        MiObj.amount = 50
    End If
    MiObj.ObjIndex = Pescado
    
    If Not MeterItemEnInventario(userIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(userIndex).pos, MiObj)
    End If
    
    Call WriteConsoleMsg(userIndex, "¡Has pescado un lindo pez!", FontTypeNames.FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(userIndex).flags.UltimoMensaje = 6 Then
      Call WriteConsoleMsg(userIndex, "¡No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
      UserList(userIndex).flags.UltimoMensaje = 6
    End If
    '[/CDT]
End If

Call SubirSkill(userIndex, Pesca)

UserList(userIndex).Counters.Trabajando = UserList(userIndex).Counters.Trabajando + 1

Exit Sub

errhandler:
    Call LogError("Error en DoPescar")
End Sub

Public Sub DoPescarRed(ByVal userIndex As Integer)
On Error GoTo errhandler

Dim iSkill As Integer
Dim Suerte As Integer
Dim res As Integer
Dim EsPescador As Boolean

If UserList(userIndex).Clase = eClass.Fisher Then
    Call QuitarSta(userIndex, EsfuerzoPescarPescador)
    EsPescador = True
Else
    Call QuitarSta(userIndex, EsfuerzoPescarGeneral)
    EsPescador = False
End If

iSkill = UserList(userIndex).Stats.UserSkills(eSkill.Pesca)

' m = (60-11)/(1-10)
' y = mx - m*10 + 11

Suerte = Int(-0.00125 * iSkill * iSkill - 0.3 * iSkill + 49)

If Suerte > 0 Then
    res = RandomNumber(1, Suerte)
    
    If res < 6 Then
        Dim nPos As WorldPos
        Dim MiObj As Obj
        Dim PecesPosibles(1 To 4) As Integer
        
        PecesPosibles(1) = PESCADO1
        PecesPosibles(2) = PESCADO2
        PecesPosibles(3) = PESCADO3
        PecesPosibles(4) = PESCADO4
        
        If EsPescador = True Then
            MiObj.amount = RandomNumber(150, 300)
        Else
            MiObj.amount = 50
        End If
        MiObj.ObjIndex = PecesPosibles(RandomNumber(LBound(PecesPosibles), UBound(PecesPosibles)))
        
        If Not MeterItemEnInventario(userIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(userIndex).pos, MiObj)
        End If
        
        Call WriteConsoleMsg(userIndex, "¡Has pescado algunos peces!", FontTypeNames.FONTTYPE_INFO)
        
    Else
        Call WriteConsoleMsg(userIndex, "¡No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
    End If
    
    Call SubirSkill(userIndex, Pesca)
End If
        
Exit Sub

errhandler:
    Call LogError("Error en DoPescarRed")
End Sub

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal victimaindex As Integer)

If Not userCanAttackUser(LadrOnIndex, victimaindex) Then Exit Sub

If Not MapInfo(UserList(victimaindex).pos.Map).Pk Then Exit Sub

If TriggerZonaPelea(LadrOnIndex, victimaindex) <> TRIGGER6_AUSENTE Then Exit Sub

If (UserList(victimaindex).Faccion.Alineacion = UserList(LadrOnIndex).Faccion.Alineacion) And Not UserList(LadrOnIndex).Faccion.Alineacion = e_Alineacion.Neutro Then
    Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a miembros de tu faccion.", FontTypeNames.FONTTYPE_FIGHT)
    Exit Sub
End If

Call QuitarSta(LadrOnIndex, 15)

Dim GuantesHurto As Boolean
'Tiene los Guantes de Hurto equipados?
GuantesHurto = True
If UserList(LadrOnIndex).Invent.AnilloEqpObjIndex = 0 Then
    GuantesHurto = False
Else
    If ObjData(UserList(LadrOnIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMin <> 0 Then GuantesHurto = False
    If ObjData(UserList(LadrOnIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMax <> 0 Then GuantesHurto = False
End If


If UserList(victimaindex).flags.Privilegios And PlayerType.User Then
    Dim Suerte As Integer
    Dim res As Integer
    
    If UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 10 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= -1 Then
                        Suerte = 35
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 20 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 11 Then
                        Suerte = 30
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 30 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 21 Then
                        Suerte = 28
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 40 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 31 Then
                        Suerte = 24
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 50 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 41 Then
                        Suerte = 22
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 60 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 51 Then
                        Suerte = 20
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 70 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 61 Then
                        Suerte = 18
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 80 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 71 Then
                        Suerte = 15
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) <= 90 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 81 Then
                        Suerte = 10
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) < 100 _
       And UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) >= 91 Then
                        Suerte = 7
    ElseIf UserList(LadrOnIndex).Stats.UserSkills(eSkill.Robar) = 100 Then
                        Suerte = 5
    End If
    
    res = RandomNumber(1, Suerte)
        
    If res < 3 Then 'Exito robo
        If (RandomNumber(1, 50) < 25) And (UserList(LadrOnIndex).Clase = eClass.Thief) Then
            If TieneObjetosRobables(victimaindex) Then
                Call RobarObjeto(LadrOnIndex, victimaindex)
            Else
                Call WriteConsoleMsg(LadrOnIndex, UserList(victimaindex).name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else 'Roba oro
            If UserList(victimaindex).Stats.GLD > 0 Then
                Dim N As Integer
                
                If UserList(LadrOnIndex).Clase = eClass.Thief Then
                ' Si no tine puestos los guantes de hurto roba un 20% menos. Pablo (ToxicWaste)
                    If GuantesHurto Then
                        N = RandomNumber(100, 1000)
                    Else
                        N = RandomNumber(80, 800)
                    End If
                Else
                    N = RandomNumber(1, 100)
                End If
                If N > UserList(victimaindex).Stats.GLD Then N = UserList(victimaindex).Stats.GLD
                UserList(victimaindex).Stats.GLD = UserList(victimaindex).Stats.GLD - N
                
                UserList(LadrOnIndex).Stats.GLD = UserList(LadrOnIndex).Stats.GLD + N
                If UserList(LadrOnIndex).Stats.GLD > MAXORO Then _
                    UserList(LadrOnIndex).Stats.GLD = MAXORO
                
                Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & N & " monedas de oro a " & UserList(victimaindex).name, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(LadrOnIndex, UserList(victimaindex).name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)
            End If
        End If
    Else
        Call WriteConsoleMsg(LadrOnIndex, "¡No has logrado robar nada!", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(victimaindex, "¡" & UserList(LadrOnIndex).name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
        Call FlushBuffer(victimaindex)
    End If

    Call SubirSkill(LadrOnIndex, Robar)
End If


End Sub


Public Function ObjEsRobable(ByVal victimaindex As Integer, ByVal Slot As Integer) As Boolean
' Agregué los barcos
' Esta funcion determina qué objetos son robables.

Dim OI As Integer

OI = UserList(victimaindex).Invent.Object(Slot).ObjIndex

ObjEsRobable = _
ObjData(OI).OBJType <> eOBJType.otLlaves And _
UserList(victimaindex).Invent.Object(Slot).Equipped = 0 And _
ObjData(OI).Real = 0 And _
ObjData(OI).Caos = 0 And _
ObjData(OI).OBJType <> eOBJType.otBarcos

End Function

Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal victimaindex As Integer)
'Call LogTarea("Sub RobarObjeto")
Dim flag As Boolean
Dim i As Integer
flag = False

If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
    i = 1
    Do While Not flag And i <= MAX_INVENTORY_SLOTS
        'Hay objeto en este slot?
        If UserList(victimaindex).Invent.Object(i).ObjIndex > 0 Then
           If ObjEsRobable(victimaindex, i) Then
                 If RandomNumber(1, 10) < 4 Then flag = True
           End If
        End If
        If Not flag Then i = i + 1
    Loop
Else
    i = 20
    Do While Not flag And i > 0
      'Hay objeto en este slot?
      If UserList(victimaindex).Invent.Object(i).ObjIndex > 0 Then
         If ObjEsRobable(victimaindex, i) Then
               If RandomNumber(1, 10) < 4 Then flag = True
         End If
      End If
      If Not flag Then i = i - 1
    Loop
End If

If flag Then
    Dim MiObj As Obj
    Dim num As Byte
    'Cantidad al azar
    num = RandomNumber(1, 5)
                
    If num > UserList(victimaindex).Invent.Object(i).amount Then
         num = UserList(victimaindex).Invent.Object(i).amount
    End If
                
    MiObj.amount = num
    MiObj.ObjIndex = UserList(victimaindex).Invent.Object(i).ObjIndex
    
    UserList(victimaindex).Invent.Object(i).amount = UserList(victimaindex).Invent.Object(i).amount - num
                
    If UserList(victimaindex).Invent.Object(i).amount <= 0 Then
          Call QuitarUserInvItem(victimaindex, CByte(i), 1)
    End If
            
    Call UpdateUserInv(False, victimaindex, CByte(i))
                
    If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(LadrOnIndex).pos, MiObj)
    End If
    
    If UserList(LadrOnIndex).Clase = eClass.Thief Then
        Call WriteConsoleMsg(LadrOnIndex, "Has robado " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).name, FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(LadrOnIndex, "Has hurtado " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).name, FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar ningún objeto.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Public Sub DoApuñalar(ByVal userIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
'***************************************************
'Autor: Nacho (Integer) & Unknown (orginal version)
'Last Modification: 07/20/06
'***************************************************
Dim Suerte As Integer
Dim Skill As Integer

Skill = UserList(userIndex).Stats.UserSkills(eSkill.Apuñalar)

Select Case UserList(userIndex).Clase
    Case eClass.Assasin
        Suerte = Int((((0.0000003 * Skill - 0.00002) * Skill + 0.00098) * Skill + 0.0425) * 100)
    
    Case eClass.Cleric, eClass.Paladin
        Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0493) * 100)
    
    Case eClass.Bard
        Suerte = Int((((0.00000002 * Skill + 0.000002) * Skill + 0.00032) * Skill + 0.0481) * 100)
    
    Case Else
        Suerte = Int((0.000361 * Skill + 0.0439) * 100)
End Select


If RandomNumber(0, 100) < Suerte Then
    If VictimUserIndex <> 0 Then
        If UserList(userIndex).Clase = eClass.Assasin Then
            daño = Int(daño * 1.5)
        Else
            daño = Int(daño * 1.4)
        End If
        
        UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - daño
        Call WriteConsoleMsg(userIndex, "Has apuñalado a " & UserList(VictimUserIndex).name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(VictimUserIndex, "Te ha apuñalado " & UserList(userIndex).name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
    Else
        Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - Int(daño * 2)
        Call WriteConsoleMsg(userIndex, "Has apuñalado la criatura por " & Int(daño * 2), FontTypeNames.FONTTYPE_FIGHT)
        Call SubirSkill(userIndex, Apuñalar)
        '[Alejo]
        Call CalcularDarExp(userIndex, VictimNpcIndex, Int(daño * 2))
    End If
Else
    Call WriteConsoleMsg(userIndex, "¡No has logrado apuñalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)
End If

'Pablo (ToxicWaste): Revisar, saque este porque hacía que se me cuelgue.
'Call FlushBuffer(VictimUserIndex)
End Sub

Public Sub DoGolpeCritico(ByVal userIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 28/01/2007
'***************************************************
Dim Suerte As Integer
Dim Skill As Integer

If UserList(userIndex).Clase <> eClass.Bandit Then Exit Sub
If UserList(userIndex).Invent.WeaponEqpSlot = 0 Then Exit Sub
If ObjData(UserList(userIndex).Invent.WeaponEqpObjIndex).name <> "Espada Vikinga" Then Exit Sub


Skill = UserList(userIndex).Stats.UserSkills(eSkill.Wrestling)

Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0493) * 100)

If RandomNumber(0, 100) < Suerte Then
    daño = Int(daño * 0.5)
    If VictimUserIndex <> 0 Then
        UserList(VictimUserIndex).Stats.MinHP = UserList(VictimUserIndex).Stats.MinHP - daño
        Call WriteConsoleMsg(userIndex, "Has golpeado críticamente a " & UserList(VictimUserIndex).name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(VictimUserIndex, UserList(userIndex).name & " te ha golpeado críticamente por " & daño, FontTypeNames.FONTTYPE_FIGHT)
    Else
        Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - daño
        Call WriteConsoleMsg(userIndex, "Has golpeado críticamente a la criatura por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        '[Alejo]
        Call CalcularDarExp(userIndex, VictimNpcIndex, daño)
    End If
End If

End Sub

Public Sub QuitarSta(ByVal userIndex As Integer, ByVal Cantidad As Integer)
    UserList(userIndex).Stats.MinSta = UserList(userIndex).Stats.MinSta - Cantidad
    If UserList(userIndex).Stats.MinSta < 0 Then UserList(userIndex).Stats.MinSta = 0
    Call WriteUpdateSta(userIndex)
End Sub

Public Sub DoTalar(ByVal userIndex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer


If UserList(userIndex).Clase = eClass.Lumberjack Then
    Call QuitarSta(userIndex, EsfuerzoTalarLeñador)
Else
    Call QuitarSta(userIndex, EsfuerzoTalarGeneral)
End If

Dim Skill As Integer
Skill = UserList(userIndex).Stats.UserSkills(eSkill.Talar)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res < 6 Then
    Dim nPos As WorldPos
    Dim MiObj As Obj
    
    If UserList(userIndex).Clase = eClass.Lumberjack Then
        MiObj.amount = RandomNumber(4, 8)
    Else
        MiObj.amount = 4
    End If
    
    MiObj.ObjIndex = Leña
    
    
    If Not MeterItemEnInventario(userIndex, MiObj) Then
        
        Call TirarItemAlPiso(UserList(userIndex).pos, MiObj)
        
    End If
    
    Call WriteConsoleMsg(userIndex, "¡Has conseguido algo de leña!", FontTypeNames.FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(userIndex).flags.UltimoMensaje = 8 Then
        Call WriteConsoleMsg(userIndex, "¡No has obtenido leña!", FontTypeNames.FONTTYPE_INFO)
        UserList(userIndex).flags.UltimoMensaje = 8
    End If
    '[/CDT]
End If

Call SubirSkill(userIndex, Talar)

UserList(userIndex).Counters.Trabajando = UserList(userIndex).Counters.Trabajando + 1

Exit Sub

errhandler:
    Call LogError("Error en DoTalar")

End Sub

Public Sub DoMineria(ByVal userIndex As Integer)
On Error GoTo errhandler

Dim Suerte As Integer
Dim res As Integer
Dim metal As Integer

If UserList(userIndex).Clase = eClass.Miner Then
    Call QuitarSta(userIndex, EsfuerzoExcavarMinero)
Else
    Call QuitarSta(userIndex, EsfuerzoExcavarGeneral)
End If

Dim Skill As Integer
Skill = UserList(userIndex).Stats.UserSkills(eSkill.Mineria)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res <= 5 Then
    Dim MiObj As Obj
    Dim nPos As WorldPos
    
    If UserList(userIndex).flags.TargetObj = 0 Then Exit Sub
    
    MiObj.ObjIndex = ObjData(UserList(userIndex).flags.TargetObj).MineralIndex
    
    If UserList(userIndex).Clase = eClass.Miner Then
        MiObj.amount = RandomNumber(4, 8)
    Else
        MiObj.amount = 4
    End If
    
    If Not MeterItemEnInventario(userIndex, MiObj) Then _
        Call TirarItemAlPiso(UserList(userIndex).pos, MiObj)
    
    Call WriteConsoleMsg(userIndex, "¡Has extraido algunos minerales!", FontTypeNames.FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(userIndex).flags.UltimoMensaje = 9 Then
        Call WriteConsoleMsg(userIndex, "¡No has conseguido nada!", FontTypeNames.FONTTYPE_INFO)
        UserList(userIndex).flags.UltimoMensaje = 9
    End If
    '[/CDT]
End If

Call SubirSkill(userIndex, Mineria)

UserList(userIndex).Counters.Trabajando = UserList(userIndex).Counters.Trabajando + 1

Exit Sub

errhandler:
    Call LogError("Error en Sub DoMineria")

End Sub

Public Sub DoMeditar(ByVal userIndex As Integer)

UserList(userIndex).Counters.IdleCount = 0

Dim Suerte As Integer
Dim res As Integer
Dim cant As Integer

'Barrin 3/10/03
'Esperamos a que se termine de concentrar
Dim TActual As Long
TActual = GetTickCount() And &H7FFFFFFF
If TActual - UserList(userIndex).Counters.tInicioMeditar < TIEMPO_INICIOMEDITAR Then
    Exit Sub
End If

If UserList(userIndex).Counters.bPuedeMeditar = False Then
    UserList(userIndex).Counters.bPuedeMeditar = True
End If

If UserList(userIndex).Stats.MinMAN >= UserList(userIndex).Stats.MaxMAN Then
    Call WriteConsoleMsg(userIndex, "Has terminado de meditar.", FontTypeNames.FONTTYPE_INFO)
    Call WriteMeditateToggle(userIndex)
    UserList(userIndex).flags.Meditando = False
    UserList(userIndex).Char.FX = 0
    UserList(userIndex).Char.loops = 0
    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCreateFX(UserList(userIndex).Char.CharIndex, 0, 0))
    Exit Sub
End If

If UserList(userIndex).Stats.UserSkills(eSkill.Meditar) <= 10 _
   And UserList(userIndex).Stats.UserSkills(eSkill.Meditar) >= -1 Then
                    Suerte = 35
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Meditar) <= 20 _
   And UserList(userIndex).Stats.UserSkills(eSkill.Meditar) >= 11 Then
                    Suerte = 30
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Meditar) <= 30 _
   And UserList(userIndex).Stats.UserSkills(eSkill.Meditar) >= 21 Then
                    Suerte = 28
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Meditar) <= 40 _
   And UserList(userIndex).Stats.UserSkills(eSkill.Meditar) >= 31 Then
                    Suerte = 24
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Meditar) <= 50 _
   And UserList(userIndex).Stats.UserSkills(eSkill.Meditar) >= 41 Then
                    Suerte = 22
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Meditar) <= 60 _
   And UserList(userIndex).Stats.UserSkills(eSkill.Meditar) >= 51 Then
                    Suerte = 20
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Meditar) <= 70 _
   And UserList(userIndex).Stats.UserSkills(eSkill.Meditar) >= 61 Then
                    Suerte = 18
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Meditar) <= 80 _
   And UserList(userIndex).Stats.UserSkills(eSkill.Meditar) >= 71 Then
                    Suerte = 15
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Meditar) <= 90 _
   And UserList(userIndex).Stats.UserSkills(eSkill.Meditar) >= 81 Then
                    Suerte = 10
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Meditar) < 100 _
   And UserList(userIndex).Stats.UserSkills(eSkill.Meditar) >= 91 Then
                    Suerte = 7
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Meditar) = 100 Then
                    Suerte = 5
End If
res = RandomNumber(1, Suerte)

If res = 1 Then
    
    cant = Porcentaje(UserList(userIndex).Stats.MaxMAN, PorcentajeRecuperoMana)
    If cant <= 0 Then cant = 1
    UserList(userIndex).Stats.MinMAN = UserList(userIndex).Stats.MinMAN + cant
    If UserList(userIndex).Stats.MinMAN > UserList(userIndex).Stats.MaxMAN Then _
        UserList(userIndex).Stats.MinMAN = UserList(userIndex).Stats.MaxMAN
    
    If Not UserList(userIndex).flags.UltimoMensaje = 22 Then
        Call WriteConsoleMsg(userIndex, "¡Has recuperado " & cant & " puntos de mana!", FontTypeNames.FONTTYPE_INFO)
        UserList(userIndex).flags.UltimoMensaje = 22
    End If
    
    Call WriteUpdateMana(userIndex)
    Call SubirSkill(userIndex, Meditar)
End If

End Sub

Public Sub DoHurtar(ByVal userIndex As Integer, ByVal victimaindex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modif: 28/01/2007
'Implements the pick pocket skill of the Bandit :)
'***************************************************
If UserList(userIndex).Clase <> eClass.Bandit Then Exit Sub
'Esto es precario y feo, pero por ahora no se me ocurrió nada mejor.
'Uso el slot de los anillos para "equipar" los guantes.
'Y los reconozco porque les puse DefensaMagicaMin y Max = 0
If UserList(userIndex).Invent.AnilloEqpObjIndex = 0 Then
    Exit Sub
Else
    If ObjData(UserList(userIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMin <> 0 Then Exit Sub
    If ObjData(UserList(userIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMax <> 0 Then Exit Sub
End If

Dim res As Integer
res = RandomNumber(1, 100)
If (res < 20) Then
    If TieneObjetosRobables(victimaindex) Then
        Call RobarObjeto(userIndex, victimaindex)
        Call WriteConsoleMsg(victimaindex, "¡" & UserList(userIndex).name & " es un Bandido!", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(userIndex, UserList(victimaindex).name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
    End If
End If

End Sub

Public Sub DoHandInmo(ByVal userIndex As Integer, ByVal victimaindex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modif: 17/02/2007
'Implements the special Skill of the Thief
'***************************************************
If UserList(victimaindex).flags.Paralizado = 1 Then Exit Sub
If UserList(userIndex).Clase <> eClass.Thief Then Exit Sub
    
'una vez más, la forma de reconocer los guantes es medio patética.
If UserList(userIndex).Invent.AnilloEqpObjIndex = 0 Then
    Exit Sub
Else
    If ObjData(UserList(userIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMin <> 0 Then Exit Sub
    If ObjData(UserList(userIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMax <> 0 Then Exit Sub
End If

    
Dim res As Integer
res = RandomNumber(0, 100)
If res < (UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) / 4) Then
    UserList(victimaindex).flags.Paralizado = 1
    UserList(victimaindex).Counters.Paralisis = IntervaloParalizado / 2
    Call WriteParalizeOK(victimaindex)
    Call WriteConsoleMsg(userIndex, "Tu golpe ha dejado inmóvil a tu oponente", FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(victimaindex, "¡El golpe te ha dejado inmóvil!", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Public Sub Desarmar(ByVal userIndex As Integer, ByVal VictimIndex As Integer)

Dim Suerte As Integer
Dim res As Integer

If UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) <= 10 _
   And UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) >= -1 Then
                    Suerte = 35
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) <= 20 _
   And UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) >= 11 Then
                    Suerte = 30
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) <= 30 _
   And UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) >= 21 Then
                    Suerte = 28
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) <= 40 _
   And UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) >= 31 Then
                    Suerte = 24
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) <= 50 _
   And UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) >= 41 Then
                    Suerte = 22
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) <= 60 _
   And UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) >= 51 Then
                    Suerte = 20
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) <= 70 _
   And UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) >= 61 Then
                    Suerte = 18
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) <= 80 _
   And UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) >= 71 Then
                    Suerte = 15
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) <= 90 _
   And UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) >= 81 Then
                    Suerte = 10
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) < 100 _
   And UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) >= 91 Then
                    Suerte = 7
ElseIf UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) = 100 Then
                    Suerte = 5
End If
res = RandomNumber(1, Suerte)

If res <= 2 Then
        Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot)
        Call WriteConsoleMsg(userIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
        If UserList(VictimIndex).Stats.ELV < 20 Then
            Call WriteConsoleMsg(VictimIndex, "Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)
        End If
        Call FlushBuffer(VictimIndex)
    End If
End Sub

