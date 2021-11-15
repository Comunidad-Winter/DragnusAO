Attribute VB_Name = "modGuilds"
'**************************************************************
' modGuilds.bas - Module to allow the usage of areas instead of maps.
' Saves a lot of bandwidth.
'
' Implemented by Mariano Barrou (El Oso)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

'guilds nueva version. Hecho por el oso, eliminando los problemas
'de sincronizacion con los datos en el HD... entre varios otros
'º¬

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DECLARACIOENS PUBLICAS CONCERNIENTES AL JUEGO
'Y CONFIGURACION DEL SISTEMA DE CLANES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private GUILDINFOFILE   As String
'archivo .\guilds\guildinfo.ini o similar

Private Const MAX_GUILDS As Integer = 1000
'cantidad maxima de guilds en el servidor

Private Const ORDENARLISTADECLANES = True
'True si se envia la lista ordenada por alineacion

Public CANTIDADDECLANES As Integer
'cantidad actual de clanes en el servidor

Private guilds(1 To MAX_GUILDS) As clsClan
'array global de guilds, se indexa por userlist().guildindex

Private Const CANTIDADMAXIMACODEX As Byte = 8
'cantidad maxima de codecs que se pueden definir

Public Const MAXASPIRANTES As Byte = 10
'cantidad maxima de aspirantes que puede tener un clan acumulados a la vez

Private Const MAXANTIFACCION As Byte = 5
'puntos maximos de antifaccion que un clan tolera antes de ser cambiada su alineacion

Private GMsEscuchando As New Collection

Public Enum ALINEACION_GUILD
    ALINEACION_ARMADA = 0
    ALINEACION_LEGION = 1
    ALINEACION_NEUTRO = 2
    ALINEACION_MASTER = 3
End Enum
'alineaciones permitidas

Public Enum SONIDOS_GUILD
    SND_CREACIONCLAN = 44
    SND_ACEPTADOCLAN = 43
    SND_DECLAREWAR = 45
End Enum
'numero de .wav del cliente

Public Enum RELACIONES_GUILD
    GUERRA = -1
    PAZ = 0
    ALIADOS = 1
End Enum
'estado entre clanes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub LoadGuildsDB()

Dim CantClanes  As String
Dim i           As Integer
Dim TempStr     As String
Dim Alin        As ALINEACION_GUILD
    
    GUILDINFOFILE = App.Path & "\guilds\guildsinfo.inf"

    CantClanes = GetVar(GUILDINFOFILE, "INIT", "nroGuilds")
    
    If IsNumeric(CantClanes) Then
        CANTIDADDECLANES = CInt(CantClanes)
    Else
        CANTIDADDECLANES = 0
    End If
    
    For i = 1 To CANTIDADDECLANES
        Set guilds(i) = New clsClan
        TempStr = GetVar(GUILDINFOFILE, "GUILD" & i, "GUILDNAME")
        Alin = String2Alineacion(GetVar(GUILDINFOFILE, "GUILD" & i, "Alineacion"))
        Call guilds(i).Inicializar(TempStr, i, Alin)
    Next i
    
End Sub

Public Function m_ConectarMiembroAClan(ByVal userIndex As Integer, ByVal GuildIndex As Integer) As Boolean


    If GuildIndex > CANTIDADDECLANES Or GuildIndex <= 0 Then Exit Function 'x las dudas...
    If m_EstadoPermiteEntrar(userIndex, GuildIndex) Then
        Call guilds(GuildIndex).ConectarMiembro(userIndex)
        UserList(userIndex).GuildIndex = GuildIndex
        m_ConectarMiembroAClan = True
    Else
        m_ConectarMiembroAClan = m_ValidarPermanencia(userIndex)
    End If

End Function


Public Function m_ValidarPermanencia(ByVal userIndex As Integer) As Boolean
Dim CambioAlineacion As Boolean
Dim GuildIndex  As Integer
Dim i           As Integer
Dim list()      As String

    m_ValidarPermanencia = True
    GuildIndex = UserList(userIndex).GuildIndex
    If GuildIndex > CANTIDADDECLANES And GuildIndex <= 0 Then Exit Function
    
    Call LogClanes(UserList(userIndex).name & " de " & guilds(GuildIndex).GuildName & " es expulsado en validar permanencia")

    m_ValidarPermanencia = False
    
    Call LogClanes(UserList(userIndex).name & " de " & guilds(GuildIndex).GuildName & IIf(CambioAlineacion, " SI ", " NO ") & "provoca cambio de alinaecion. MAXANT:" & (guilds(GuildIndex).PuntosAntifaccion = MAXANTIFACCION) & ", GUILDFOU:" & m_EsGuildFounder(UserList(userIndex).name, GuildIndex))
    
    CambioAlineacion = m_EsGuildFounder(UserList(userIndex).name, GuildIndex)
    
    If CambioAlineacion Then
        'Si te saliste de alguna faccion, no pasa nada, el tema es cuando entras a otra.
        If UserList(userIndex).Faccion.Alineacion = e_Alineacion.Caos Then
            Call guilds(GuildIndex).CambiarAlineacion(ALINEACION_GUILD.ALINEACION_LEGION)
        ElseIf UserList(userIndex).Faccion.Alineacion = e_Alineacion.Real Then
            Call guilds(GuildIndex).CambiarAlineacion(ALINEACION_GUILD.ALINEACION_ARMADA)
        Else
            Call guilds(GuildIndex).CambiarAlineacion(ALINEACION_GUILD.ALINEACION_NEUTRO)
        End If
        
        guilds(GuildIndex).SetGuildNews "Clan cambio alineacion: " & Alineacion2String(guilds(GuildIndex).Alineacion)
        
        list = guilds(GuildIndex).GetMemberList
        For i = 0 To UBound(list)
            If NameIndex(list(i)) Then
                If Not m_EstadoPermiteEntrar(NameIndex(list(i)), GuildIndex) Then
                    If list(i) = guilds(GuildIndex).GetLeader Then
                        guilds(GuildIndex).SetLeader (guilds(GuildIndex).Fundador)
                    End If
                    Call m_EcharMiembroDeClan(-1, list(i))
                End If
            Else
                If Not m_EstadoPermiteEntrarChar(list(i), GuildIndex) Then
                    If list(i) = guilds(GuildIndex).GetLeader Then
                        guilds(GuildIndex).SetLeader (guilds(GuildIndex).Fundador)
                    End If
                    Call m_EcharMiembroDeClan(-1, list(i))
                End If
            End If
        Next i
    Else
        If Not m_EstadoPermiteEntrar(userIndex, GuildIndex) Then
            If m_EsGuildLeader(UserList(userIndex).name, GuildIndex) Then
                guilds(GuildIndex).SetLeader guilds(GuildIndex).Fundador
            End If
            Call m_EcharMiembroDeClan(-1, UserList(userIndex).name)   'y lo echamos
        End If
    End If

End Function

Public Sub m_DesconectarMiembroDelClan(ByVal userIndex As Integer, ByVal GuildIndex As Integer)
    If UserList(userIndex).GuildIndex > CANTIDADDECLANES Then Exit Sub
    Call guilds(GuildIndex).DesConectarMiembro(userIndex)
End Sub

Private Function m_EsGuildLeader(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
    m_EsGuildLeader = (UCase$(PJ) = UCase$(Trim$(guilds(GuildIndex).GetLeader)))
End Function

Private Function m_EsGuildFounder(ByRef PJ As String, ByVal GuildIndex As Integer) As Boolean
    m_EsGuildFounder = (UCase$(PJ) = UCase$(Trim$(guilds(GuildIndex).Fundador)))
End Function

Public Function m_EcharMiembroDeClan(ByVal Expulsador As Integer, ByVal Expulsado As String) As Integer
'UI echa a Expulsado del clan de Expulsado
Dim userIndex   As Integer
Dim GI          As Integer
    
    m_EcharMiembroDeClan = 0

    userIndex = NameIndex(Expulsado)
    If userIndex > 0 Then
        'pj online
        GI = UserList(userIndex).GuildIndex
        If GI > 0 Then
            If m_PuedeSalirDeClan(Expulsado, GI, Expulsador) Then
                If m_EsGuildLeader(Expulsado, GI) Then guilds(GI).SetLeader (guilds(GI).Fundador)
                Call guilds(GI).DesConectarMiembro(userIndex)
                Call guilds(GI).ExpulsarMiembro(Expulsado)
                Call LogClanes(Expulsado & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)
                UserList(userIndex).GuildIndex = 0
                Call RefreshCharStatus(userIndex)
                m_EcharMiembroDeClan = GI
            Else
                m_EcharMiembroDeClan = 0
            End If
        Else
            m_EcharMiembroDeClan = 0
        End If
    Else
        'pj offline
        GI = GetGuildIndexFromChar(Expulsado)
        If GI > 0 Then
            If m_PuedeSalirDeClan(Expulsado, GI, Expulsador) Then
                If m_EsGuildLeader(Expulsado, GI) Then guilds(GI).SetLeader (guilds(GI).Fundador)
                Call guilds(GI).ExpulsarMiembro(Expulsado)
                Call LogClanes(Expulsado & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)
                m_EcharMiembroDeClan = GI
            Else
                m_EcharMiembroDeClan = 0
            End If
        Else
            m_EcharMiembroDeClan = 0
        End If
    End If

End Function

Public Sub ActualizarWebSite(ByVal userIndex As Integer, ByRef Web As String)
Dim GI As Integer

    GI = UserList(userIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
    
    If Not m_EsGuildLeader(UserList(userIndex).name, GI) Then Exit Sub
    
    Call guilds(GI).SetURL(Web)
    
End Sub


Public Sub ChangeCodexAndDesc(ByRef Desc As String, ByRef codex() As String, ByVal GuildIndex As Integer)
    Dim i As Long
    
    If GuildIndex < 1 Or GuildIndex > CANTIDADDECLANES Then Exit Sub
    
    With guilds(GuildIndex)
        Call .SetDesc(Desc)
        
        For i = 0 To UBound(codex())
            Call .SetCodex(i, codex(i))
        Next i
        
        For i = i To CANTIDADMAXIMACODEX
            Call .SetCodex(i, vbNullString)
        Next i
    End With
End Sub

Public Sub ActualizarNoticias(ByVal userIndex As Integer, ByRef Datos As String)
Dim GI              As Integer

    GI = UserList(userIndex).GuildIndex
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then Exit Sub
    
    If Not m_EsGuildLeader(UserList(userIndex).name, GI) Then Exit Sub
    
    Call guilds(GI).SetGuildNews(Datos)
        
End Sub

Public Function CrearNuevoClan(ByVal FundadorIndex As Integer, ByRef Desc As String, ByRef GuildName As String, ByRef URL As String, ByRef codex() As String, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean
Dim CantCodex       As Integer
Dim i               As Integer
Dim DummyString     As String

    CrearNuevoClan = False
    If Not PuedeFundarUnClan(FundadorIndex, Alineacion, DummyString) Then
        refError = DummyString
        Exit Function
    End If

    If GuildName = vbNullString Or Not GuildNameValido(GuildName) Then
        refError = "Nombre de clan inválido."
        Exit Function
    End If
    
    If YaExiste(GuildName) Then
        refError = "Ya existe un clan con ese nombre."
        Exit Function
    End If

    CantCodex = UBound(codex()) + 1

    'tenemos todo para fundar ya
    If CANTIDADDECLANES < UBound(guilds) Then
        CANTIDADDECLANES = CANTIDADDECLANES + 1
        'ReDim Preserve Guilds(1 To CANTIDADDECLANES) As clsClan

        'constructor custom de la Clase clan
        Set guilds(CANTIDADDECLANES) = New clsClan
        Call guilds(CANTIDADDECLANES).Inicializar(GuildName, CANTIDADDECLANES, Alineacion)
        
        'Damos de alta al clan como nuevo inicializando sus archivos
        Call guilds(CANTIDADDECLANES).InicializarNuevoClan(UserList(FundadorIndex).name)
        
        'seteamos codex y descripcion
        For i = 1 To CantCodex
            Call guilds(CANTIDADDECLANES).SetCodex(i, codex(i - 1))
        Next i
        Call guilds(CANTIDADDECLANES).SetDesc(Desc)
        Call guilds(CANTIDADDECLANES).SetGuildNews("Clan creado con alineación : " & Alineacion2String(Alineacion))
        Call guilds(CANTIDADDECLANES).SetLeader(UserList(FundadorIndex).name)
        Call guilds(CANTIDADDECLANES).SetURL(URL)
        
        '"conectamos" al nuevo miembro a la lista de la Clase
        Call guilds(CANTIDADDECLANES).AceptarNuevoMiembro(UserList(FundadorIndex).name)
        Call guilds(CANTIDADDECLANES).ConectarMiembro(FundadorIndex)
        UserList(FundadorIndex).GuildIndex = CANTIDADDECLANES
        Call RefreshCharStatus(FundadorIndex)
        
        For i = 1 To CANTIDADDECLANES - 1
            Call guilds(i).ProcesarFundacionDeOtroClan
        Next i
    Else
        refError = "No hay mas slots para fundar clanes. Consulte a un administrador."
        Exit Function
    End If
    
    CrearNuevoClan = True
End Function
Public Sub SendGuildNews(ByVal userIndex As Integer)
Dim GuildIndex  As Integer
Dim enemiesCount    As Integer
Dim i               As Integer
Dim go As Integer

    GuildIndex = UserList(userIndex).GuildIndex
    If GuildIndex = 0 Then Exit Sub

    Dim enemies() As String
    
    If guilds(GuildIndex).CantidadEnemys Then
        ReDim enemies(0 To guilds(GuildIndex).CantidadEnemys - 1) As String
    Else
        ReDim enemies(0)
    End If
    
    Dim allies() As String
    
    If guilds(GuildIndex).CantidadAllies Then
        ReDim allies(0 To guilds(GuildIndex).CantidadAllies - 1) As String
    Else
        ReDim allies(0)
    End If
    
    i = guilds(GuildIndex).Iterador_ProximaRelacion(RELACIONES_GUILD.GUERRA)
    go = 0
    
    While i > 0
        enemies(go) = guilds(i).GuildName
        i = guilds(GuildIndex).Iterador_ProximaRelacion(RELACIONES_GUILD.GUERRA)
        go = go + 1
    Wend
    
    i = guilds(GuildIndex).Iterador_ProximaRelacion(RELACIONES_GUILD.ALIADOS)
    go = 0
    
    While i > 0
        enemies(go) = guilds(i).GuildName
        i = guilds(GuildIndex).Iterador_ProximaRelacion(RELACIONES_GUILD.ALIADOS)
    Wend

    Call WriteGuildNews(userIndex, guilds(GuildIndex).GetGuildNews, enemies, allies)

    If guilds(GuildIndex).EleccionesAbiertas Then
        Call WriteConsoleMsg(userIndex, "Hoy es la votacion para elegir un nuevo líder para el clan!!.", FontTypeNames.FONTTYPE_GUILD)
        Call WriteConsoleMsg(userIndex, "La eleccion durara 24 horas, se puede votar a cualquier miembro del clan.", FontTypeNames.FONTTYPE_GUILD)
        Call WriteConsoleMsg(userIndex, "Para votar escribe /VOTO NICKNAME.", FontTypeNames.FONTTYPE_GUILD)
        Call WriteConsoleMsg(userIndex, "Solo se computara un voto por miembro. Tu voto no puede ser cambiado.", FontTypeNames.FONTTYPE_GUILD)
    End If

End Sub

Public Function m_PuedeSalirDeClan(ByRef Nombre As String, ByVal GuildIndex As Integer, ByVal QuienLoEchaUI As Integer) As Boolean
'sale solo si no es fundador del clan.

    m_PuedeSalirDeClan = False
    If GuildIndex = 0 Then Exit Function
    
    'esto es un parche, si viene en -1 es porque la invoca la rutina de expulsion automatica de clanes x antifacciones
    If QuienLoEchaUI = -1 Then
        m_PuedeSalirDeClan = True
        Exit Function
    End If

    'cuando UI no puede echar a nombre?
    'si no es gm Y no es lider del clan del pj Y no es el mismo que se va voluntariamente
    If UserList(QuienLoEchaUI).flags.Privilegios And PlayerType.User Then
        If Not m_EsGuildLeader(UCase$(UserList(QuienLoEchaUI).name), GuildIndex) Then
            If UCase$(UserList(QuienLoEchaUI).name) <> UCase$(Nombre) Then      'si no sale voluntariamente...
                Exit Function
            End If
        End If
    End If

    m_PuedeSalirDeClan = UCase$(guilds(GuildIndex).Fundador) <> UCase$(Nombre)

End Function

Public Function PuedeFundarUnClan(ByVal userIndex As Integer, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean

    PuedeFundarUnClan = False
    If UserList(userIndex).GuildIndex > 0 Then
        refError = "Ya perteneces a un clan, no puedes fundar otro"
        Exit Function
    End If
    
    If UserList(userIndex).Stats.ELV < 25 Or UserList(userIndex).Stats.UserSkills(eSkill.Liderazgo) < 90 Then
        refError = "Para fundar un clan debes ser nivel 25 y tener 90 en liderazgo."
        Exit Function
    End If
    
    Select Case Alineacion
        Case ALINEACION_GUILD.ALINEACION_ARMADA
            If UserList(userIndex).Faccion.Alineacion <> e_Alineacion.Real And UserList(userIndex).Faccion.Alineacion <> e_Alineacion.Neutro Then
                refError = "Para fundar un clan real debes ser miembro de la armada."
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_LEGION
            If UserList(userIndex).Faccion.Alineacion <> e_Alineacion.Caos And UserList(userIndex).Faccion.Alineacion <> e_Alineacion.Neutro Then
                refError = "Para fundar un clan del mal debes pertenecer a la legión oscura"
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_MASTER
            If UserList(userIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
                refError = "Para fundar un clan sin alineación debes ser un dios."
                Exit Function
            End If
    End Select
    
    PuedeFundarUnClan = True
    
End Function

Private Function m_EstadoPermiteEntrarChar(ByRef Personaje As String, ByVal GuildIndex As Integer) As Boolean
Dim ELV         As Integer
Dim f           As Byte

    m_EstadoPermiteEntrarChar = False
    
    If InStrB(Personaje, "\") <> 0 Then
        Personaje = Replace(Personaje, "\", vbNullString)
    End If
    If InStrB(Personaje, "/") <> 0 Then
        Personaje = Replace(Personaje, "/", vbNullString)
    End If
    If InStrB(Personaje, ".") <> 0 Then
        Personaje = Replace(Personaje, ".", vbNullString)
    End If
    
    If FileExist(CharPath & Personaje & ".chr") Then
        Select Case guilds(GuildIndex).Alineacion
            Case ALINEACION_GUILD.ALINEACION_ARMADA
                ELV = CInt(GetVar(CharPath & Personaje & ".chr", "Stats", "ELV"))
                If ELV >= 25 Then
                    f = CByte(GetVar(CharPath & Personaje & ".chr", "Facciones", "Alineacion"))
                End If
                m_EstadoPermiteEntrarChar = ELV >= 25 And (f = e_Alineacion.Real)
            Case ALINEACION_GUILD.ALINEACION_LEGION
                ELV = CInt(GetVar(CharPath & Personaje & ".chr", "Stats", "ELV"))
                If ELV >= 25 Then
                    f = CByte(GetVar(CharPath & Personaje & ".chr", "Facciones", "Alineacion"))
                End If
                m_EstadoPermiteEntrarChar = ELV >= 25 And (f = e_Alineacion.Caos)
            Case Else
                m_EstadoPermiteEntrarChar = True
        End Select
    End If
End Function

Private Function m_EstadoPermiteEntrar(ByVal userIndex As Integer, ByVal GuildIndex As Integer) As Boolean
    Select Case guilds(GuildIndex).Alineacion
        Case ALINEACION_GUILD.ALINEACION_ARMADA
            m_EstadoPermiteEntrar = UserList(userIndex).Stats.ELV >= 25 And UserList(userIndex).Faccion.Alineacion = e_Alineacion.Real
        Case ALINEACION_GUILD.ALINEACION_LEGION
            m_EstadoPermiteEntrar = UserList(userIndex).Stats.ELV >= 25 And UserList(userIndex).Faccion.Alineacion = e_Alineacion.Caos
        Case Else   'game masters
            m_EstadoPermiteEntrar = True
    End Select
End Function


Public Function String2Alineacion(ByRef S As String) As ALINEACION_GUILD
    Select Case S
        Case "Legión oscura"
            String2Alineacion = ALINEACION_LEGION
        Case "Armada Real"
            String2Alineacion = ALINEACION_ARMADA
        Case "Game Masters"
            String2Alineacion = ALINEACION_MASTER
        Case "Neutro"
            String2Alineacion = ALINEACION_NEUTRO
    End Select
End Function

Public Function Alineacion2String(ByVal Alineacion As ALINEACION_GUILD) As String
    Select Case Alineacion
        Case ALINEACION_GUILD.ALINEACION_LEGION
            Alineacion2String = "Legión oscura"
        Case ALINEACION_GUILD.ALINEACION_ARMADA
            Alineacion2String = "Armada Real"
        Case ALINEACION_GUILD.ALINEACION_MASTER
            Alineacion2String = "Game Masters"
        Case ALINEACION_GUILD.ALINEACION_NEUTRO
            Alineacion2String = "Neutro"
    End Select
End Function

Public Function Relacion2String(ByVal Relacion As RELACIONES_GUILD) As String
    Select Case Relacion
        Case RELACIONES_GUILD.ALIADOS
            Relacion2String = "A"
        Case RELACIONES_GUILD.GUERRA
            Relacion2String = "G"
        Case RELACIONES_GUILD.PAZ
            Relacion2String = "P"
        Case RELACIONES_GUILD.ALIADOS
            Relacion2String = "?"
    End Select
End Function

Public Function String2Relacion(ByVal S As String) As RELACIONES_GUILD
    Select Case UCase$(Trim$(S))
        Case vbNullString, "P"
            String2Relacion = RELACIONES_GUILD.PAZ
        Case "G"
            String2Relacion = RELACIONES_GUILD.GUERRA
        Case "A"
            String2Relacion = RELACIONES_GUILD.ALIADOS
        Case Else
            String2Relacion = RELACIONES_GUILD.PAZ
    End Select
End Function

Private Function GuildNameValido(ByVal cad As String) As Boolean
Dim car     As Byte
Dim i       As Integer

'old function by morgo

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))

    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        GuildNameValido = False
        Exit Function
    End If
    
Next i

GuildNameValido = True

End Function

Private Function YaExiste(ByVal GuildName As String) As Boolean
Dim i   As Integer

YaExiste = False
GuildName = UCase$(GuildName)

For i = 1 To CANTIDADDECLANES
    YaExiste = (UCase$(guilds(i).GuildName) = GuildName)
    If YaExiste Then Exit Function
Next i



End Function

Public Function v_AbrirElecciones(ByVal userIndex As Integer, ByRef refError As String) As Boolean
Dim GuildIndex      As Integer

    v_AbrirElecciones = False
    GuildIndex = UserList(userIndex).GuildIndex
    
    If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
        refError = "Tu no perteneces a ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(userIndex).name, GuildIndex) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If guilds(GuildIndex).EleccionesAbiertas Then
        refError = "Las elecciones ya están abiertas"
        Exit Function
    End If
    
    v_AbrirElecciones = True
    Call guilds(GuildIndex).AbrirElecciones
    
End Function

Public Function v_UsuarioVota(ByVal userIndex As Integer, ByRef Votado As String, ByRef refError As String) As Boolean
Dim GuildIndex      As Integer
Dim list()          As String
Dim i As Long

    v_UsuarioVota = False
    GuildIndex = UserList(userIndex).GuildIndex
    
    If GuildIndex = 0 Or GuildIndex > CANTIDADDECLANES Then
        refError = "Tu no perteneces a ningún clan"
        Exit Function
    End If

    If Not guilds(GuildIndex).EleccionesAbiertas Then
        refError = "No hay elecciones abiertas en tu clan."
        Exit Function
    End If
    
    
    list = guilds(GuildIndex).GetMemberList()
    For i = 0 To UBound(list())
        If UCase$(Votado) = UCase$(list(i)) Then Exit For
    Next i
    
    If i > UBound(list()) Then
        refError = Votado & " no pertenece al clan"
        Exit Function
    End If
    
    
    If guilds(GuildIndex).YaVoto(UserList(userIndex).name) Then
        refError = "Ya has votado, no puedes cambiar tu voto"
        Exit Function
    End If
    
    Call guilds(GuildIndex).ContabilizarVoto(UserList(userIndex).name, Votado)
    v_UsuarioVota = True

End Function

Public Sub v_RutinaElecciones()
Dim i       As Integer

On Error GoTo errh
    Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Servidor> Revisando elecciones", FontTypeNames.FONTTYPE_SERVER))
    For i = 1 To CANTIDADDECLANES
        If Not guilds(i) Is Nothing Then
            If guilds(i).RevisarElecciones Then
                Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Servidor> " & guilds(i).GetLeader & " es el nuevo lider de " & guilds(i).GuildName & "!", FontTypeNames.FONTTYPE_SERVER))
            End If
        End If
proximo:
    Next i
    Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Servidor> Elecciones revisadas", FontTypeNames.FONTTYPE_SERVER))
Exit Sub
errh:
    Call LogError("modGuilds.v_RutinaElecciones():" & Err.description)
    Resume proximo
End Sub

Private Function GetGuildIndexFromChar(ByRef PlayerName As String) As Integer
'aca si que vamos a violar las capas deliveradamente ya que
'visual basic no permite declarar metodos de Clase
Dim i       As Integer
Dim Temps   As String
    If InStrB(PlayerName, "\") <> 0 Then
        PlayerName = Replace(PlayerName, "\", vbNullString)
    End If
    If InStrB(PlayerName, "/") <> 0 Then
        PlayerName = Replace(PlayerName, "/", vbNullString)
    End If
    If InStrB(PlayerName, ".") <> 0 Then
        PlayerName = Replace(PlayerName, ".", vbNullString)
    End If
    Temps = GetVar(CharPath & PlayerName & ".chr", "GUILD", "GUILDINDEX")
    If IsNumeric(Temps) Then
        GetGuildIndexFromChar = CInt(Temps)
    Else
        GetGuildIndexFromChar = 0
    End If
End Function

Public Function GuildIndex(ByRef GuildName As String) As Integer
'me da el indice del guildname
Dim i As Integer

    GuildIndex = 0
    GuildName = UCase$(GuildName)
    For i = 1 To CANTIDADDECLANES
        If UCase$(guilds(i).GuildName) = GuildName Then
            GuildIndex = i
            Exit Function
        End If
    Next i
End Function

Public Function m_ListaDeMiembrosOnline(ByVal userIndex As Integer, ByVal GuildIndex As Integer) As String
Dim i As Integer
    
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
        While i > 0
            'No mostramos dioses y admins
            If i <> userIndex And ((UserList(i).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or (UserList(userIndex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0)) Then _
                m_ListaDeMiembrosOnline = m_ListaDeMiembrosOnline & UserList(i).name & ","
            i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
        Wend
    End If
    If Len(m_ListaDeMiembrosOnline) > 0 Then
        m_ListaDeMiembrosOnline = Left$(m_ListaDeMiembrosOnline, Len(m_ListaDeMiembrosOnline) - 1)
    End If
End Function

Public Function m_ListaDeMiembros(ByVal userIndex As Integer, ByVal GuildIndex As Integer) As String
    Dim i As Integer
    Dim mList() As String
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        mList = guilds(GuildIndex).GetMemberList
        For i = 0 To UBound(mList)
            m_ListaDeMiembros = m_ListaDeMiembros & mList(i) & ","
        Next i

    End If
    If Len(m_ListaDeMiembros) > 0 Then
        m_ListaDeMiembros = Left$(m_ListaDeMiembros, Len(m_ListaDeMiembros) - 1)
    End If
End Function

Public Function PrepareGuildsList() As String()
    Dim tStr() As String
    Dim i As Long
    
    If CANTIDADDECLANES = 0 Then
        ReDim tStr(0) As String
    Else
        ReDim tStr(CANTIDADDECLANES - 1) As String
        
        For i = 1 To CANTIDADDECLANES
            tStr(i - 1) = guilds(i).GuildName
        Next i
    End If
    
    PrepareGuildsList = tStr
End Function

Public Sub SendGuildDetails(ByVal userIndex As Integer, ByRef GuildName As String)
    Dim codex(CANTIDADMAXIMACODEX - 1)  As String
    Dim GI      As Integer
    Dim i       As Long

    GI = GuildIndex(GuildName)
    If GI = 0 Then Exit Sub
    
    With guilds(GI)
        For i = 1 To CANTIDADMAXIMACODEX
            codex(i - 1) = .GetCodex(i)
        Next i
        
        Call Protocol.WriteGuildDetails(userIndex, GuildName, .Fundador, .GetFechaFundacion, .GetLeader, _
                                    .GetURL, .CantidadDeMiembros, .EleccionesAbiertas, Alineacion2String(.Alineacion), _
                                    .CantidadEnemys, .CantidadAllies, _
                                    codex, .GetDesc, .Reputacion)
    End With
End Sub

Public Sub SendGuildLeaderInfo(ByVal userIndex As Integer)
'***************************************************
'Autor: Mariano Barrou (El Oso)
'Last Modification: 12/10/06
'Las Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'***************************************************
    Dim GI      As Integer
    Dim guildList() As String
    Dim MemberList() As String
    Dim aspirantsList() As String

    With UserList(userIndex)
        GI = .GuildIndex
        
        guildList = PrepareGuildsList()
        
        If GI <= 0 Or GI > CANTIDADDECLANES Then
            'Send the guild list instead
            Call Protocol.WriteGuildList(userIndex, guildList)
            Exit Sub
        End If
        
        If Not m_EsGuildLeader(.name, GI) Then
            'Send the guild list instead
            Call Protocol.WriteGuildList(userIndex, guildList)
            Exit Sub
        End If
        
        MemberList = guilds(GI).GetMemberList()
        aspirantsList = guilds(GI).GetAspirantes()
        
        Call WriteGuildLeaderInfo(userIndex, guildList, MemberList, guilds(GI).GetGuildNews(), aspirantsList, modGuilds.r_ListaDePropuestas(userIndex, RELACIONES_GUILD.PAZ))
    End With
End Sub


Public Function m_Iterador_ProximoUserIndex(ByVal GuildIndex As Integer) As Integer
    'itera sobre los onlinemembers
    m_Iterador_ProximoUserIndex = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        m_Iterador_ProximoUserIndex = guilds(GuildIndex).m_Iterador_ProximoUserIndex()
    End If
End Function

Public Function Iterador_ProximoGM(ByVal GuildIndex As Integer) As Integer
    'itera sobre los gms escuchando este clan
    Iterador_ProximoGM = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        Iterador_ProximoGM = guilds(GuildIndex).Iterador_ProximoGM()
    End If
End Function

Public Function r_Iterador_ProximaPropuesta(ByVal GuildIndex As Integer, ByVal Tipo As RELACIONES_GUILD) As Integer
    'itera sobre las propuestas
    r_Iterador_ProximaPropuesta = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        r_Iterador_ProximaPropuesta = guilds(GuildIndex).Iterador_ProximaPropuesta(Tipo)
    End If
End Function

Public Function GMEscuchaClan(ByVal userIndex As Integer, ByVal GuildName As String) As Integer
Dim GI As Integer

    'listen to no guild at all
    If LenB(GuildName) = 0 And UserList(userIndex).EscucheClan <> 0 Then
        'Quit listening to previous guild!!
        Call WriteConsoleMsg(userIndex, "Dejas de escuchar a : " & guilds(UserList(userIndex).EscucheClan).GuildName, FontTypeNames.FONTTYPE_GUILD)
        guilds(UserList(userIndex).EscucheClan).DesconectarGM (userIndex)
        Exit Function
    End If
    
'devuelve el guildindex
    GI = GuildIndex(GuildName)
    If GI > 0 Then
        If UserList(userIndex).EscucheClan <> 0 Then
            If UserList(userIndex).EscucheClan = GI Then
                'Already listening to them...
                Call WriteConsoleMsg(userIndex, "Conectado a : " & GuildName, FontTypeNames.FONTTYPE_GUILD)
                GMEscuchaClan = GI
                Exit Function
            Else
                'Quit listening to previous guild!!
                Call WriteConsoleMsg(userIndex, "Dejas de escuchar a : " & guilds(UserList(userIndex).EscucheClan).GuildName, FontTypeNames.FONTTYPE_GUILD)
                guilds(UserList(userIndex).EscucheClan).DesconectarGM (userIndex)
            End If
        End If
        
        Call guilds(GI).ConectarGM(userIndex)
        Call WriteConsoleMsg(userIndex, "Conectado a : " & GuildName, FontTypeNames.FONTTYPE_GUILD)
        GMEscuchaClan = GI
        UserList(userIndex).EscucheClan = GI
    Else
        Call WriteConsoleMsg(userIndex, "Error, el clan no existe", FontTypeNames.FONTTYPE_GUILD)
        GMEscuchaClan = 0
    End If
    
End Function

Public Sub GMDejaDeEscucharClan(ByVal userIndex As Integer, ByVal GuildIndex As Integer)
'el index lo tengo que tener de cuando me puse a escuchar
    UserList(userIndex).EscucheClan = 0
    Call guilds(GuildIndex).DesconectarGM(userIndex)
End Sub
Public Function r_DeclararGuerra(ByVal userIndex As Integer, ByRef GuildGuerra As String, ByRef refError As String) As Integer
Dim GI  As Integer
Dim GIG As Integer

    r_DeclararGuerra = 0
    GI = UserList(userIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(userIndex).name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildGuerra) = vbNullString Then
        refError = "No has seleccionado ningún clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildGuerra)
    
    If GI = GIG Then
        refError = "No puedes declarar la guerra a tu mismo clan"
        Exit Function
    End If

    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_DeclararGuerra: " & GI & " declara a " & GuildGuerra)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If

    Call guilds(GI).AnularPropuestas(GIG)
    Call guilds(GIG).AnularPropuestas(GI)
    Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.GUERRA)
    Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.GUERRA)

    r_DeclararGuerra = GIG

End Function


Public Function r_AceptarPropuestaDePaz(ByVal userIndex As Integer, ByRef GuildPaz As String, ByRef refError As String) As Integer
'el clan de UserIndex acepta la propuesta de paz de guildpaz, con quien esta en guerra
Dim GI      As Integer
Dim GIG     As Integer

    GI = UserList(userIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(userIndex).name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildPaz) = vbNullString Then
        refError = "No has seleccionado ningún clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildPaz)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_AceptarPropuestaDePaz: " & GI & " acepta de " & GuildPaz)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If

    If guilds(GI).GetRelacion(GIG) <> RELACIONES_GUILD.GUERRA Then
        refError = "No estás en guerra con ese clan"
        Exit Function
    End If
    
    If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.PAZ) Then
        refError = "No hay ninguna propuesta de paz para aceptar"
        Exit Function
    End If

    Call guilds(GI).AnularPropuestas(GIG)
    Call guilds(GIG).AnularPropuestas(GI)
    Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.PAZ)
    Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.PAZ)
    
    r_AceptarPropuestaDePaz = GIG
End Function

Public Function r_RechazarPropuestaDeAlianza(ByVal userIndex As Integer, ByRef GuildPro As String, ByRef refError As String) As Integer
'devuelve el index al clan guildPro
Dim GI      As Integer
Dim GIG     As Integer

    r_RechazarPropuestaDeAlianza = 0
    GI = UserList(userIndex).GuildIndex
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(userIndex).name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildPro) = vbNullString Then
        refError = "No has seleccionado ningún clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildPro)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_RechazarPropuestaDeAlianza: " & GI & " acepta de " & GuildPro)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If
    
    If Not guilds(GI).HayPropuesta(GIG, ALIADOS) Then
        refError = "No hay propuesta de alianza del clan " & GuildPro
        Exit Function
    End If
    
    Call guilds(GI).AnularPropuestas(GIG)
    'avisamos al otro clan
    Call guilds(GIG).SetGuildNews(guilds(GI).GuildName & " ha rechazado nuestra propuesta de alianza. " & guilds(GIG).GetGuildNews())
    r_RechazarPropuestaDeAlianza = GIG

End Function


Public Function r_RechazarPropuestaDePaz(ByVal userIndex As Integer, ByRef GuildPro As String, ByRef refError As String) As Integer
'devuelve el index al clan guildPro
Dim GI      As Integer
Dim GIG     As Integer

    r_RechazarPropuestaDePaz = 0
    GI = UserList(userIndex).GuildIndex
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(userIndex).name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildPro) = vbNullString Then
        refError = "No has seleccionado ningún clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildPro)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_RechazarPropuestaDePaz: " & GI & " acepta de " & GuildPro)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If
    
    If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.PAZ) Then
        refError = "No hay propuesta de paz del clan " & GuildPro
        Exit Function
    End If
    
    Call guilds(GI).AnularPropuestas(GIG)
    'avisamos al otro clan
    Call guilds(GIG).SetGuildNews(guilds(GI).GuildName & " ha rechazado nuestra propuesta de paz. " & guilds(GIG).GetGuildNews())
    r_RechazarPropuestaDePaz = GIG

End Function


Public Function r_AceptarPropuestaDeAlianza(ByVal userIndex As Integer, ByRef GuildAllie As String, ByRef refError As String) As Integer
'el clan de UserIndex acepta la propuesta de paz de guildpaz, con quien esta en guerra
Dim GI      As Integer
Dim GIG     As Integer

    r_AceptarPropuestaDeAlianza = 0
    GI = UserList(userIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(userIndex).name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildAllie) = vbNullString Then
        refError = "No has seleccionado ningún clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildAllie)
    
    If GIG < 1 Or GIG > CANTIDADDECLANES Then
        Call LogError("ModGuilds.r_AceptarPropuestaDeAlianza: " & GI & " acepta de " & GuildAllie)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If

    If guilds(GI).GetRelacion(GIG) <> RELACIONES_GUILD.PAZ Then
        refError = "No estás en paz con el clan, solo puedes aceptar propuesas de alianzas con alguien que estes en paz."
        Exit Function
    End If
    
    If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.ALIADOS) Then
        refError = "No hay ninguna propuesta de alianza para aceptar."
        Exit Function
    End If

    Call guilds(GI).AnularPropuestas(GIG)
    Call guilds(GIG).AnularPropuestas(GI)
    Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.ALIADOS)
    Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.ALIADOS)
    
    r_AceptarPropuestaDeAlianza = GIG

End Function


Public Function r_ClanGeneraPropuesta(ByVal userIndex As Integer, ByRef OtroClan As String, ByVal Tipo As RELACIONES_GUILD, ByRef Detalle As String, ByRef refError As String) As Boolean
Dim OtroClanGI      As Integer
Dim GI              As Integer

    r_ClanGeneraPropuesta = False
    
    GI = UserList(userIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    OtroClanGI = GuildIndex(OtroClan)
    
    If OtroClanGI = GI Then
        refError = "No puedes declarar relaciones con tu propio clan"
        Exit Function
    End If
    
    If OtroClanGI <= 0 Or OtroClanGI > CANTIDADDECLANES Then
        refError = "El sistema de clanes esta inconsistente, el otro clan no existe!"
        Exit Function
    End If
    
    If guilds(OtroClanGI).HayPropuesta(GI, Tipo) Then
        refError = "Ya hay propuesta de " & Relacion2String(Tipo) & " con " & OtroClan
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(userIndex).name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    'de acuerdo al tipo procedemos validando las transiciones
    If Tipo = RELACIONES_GUILD.PAZ Then
        If guilds(GI).GetRelacion(OtroClanGI) <> RELACIONES_GUILD.GUERRA Then
            refError = "No estás en guerra con " & OtroClan
            Exit Function
        End If
    ElseIf Tipo = RELACIONES_GUILD.GUERRA Then
        'por ahora no hay propuestas de guerra
    ElseIf Tipo = RELACIONES_GUILD.ALIADOS Then
        If guilds(GI).GetRelacion(OtroClanGI) <> RELACIONES_GUILD.PAZ Then
            refError = "Para solicitar alianza no debes estar ni aliado ni en guerra con " & OtroClan
            Exit Function
        End If
    End If
    
    Call guilds(OtroClanGI).SetPropuesta(Tipo, GI, Detalle)
    r_ClanGeneraPropuesta = True

End Function

Public Function r_VerPropuesta(ByVal userIndex As Integer, ByRef OtroGuild As String, ByVal Tipo As RELACIONES_GUILD, ByRef refError As String) As String
Dim OtroClanGI      As Integer
Dim GI              As Integer
    
    r_VerPropuesta = vbNullString
    refError = vbNullString
    
    GI = UserList(userIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(userIndex).name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    OtroClanGI = GuildIndex(OtroGuild)
    
    If Not guilds(GI).HayPropuesta(OtroClanGI, Tipo) Then
        refError = "No existe la propuesta solicitada"
        Exit Function
    End If
    
    r_VerPropuesta = guilds(GI).GetPropuesta(OtroClanGI, Tipo)
    
End Function

Public Function r_ListaDePropuestas(ByVal userIndex As Integer, ByVal Tipo As RELACIONES_GUILD) As String()

    Dim GI  As Integer
    Dim i   As Integer
    Dim proposalCount As Integer
    Dim proposals() As String
    
    GI = UserList(userIndex).GuildIndex
    
    If GI > 0 And GI <= CANTIDADDECLANES Then
        With guilds(GI)
            proposalCount = .CantidadPropuestas(Tipo)
            
            'Resize array to contain all proposals
            If proposalCount > 0 Then
                ReDim proposals(proposalCount - 1) As String
            Else
                ReDim proposals(0) As String
            End If
            
            'Store each guild name
            For i = 0 To proposalCount - 1
                proposals(i) = guilds(.Iterador_ProximaPropuesta(Tipo)).GuildName
            Next i
        End With
    End If
    
    r_ListaDePropuestas = proposals
End Function

Public Sub a_RechazarAspiranteChar(ByRef Aspirante As String, ByVal guild As Integer, ByRef Detalles As String)
    If InStrB(Aspirante, "\") <> 0 Then
        Aspirante = Replace(Aspirante, "\", "")
    End If
    If InStrB(Aspirante, "/") <> 0 Then
        Aspirante = Replace(Aspirante, "/", "")
    End If
    If InStrB(Aspirante, ".") <> 0 Then
        Aspirante = Replace(Aspirante, ".", "")
    End If
    Call guilds(guild).InformarRechazoEnChar(Aspirante, Detalles)
End Sub

Public Function a_ObtenerRechazoDeChar(ByRef Aspirante As String) As String
    If InStrB(Aspirante, "\") <> 0 Then
        Aspirante = Replace(Aspirante, "\", "")
    End If
    If InStrB(Aspirante, "/") <> 0 Then
        Aspirante = Replace(Aspirante, "/", "")
    End If
    If InStrB(Aspirante, ".") <> 0 Then
        Aspirante = Replace(Aspirante, ".", "")
    End If
    a_ObtenerRechazoDeChar = GetVar(CharPath & Aspirante & ".chr", "GUILD", "MotivoRechazo")
    Call WriteVar(CharPath & Aspirante & ".chr", "GUILD", "MotivoRechazo", vbNullString)
End Function

Public Function a_RechazarAspirante(ByVal userIndex As Integer, ByRef Nombre As String, ByRef motivo As String, ByRef refError As String) As Boolean
'CHECK: El parámetro motivo, no se utiliza ¬¬
Dim GI              As Integer
Dim UI              As Integer
Dim NroAspirante    As Integer

    a_RechazarAspirante = False
    GI = UserList(userIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No perteneces a ningún clan"
        Exit Function
    End If

    NroAspirante = guilds(GI).NumeroDeAspirante(Nombre)

    If NroAspirante = 0 Then
        refError = Nombre & " no es aspirante a tu clan"
        Exit Function
    End If

    Call guilds(GI).RetirarAspirante(Nombre, NroAspirante)
    refError = "Fue rechazada tu solicitud de ingreso a " & guilds(GI).GuildName
    a_RechazarAspirante = True

End Function

Public Function a_DetallesAspirante(ByVal userIndex As Integer, ByRef Nombre As String) As String
Dim GI              As Integer
Dim NroAspirante    As Integer

    GI = UserList(userIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(userIndex).name, GI) Then
        Exit Function
    End If
    
    NroAspirante = guilds(GI).NumeroDeAspirante(Nombre)
    If NroAspirante > 0 Then
        a_DetallesAspirante = guilds(GI).DetallesSolicitudAspirante(NroAspirante)
    End If
    
End Function

Public Sub SendDetallesPersonaje(ByVal userIndex As Integer, ByRef Personaje As String)
    Dim GI          As Integer
    Dim NroAsp      As Integer
    Dim GuildName   As String
    Dim UserFile    As clsIniReader
    Dim Miembro     As String
    Dim GuildActual As Integer
    Dim list()      As String
    Dim i           As Long
    
    GI = UserList(userIndex).GuildIndex
    
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        Call Protocol.WriteConsoleMsg(userIndex, "No perteneces a ningún clan", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Not m_EsGuildLeader(UserList(userIndex).name, GI) Then
        Call Protocol.WriteConsoleMsg(userIndex, "No eres el líder de tu clan", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If InStrB(Personaje, "\") <> 0 Then
        Personaje = Replace$(Personaje, "\", vbNullString)
    End If
    If InStrB(Personaje, "/") <> 0 Then
        Personaje = Replace$(Personaje, "/", vbNullString)
    End If
    If InStrB(Personaje, ".") <> 0 Then
        Personaje = Replace$(Personaje, ".", vbNullString)
    End If
    
    NroAsp = guilds(GI).NumeroDeAspirante(Personaje)
    
    If NroAsp = 0 Then
        list = guilds(GI).GetMemberList()
        
        For i = 0 To UBound(list())
            If Personaje = list(i) Then Exit For
        Next i
        
        If i > UBound(list()) Then
            Call Protocol.WriteConsoleMsg(userIndex, "El personaje no es ni aspirante ni miembro del clan", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    'ahora traemos la info
    
    Set UserFile = New clsIniReader
    
    With UserFile
        .Initialize (CharPath & Personaje & ".chr")
        
        ' Get the character's current guild
        GuildActual = val(.GetValue("GUILD", "GuildIndex"))
        If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
            GuildName = "<" & guilds(GuildActual).GuildName & ">"
        Else
            GuildName = "Ninguno"
        End If
        
        'Get previous guilds
        Miembro = .GetValue("GUILD", "Miembro")
        If Len(Miembro) > 400 Then
            Miembro = ".." & Right$(Miembro, 400)
        End If
        
        Call Protocol.WriteCharacterInfo(userIndex, Personaje, .GetValue("INIT", "Raza"), .GetValue("INIT", "Clase"), _
                                .GetValue("INIT", "Genero"), .GetValue("STATS", "ELV"), .GetValue("STATS", "GLD"), _
                                .GetValue("STATS", "Banco"), .GetValue("REP", "Promedio"), .GetValue("GUILD", "Pedidos"), _
                                GuildName, Miembro, .GetValue("FACCIONES", "EjercitoReal"), .GetValue("FACCIONES", "EjercitoCaos"), _
                                .GetValue("FACCIONES", "CiudMatados"), .GetValue("FACCIONES", "CrimMatados"))
    End With
    
    Set UserFile = Nothing
End Sub

Public Function a_NuevoAspirante(ByVal userIndex As Integer, ByRef clan As String, ByRef Solicitud As String, ByRef refError As String) As Boolean
Dim ViejoSolicitado     As String
Dim ViejoGuildINdex     As Integer
Dim ViejoNroAspirante   As Integer
Dim NuevoGuildIndex     As Integer

    a_NuevoAspirante = False

    If UserList(userIndex).GuildIndex > 0 Then
        refError = "Ya perteneces a un clan, debes salir del mismo antes de solicitar ingresar a otro"
        Exit Function
    End If
    
    If EsNewbie(userIndex) Then
        refError = "Los newbies no tienen derecho a entrar a un clan."
        Exit Function
    End If

    NuevoGuildIndex = GuildIndex(clan)
    If NuevoGuildIndex = 0 Then
        refError = "Ese clan no existe! Avise a un administrador."
        Exit Function
    End If
    
    If Not m_EstadoPermiteEntrar(userIndex, NuevoGuildIndex) Then
        refError = "Tu no puedes entrar a un clan de alineación " & Alineacion2String(guilds(NuevoGuildIndex).Alineacion)
        Exit Function
    End If

    If guilds(NuevoGuildIndex).CantidadAspirantes >= MAXASPIRANTES Then
        refError = "El clan tiene demasiados aspirantes. Contáctate con un miembro para que procese las solicitudes."
        Exit Function
    End If

    ViejoSolicitado = GetVar(CharPath & UserList(userIndex).name & ".chr", "GUILD", "ASPIRANTEA")

    If LenB(ViejoSolicitado) <> 0 Then
        'borramos la vieja solicitud
        ViejoGuildINdex = CInt(ViejoSolicitado)
        If ViejoGuildINdex <> 0 Then
            ViejoNroAspirante = guilds(ViejoGuildINdex).NumeroDeAspirante(UserList(userIndex).name)
            If ViejoNroAspirante > 0 Then
                Call guilds(ViejoGuildINdex).RetirarAspirante(UserList(userIndex).name, ViejoNroAspirante)
            End If
        Else
            'RefError = "Inconsistencia en los clanes, avise a un administrador"
            'Exit Function
        End If
    End If
    
    Call guilds(NuevoGuildIndex).NuevoAspirante(UserList(userIndex).name, Solicitud)
    a_NuevoAspirante = True
End Function

Public Function a_AceptarAspirante(ByVal userIndex As Integer, ByRef Aspirante As String, ByRef refError As String) As Boolean
Dim GI              As Integer
Dim NroAspirante    As Integer
Dim AspiranteUI     As Integer

    'un pj ingresa al clan :D

    a_AceptarAspirante = False
    
    GI = UserList(userIndex).GuildIndex
    If GI <= 0 Or GI > CANTIDADDECLANES Then
        refError = "No perteneces a ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(userIndex).name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    NroAspirante = guilds(GI).NumeroDeAspirante(Aspirante)
    
    If NroAspirante = 0 Then
        refError = "El Pj no es aspirante al clan"
        Exit Function
    End If
    
    AspiranteUI = NameIndex(Aspirante)
    If AspiranteUI > 0 Then
        'pj Online
        If Not m_EstadoPermiteEntrar(AspiranteUI, GI) Then
            refError = Aspirante & " no puede entrar a un clan " & Alineacion2String(guilds(GI).Alineacion)
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        ElseIf Not UserList(AspiranteUI).GuildIndex = 0 Then
            refError = Aspirante & " ya es parte de otro clan."
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        End If
    Else
        If Not m_EstadoPermiteEntrarChar(Aspirante, GI) Then
            refError = Aspirante & " no puede entrar a un clan " & Alineacion2String(guilds(GI).Alineacion)
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        ElseIf GetGuildIndexFromChar(Aspirante) Then
            refError = Aspirante & " ya es parte de otro clan."
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        End If
    End If
    'el pj es aspirante al clan y puede entrar
    
    Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
    Call guilds(GI).AceptarNuevoMiembro(Aspirante)
    
    ' If player is online, update tag
    If AspiranteUI > 0 Then
        Call RefreshCharStatus(AspiranteUI)
    End If
    
    a_AceptarAspirante = True
End Function

Public Function GuildName(ByVal GuildIndex As Integer) As String
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then _
        Exit Function
    
    GuildName = guilds(GuildIndex).GuildName
End Function

Public Function GuildLeader(ByVal GuildIndex As Integer) As String
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then _
        Exit Function
    
    GuildLeader = guilds(GuildIndex).GetLeader
End Function

Public Function GuildAlignment(ByVal GuildIndex As Integer) As String
    If GuildIndex <= 0 Or GuildIndex > CANTIDADDECLANES Then _
        Exit Function
    
    GuildAlignment = Alineacion2String(guilds(GuildIndex).Alineacion)
End Function





