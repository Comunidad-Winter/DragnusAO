Attribute VB_Name = "TCP"
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

#If UsarQueSocket = 0 Then
' General constants used with most of the controls
Public Const INVALID_HANDLE As Integer = -1
Public Const CONTROL_ERRIGNORE As Integer = 0
Public Const CONTROL_ERRDISPLAY As Integer = 1


' SocietWrench Control Actions
Public Const SOCKET_OPEN As Integer = 1
Public Const SOCKET_CONNECT As Integer = 2
Public Const SOCKET_LISTEN As Integer = 3
Public Const SOCKET_ACCEPT As Integer = 4
Public Const SOCKET_CANCEL As Integer = 5
Public Const SOCKET_FLUSH As Integer = 6
Public Const SOCKET_CLOSE As Integer = 7
Public Const SOCKET_DISCONNECT As Integer = 7
Public Const SOCKET_ABORT As Integer = 8

' SocketWrench Control States
Public Const SOCKET_NONE As Integer = 0
Public Const SOCKET_IDLE As Integer = 1
Public Const SOCKET_LISTENING As Integer = 2
Public Const SOCKET_CONNECTING As Integer = 3
Public Const SOCKET_ACCEPTING As Integer = 4
Public Const SOCKET_RECEIVING As Integer = 5
Public Const SOCKET_SENDING As Integer = 6
Public Const SOCKET_CLOSING As Integer = 7

' Societ Address Families
Public Const AF_UNSPEC As Integer = 0
Public Const AF_UNIX As Integer = 1
Public Const AF_INET As Integer = 2

' Societ Types
Public Const SOCK_STREAM As Integer = 1
Public Const SOCK_DGRAM As Integer = 2
Public Const SOCK_RAW As Integer = 3
Public Const SOCK_RDM As Integer = 4
Public Const SOCK_SEQPACKET As Integer = 5

' Protocol Types
Public Const IPPROTO_IP As Integer = 0
Public Const IPPROTO_ICMP As Integer = 1
Public Const IPPROTO_GGP As Integer = 2
Public Const IPPROTO_TCP As Integer = 6
Public Const IPPROTO_PUP As Integer = 12
Public Const IPPROTO_UDP As Integer = 17
Public Const IPPROTO_IDP As Integer = 22
Public Const IPPROTO_ND As Integer = 77
Public Const IPPROTO_RAW As Integer = 255
Public Const IPPROTO_MAX As Integer = 256


' Network Addpesses
Public Const INADDR_ANY As String = "0.0.0.0"
Public Const INADDR_LOOPBACK As String = "127.0.0.1"
Public Const INADDR_NONE As String = "255.055.255.255"

' Shutdown Values
Public Const SOCKET_READ As Integer = 0
Public Const SOCKET_WRITE As Integer = 1
Public Const SOCKET_READWRITE As Integer = 2

' SocketWrench Error Pesponse
Public Const SOCKET_ERRIGNORE As Integer = 0
Public Const SOCKET_ERRDISPLAY As Integer = 1

' SocketWrench Error Codes
Public Const WSABASEERR As Integer = 24000
Public Const WSAEINTR As Integer = 24004
Public Const WSAEBADF As Integer = 24009
Public Const WSAEACCES As Integer = 24013
Public Const WSAEFAULT As Integer = 24014
Public Const WSAEINVAL As Integer = 24022
Public Const WSAEMFILE As Integer = 24024
Public Const WSAEWOULDBLOCK As Integer = 24035
Public Const WSAEINPROGRESS As Integer = 24036
Public Const WSAEALREADY As Integer = 24037
Public Const WSAENOTSOCK As Integer = 24038
Public Const WSAEDESTADDRREQ As Integer = 24039
Public Const WSAEMSGSIZE As Integer = 24040
Public Const WSAEPROTOTYPE As Integer = 24041
Public Const WSAENOPROTOOPT As Integer = 24042
Public Const WSAEPROTONOSUPPORT As Integer = 24043
Public Const WSAESOCKTNOSUPPORT As Integer = 24044
Public Const WSAEOPNOTSUPP As Integer = 24045
Public Const WSAEPFNOSUPPORT As Integer = 24046
Public Const WSAEAFNOSUPPORT As Integer = 24047
Public Const WSAEADDRINUSE As Integer = 24048
Public Const WSAEADDRNOTAVAIL As Integer = 24049
Public Const WSAENETDOWN As Integer = 24050
Public Const WSAENETUNREACH As Integer = 24051
Public Const WSAENETRESET As Integer = 24052
Public Const WSAECONNABORTED As Integer = 24053
Public Const WSAECONNRESET As Integer = 24054
Public Const WSAENOBUFS As Integer = 24055
Public Const WSAEISCONN As Integer = 24056
Public Const WSAENOTCONN As Integer = 24057
Public Const WSAESHUTDOWN As Integer = 24058
Public Const WSAETOOMANYREFS As Integer = 24059
Public Const WSAETIMEDOUT As Integer = 24060
Public Const WSAECONNREFUSED As Integer = 24061
Public Const WSAELOOP As Integer = 24062
Public Const WSAENAMETOOLONG As Integer = 24063
Public Const WSAEHOSTDOWN As Integer = 24064
Public Const WSAEHOSTUNREACH As Integer = 24065
Public Const WSAENOTEMPTY As Integer = 24066
Public Const WSAEPROCLIM As Integer = 24067
Public Const WSAEUSERS As Integer = 24068
Public Const WSAEDQUOT As Integer = 24069
Public Const WSAESTALE As Integer = 24070
Public Const WSAEREMOTE As Integer = 24071
Public Const WSASYSNOTREADY As Integer = 24091
Public Const WSAVERNOTSUPPORTED As Integer = 24092
Public Const WSANOTINITIALISED As Integer = 24093
Public Const WSAHOST_NOT_FOUND As Integer = 25001
Public Const WSATRY_AGAIN As Integer = 25002
Public Const WSANO_RECOVERY As Integer = 25003
Public Const WSANO_DATA As Integer = 25004
Public Const WSANO_ADDRESS As Integer = 2500
#End If

Sub DarCuerpoYCabeza(ByVal userIndex As Integer)
'*************************************************
'Author: Nacho (Integer)
'Last modified: 14/03/2007
'Elije una cabeza para el usuario y le da un Body
'*************************************************
Dim NewBody As Integer
Dim NewHead As Integer
Dim UserRaza As Byte
Dim UserGenero As Byte
UserGenero = UserList(userIndex).Genero
UserRaza = UserList(userIndex).Raza
Select Case UserGenero
   Case eGenero.Hombre
        Select Case UserRaza
            Case eRaza.Humano
                NewHead = RandomNumber(1, 40)
                NewBody = 1
            Case eRaza.Elfo
                NewHead = RandomNumber(101, 112)
                NewBody = 2
            Case eRaza.ElfoOscuro
                NewHead = RandomNumber(200, 210)
                NewBody = 3
            Case eRaza.Enano
                NewHead = RandomNumber(300, 306)
                NewBody = 300
            Case eRaza.Gnomo
                NewHead = RandomNumber(401, 406)
                NewBody = 300
            Case eRaza.Orco
                NewHead = 509 + RandomNumber(1, 7) 'no se sabe,blizzard
                NewBody = 401
        End Select
   Case eGenero.Mujer
        Select Case UserRaza
            Case eRaza.Humano
                NewHead = RandomNumber(70, 79)
                NewBody = 1
            Case eRaza.Elfo
                NewHead = RandomNumber(170, 178)
                NewBody = 2
            Case eRaza.ElfoOscuro
                NewHead = RandomNumber(270, 278)
                NewBody = 3
            Case eRaza.Gnomo
                NewHead = RandomNumber(370, 372)
                NewBody = 300
            Case eRaza.Enano
                NewHead = RandomNumber(470, 476)
                NewBody = 300
            Case eRaza.Orco
                NewHead = 516 + RandomNumber(1, 2) 'no se sabe,blizzard
                NewBody = 402 'no se sabe,blizzard
        End Select
End Select
UserList(userIndex).Char.head = NewHead
UserList(userIndex).Char.body = NewBody
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
    
Next i

AsciiValidos = True

End Function

Function Numeric(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mid$(cad, i, 1))
    
    If (car < 48 Or car > 57) Then
        Numeric = False
        Exit Function
    End If
    
Next i

Numeric = True

End Function


Function NombrePermitido(ByVal Nombre As String) As Boolean
Dim i As Integer

For i = 1 To UBound(ForbidenNames)
    If InStr(Nombre, ForbidenNames(i)) Then
            NombrePermitido = False
            Exit Function
    End If
Next i

NombrePermitido = True

End Function

Function ValidateSkills(ByVal userIndex As Integer) As Boolean

Dim LoopC As Integer

For LoopC = 1 To NUMSKILLS
    If UserList(userIndex).Stats.UserSkills(LoopC) < 0 Then
        Exit Function
        If UserList(userIndex).Stats.UserSkills(LoopC) > 100 Then UserList(userIndex).Stats.UserSkills(LoopC) = 100
    End If
Next LoopC

ValidateSkills = True
    
End Function

Sub ConnectNewUser(ByVal userIndex As Integer, ByRef name As String, ByVal UserRaza As eRaza, ByVal UserSexo As eGenero, ByVal UserClase As eClass, _
                    ByRef skills() As Byte)
'*************************************************
'Author: Unknown
'Last modified: 20/4/2007
'Conecta un nuevo Usuario
'23/01/2007 Pablo (ToxicWaste) - Agregué ResetFaccion al crear usuario
'24/01/2007 Pablo (ToxicWaste) - Agregué el nuevo mana inicial de los magos.
'12/02/2007 Pablo (ToxicWaste) - Puse + 1 de const al Elfo normal.
'20/04/2007 Pablo (ToxicWaste) - Puse -1 de fuerza al Elfo.
'*************************************************

If Not AsciiValidos(name) Or LenB(name) = 0 Then
    Call WriteErrorMsg(userIndex, "Nombre invalido.")
    Exit Sub
End If

Dim LoopC As Long
Dim totalskpts As Long

UserList(userIndex).flags.Muerto = 0
UserList(userIndex).flags.Escondido = 0

UserList(userIndex).name = name
UserList(userIndex).Clase = UserClase
UserList(userIndex).Raza = UserRaza
UserList(userIndex).Genero = UserSexo
'UserList(UserIndex).email = UserEmail

Select Case UserRaza
    Case eRaza.Humano
        UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) + 1
        UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad) + 1
        UserList(userIndex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(userIndex).Stats.UserAtributos(eAtributos.Constitucion) + 2
    Case eRaza.Elfo
        UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) - 1
        UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad) + 4
        UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia) + 2
        UserList(userIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(userIndex).Stats.UserAtributos(eAtributos.Carisma) + 2
        UserList(userIndex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(userIndex).Stats.UserAtributos(eAtributos.Constitucion) + 1
    Case eRaza.ElfoOscuro
        UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) + 2
        UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad) + 2
        UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia) + 2
        UserList(userIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(userIndex).Stats.UserAtributos(eAtributos.Carisma) - 3
    Case eRaza.Enano
        UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) + 3
        UserList(userIndex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(userIndex).Stats.UserAtributos(eAtributos.Constitucion) + 3
        UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia) - 6
        UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad) - 1
        UserList(userIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(userIndex).Stats.UserAtributos(eAtributos.Carisma) - 2
    Case eRaza.Gnomo
        UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) - 4
        UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia) + 3
        UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad) + 3
        UserList(userIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(userIndex).Stats.UserAtributos(eAtributos.Carisma) + 1
    Case eRaza.Orco
        UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) = UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) + 3
        UserList(userIndex).Stats.UserAtributos(eAtributos.Constitucion) = UserList(userIndex).Stats.UserAtributos(eAtributos.Constitucion) + 3
        UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia) = UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia) - 6
        UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad) = UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad) - 2
        UserList(userIndex).Stats.UserAtributos(eAtributos.Carisma) = UserList(userIndex).Stats.UserAtributos(eAtributos.Carisma) - 2
End Select


For LoopC = 1 To NUMSKILLS
    UserList(userIndex).Stats.UserSkills(LoopC) = skills(LoopC - 1)
    totalskpts = totalskpts + Abs(UserList(userIndex).Stats.UserSkills(LoopC))
Next LoopC


If totalskpts > 10 Then
    Call LogHackAttemp(UserList(userIndex).name & " intento hackear los skills.")
    Call BorrarUsuario(UserList(userIndex).name)
    Call closeConnection(userIndex)
    Exit Sub
End If
'%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%

UserList(userIndex).Char.Heading = eHeading.SOUTH

Call DarCuerpoYCabeza(userIndex)
UserList(userIndex).OrigChar = UserList(userIndex).Char
   
 
UserList(userIndex).Char.WeaponAnim = NingunArma
UserList(userIndex).Char.ShieldAnim = NingunEscudo
UserList(userIndex).Char.CascoAnim = NingunCasco

Dim MiInt As Long
MiInt = RandomNumber(1, UserList(userIndex).Stats.UserAtributos(eAtributos.Constitucion) \ 3)

UserList(userIndex).Stats.MaxHP = 15 + MiInt
UserList(userIndex).Stats.MinHP = 15 + MiInt

MiInt = RandomNumber(1, UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad) \ 6)
If MiInt = 1 Then MiInt = 2

UserList(userIndex).Stats.MaxSta = 20 * MiInt
UserList(userIndex).Stats.MinSta = 20 * MiInt


UserList(userIndex).Stats.MaxAGU = 100
UserList(userIndex).Stats.MinAGU = 100

UserList(userIndex).Stats.MaxHam = 100
UserList(userIndex).Stats.MinHam = 100


'<-----------------MANA----------------------->
If UserClase = eClass.Mage Then 'Cambio en mana inicial (ToxicWaste)
    MiInt = UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia) * 3
    UserList(userIndex).Stats.MaxMAN = MiInt
    UserList(userIndex).Stats.MinMAN = MiInt
ElseIf UserClase = eClass.Cleric Or UserClase = eClass.Druid _
    Or UserClase = eClass.Bard Or UserClase = eClass.Assasin Then
        UserList(userIndex).Stats.MaxMAN = 50
        UserList(userIndex).Stats.MinMAN = 50
ElseIf UserClase = eClass.Bandit Then 'Mana Inicial del Bandido (ToxicWaste)
        UserList(userIndex).Stats.MaxMAN = 150
        UserList(userIndex).Stats.MinMAN = 150
Else
    UserList(userIndex).Stats.MaxMAN = 0
    UserList(userIndex).Stats.MinMAN = 0
End If

If UserClase = eClass.Mage Or UserClase = eClass.Cleric Or _
   UserClase = eClass.Druid Or UserClase = eClass.Bard Or _
   UserClase = eClass.Assasin Then
        UserList(userIndex).Stats.UserHechizos(1) = 2
End If

UserList(userIndex).Stats.MaxHIT = 2
UserList(userIndex).Stats.MinHIT = 1

UserList(userIndex).Stats.GLD = 0

UserList(userIndex).Stats.Exp = 0
UserList(userIndex).Stats.ELU = 300
UserList(userIndex).Stats.ELV = 1

'???????????????? INVENTARIO ¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿¿ PONER ITEMS ESPECIFICOS DE CADA Clase
UserList(userIndex).Invent.NroItems = 4

UserList(userIndex).Invent.Object(1).ObjIndex = 467
UserList(userIndex).Invent.Object(1).amount = 100

UserList(userIndex).Invent.Object(2).ObjIndex = 468
UserList(userIndex).Invent.Object(2).amount = 100

UserList(userIndex).Invent.Object(3).ObjIndex = 460
UserList(userIndex).Invent.Object(3).amount = 1
UserList(userIndex).Invent.Object(3).Equipped = 1

Select Case UserRaza
    Case eRaza.Humano
        UserList(userIndex).Invent.Object(4).ObjIndex = 463
    Case eRaza.Elfo
        UserList(userIndex).Invent.Object(4).ObjIndex = 464
    Case eRaza.ElfoOscuro
        UserList(userIndex).Invent.Object(4).ObjIndex = 465
    Case eRaza.Enano
        UserList(userIndex).Invent.Object(4).ObjIndex = 466
    Case eRaza.Gnomo
        UserList(userIndex).Invent.Object(4).ObjIndex = 466
    Case eRaza.Orco
        UserList(userIndex).Invent.Object(4).ObjIndex = 463
End Select

UserList(userIndex).Invent.Object(4).amount = 1
UserList(userIndex).Invent.Object(4).Equipped = 1

UserList(userIndex).Invent.ArmourEqpSlot = 4
UserList(userIndex).Invent.ArmourEqpObjIndex = UserList(userIndex).Invent.Object(4).ObjIndex

UserList(userIndex).Invent.WeaponEqpObjIndex = UserList(userIndex).Invent.Object(3).ObjIndex
UserList(userIndex).Invent.WeaponEqpSlot = 3

UserList(userIndex).Pos.Map = DungeonNewbie.Map
UserList(userIndex).Pos.X = DungeonNewbie.X
UserList(userIndex).Pos.Y = DungeonNewbie.Y


#If ConUpTime Then
    UserList(userIndex).LogOnTime = Now
    UserList(userIndex).UpTime = 0
#End If

'Valores Default de facciones al Activar nuevo usuario
Call ResetFacciones(userIndex)

'Guardamos el usuario.
Call dbSaveCharData(userIndex)

UserList(userIndex).UserAccount.CharCount = UserList(userIndex).UserAccount.CharCount + 1
UserList(userIndex).UserAccount.Chars(UserList(userIndex).UserAccount.CharCount) = UCase(name)
'Si se guardo correctamente, guardamos la cuenta
Call dbSaveAccountData(userIndex)

'Open User
'Call ConnectUser(UserIndex, Name)

'Re open account
Call WriteAccountLogged(userIndex)
  
End Sub
Public Sub closeConnection(ByVal userIndex As Integer, Optional ByVal forced As Boolean = False)

    Debug.Print UserList(userIndex).name & " Start closeConnection"
    'Si se corto la coneccion (Salio sin /SALIR)
    'Se cierra la cuenta y la coneccion, pero el usuario queda logueado dependiendo de su posicion y privilegios.
    Call closeAccount(userIndex)
    
    Debug.Print UserList(userIndex).name & " Account Closed"
    
    If UserList(userIndex).ConnID <> -1 And UserList(userIndex).ConnIDValida Then
        Call BorraSlotSock(UserList(userIndex).ConnID)
        Call WSApicloseConnection(UserList(userIndex).ConnID)
        UserList(userIndex).ConnID = -1
        UserList(userIndex).ConnIDValida = False
        UserList(userIndex).NumeroPaquetesPorMiliSec = 0
    End If
    
    Debug.Print UserList(userIndex).name & " Connection Closed"
    
    'Hay un user logueado??
    If UserList(userIndex).flags.UserLogged Then
        If forced Then
            Call closeChar(userIndex)
        Else
            Call Cerrar_Usuario(userIndex)
        End If
    Else
        Call closeSocket(userIndex)
    End If
    
    Debug.Print UserList(userIndex).name & " End closeConnection"
End Sub


Public Sub closeAccount(ByVal userIndex As Integer)
    'Guardamos la cuenta.
    Call dbSaveAccountData(userIndex)

    'Deslogueamos la cuenta.
    UserList(userIndex).UserAccount.Logged = False
    UserList(userIndex).UserAccount.name = ""
    
    Call ResetAccountBoveda(userIndex)
    
End Sub


Public Sub closeChar(ByVal userIndex As Integer)
    'Es el mismo user al que está revisando el centinela??
    'IMPORTANTE!!! hacerlo antes de resetear así todavía sabemos el nombre del user
    ' y lo podemos loguear
    
    
    
    If Centinela.RevisandoUserIndex = userIndex Then _
        Call modCentinela.CentinelaUserLogout
    
    'mato los comercios seguros
    If UserList(userIndex).ComUsu.DestUsu > 0 Then
        If UserList(UserList(userIndex).ComUsu.DestUsu).flags.UserLogged Then
            If UserList(UserList(userIndex).ComUsu.DestUsu).ComUsu.DestUsu = userIndex Then
                Call WriteConsoleMsg(UserList(userIndex).ComUsu.DestUsu, "Comercio cancelado por el otro usuario", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(userIndex).ComUsu.DestUsu)
                Call FlushBuffer(UserList(userIndex).ComUsu.DestUsu)
            End If
        End If
    End If
    
    If UserList(userIndex).flags.EnDuelo = 1 Then
        Call WarpUserChar(userIndex, 26, 50, 50)
        UserList(userIndex).flags.EnDuelo = 0
    Else
        If UserList(userIndex).Pos.Map = 173 Then
            Call WarpUserChar(userIndex, 26, 50, 50)
        End If
    End If
   
    If ColaTorneo.Existe(UserList(userIndex).name) Then ColaTorneo.Quitar (UserList(userIndex).name)
    
    If UserList(userIndex).Pos.Map = MAPAESPERA Or UserList(userIndex).Pos.Map = MAPATORNEO Then
        Call WarpUserChar(userIndex, 26, 50, 50)
    End If

    'Subastas
    If userIndex = MayorOfertaUserIndex Then
        Dim i As Integer
        UserList(userIndex).Stats.GLD = UserList(userIndex).Stats.GLD + MayorOferta
        MayorOferta = 0
        MayorOfertaUserIndex = 0
        MayorOferta = SubastaMinimo
        For i = 1 To LastUser
            Call WriteConsoleMsg(i, "El usuario que habia hecho la mejor oferta se ha ido!.", FontTypeNames.FONTTYPE_INFO)
        Next i
    End If
    
    If userIndex = SubastaUserIndex Then
        Dim BlizzObj As Obj
        BlizzObj.ObjIndex = SubastaObjIndex
        BlizzObj.amount = 1

        If MayorOfertaUserIndex <> 0 Then
            UserList(SubastaUserIndex).Stats.GLD = UserList(SubastaUserIndex).Stats.GLD + MayorOferta
            If Not MeterItemEnInventario(MayorOfertaUserIndex, BlizzObj) Then
                Call TirarItemAlPiso(UserList(MayorOfertaUserIndex).Pos, BlizzObj)
            End If
        Else
            If Not MeterItemEnInventario(SubastaUserIndex, BlizzObj) Then
                Call TirarItemAlPiso(UserList(SubastaUserIndex).Pos, BlizzObj)
            End If
        End If
        
        frmMain.TSubasta.Enabled = False
        MinutosSubasta = 0
        SubastaEnCurso = 0
        SubastaUserIndex = 0
        MayorOferta = 0
        MayorOfertaUserIndex = 0
    End If
    'End Subastas
    
    If UserList(userIndex).flags.UserLogged Then
        If NumUsers > 0 Then NumUsers = NumUsers - 1
        Call closeUser(userIndex)
        Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    End If
    
    If UserList(userIndex).ConnID = -1 Then
        Call closeSocket(userIndex)
    Else
        Call WriteDisconnect(userIndex)
        Call WriteAccountLogged(userIndex)
    End If
    
End Sub


#If UsarQueSocket = 1 Or UsarQueSocket = 2 Then

Sub closeSocket(ByVal userIndex As Integer, Optional ByVal cerrarlo As Boolean = True)
Dim LoopC As Integer

On Error GoTo errhandler

    'Empty buffer for reuse
    Call UserList(userIndex).incomingData.ReadASCIIStringFixed(UserList(userIndex).incomingData.length)
    
    Call ResetUserSlot(userIndex)
    
    'Deslogueamos el user y si se desconecto buscamos el ultimo elemento del array.
    If userIndex = LastUser Then
        Do Until (UserList(LastUser).flags.UserLogged Or UserList(LastUser).UserAccount.Logged)
            LastUser = LastUser - 1
            If LastUser < 1 Then Exit Do
        Loop
    End If
    
Exit Sub

errhandler:
    UserList(userIndex).ConnID = -1
    UserList(userIndex).ConnIDValida = False
    UserList(userIndex).NumeroPaquetesPorMiliSec = 0
    Call ResetUserSlot(userIndex)

    Call LogError("closeConnection - Error = " & Err.Number & " - Descripción = " & Err.description & " - UserIndex = " & userIndex)
End Sub

#ElseIf UsarQueSocket = 0 Then

Sub closeSocket(ByVal userIndex As Integer)
On Error GoTo errhandler
    
    
    
    UserList(userIndex).ConnID = -1
    UserList(userIndex).NumeroPaquetesPorMiliSec = 0

    If userIndex = LastUser And LastUser > 1 Then
        Do Until UserList(LastUser).flags.UserLogged
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop
    End If

    If UserList(userIndex).flags.UserLogged Then
            If NumUsers <> 0 Then NumUsers = NumUsers - 1
            Call closeUser(userIndex)
    End If

    frmMain.Socket2(userIndex).Cleanup
    Unload frmMain.Socket2(userIndex)
    Call ResetUserSlot(userIndex)

Exit Sub

errhandler:
    UserList(userIndex).ConnID = -1
    UserList(userIndex).NumeroPaquetesPorMiliSec = 0
    Call ResetUserSlot(userIndex)
End Sub







#ElseIf UsarQueSocket = 3 Then

Sub closeSocket(ByVal userIndex As Integer, Optional ByVal cerrarlo As Boolean = True)

On Error GoTo errhandler

Dim NURestados As Boolean
Dim CoNnEcTiOnId As Long


    NURestados = False
    CoNnEcTiOnId = UserList(userIndex).ConnID
    
    'call logindex(UserIndex, "******> Sub closeConnection. ConnId: " & CoNnEcTiOnId & " Cerrarlo: " & Cerrarlo)
    
    
  
    UserList(userIndex).ConnID = -1 'inabilitamos operaciones en socket
    UserList(userIndex).NumeroPaquetesPorMiliSec = 0

    If userIndex = LastUser And LastUser > 1 Then
        Do
            LastUser = LastUser - 1
            If LastUser <= 1 Then Exit Do
        Loop While UserList(LastUser).flags.UserLogged = True
    End If

    If UserList(userIndex).flags.UserLogged Then
            If NumUsers <> 0 Then NumUsers = NumUsers - 1
            NURestados = True
            Call closeUser(userIndex)
    End If
    
    Call ResetUserSlot(userIndex)
    
    'limpiada la userlist... reseteo el socket, si me lo piden
    'Me lo piden desde: cerrada intecional del servidor (casi todas
    'las llamadas a closeConnection del codigo)
    'No me lo piden desde: disconnect remoto (el on_close del control
    'de alejo realiza la desconexion automaticamente). Esto puede pasar
    'por ejemplo, si el cliente cierra el AO.
    If cerrarlo Then Call frmMain.TCPServ.CerrarSocket(CoNnEcTiOnId)

Exit Sub

errhandler:
    UserList(userIndex).NumeroPaquetesPorMiliSec = 0
    Call LogError("closeConnectionERR: " & Err.description & " UI:" & userIndex)
    
    If Not NURestados Then
        If UserList(userIndex).flags.UserLogged Then
            If NumUsers > 0 Then
                NumUsers = NumUsers - 1
            End If
            Call LogError("Cerre sin grabar a: " & UserList(userIndex).name)
        End If
    End If
    
    Call LogError("El usuario no guardado tenia connid " & CoNnEcTiOnId & ". Socket no liberado.")
    Call ResetUserSlot(userIndex)

End Sub


#End If

'[Alejo-21-5]: Cierra un socket sin limpiar el slot
Sub closeSocketSL(ByVal userIndex As Integer)

#If UsarQueSocket = 1 Then

If UserList(userIndex).ConnID <> -1 And UserList(userIndex).ConnIDValida Then
    Call BorraSlotSock(UserList(userIndex).ConnID)
    Call WSApicloseConnection(UserList(userIndex).ConnID)
    UserList(userIndex).ConnID = -1
    UserList(userIndex).ConnIDValida = False
    UserList(userIndex).NumeroPaquetesPorMiliSec = 0
End If

#ElseIf UsarQueSocket = 0 Then

If UserList(userIndex).ConnID <> -1 And UserList(userIndex).ConnIDValida Then
    frmMain.Socket2(userIndex).Cleanup
    Unload frmMain.Socket2(userIndex)
    UserList(userIndex).ConnIDValida = False
End If

#ElseIf UsarQueSocket = 2 Then

If UserList(userIndex).ConnID <> -1 And UserList(userIndex).ConnIDValida Then
    Call frmMain.Serv.CerrarSocket(UserList(userIndex).ConnID)
    UserList(userIndex).ConnIDValida = False
End If

#End If
End Sub

''
' Send an string to a Slot
'
' @param UserIndex The index of the User
' @param Datos The string that will be send
' @remarks If UsarQueSocket is 3 it won`t use the clsByteQueue

Public Function EnviarDatosASlot(ByVal userIndex As Integer, ByRef Datos As String) As Long
'***************************************************
'Author: Unknown
'Last Modification: 01/10/07
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
'Now it uses the clsByteQueue class and don`t make a FIFO Queue of String
'***************************************************

#If UsarQueSocket = 1 Then '**********************************************
    On Error GoTo Err
    
    Dim Ret As Long
    
    Ret = WsApiEnviar(userIndex, Datos)
    
    If Ret <> 0 And Ret <> WSAEWOULDBLOCK Then
        ' Close the socket avoiding any critical error
        Call closeConnection(userIndex)
    End If
Exit Function
    
Err:
        'If frmMain.SUPERLOG.Value = 1 Then LogCustom ("EnviarDatosASlot:: ERR Handler. UserIndex=" & UserIndex & " datos=" & Datos & " UL?/CId/CIdV?=" & UserList(UserIndex).flags.UserLogged & "/" & UserList(UserIndex).ConnID & "/" & UserList(UserIndex).ConnIDValida & " ERR: " & Err.Description)

#ElseIf UsarQueSocket = 0 Then '**********************************************
    
    If frmMain.Socket2(userIndex).Write(Datos, Len(Datos)) < 0 Then
        If frmMain.Socket2(userIndex).LastError = WSAEWOULDBLOCK Then
            ' WSAEWOULDBLOCK, put the data again in the outgoingData Buffer
            Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(Datos)
        Else
            'Close the socket avoiding any critical error
            Call Cerrar_Usuario(userIndex)
        End If
    End If
#ElseIf UsarQueSocket = 2 Then '**********************************************

    'Return value for this Socket:
    '--0) OK
    '--1) WSAEWOULDBLOCK
    '--2) ERROR
    
    Dim Ret As Long

    Ret = frmMain.Serv.Enviar(.ConnID, Datos, Len(Datos))
            
    If Ret = 1 Then
        ' WSAEWOULDBLOCK, put the data again in the outgoingData Buffer
        Call .outgoingData.WriteASCIIStringFixed(Datos)
    ElseIf Ret = 2 Then
        'Close socket avoiding any critical error
        Call closeConnection(userIndex)
    End If
    

#ElseIf UsarQueSocket = 3 Then
    'THIS SOCKET DOESN`T USE THE BYTE QUEUE CLASS
    Dim rv As Long
    'al carajo, esto encola solo!!! che, me aprobará los
    'parciales también?, este control hace todo solo!!!!
    On Error GoTo ErrorHandler
        
        If UserList(userIndex).ConnID = -1 Then
            Call LogError("TCP::EnviardatosASlot, se intento enviar datos a un UserIndex con ConnId=-1")
            Exit Function
        End If
        
        If frmMain.TCPServ.Enviar(UserList(userIndex).ConnID, Datos, Len(Datos)) = 2 Then Call closeConnection(userIndex)

Exit Function
ErrorHandler:
    Call LogError("TCP::EnviarDatosASlot. UI/ConnId/Datos: " & userIndex & "/" & UserList(userIndex).ConnID & "/" & Datos)
#End If '**********************************************

End Function
Function EstaPCarea(Index As Integer, Index2 As Integer) As Boolean


Dim X As Integer, Y As Integer
For Y = UserList(Index).Pos.Y - MinYBorder + 1 To UserList(Index).Pos.Y + MinYBorder - 1
        For X = UserList(Index).Pos.X - MinXBorder + 1 To UserList(Index).Pos.X + MinXBorder - 1

            If MapData(UserList(Index).Pos.Map, X, Y).userIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
        Next X
Next Y
EstaPCarea = False
End Function

Function HayPCarea(Pos As WorldPos) As Boolean


Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                If MapData(Pos.Map, X, Y).userIndex > 0 Then
                    HayPCarea = True
                    Exit Function
                End If
            End If
        Next X
Next Y
HayPCarea = False
End Function

Function HayOBJarea(Pos As WorldPos, ObjIndex As Integer) As Boolean


Dim X As Integer, Y As Integer
For Y = Pos.Y - MinYBorder + 1 To Pos.Y + MinYBorder - 1
        For X = Pos.X - MinXBorder + 1 To Pos.X + MinXBorder - 1
            If MapData(Pos.Map, X, Y).ObjInfo.ObjIndex = ObjIndex Then
                HayOBJarea = True
                Exit Function
            End If
        
        Next X
Next Y
HayOBJarea = False
End Function
Function ValidateChr(ByVal userIndex As Integer) As Boolean

ValidateChr = UserList(userIndex).Char.head <> 0 _
                And UserList(userIndex).Char.body <> 0 _
                And ValidateSkills(userIndex)

End Function

Sub ConnectUser(ByVal userIndex As Integer, ByRef name As String)
    Dim N As Integer
    Dim tStr As String

    'Cargamos los datos del personaje
    Call dbLoadCharData(name, userIndex)
    
    If Not ValidateChr(userIndex) Then
        Call WriteErrorMsg(userIndex, "Error en el personaje.")
        Exit Sub
    End If
    
    'Reseteamos los FLAGS
    UserList(userIndex).flags.Escondido = 0
    UserList(userIndex).flags.TargetNPC = 0
    UserList(userIndex).flags.TargetNpcTipo = eNPCType.Comun
    UserList(userIndex).flags.TargetObj = 0
    UserList(userIndex).flags.TargetUser = 0
    UserList(userIndex).Char.FX = 0
    
    If UserList(userIndex).Invent.EscudoEqpSlot = 0 Then UserList(userIndex).Char.ShieldAnim = NingunEscudo
    If UserList(userIndex).Invent.CascoEqpSlot = 0 Then UserList(userIndex).Char.CascoAnim = NingunCasco
    If UserList(userIndex).Invent.WeaponEqpSlot = 0 Then UserList(userIndex).Char.WeaponAnim = NingunArma
    
    Call UpdateUserInv(True, userIndex, 0)
    Call UpdateUserHechizos(True, userIndex, 0)
    
    If UserList(userIndex).flags.Paralizado Then
        Call WriteParalizeOK(userIndex)
    End If
    
    ''
    'TODO : Feo, esto tiene que ser parche cliente
    If UserList(userIndex).flags.Estupidez = 0 Then
        Call WriteDumbNoMore(userIndex)
    End If
    
    If Not MapaValido(UserList(userIndex).Pos.Map) Then
        UserList(userIndex).Pos = Ullathorpe
        If Not MapaValido(UserList(userIndex).Pos.Map) Then
            Call WriteErrorMsg(userIndex, "EL PJ se encuenta en un mapa invalido.")
            Call FlushBuffer(userIndex)
            Call closeConnection(userIndex)
            Exit Sub
        End If
    End If
    
    
    
    
    'Tratamos de evitar en lo posible el "Telefrag". Solo 1 intento de loguear en pos adjacentes.
    'Codigo por Pablo (ToxicWaste) y revisado por Nacho (Integer), corregido para que realmetne ande y no tire el server por Juan Martín Sotuyo Dodero (Maraxus)
    If MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).userIndex <> 0 Or MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).NpcIndex <> 0 Then
        Dim FoundPlace As Boolean
        Dim tX As Long
        Dim tY As Long
        
        FoundPlace = False
        
        For tY = UserList(userIndex).Pos.Y - 1 To UserList(userIndex).Pos.Y + 1
            For tX = UserList(userIndex).Pos.X - 1 To UserList(userIndex).Pos.X + 1
                'reviso que sea pos legal en tierra, que no haya User ni NPC para poder loguear.
                If LegalPos(UserList(userIndex).Pos.Map, tX, tY, False, True) Then
                    FoundPlace = True
                    Exit For
                End If
            Next tX
            
            If FoundPlace Then _
                Exit For
        Next tY
        
        If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
            UserList(userIndex).Pos.X = tX
            UserList(userIndex).Pos.Y = tY
        Else
            'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
            If MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).userIndex <> 0 Then
                'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                If UserList(MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).userIndex).ComUsu.DestUsu > 0 Then
                    'Le avisamos al que estaba comerciando que se tuvo que ir.
                    If UserList(UserList(MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).userIndex).ComUsu.DestUsu).flags.UserLogged Then
                        Call FinComerciarUsu(UserList(MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).userIndex).ComUsu.DestUsu)
                        Call WriteConsoleMsg(UserList(MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).userIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                        Call FlushBuffer(UserList(MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).userIndex).ComUsu.DestUsu)
                    End If
                    'Lo sacamos.
                    If UserList(MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).userIndex).flags.UserLogged Then
                        Call FinComerciarUsu(MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).userIndex)
                        Call WriteErrorMsg(MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).userIndex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                        Call FlushBuffer(MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).userIndex)
                    End If
                End If
                
                Call closeConnection(MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).userIndex)
            End If
        End If
    End If
    
    If UserList(userIndex).flags.Muerto = 1 Then
        Call Empollando(userIndex)
    End If
    
    'Nombre de sistema
    UserList(userIndex).name = name
    
    UserList(userIndex).showName = True 'Por default los nombres son visibles
    
    'If in the water, and has a boat, equip it!
    If UserList(userIndex).Invent.BarcoObjIndex > 0 And HayAgua(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y) Then
            If UserList(userIndex).flags.Muerto = 0 Then
            UserList(userIndex).Char.body = ObjData(UserList(userIndex).Invent.BarcoObjIndex).Ropaje
            UserList(userIndex).Char.head = 0
            UserList(userIndex).Char.WeaponAnim = NingunArma
            UserList(userIndex).Char.ShieldAnim = NingunEscudo
            UserList(userIndex).Char.CascoAnim = NingunCasco
            UserList(userIndex).flags.Navegando = 1
         Else
            UserList(userIndex).Char.body = iFragataFantasmal
            UserList(userIndex).Char.head = 0
            UserList(userIndex).Char.WeaponAnim = NingunArma
            UserList(userIndex).Char.ShieldAnim = NingunEscudo
            UserList(userIndex).Char.CascoAnim = NingunCasco
            UserList(userIndex).flags.Navegando = 1
         End If
    End If
    
    If UserList(userIndex).Invent.MonturaObjIndex > 0 Then
         UserList(userIndex).Char.body = ObjData(UserList(userIndex).Invent.MonturaObjIndex).Ropaje
         UserList(userIndex).Char.WeaponAnim = NingunArma
         UserList(userIndex).Char.ShieldAnim = NingunEscudo
         UserList(userIndex).Char.CascoAnim = NingunCasco
         UserList(userIndex).flags.Montado = 1
         Call WriteMontuToggle(userIndex)
    End If
    
    'Info
    Call WriteUserIndexInServer(userIndex) 'Enviamos el User index
    Call WriteChangeMap(userIndex, UserList(userIndex).Pos.Map, MapInfo(UserList(userIndex).Pos.Map).MapVersion) 'Carga el mapa
    Call WritePlayMidi(userIndex, val(ReadField(1, MapInfo(UserList(userIndex).Pos.Map).Music, 45)))
    
    'Reseteamos los privilegios
    UserList(userIndex).flags.Privilegios = 0
    
    'Vemos que Clase de user es (se lo usa para setear los privilegios alcrear el PJ)
    If EsAdmin(name) Then
        UserList(userIndex).flags.Privilegios = UserList(userIndex).flags.Privilegios Or PlayerType.Admin
        Call LogGM(UserList(userIndex).name, "Se conecto con ip:" & UserList(userIndex).ip)
    ElseIf EsDios(name) Then
        UserList(userIndex).flags.Privilegios = UserList(userIndex).flags.Privilegios Or PlayerType.Dios
        Call LogGM(UserList(userIndex).name, "Se conecto con ip:" & UserList(userIndex).ip)
    ElseIf EsSemiDios(name) Then
        UserList(userIndex).flags.Privilegios = UserList(userIndex).flags.Privilegios Or PlayerType.SemiDios
        Call LogGM(UserList(userIndex).name, "Se conecto con ip:" & UserList(userIndex).ip)
    ElseIf EsConsejero(name) Then
        UserList(userIndex).flags.Privilegios = UserList(userIndex).flags.Privilegios Or PlayerType.Consejero
        Call LogGM(UserList(userIndex).name, "Se conecto con ip:" & UserList(userIndex).ip)
    Else
        UserList(userIndex).flags.Privilegios = UserList(userIndex).flags.Privilegios Or PlayerType.User
        UserList(userIndex).flags.AdminPerseguible = True
    End If
    
    'Add RM flag if needed
    If EsRolesMaster(name) Then
        UserList(userIndex).flags.Privilegios = UserList(userIndex).flags.Privilegios Or PlayerType.RoleMaster
    End If
    
    If UserList(userIndex).flags.Privilegios <> PlayerType.User And UserList(userIndex).flags.Privilegios <> PlayerType.ChaosCouncil And UserList(userIndex).flags.Privilegios <> PlayerType.RoyalCouncil Then
        UserList(userIndex).flags.ChatColor = RGB(0, 255, 0)
    Else
        UserList(userIndex).flags.ChatColor = vbWhite
    End If
    
    ''[EL OSO]: TRAIGO ESTO ACA ARRIBA PARA DARLE EL IP!
    #If ConUpTime Then
        UserList(userIndex).LogOnTime = Now
    #End If
    
    'Crea  el personaje del usuario
    Call MakeUserChar(True, UserList(userIndex).Pos.Map, userIndex, UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y)
    
    Call WriteUserCharIndexInServer(userIndex)
    ''[/el oso]
    
    If EnTesting And UserList(userIndex).Stats.ELV >= 18 Then
        Call WriteErrorMsg(userIndex, "Servidor en Testing por unos minutos, conectese con PJs de nivel menor a 18. No se conecte con Pjs que puedan resultar importantes por ahora pues pueden arruinarse.")
        Call FlushBuffer(userIndex)
        Call closeConnection(userIndex)
        Exit Sub
    End If
    
    'Actualiza el Num de usuarios
    'DE ACA EN ADELANTE GRABA EL CHARFILE, OJO!
    NumUsers = NumUsers + 1
    UserList(userIndex).flags.UserLogged = True
    
    'usado para borrar Pjs
    Call WriteVar(CharPath & UserList(userIndex).name & ".chr", "INIT", "Logged", "1")
    
    Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    
    MapInfo(UserList(userIndex).Pos.Map).NumUsers = MapInfo(UserList(userIndex).Pos.Map).NumUsers + 1
    
    If UserList(userIndex).Stats.SkillPts > 0 Then
        Call WriteSendSkills(userIndex)
        Call WriteLevelUp(userIndex, UserList(userIndex).Stats.SkillPts)
    End If
    
    If NumUsers > DayStats.MaxUsuarios Then DayStats.MaxUsuarios = NumUsers
    
    
    
    If UserList(userIndex).NroMacotas > 0 Then
        Dim i As Integer
        For i = 1 To MAXMASCOTAS
            If UserList(userIndex).MascotasType(i) > 0 Then
                UserList(userIndex).MascotasIndex(i) = SpawnNpc(UserList(userIndex).MascotasType(i), UserList(userIndex).Pos, True, True)
                
                If UserList(userIndex).MascotasIndex(i) > 0 Then
                    Npclist(UserList(userIndex).MascotasIndex(i)).MaestroUser = userIndex
                    Call FollowAmo(UserList(userIndex).MascotasIndex(i))
                Else
                    UserList(userIndex).MascotasIndex(i) = 0
                End If
            End If
        Next i
    End If
    
    If UserList(userIndex).flags.Navegando = 1 Then
        Call WriteNavigateToggle(userIndex)
    End If
    
    If ServerSoloGMs > 0 Then
        If (UserList(userIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) = 0 Then
            Call WriteErrorMsg(userIndex, "Servidor restringido a administradores de jerarquia mayor o igual a: " & ServerSoloGMs & ". Por favor intente en unos momentos.")
            Call FlushBuffer(userIndex)
            Call closeConnection(userIndex)
            Exit Sub
        End If
    End If
    
    If UserList(userIndex).GuildIndex > 0 Then
        'welcome to the show baby...
        If Not modGuilds.m_ConectarMiembroAClan(userIndex, UserList(userIndex).GuildIndex) Then
            Call WriteConsoleMsg(userIndex, "Tu estado no te permite entrar al clan.", FontTypeNames.FONTTYPE_GUILD)
        End If
    End If
    
    Call WriteLoggedMessage(userIndex)
    
    Call SendMOTD(userIndex)
    
    If haciendoBK Then
        Call WritePauseToggle(userIndex)
        Call WriteConsoleMsg(userIndex, "Servidor> Por favor espera algunos segundos, WorldSave esta ejecutandose.", FontTypeNames.FONTTYPE_SERVER)
    End If
    
    If EnPausa Then
        Call WritePauseToggle(userIndex)
        Call WriteConsoleMsg(userIndex, "Servidor> Lo sentimos mucho pero el servidor se encuentra actualmente detenido. Intenta ingresar más tarde.", FontTypeNames.FONTTYPE_SERVER)
    End If
    
    If NumUsers > recordusuarios Then
        Call SendData(SendTarget.toall, 0, PrepareMessageConsoleMsg("Record de usuarios conectados simultaniamente." & "Hay " & NumUsers & " usuarios.", FontTypeNames.FONTTYPE_INFO))
        recordusuarios = NumUsers
        Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(recordusuarios))
        
        Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)
    End If
    
    Call WriteUpdateUserStats(userIndex)
    Call WriteUpdateStrengthAgility(userIndex)
    Call WriteUpdateHit(userIndex)
    Call WriteUpdateArmor(userIndex)
    Call WriteUpdateEscu(userIndex)
    Call WriteUpdateCasco(userIndex)
    Call WriteUpdateHungerAndThirst(userIndex)
    Call WriteUpdateSta(userIndex)
    
    Call modGuilds.SendGuildNews(userIndex)
    
    If UserList(userIndex).flags.NoActualizado Then
        Call WriteUpdateNeeded(userIndex)
    End If
    
    If Lloviendo Then
        Call WriteRainToggle(userIndex)
    End If
    
    Call WriteSendNight(userIndex, DeNoche)
        
    tStr = modGuilds.a_ObtenerRechazoDeChar(UserList(userIndex).name)
    
    If LenB(tStr) <> 0 Then
        Call WriteShowMessageBox(userIndex, "Tu solicitud de ingreso al clan ha sido rechazada. El clan te explica que: " & tStr)
    End If
    
    If UserList(userIndex).Pos.Map = MAPAESPERA Or UserList(userIndex).Pos.Map = MAPATORNEO Then
        If Not Torneo Then
            If Not ColaTorneo.Existe(UserList(userIndex).name) Then
                Call WarpUserChar(userIndex, 26, 50, 50)
            End If
        End If
    End If
    
    
    'Send the SpawnFX
    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCreateFX(UserList(userIndex).Char.CharIndex, FXIDs.FXWARP, 0))
    
    'Load the user statistics
    Call Statistics.UserConnected(userIndex)
    
    Call MostrarNumUsers
    
    N = FreeFile
    Open App.Path & "\logs\numusers.log" For Output As N
    Print #N, NumUsers
    Close #N
    
    N = FreeFile
    'Log
    Open App.Path & "\logs\Connect.log" For Append Shared As #N
    Print #N, UserList(userIndex).name & " ha entrado al juego. UserIndex:" & userIndex & " " & time & " " & Date
    Close #N

End Sub

Sub SendMOTD(ByVal userIndex As Integer)
    Dim j As Long
    
    Call WriteConsoleMsg(userIndex, "Mensajes de entrada:", FontTypeNames.FONTTYPE_DUELO)
    For j = 1 To MaxLines
        Call WriteConsoleMsg(userIndex, MOTD(j).texto, FontTypeNames.FONTTYPE_GUILDMSG)
    Next j
End Sub

Sub ResetFacciones(ByVal userIndex As Integer)
    With UserList(userIndex).Faccion
        .Alineacion = e_Alineacion.Neutro
        .CriminalesMatados = 0
        .NeutralesMatados = 0
        .CiudadanosMatados = 0
        .SalioFaccion = 0
        .SalioFaccionCounter = 0
    End With
End Sub


Sub ResetContadores(ByVal userIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'05/20/2007 Integer - Agregue todas las variables que faltaban.
'*************************************************
    With UserList(userIndex).Counters
        .AGUACounter = 0
        .AttackCounter = 0
        .Ceguera = 0
        .COMCounter = 0
        .Estupidez = 0
        .Frio = 0
        .HPCounter = 0
        .IdleCount = 0
        .Invisibilidad = 0
        .Paralisis = 0
        .Pasos = 0
        .Pena = 0
        .PiqueteC = 0
        .STACounter = 0
        .Veneno = 0
        .Trabajando = 0
        .Ocultando = 0
        .bPuedeMeditar = False
        .Lava = 0
        .Mimetismo = 0
        .Saliendo = False
        .Salir = 0
        .TiempoOculto = 0
        .TimerMagiaGolpe = 0
        .TimerGolpeMagia = 0
        .TimerLanzarSpell = 0
        .TimerPuedeAtacar = 0
        .TimerPuedePotear = 0
        .TimerPuedeUsarArco = 0
        .TimerPuedeTrabajar = 0
        .TimerUsar = 0
    End With
End Sub

Sub ResetCharInfo(ByVal userIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(userIndex).Char
        .body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .head = 0
        .loops = 0
        .Heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
    End With
End Sub

Sub ResetBasicUserInfo(ByVal userIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/15/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'*************************************************
    With UserList(userIndex)
        .name = vbNullString
        .modName = vbNullString
        .Desc = vbNullString
        .DescRM = vbNullString
        .Pos.Map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .ip = vbNullString
        .Clase = 0
        .email = vbNullString
        .Genero = 0
        .Raza = 0
        
        .EmpoCont = 0
        
        With .Stats
            .Banco = 0
            .ELV = 0
            .ELU = 0
            .Exp = 0
            .def = 0
            '.CriminalesMatados = 0
            .NPCsMuertos = 0
            .UsuariosMatados = 0
            .SkillPts = 0
            .GLD = 0
            .UserAtributos(1) = 0
            .UserAtributos(2) = 0
            .UserAtributos(3) = 0
            .UserAtributos(4) = 0
            .UserAtributos(5) = 0
            .UserAtributosBackUP(1) = 0
            .UserAtributosBackUP(2) = 0
            .UserAtributosBackUP(3) = 0
            .UserAtributosBackUP(4) = 0
            .UserAtributosBackUP(5) = 0
        End With
        
    End With
End Sub


Sub ResetGuildInfo(ByVal userIndex As Integer)
    If UserList(userIndex).EscucheClan > 0 Then
        Call modGuilds.GMDejaDeEscucharClan(userIndex, UserList(userIndex).EscucheClan)
        UserList(userIndex).EscucheClan = 0
    End If
    If UserList(userIndex).GuildIndex > 0 Then
        Call modGuilds.m_DesconectarMiembroDelClan(userIndex, UserList(userIndex).GuildIndex)
    End If
    UserList(userIndex).GuildIndex = 0
End Sub

Sub ResetUserFlags(ByVal userIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 03/29/2006
'Resetea todos los valores generales y las stats
'03/15/2006 Maraxus - Uso de With para mayor performance y claridad.
'03/29/2006 Maraxus - Reseteo el CentinelaOK también.
'*************************************************
    With UserList(userIndex).flags
        .Comerciando = False
        .Ban = 0
        .Escondido = 0
        .DuracionEfecto = 0
        .NpcInv = 0
        .StatsChanged = 0
        .TargetNPC = 0
        .TargetNpcTipo = eNPCType.Comun
        .TargetObj = 0
        .TargetObjMap = 0
        .TargetObjX = 0
        .TargetObjY = 0
        .TargetUser = 0
        .TipoPocion = 0
        .TomoPocion = False
        .Descuento = vbNullString
        .Hambre = 0
        .Sed = 0
        .Descansar = False
        .ModoCombate = False
        .Vuela = 0
        .Navegando = 0
        .Montado = 0
        .Oculto = 0
        .Envenenado = 0
        .invisible = 0
        .Paralizado = 0
        .Maldicion = 0
        .Bendicion = 0
        .Meditando = 0
        .Privilegios = 0
        .PuedeMoverse = 0
        .OldBody = 0
        .OldHead = 0
        .AdminInvisible = 0
        .ValCoDe = 0
        .Hechizo = 0
        .TimesWalk = 0
        .StartWalk = 0
        .CountSH = 0
        .EstaEmpo = 0
        .Silenciado = 0
        .CentinelaOK = False
        .AdminPerseguible = False
        .EnDuelo = 0
    End With
End Sub

Sub ResetUserSpells(ByVal userIndex As Integer)
    Dim LoopC As Long
    For LoopC = 1 To MAXUSERHECHIZOS
        UserList(userIndex).Stats.UserHechizos(LoopC) = 0
    Next LoopC
End Sub

Sub ResetUserPets(ByVal userIndex As Integer)
    Dim LoopC As Long
    
    UserList(userIndex).NroMacotas = 0
        
    For LoopC = 1 To MAXMASCOTAS
        UserList(userIndex).MascotasIndex(LoopC) = 0
        UserList(userIndex).MascotasType(LoopC) = 0
    Next LoopC
End Sub

Sub ResetAccountBoveda(ByVal userIndex As Integer)
    Dim LoopC As Long
    
    For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
          UserList(userIndex).UserAccount.Boveda.Object(LoopC).amount = 0
          UserList(userIndex).UserAccount.Boveda.Object(LoopC).Equipped = 0
          UserList(userIndex).UserAccount.Boveda.Object(LoopC).ObjIndex = 0
    Next LoopC
    
    UserList(userIndex).UserAccount.Boveda.NroItems = 0
End Sub

Public Sub LimpiarComercioSeguro(ByVal userIndex As Integer)
    With UserList(userIndex).ComUsu
        If .DestUsu > 0 Then
            Call FinComerciarUsu(.DestUsu)
            Call FinComerciarUsu(userIndex)
        End If
    End With
End Sub

Sub ResetUserSlot(ByVal userIndex As Integer)

'UserList(UserIndex).ConnIDValida = False
'UserList(UserIndex).ConnID = -1

Call LimpiarComercioSeguro(userIndex)
Call ResetFacciones(userIndex)
Call ResetContadores(userIndex)
Call ResetCharInfo(userIndex)
Call ResetBasicUserInfo(userIndex)
Call ResetGuildInfo(userIndex)
Call ResetUserFlags(userIndex)
Call LimpiarInventario(userIndex)
Call ResetUserSpells(userIndex)
Call ResetUserPets(userIndex)

With UserList(userIndex).ComUsu
    .Acepto = False
    .cant = 0
    .DestNick = vbNullString
    .DestUsu = 0
    .Objeto = 0
End With

End Sub

Sub closeUser(ByVal userIndex As Integer)
'Call LogTarea("CloseUser " & UserIndex)
On Error GoTo errhandler

Dim N As Integer
Dim X As Integer
Dim Y As Integer
Dim LoopC As Integer
Dim Map As Integer
Dim name As String
Dim Raza As eRaza
Dim Clase As eClass
Dim i As Integer

Dim aN As Integer

aN = UserList(userIndex).flags.AtacadoPorNpc
If aN > 0 Then
      Npclist(aN).Movement = Npclist(aN).flags.OldMovement
      Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
      Npclist(aN).flags.AttackedBy = vbNullString
End If
aN = UserList(userIndex).flags.npcAttacked
If aN > 0 Then
    If Npclist(aN).flags.AttackedFirstBy = UserList(userIndex).name Then
        Npclist(aN).flags.AttackedFirstBy = vbNullString
    End If
End If
UserList(userIndex).flags.AtacadoPorNpc = 0
UserList(userIndex).flags.npcAttacked = 0


'CHECK:: ACA SE GUARDAN UN MONTON DE COSAS QUE NO SE OCUPAN PARA NADA :S
Map = UserList(userIndex).Pos.Map
X = UserList(userIndex).Pos.X
Y = UserList(userIndex).Pos.Y
name = UCase$(UserList(userIndex).name)
Raza = UserList(userIndex).Raza
Clase = UserList(userIndex).Clase

UserList(userIndex).Char.FX = 0
UserList(userIndex).Char.loops = 0
Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCreateFX(UserList(userIndex).Char.CharIndex, 0, 0))


UserList(userIndex).flags.UserLogged = False
UserList(userIndex).Counters.Saliendo = False

'Le devolvemos el Body y head originales
If UserList(userIndex).flags.AdminInvisible = 1 Then Call DoAdminInvisible(userIndex)

'Save statistics
Call Statistics.UserDisconnected(userIndex)

' Grabamos el personaje del usuario
dbSaveCharData userIndex

'Quitar el dialogo
'If MapInfo(Map).NumUsers > 0 Then
'    Call SendToUserArea(UserIndex, "QDL" & UserList(UserIndex).Char.charindex)
'End If

If MapInfo(Map).NumUsers > 0 Then
    Call SendData(SendTarget.ToPCAreaButIndex, userIndex, PrepareMessageRemoveCharDialog(UserList(userIndex).Char.CharIndex))
End If

'Borrar el personaje
If UserList(userIndex).Char.CharIndex > 0 Then
    Call EraseUserChar(userIndex)
End If

'Borrar mascotas
For i = 1 To MAXMASCOTAS
    If UserList(userIndex).MascotasIndex(i) > 0 Then
        If Npclist(UserList(userIndex).MascotasIndex(i)).flags.NPCActive Then _
            Call QuitarNPC(UserList(userIndex).MascotasIndex(i))
    End If
Next i

'Update Map Users
MapInfo(Map).NumUsers = MapInfo(Map).NumUsers - 1

If MapInfo(Map).NumUsers < 0 Then
    MapInfo(Map).NumUsers = 0
End If

' Si el usuario habia dejado un msg en la gm's queue lo borramos
If Ayuda.Existe(UserList(userIndex).name) Then Call Ayuda.Quitar(UserList(userIndex).name)

Call MostrarNumUsers

N = FreeFile(1)
Open App.Path & "\logs\Connect.log" For Append Shared As #N
Print #N, name & " há dejado el juego. " & "User Index:" & userIndex & " " & time & " " & Date
Close #N

Exit Sub

errhandler:
Call LogError("Error en CloseUser. Número " & Err.Number & " Descripción: " & Err.description)

End Sub

Sub ReloadSokcet()
On Error GoTo errhandler
#If UsarQueSocket = 1 Then

    Call LogApiSock("ReloadSokcet() " & NumUsers & " " & LastUser & " " & MaxUsers)
    
    If NumUsers <= 0 Then
        Call WSApiReiniciarSockets
    Else
'       Call apicloseConnection(SockListen)
'       SockListen = ListenForConnect(Puerto, hWndMsg, "")
    End If

#ElseIf UsarQueSocket = 0 Then

    frmMain.Socket1.Cleanup
    Call ConfigListeningSocket(frmMain.Socket1, Puerto)
    
#ElseIf UsarQueSocket = 2 Then

    

#End If

Exit Sub
errhandler:
    Call LogError("Error en CheckSocketState " & Err.Number & ": " & Err.description)

End Sub

Public Sub EcharPjsNoPrivilegiados()
Dim LoopC As Long

For LoopC = 1 To LastUser
    If UserList(LoopC).flags.UserLogged And UserList(LoopC).ConnID >= 0 And UserList(LoopC).ConnIDValida Then
        If UserList(LoopC).flags.Privilegios And PlayerType.User Then
            Call closeConnection(LoopC)
        End If
    End If
Next LoopC

End Sub
