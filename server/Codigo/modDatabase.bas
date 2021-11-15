Attribute VB_Name = "modDatabase"
Option Explicit

' Server Database mySQL
Private db As ADODB.Connection
Public rs As New ADODB.Recordset


Public Function dbConnect() As Boolean
On Error GoTo errhandler

    Set db = New ADODB.Connection
    db.CursorLocation = adUseClient
    db.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=localhost;DATABASE=rebornao;UID=root;PWD=root; OPTION=3"
    db.Open

    dbConnect = True
Exit Function
errhandler:
    dbConnect = False
    MsgBox "No se pudo conectar a la base de datos"
End Function

Public Function dbDisconnect() As Boolean
On Error GoTo errhandler
    db.Close
    Set db = Nothing
Exit Function
errhandler:
End Function

Private Sub saveCharInit(ByVal userIndex As Integer)
    Dim sConsult As String
    
    Set rs = db.Execute("SELECT * FROM `charinit` WHERE Nombre='" & UCase(UserList(userIndex).name) & "'")
    
    If rs.BOF Or rs.EOF Then
        Call db.Execute("INSERT INTO `charinit` (Nombre) VALUES ('" & UCase(UserList(userIndex).name) & "')")
    End If
    Set rs = Nothing
    
    'charflags
    With UserList(userIndex)
        sConsult = "UPDATE `charinit` SET"
        sConsult = sConsult & " Nombre='" & UCase(UserList(userIndex).name) & "'"
        sConsult = sConsult & ",UpTime=" & .UpTime
        sConsult = sConsult & ",Genero=" & .Genero
        sConsult = sConsult & ",Clase=" & .Clase
        sConsult = sConsult & ",Raza=" & .Raza
        sConsult = sConsult & ",Ban=" & .flags.Ban
        sConsult = sConsult & ",Pena=" & .Counters.Pena
        sConsult = sConsult & ",Map=" & .Pos.Map
        sConsult = sConsult & ",X=" & .Pos.X
        sConsult = sConsult & ",Y=" & .Pos.Y
        sConsult = sConsult & ",GuildIndex=" & .GuildIndex
        'sConsult = sConsult & ",Desc= '" & .Desc & "'"
        
        'CHAR
        sConsult = sConsult & ",Heading=" & .Char.Heading
        sConsult = sConsult & ",Head=" & .Char.head
        sConsult = sConsult & ",Body=" & .Char.body
        sConsult = sConsult & ",WeaponAnim=" & .Char.WeaponAnim
        sConsult = sConsult & ",ShieldAnim=" & .Char.ShieldAnim
        sConsult = sConsult & ",CascoAnim=" & .Char.CascoAnim & " WHERE Nombre='" & UCase(UserList(userIndex).name) & "'"
    End With
    
    Call db.Execute(sConsult)
    
End Sub




Private Sub saveCharFlags(ByVal userIndex As Integer)
    Dim sConsult As String
    
    Set rs = db.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase(UserList(userIndex).name) & "'")
    
    If rs.BOF Or rs.EOF Then
        Call db.Execute("INSERT INTO `charflags` (Nombre) VALUES ('" & UCase(UserList(userIndex).name) & "')")
    End If
    Set rs = Nothing
    
    'charflags
    With UserList(userIndex)
        sConsult = "UPDATE `charflags` SET"
        sConsult = sConsult & " Nombre='" & UCase(UserList(userIndex).name) & "'"
        
        sConsult = sConsult & ",Navegando=" & .flags.Navegando
        sConsult = sConsult & ",Envenenado=" & .flags.Envenenado
        sConsult = sConsult & ",Muerto=" & .flags.Muerto
        sConsult = sConsult & ",Escondido=" & .flags.Escondido
        sConsult = sConsult & ",Hambre=" & .flags.Hambre
        sConsult = sConsult & ",Sed=" & .flags.Sed
        sConsult = sConsult & ",Desnudo=" & .flags.Desnudo
        sConsult = sConsult & ",Paralizado=" & .flags.Paralizado
        sConsult = sConsult & ",Montado=" & .flags.Montado
    End With
    
    Call db.Execute(sConsult)
End Sub
Private Sub saveCharAtributes(ByVal userIndex As Integer)
    Dim sConsult As String
    Dim i As Byte
    
    'charatrib
    Set rs = db.Execute("SELECT * FROM `charatrib` WHERE Nombre='" & UCase(UserList(userIndex).name) & "'")
    If rs.BOF Or rs.EOF Then
        Call db.Execute("INSERT INTO `charatrib` (Nombre) VALUES ('" & UCase(UserList(userIndex).name) & "')")
    End If
    Set rs = Nothing
    
    With UserList(userIndex)
        sConsult = "UPDATE `charatrib` SET "
        sConsult = sConsult & " Nombre='" & UCase(UserList(userIndex).name) & "'"
        For i = 1 To NUMATRIBUTOS
            sConsult = sConsult & ",AT" & i & "=" & .Stats.UserAtributos(i)
        Next i
        sConsult = sConsult & " WHERE Nombre='" & UCase(UserList(userIndex).name) & "'"
    End With
    
    Call db.Execute(sConsult)
End Sub
Private Sub saveCharSkills(ByVal userIndex As Integer)
    Dim sConsult As String
    Dim i As Byte
    
    'charskills
    Set rs = db.Execute("SELECT * FROM `charskills` WHERE Nombre='" & UCase(UserList(userIndex).name) & "'")
    If rs.BOF Or rs.EOF Then
        Call db.Execute("INSERT INTO `charskills` (Nombre) VALUES ('" & UCase(UserList(userIndex).name) & "')")
    End If
    Set rs = Nothing
    
    With UserList(userIndex)
        sConsult = "UPDATE `charskills` SET "
        sConsult = sConsult & " Nombre='" & UCase(UserList(userIndex).name) & "'"
        For i = 1 To NUMSKILLS
            sConsult = sConsult & ",SK" & i & "=" & .Stats.UserSkills(i)
        Next i
        sConsult = sConsult & " WHERE Nombre='" & UCase(UserList(userIndex).name) & "'"
    End With
    
    Call db.Execute(sConsult)
End Sub

Private Sub saveCharSpells(ByVal userIndex As Integer)
    Dim sConsult As String
    Dim i As Byte
    
    'charhechizos
    Set rs = db.Execute("SELECT * FROM `charhechizos` WHERE Nombre='" & UCase(UserList(userIndex).name) & "'")
    If rs.BOF Or rs.EOF Then
        Call db.Execute("INSERT INTO `charhechizos` (Nombre) VALUES ('" & UCase(UserList(userIndex).name) & "')")
    End If
    Set rs = Nothing
    
    
    With UserList(userIndex)
        sConsult = "UPDATE `charhechizos` SET "
        sConsult = sConsult & " Nombre='" & UCase(UserList(userIndex).name) & "'"
        For i = 1 To MAXUSERHECHIZOS
            sConsult = sConsult & ",H" & i & "=" & .Stats.UserHechizos(i)
        Next i
        sConsult = sConsult & " WHERE Nombre='" & UCase(UserList(userIndex).name) & "'"
    End With
    
    Call db.Execute(sConsult)
End Sub

Private Sub saveCharStats(ByVal userIndex As Integer)
    Dim sConsult As String
    
    'charstats
    Set rs = db.Execute("SELECT * FROM `charstats` WHERE Nombre='" & UCase(UserList(userIndex).name) & "'")
    If rs.BOF Or rs.EOF Then
        Call db.Execute("INSERT INTO `charstats` (Nombre) VALUES ('" & UCase(UserList(userIndex).name) & "')")
    End If
    Set rs = Nothing
    
    With UserList(userIndex)
        sConsult = "UPDATE `charstats` SET "
        sConsult = sConsult & " Nombre='" & UCase(UserList(userIndex).name) & "'"
        sConsult = sConsult & ",GLD=" & .Stats.GLD
        sConsult = sConsult & ",MaxHP=" & .Stats.MaxHP
        sConsult = sConsult & ",MinHP=" & .Stats.MinHP
        sConsult = sConsult & ",MinSta=" & .Stats.MinSta
        sConsult = sConsult & ",Banco=" & .Stats.Banco
        sConsult = sConsult & ",MaxMAN=" & .Stats.MaxMAN
        sConsult = sConsult & ",MinMAN=" & .Stats.MinMAN
        sConsult = sConsult & ",MaxHIT=" & .Stats.MaxHIT
        sConsult = sConsult & ",MinHIT=" & .Stats.MinHIT
        sConsult = sConsult & ",MinAGU=" & .Stats.MinAGU
        sConsult = sConsult & ",MinHAM=" & .Stats.MinHam
        sConsult = sConsult & ",MaxSTA=" & .Stats.MaxSta
        sConsult = sConsult & ",MaxAGU=" & .Stats.MaxAGU
        sConsult = sConsult & ",MaxHAM=" & .Stats.MaxHam
        sConsult = sConsult & ",SkillPtsLibres=" & .Stats.SkillPts
        sConsult = sConsult & ",Exp=" & .Stats.Exp
        sConsult = sConsult & ",ELV=" & .Stats.ELV
        sConsult = sConsult & ",ELU=" & .Stats.ELU
        sConsult = sConsult & ",NpcsMuertes=" & .Stats.NPCsMuertos
        sConsult = sConsult & ",UsuariosMatados=" & .Stats.UsuariosMatados
        'sConsult = sConsult & ",VecesMurioUsuario=" & .Stats.VecesMurioUsuario
        sConsult = sConsult & " WHERE Nombre='" & UCase(UserList(userIndex).name) & "'"
    End With
    
    Call db.Execute(sConsult)
End Sub

Private Sub saveCharInventory(ByVal userIndex As Integer)
    Dim sConsult As String
    Dim i As Byte
    
    'USERINVENTORY
    Set rs = db.Execute("SELECT * FROM `charinvent` WHERE Nombre='" & UCase(UserList(userIndex).name) & "'")
    If rs.BOF Or rs.EOF Then
        Call db.Execute("INSERT INTO `charinvent` (Nombre) VALUES ('" & UCase(UserList(userIndex).name) & "')")
    End If
    Set rs = Nothing
    
    With UserList(userIndex)
        sConsult = "UPDATE `charinvent` SET "
        sConsult = sConsult & " Nombre='" & UCase(UserList(userIndex).name) & "'"
        For i = 1 To MAX_INVENTORY_SLOTS
            sConsult = sConsult & ",OBJ" & i & "=" & .Invent.Object(i).ObjIndex & ",CANT" & i & "=" & .Invent.Object(i).amount & ",EQP" & i & "=" & .Invent.Object(i).Equipped
        Next i
        
        sConsult = sConsult & ",WEAPONSLOT=" & .Invent.WeaponEqpSlot
        sConsult = sConsult & ",CASCOSLOT=" & .Invent.CascoEqpSlot
        sConsult = sConsult & ",ARMORSLOT=" & .Invent.ArmourEqpSlot
        sConsult = sConsult & ",SHIELDSLOT=" & .Invent.EscudoEqpSlot
        'sConsult = sConsult & ",HERRAMIENTASLOT=" & .Invent.HerramientaEqpslot
        sConsult = sConsult & ",MUNICIONSLOT=" & .Invent.MunicionEqpSlot
        sConsult = sConsult & ",BarcoSlot=" & .Invent.BarcoSlot
        
        sConsult = sConsult & " WHERE Nombre='" & UCase(UserList(userIndex).name) & "'"
    End With
    
    Call db.Execute(sConsult)
End Sub

Private Sub saveCharPriviledges(ByVal userIndex As Integer)
    Dim sConsult As String
    
    'USER PRIVILEDGES
    Set rs = db.Execute("SELECT * FROM `charprivs` WHERE Nombre='" & UCase(UserList(userIndex).name) & "'")
    If rs.BOF Or rs.EOF Then
        Call db.Execute("INSERT INTO `charprivs` (Nombre) VALUES ('" & UCase(UserList(userIndex).name) & "')")
    End If
    Set rs = Nothing
    
    With UserList(userIndex)
        sConsult = "UPDATE `charprivs` SET "
        sConsult = sConsult & "CONSEJOCAOS=" & 0
        sConsult = sConsult & ",CONSEJOREAL=" & 0
        sConsult = sConsult & " WHERE Nombre='" & UCase(UserList(userIndex).name) & "'"
    End With
    
    Call db.Execute(sConsult)
End Sub

Private Sub saveCharFaction(ByVal userIndex As Integer)
    Dim sConsult As String
    
    'USER FACCIONES
    Set rs = db.Execute("SELECT * FROM `charfaccion` WHERE Nombre='" & UCase(UserList(userIndex).name) & "'")
    If rs.BOF Or rs.EOF Then
        Call db.Execute("INSERT INTO `charfaccion` (Nombre) VALUES ('" & UCase(UserList(userIndex).name) & "')")
    End If
    Set rs = Nothing
    
    With UserList(userIndex)
        sConsult = "UPDATE `charfaccion` SET "
        sConsult = sConsult & "Alineacion=" & .Faccion.Alineacion
        sConsult = sConsult & ",CiudadanosMatados=" & .Faccion.CiudadanosMatados
        sConsult = sConsult & ",CriminalesMatados=" & .Faccion.CriminalesMatados
        sConsult = sConsult & ",NeutralesMatados=" & .Faccion.NeutralesMatados
        sConsult = sConsult & ",RangoFaccionario=" & .Faccion.RangoFaccionario
        sConsult = sConsult & " WHERE Nombre='" & UCase(UserList(userIndex).name) & "'"
    End With
    
    Call db.Execute(sConsult)
End Sub
Public Sub dbSaveCharData(ByVal userIndex As Integer)
On Error GoTo errhandler:
    
    Call saveCharFlags(userIndex)
    Call saveCharInit(userIndex)
    Call saveCharStats(userIndex)
    Call saveCharSpells(userIndex)
    Call saveCharSkills(userIndex)
    Call saveCharAtributes(userIndex)
    Call saveCharInventory(userIndex)
    Call saveCharFaction(userIndex)
    Call saveCharPriviledges(userIndex)
    
    Set rs = Nothing
Exit Sub
errhandler:
    Set rs = Nothing
    Debug.Print "ERROR: dbSaveCharData " & Err.description
End Sub

Private Function loadCharFlags(ByVal sCharName As String, ByVal userIndex As Integer) As Boolean
On Error GoTo errhandler:
    loadCharFlags = True
    'USER FLAGS
    Set rs = db.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(sCharName) & "'")
    'Existe el personaje?
    If rs.BOF Or rs.EOF Then
        loadCharFlags = False
        Exit Function
    End If
    
    UserList(userIndex).name = sCharName

    With UserList(userIndex)
        .flags.Navegando = rs!Navegando
        .flags.Envenenado = rs!Envenenado
        .flags.Muerto = rs!Muerto
        .flags.Escondido = rs!Escondido
        .flags.Hambre = rs!Hambre
        .flags.Sed = rs!Sed
        .flags.Desnudo = rs!Desnudo
        .flags.Paralizado = rs!Paralizado
        .flags.Montado = rs!Montado
        
        If .flags.Paralizado = 1 Then
            .Counters.Paralisis = IntervaloParalizado
        End If
        
    End With
Exit Function
errhandler:
    loadCharFlags = False
End Function


Private Function loadCharInit(ByVal sCharName As String, ByVal userIndex As Integer) As Boolean
On Error GoTo errhandler:
    loadCharInit = True
    'USER FLAGS
    Set rs = db.Execute("SELECT * FROM `charinit` WHERE Nombre='" & UCase$(sCharName) & "'")
    'Existe el personaje?
    If rs.BOF Or rs.EOF Then
        loadCharInit = False
        Exit Function
    End If
    
    UserList(userIndex).name = sCharName

    With UserList(userIndex)
#If ConUpTime Then
        .UpTime = rs!UpTime
#End If
        .Genero = rs!Genero
        .Clase = rs!Clase
        .Raza = rs!Raza
        
        .flags.Ban = rs!Ban
        .Counters.Pena = rs!Pena
        
        .OrigChar.Heading = rs!Heading
        .OrigChar.head = rs!head
        .OrigChar.body = rs!body
        .OrigChar.WeaponAnim = rs!WeaponAnim
        .OrigChar.ShieldAnim = rs!ShieldAnim
        .OrigChar.CascoAnim = rs!CascoAnim
        
        .Char.Heading = eHeading.SOUTH
        
        If .flags.Muerto = 0 Then
           .Char = .OrigChar
        Else
            .Char.body = iCuerpoMuerto
            .Char.head = iCabezaMuerto
            .Char.WeaponAnim = NingunArma
            .Char.ShieldAnim = NingunEscudo
            .Char.CascoAnim = NingunCasco
        End If
        
        '.Desc = rs!Desc

        .Pos.Map = rs!Map
        .Pos.X = rs!X
        .Pos.Y = rs!Y
        
        .NroMacotas = 0

        .GuildIndex = rs!GuildIndex
    End With
Exit Function
errhandler:
    loadCharInit = False
End Function

Private Function loadCharAtributes(ByVal sCharName As String, ByVal userIndex As Integer) As Boolean
On Error GoTo errhandler:
    
    Dim i As Byte
    
    loadCharAtributes = True
    
    'USER ATRIBUTES
    Set rs = db.Execute("SELECT * FROM `charatrib` WHERE Nombre='" & UCase$(sCharName) & "'")
    'Existe el personaje?
    If rs.BOF Or rs.EOF Then
        loadCharAtributes = False
        Exit Function
    End If
    
    With UserList(userIndex)
        For i = 1 To NUMATRIBUTOS
            .Stats.UserAtributos(i) = rs.Fields("AT" & i)
            .Stats.UserAtributosBackUP(i) = .Stats.UserAtributos(i)
        Next i
    End With

Exit Function
errhandler:
    loadCharAtributes = False
End Function

Private Function loadCharSkills(ByVal sCharName As String, ByVal userIndex As Integer) As Boolean
On Error GoTo errhandler:
    Dim i As Byte
    
    loadCharSkills = True
    
    'USER SKILLS
    Set rs = db.Execute("SELECT * FROM `charskills` WHERE Nombre='" & UCase$(sCharName) & "'")
    'Existe el personaje?
    If rs.BOF Or rs.EOF Then
        loadCharSkills = False
        Exit Function
    End If
    
    With UserList(userIndex)
        For i = 1 To NUMSKILLS
            .Stats.UserSkills(i) = rs.Fields("SK" & i)
        Next i
    End With

Exit Function
errhandler:
    loadCharSkills = False
End Function

Private Function loadCharSpells(ByVal sCharName As String, ByVal userIndex As Integer) As Boolean
On Error GoTo errhandler:
    Dim i As Byte

    loadCharSpells = True
    
    'USER HECHIZOS
    Set rs = db.Execute("SELECT * FROM `charhechizos` WHERE Nombre='" & UCase$(sCharName) & "'")
    'Existe el personaje?
    If rs.BOF Or rs.EOF Then
        loadCharSpells = False
        Exit Function
    End If
    
    With UserList(userIndex)
        For i = 1 To MAXUSERHECHIZOS
            .Stats.UserHechizos(i) = rs.Fields("H" & i)
        Next i
    End With

Exit Function
errhandler:
    loadCharSpells = False
End Function

Private Function loadCharStats(ByVal sCharName As String, ByVal userIndex As Integer) As Boolean
On Error GoTo errhandler:
    loadCharStats = True
    
    'USER STATS
    Set rs = db.Execute("SELECT * FROM `charstats` WHERE Nombre='" & UCase$(sCharName) & "'")
    'Existe el personaje?
    If rs.BOF Or rs.EOF Then
        loadCharStats = False
        Exit Function
    End If
    
    With UserList(userIndex)
        .Stats.GLD = rs!GLD
        '.Stats.Banco = rs!Banco
        .Stats.MaxHP = rs!MaxHP
        .Stats.MinHP = rs!MinHP
        .Stats.MinSta = rs!MinSta
        .Stats.MaxMAN = rs!MaxMAN
        .Stats.MinMAN = rs!MinMAN
        .Stats.MaxHIT = rs!MaxHIT
        .Stats.MinHIT = rs!MinHIT
        .Stats.MinAGU = rs!MinAGU
        .Stats.MaxSta = rs!MaxSta
        .Stats.MinHam = rs!MinHam
        .Stats.SkillPts = rs!SkillPtsLibres
        '.Stats.VecesMurioUsuario = rs!VecesMurioUsuario
        .Stats.Exp = rs!Exp
        .Stats.ELV = rs!ELV
        .Stats.ELU = rs!ELU
        .Stats.NPCsMuertos = rs!NpcsMuertes
        .Stats.UsuariosMatados = rs!UsuariosMatados
        
        .Stats.MaxHam = 100
        .Stats.MaxAGU = 100
        
        If .Stats.MinAGU < 1 Then .flags.Sed = 1
        If .Stats.MinHam < 1 Then .flags.Hambre = 1
        If .Stats.MinHP < 1 Then .flags.Muerto = 1
    End With

Exit Function
errhandler:
    loadCharStats = False
End Function

Private Function loadCharInventory(ByVal sCharName As String, ByVal userIndex As Integer) As Boolean
On Error GoTo errhandler:
    
    Dim i As Byte

    loadCharInventory = True
    
    'USER INVENTORY
    Set rs = db.Execute("SELECT * FROM `charinvent` WHERE Nombre='" & UCase$(sCharName) & "'")
    'Existe el personaje?
    If rs.BOF Or rs.EOF Then
        loadCharInventory = False
        Exit Function
    End If
    
    With UserList(userIndex)
        .Invent.NroItems = rs!Items
        For i = 1 To MAX_INVENTORY_SLOTS
            .Invent.Object(i).ObjIndex = rs.Fields("OBJ" & i)
            .Invent.Object(i).amount = rs.Fields("CANT" & i)
            .Invent.Object(i).Equipped = rs.Fields("EQP" & i)
        Next i
        .Invent.CascoEqpSlot = rs!CASCOSLOT
        .Invent.ArmourEqpSlot = rs!ARMORSLOT
        .Invent.EscudoEqpSlot = rs!SHIELDSLOT
        .Invent.WeaponEqpSlot = rs!WEAPONSLOT
        '.Invent.HerramientaEqpslot = rs!HERRAMIENTASLOT
        .Invent.MunicionEqpSlot = rs!MUNICIONSLOT
        .Invent.BarcoSlot = rs!BarcoSlot
    End With
    'Asignamos los objIndex correspondientes a lo equipado.
    Call CharEquipSet(userIndex)

Exit Function
errhandler:
    loadCharInventory = False
End Function

Private Function loadCharPriviledges(ByVal sCharName As String, ByVal userIndex As Integer) As Boolean
On Error GoTo errhandler:
    loadCharPriviledges = True
    
    'USER PRIVILEDGES
    Set rs = db.Execute("SELECT * FROM `charprivs` WHERE Nombre='" & UCase$(sCharName) & "'")
    'Existe el personaje?
    If rs.BOF Or rs.EOF Then
        loadCharPriviledges = False
        Exit Function
    End If
    With UserList(userIndex)
        If rs!CONSEJOCAOS Then _
            .flags.Privilegios = .flags.Privilegios Or PlayerType.ChaosCouncil
        If rs!CONSEJOREAL Then _
            .flags.Privilegios = .flags.Privilegios Or PlayerType.RoyalCouncil
    End With

Exit Function
errhandler:
    loadCharPriviledges = False
End Function

Private Function loadCharFaction(ByVal sCharName As String, ByVal userIndex As Integer) As Boolean
On Error GoTo errhandler:
    loadCharFaction = True
    
    'USER FACCIONES
    Set rs = db.Execute("SELECT * FROM `charfaccion` WHERE Nombre='" & UCase$(sCharName) & "'")
    'Existe el personaje?
    If rs.BOF Or rs.EOF Then
        loadCharFaction = False
        Exit Function
    End If
    With UserList(userIndex)
        .Faccion.Alineacion = rs!Alineacion
        .Faccion.CiudadanosMatados = rs!CiudadanosMatados
        .Faccion.CriminalesMatados = rs!CriminalesMatados
        .Faccion.NeutralesMatados = rs!NeutralesMatados
        .Faccion.RangoFaccionario = rs!RangoFaccionario
    End With
    
    Set rs = Nothing

Exit Function
errhandler:
    loadCharFaction = False
End Function

Public Function dbLoadCharData(ByVal sCharName As String, ByVal userIndex As Integer) As Boolean
On Error GoTo errhandler:

    dbLoadCharData = loadCharStats(sCharName, userIndex)
    dbLoadCharData = dbLoadCharData And loadCharFlags(sCharName, userIndex)
    dbLoadCharData = dbLoadCharData And loadCharFaction(sCharName, userIndex)
    dbLoadCharData = dbLoadCharData And loadCharPriviledges(sCharName, userIndex)
    dbLoadCharData = dbLoadCharData And loadCharInventory(sCharName, userIndex)
    dbLoadCharData = dbLoadCharData And loadCharAtributes(sCharName, userIndex)
    dbLoadCharData = dbLoadCharData And loadCharSkills(sCharName, userIndex)
    dbLoadCharData = dbLoadCharData And loadCharSpells(sCharName, userIndex)
    dbLoadCharData = dbLoadCharData And loadCharInit(sCharName, userIndex)
    
    If Not dbLoadCharData Then GoTo errhandler
    
Exit Function
errhandler:
    Debug.Print "ERROR: dbLoadCharData " & Err.description
End Function

Sub CharEquipSet(ByVal userIndex As Integer)
    'Obtiene el indice-objeto del arma
    If UserList(userIndex).Invent.WeaponEqpSlot > 0 Then
        UserList(userIndex).Invent.WeaponEqpObjIndex = UserList(userIndex).Invent.Object(UserList(userIndex).Invent.WeaponEqpSlot).ObjIndex
        If ObjData(UserList(userIndex).Invent.WeaponEqpObjIndex).Aura Then
            UserList(userIndex).Char.Aura = ObjData(UserList(userIndex).Invent.WeaponEqpObjIndex).Aura
        End If
    End If
    
    'Obtiene el indice-objeto del armadura
    If UserList(userIndex).Invent.ArmourEqpSlot > 0 Then
        UserList(userIndex).Invent.ArmourEqpObjIndex = UserList(userIndex).Invent.Object(UserList(userIndex).Invent.ArmourEqpSlot).ObjIndex
        UserList(userIndex).flags.Desnudo = 0
    Else
        UserList(userIndex).flags.Desnudo = 1
    End If
    
    'Obtiene el indice-objeto del escudo
    If UserList(userIndex).Invent.EscudoEqpSlot > 0 Then
        UserList(userIndex).Invent.EscudoEqpObjIndex = UserList(userIndex).Invent.Object(UserList(userIndex).Invent.EscudoEqpSlot).ObjIndex
    End If
    
    'Obtiene el indice-objeto del casco
    If UserList(userIndex).Invent.CascoEqpSlot > 0 Then
        UserList(userIndex).Invent.CascoEqpObjIndex = UserList(userIndex).Invent.Object(UserList(userIndex).Invent.CascoEqpSlot).ObjIndex
    End If
    
    'Obtiene el indice-objeto barco
    If UserList(userIndex).Invent.BarcoSlot > 0 Then
        UserList(userIndex).Invent.BarcoObjIndex = UserList(userIndex).Invent.Object(UserList(userIndex).Invent.BarcoSlot).ObjIndex
    End If
    
    If UserList(userIndex).Invent.MonturaSlot > 0 Then
        UserList(userIndex).Invent.MonturaObjIndex = UserList(userIndex).Invent.Object(UserList(userIndex).Invent.MonturaSlot).ObjIndex
    End If
    
    'Obtiene el indice-objeto municion
    If UserList(userIndex).Invent.MunicionEqpSlot > 0 Then
        UserList(userIndex).Invent.MunicionEqpObjIndex = UserList(userIndex).Invent.Object(UserList(userIndex).Invent.MunicionEqpSlot).ObjIndex
    End If
    
    '[Alejo]
    'Obtiene el indice-objeto anilo
    If UserList(userIndex).Invent.AnilloEqpSlot > 0 Then
        UserList(userIndex).Invent.AnilloEqpObjIndex = UserList(userIndex).Invent.Object(UserList(userIndex).Invent.AnilloEqpSlot).ObjIndex
    End If
End Sub

Public Function dbSaveAccountData(ByVal userIndex As Integer) As Boolean
On Error GoTo errhandler:
    'SAVE NEW CHARS AND WAREHOUSE
    Dim sConsult As String
    Dim i As Byte
    Set rs = db.Execute("SELECT * FROM `cuentas` WHERE Nombre='" & UCase$(UserList(userIndex).UserAccount.name) & "'")
    dbSaveAccountData = Not (rs.BOF Or rs.EOF)
    If dbSaveAccountData Then
        If UserList(userIndex).UserAccount.CharCount > 0 Then
            sConsult = "UPDATE `cuentas` SET "
            sConsult = sConsult & "PJS=" & UserList(userIndex).UserAccount.CharCount
            For i = 1 To UserList(userIndex).UserAccount.CharCount
                If UserList(userIndex).UserAccount.Chars(i) <> "" Then
                    sConsult = sConsult & ",PJ" & i & "='" & UCase(UserList(userIndex).UserAccount.Chars(i)) & "'"
                End If
            Next i
            sConsult = sConsult & " WHERE Nombre='" & UCase(UserList(userIndex).UserAccount.name) & "'"
            Call db.Execute(sConsult)
        End If
        
        Call saveAccountBoveda(userIndex)
    End If

    Set rs = Nothing
Exit Function
errhandler:
    Debug.Print "ERROR: dbSaveAccountData " & Err.description
End Function
Public Sub saveAccountBoveda(ByVal userIndex As Integer)
    Dim sConsult As String
    Dim i As Byte
    
    Set rs = db.Execute("SELECT * FROM `boveda` WHERE Account='" & UCase$(UserList(userIndex).UserAccount.name) & "'")

    'Si no existe la boveda, la creamos.
    If rs.EOF Or rs.BOF Then
        Call db.Execute("INSERT INTO `boveda` (Account) VALUES ('" & UCase(UserList(userIndex).UserAccount.name) & "')")
    End If
    
    sConsult = "UPDATE `boveda` SET "
    sConsult = sConsult & "Gold=" & UserList(userIndex).Stats.Banco
    sConsult = sConsult & ",Items=" & UserList(userIndex).UserAccount.Boveda.NroItems
    For i = 1 To UserList(userIndex).UserAccount.Boveda.NroItems
        sConsult = sConsult & ",OBJ" & i & "=" & UserList(userIndex).UserAccount.Boveda.Object(i).ObjIndex
        sConsult = sConsult & ",CANT" & i & "=" & UserList(userIndex).UserAccount.Boveda.Object(i).amount
    Next i
    sConsult = sConsult & " WHERE Account='" & UCase(UserList(userIndex).UserAccount.name) & "'"
    Call db.Execute(sConsult)
End Sub
Public Function dbCheckAccountData(ByVal sAccountName As String, ByVal sAccountPassword As String) As Boolean
    'Devuelve true si la cuenta y la contraseña son correctas.
    Set rs = db.Execute("SELECT * FROM `cuentas` WHERE Nombre='" & UCase$(sAccountName) & "'" & "AND Password='" & sAccountPassword & "'")
    dbCheckAccountData = Not (rs.BOF Or rs.EOF)
    Set rs = Nothing
End Function

Public Function dbCharCheck(ByVal sName As String) As Boolean
    Dim i As Byte
    dbCharCheck = False
    
    For i = 1 To MAX_ACCOUNT_CHARS
        'Devuelve true si el pj existe.
        Set rs = db.Execute("SELECT * FROM `cuentas` WHERE PJ" & i & "='" & UCase$(sName) & "'")
        dbCharCheck = Not (rs.BOF Or rs.EOF)
        If dbCharCheck Then Exit For
    Next i
    
    Set rs = Nothing
End Function

Public Function dbLoadAccountData(ByVal userIndex As Integer) As Boolean
On Error GoTo errhandler:
    Dim i As Byte
    Set rs = db.Execute("SELECT * FROM `cuentas` WHERE Nombre='" & UCase$(UserList(userIndex).UserAccount.name) & "'")
    dbLoadAccountData = Not (rs.BOF Or rs.EOF)
    If dbLoadAccountData Then
        UserList(userIndex).UserAccount.CharCount = rs!PJS
        For i = 1 To UserList(userIndex).UserAccount.CharCount
            UserList(userIndex).UserAccount.Chars(i) = rs.Fields("PJ" & i)
        Next i
    End If
    
    Set rs = db.Execute("SELECT * FROM `boveda` WHERE Account='" & UCase$(UserList(userIndex).UserAccount.name) & "'")
    dbLoadAccountData = dbLoadAccountData And Not (rs.BOF Or rs.EOF)
    If dbLoadAccountData Then
        UserList(userIndex).Stats.Banco = rs.Fields("Gold")
        UserList(userIndex).UserAccount.Boveda.NroItems = rs.Fields("Items")
        For i = 1 To UserList(userIndex).UserAccount.Boveda.NroItems
            UserList(userIndex).UserAccount.Boveda.Object(i).ObjIndex = rs.Fields("OBJ" & i)
            UserList(userIndex).UserAccount.Boveda.Object(i).amount = rs.Fields("CANT" & i)
        Next i
    End If
    
Exit Function
errhandler:
    Debug.Print "ERROR: dbLoadAccountData " & Err.description
End Function

Public Function dbReadInteger(ByVal sTable As String, ByVal sField As String, ByVal searchBy As String, ByVal searchID As String) As Integer
    Set rs = db.Execute("SELECT * FROM `" & sTable & "` WHERE " & searchBy & "='" & searchID & "'")
    If Not (rs.BOF Or rs.EOF) Then
        dbReadInteger = val(rs.Fields(sField))
    End If
    Set rs = Nothing
End Function

Public Function dbGetAccountCharInfo(ByVal sName As String, iBody As Integer, iHead As Integer, iWeapon As Integer, iShield As Integer, iHelm As Integer) As Boolean
    iBody = dbReadInteger("charinit", "Body", "Nombre", sName)
    iShield = dbReadInteger("charinit", "ShieldAnim", "Nombre", sName)
    iHelm = dbReadInteger("charinit", "CascoAnim", "Nombre", sName)
    iWeapon = dbReadInteger("charinit", "WeaponAnim", "Nombre", sName)
    iHead = dbReadInteger("charinit", "Head", "Nombre", sName)
End Function

Public Function dbReadString(ByVal sTable As String, ByVal sField As String, ByVal searchBy As String, ByVal searchID As String) As String
    Set rs = db.Execute("SELECT * FROM `" & sTable & "` WHERE " & searchBy & "='" & searchID & "'")
    If Not (rs.BOF Or rs.EOF) Then
        dbReadString = str(rs.Fields(sField))
    End If
    Set rs = Nothing
End Function

Public Sub dbWriteString(ByVal table As String, ByVal field As String, ByVal userID As String, ByVal value As String)
    Dim consult As String
    
    consult = "UPDATE `" & table & "` SET " & field & "=" & value & " WHERE Nombre='" & userID & "'"
End Sub

Public Sub dbWriteInteger(ByVal table As String, ByVal field As String, ByVal userID As String, ByVal value As Integer)
    Dim consult As String
    
    consult = "UPDATE `" & table & "` SET " & field & "=" & value & " WHERE Nombre='" & userID & "'"
End Sub
