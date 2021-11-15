Attribute VB_Name = "UsUaRiOs"
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

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo Usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Rutinas de los usuarios
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Sub ActStats(ByVal VictimIndex As Integer, ByVal attackerIndex As Integer)

Dim DaExp As Integer

DaExp = CInt(UserList(VictimIndex).Stats.ELV * 2)

UserList(attackerIndex).Stats.Exp = UserList(attackerIndex).Stats.Exp + DaExp
If UserList(attackerIndex).Stats.Exp > MAXEXP Then _
    UserList(attackerIndex).Stats.Exp = MAXEXP

'Lo mata
Call WriteConsoleMsg(attackerIndex, "Has matado a " & UserList(VictimIndex).name & "!", FontTypeNames.FONTTYPE_FIGHT)
Call WriteConsoleMsg(attackerIndex, "Has ganado " & DaExp & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)
      
Call WriteConsoleMsg(VictimIndex, UserList(attackerIndex).name & " te ha matado!", FontTypeNames.FONTTYPE_FIGHT)

If UserList(attackerIndex).flags.EnDuelo = 1 Then
    Call WriteConsoleMsg(attackerIndex, "Has ganado el duelo!.", FontTypeNames.FONTTYPE_INFO)
    UserList(attackerIndex).Stats.GLD = UserList(attackerIndex).Stats.GLD + 30000
End If

Call UserDie(VictimIndex)

If UserList(attackerIndex).Stats.UsuariosMatados < MAXUSERMATADOS Then _
    UserList(attackerIndex).Stats.UsuariosMatados = UserList(attackerIndex).Stats.UsuariosMatados + 1

Call FlushBuffer(VictimIndex)

'Log
Call LogAsesinato(UserList(attackerIndex).name & " asesino a " & UserList(VictimIndex).name)

End Sub


Sub RevivirUsuario(ByVal userIndex As Integer)

UserList(userIndex).flags.Muerto = 0
UserList(userIndex).Stats.MinHP = UserList(userIndex).Stats.UserAtributos(eAtributos.Constitucion)

'If he died, venom should fade away
UserList(userIndex).flags.Envenenado = 0

'No puede estar empollando
UserList(userIndex).flags.EstaEmpo = 0
UserList(userIndex).EmpoCont = 0

If UserList(userIndex).Stats.MinHP > UserList(userIndex).Stats.MaxHP Then
    UserList(userIndex).Stats.MinHP = UserList(userIndex).Stats.MaxHP
End If

Call DarCuerpoDesnudo(userIndex)
Call ChangeUserChar(userIndex, UserList(userIndex).Char.body, UserList(userIndex).OrigChar.head, UserList(userIndex).Char.Heading, UserList(userIndex).Char.WeaponAnim, UserList(userIndex).Char.ShieldAnim, UserList(userIndex).Char.CascoAnim)
Call WriteUpdateUserStats(userIndex)

End Sub

Sub ChangeUserChar(ByVal userIndex As Integer, Optional ByVal body As Integer = -1, Optional ByVal head As Integer = -1, Optional ByVal Heading As Integer = -1, _
                    Optional ByVal Arma As Integer = -1, Optional ByVal Escudo As Integer = -1, Optional ByVal casco As Integer = -1, Optional ByVal Aura As Integer = -1)

    With UserList(userIndex).Char
        If body > -1 Then _
            .body = body
        If head > -1 Then _
            .head = head
        If Heading > -1 Then _
            .Heading = Heading
        If Arma > -1 Then _
            .WeaponAnim = Arma
        If Escudo > -1 Then _
            .ShieldAnim = Escudo
        If casco > -1 Then _
            .CascoAnim = casco
        If Aura > -1 Then _
            .Aura = Aura
            
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCharacterChange(.body, .head, .Heading, UserList(userIndex).Char.CharIndex, .WeaponAnim, .ShieldAnim, UserList(userIndex).Char.FX, UserList(userIndex).Char.loops, .CascoAnim, .Aura))
    End With
    
End Sub



Sub EraseUserChar(ByVal userIndex As Integer)

On Error GoTo ErrorHandler
   
    CharList(UserList(userIndex).Char.CharIndex) = 0
    
    If UserList(userIndex).Char.CharIndex = LastChar Then
        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1
            If LastChar <= 1 Then Exit Do
        Loop
    End If
    
    'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCharacterRemove(UserList(userIndex).Char.CharIndex))
    Call QuitarUser(userIndex, UserList(userIndex).Pos.Map)
    
    MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).userIndex = 0
    UserList(userIndex).Char.CharIndex = 0
    
    NumChars = NumChars - 1
Exit Sub
    
ErrorHandler:
        Call LogError("Error en EraseUserchar " & Err.Number & ": " & Err.description)
End Sub

Sub RefreshCharStatus(ByVal userIndex As Integer)
'*************************************************
'Author: Tararira
'Last modified: 6/04/2007
'Refreshes the status and tag of UserIndex.
'*************************************************
    Dim klan As String
    If UserList(userIndex).GuildIndex > 0 Then
        klan = modGuilds.GuildName(UserList(userIndex).GuildIndex)
        klan = " <" & klan & ">"
    End If
    
    If UserList(userIndex).showName Then
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageUpdateTagAndStatus(userIndex, Faccion(userIndex), UserList(userIndex).name & klan))
    Else
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageUpdateTagAndStatus(userIndex, Faccion(userIndex), vbNullString))
    End If
End Sub

Sub MakeUserChar(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal userIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

On Error GoTo hayerror
    Dim CharIndex As Integer

    If InMapBounds(Map, X, Y) Then
        'If needed make a new character in list
        If UserList(userIndex).Char.CharIndex = 0 Then
            CharIndex = NextOpenCharIndex
            UserList(userIndex).Char.CharIndex = CharIndex
            CharList(CharIndex) = userIndex
        End If
        
        'Place character on map if needed
        If toMap Then _
            MapData(Map, X, Y).userIndex = userIndex
        
        'Send make character command to clients
        Dim klan As String
        If UserList(userIndex).GuildIndex > 0 Then
            klan = modGuilds.GuildName(UserList(userIndex).GuildIndex)
        End If
        
        Dim bCr As Byte
        
        bCr = Faccion(userIndex)
        
        If LenB(klan) <> 0 Then
            If Not toMap Then
                If UserList(userIndex).showName Then
                    Call WriteCharacterCreate(sndIndex, UserList(userIndex).Char.body, UserList(userIndex).Char.head, UserList(userIndex).Char.Heading, UserList(userIndex).Char.CharIndex, X, Y, UserList(userIndex).Char.WeaponAnim, UserList(userIndex).Char.ShieldAnim, UserList(userIndex).Char.FX, 999, UserList(userIndex).Char.CascoAnim, UserList(userIndex).name & " <" & klan & ">", bCr, UserList(userIndex).flags.Privilegios, UserList(userIndex).Char.Aura)
                Else
                    'Hide the name and clan - set privs as normal user
                    Call WriteCharacterCreate(sndIndex, UserList(userIndex).Char.body, UserList(userIndex).Char.head, UserList(userIndex).Char.Heading, UserList(userIndex).Char.CharIndex, X, Y, UserList(userIndex).Char.WeaponAnim, UserList(userIndex).Char.ShieldAnim, UserList(userIndex).Char.FX, 999, UserList(userIndex).Char.CascoAnim, vbNullString, bCr, PlayerType.User, UserList(userIndex).Char.Aura)
                End If
            Else
                Call AgregarUser(userIndex, UserList(userIndex).Pos.Map)
            End If
        Else 'if tiene clan
            If Not toMap Then
                If UserList(userIndex).showName Then
                    Call WriteCharacterCreate(sndIndex, UserList(userIndex).Char.body, UserList(userIndex).Char.head, UserList(userIndex).Char.Heading, UserList(userIndex).Char.CharIndex, X, Y, UserList(userIndex).Char.WeaponAnim, UserList(userIndex).Char.ShieldAnim, UserList(userIndex).Char.FX, 999, UserList(userIndex).Char.CascoAnim, UserList(userIndex).name, bCr, UserList(userIndex).flags.Privilegios, UserList(userIndex).Char.Aura)
                Else
                    Call WriteCharacterCreate(sndIndex, UserList(userIndex).Char.body, UserList(userIndex).Char.head, UserList(userIndex).Char.Heading, UserList(userIndex).Char.CharIndex, X, Y, UserList(userIndex).Char.WeaponAnim, UserList(userIndex).Char.ShieldAnim, UserList(userIndex).Char.FX, 999, UserList(userIndex).Char.CascoAnim, vbNullString, bCr, PlayerType.User, UserList(userIndex).Char.Aura)
                End If
            Else
                Call AgregarUser(userIndex, UserList(userIndex).Pos.Map)
            End If
        End If 'if clan
    End If
Exit Sub

hayerror:
    LogError ("MakeUserChar: num: " & Err.Number & " desc: " & Err.description)
    'Resume Next
    Call closeConnection(userIndex)
End Sub

Sub CheckUserLevel(ByVal userIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 01/10/2007
'Chequea que el usuario no halla alcanzado el siguiente nivel,
'de lo contrario le da la vida, mana, etc, correspodiente.
'07/08/2006 Integer - Modificacion de los valores
'01/10/2007 Tavo - Corregido el BUG de STAT_MAXELV
'24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones en ELU al subir de nivel.
'24/01/2007 Pablo (ToxicWaste) - Agrego modificaciones de la subida de mana de los magos por lvl.
'13/03/2007 Pablo (ToxicWaste) - Agrego diferencias entre el 18 y el 19 en Constitución.
'*************************************************

On Error GoTo errhandler

Dim Pts As Integer
Dim Constitucion As Integer
Dim AumentoHIT As Integer
Dim AumentoMANA As Integer
Dim AumentoSTA As Integer
Dim AumentoHP As Integer
Dim WasNewbie As Boolean

'¿Alcanzo el maximo nivel?
If UserList(userIndex).Stats.ELV >= STAT_MAXELV Then
    UserList(userIndex).Stats.Exp = 0
    UserList(userIndex).Stats.ELU = 0
    Exit Sub
End If
    
WasNewbie = EsNewbie(userIndex)

Do While UserList(userIndex).Stats.Exp >= UserList(userIndex).Stats.ELU
    
    'Checkea otra vez, esto sucede si tiene mas EXP y puede saltarse el maximo
    'nivel
    If UserList(userIndex).Stats.ELV >= STAT_MAXELV Then
        UserList(userIndex).Stats.Exp = 0
        UserList(userIndex).Stats.ELU = 0
        Exit Sub
    End If
    
    
    'Store it!
    Call Statistics.UserLevelUp(userIndex)
    
    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_NIVEL))
    Call WriteConsoleMsg(userIndex, "¡Has subido de nivel!", FontTypeNames.FONTTYPE_INFO)
    
    If UserList(userIndex).Stats.ELV = 1 Then
        Pts = 10
    Else
        'For multiple levels being rised at once
        Pts = Pts + 5
    End If
    
    UserList(userIndex).Stats.ELV = UserList(userIndex).Stats.ELV + 1
    
    UserList(userIndex).Stats.Exp = UserList(userIndex).Stats.Exp - UserList(userIndex).Stats.ELU
    
    'Nueva subida de exp x lvl. Pablo (ToxicWaste)
    If UserList(userIndex).Stats.ELV < 15 Then
        UserList(userIndex).Stats.ELU = UserList(userIndex).Stats.ELU * 1.4
    ElseIf UserList(userIndex).Stats.ELV < 21 Then
        UserList(userIndex).Stats.ELU = UserList(userIndex).Stats.ELU * 1.35
    ElseIf UserList(userIndex).Stats.ELV < 33 Then
        UserList(userIndex).Stats.ELU = UserList(userIndex).Stats.ELU * 1.3
    ElseIf UserList(userIndex).Stats.ELV < 45 Then
        UserList(userIndex).Stats.ELU = UserList(userIndex).Stats.ELU * 1.225
    'Else
    '    UserList(UserIndex).Stats.ELU = UserList(UserIndex).Stats.ELU * 1.8
    End If
    
    Constitucion = UserList(userIndex).Stats.UserAtributos(eAtributos.Constitucion)
    
    Select Case UserList(userIndex).Clase
        Case eClass.Warrior
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(8, 12)
                Case 20
                    AumentoHP = RandomNumber(8, 11)
                Case 19
                    AumentoHP = RandomNumber(7, 11)
                Case 18
                    AumentoHP = RandomNumber(7, 10)
                Case 17
                    AumentoHP = RandomNumber(6, 10)
                Case 16
                    AumentoHP = RandomNumber(6, 9)
                Case 15
                    AumentoHP = RandomNumber(5, 9)
                Case 14
                    AumentoHP = RandomNumber(5, 8)
                Case 13
                    AumentoHP = RandomNumber(4, 8)
                Case 12
                    AumentoHP = RandomNumber(4, 7)
            End Select
            AumentoHIT = IIf(UserList(userIndex).Stats.ELV > 35, 2, 3)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Hunter
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(8, 11)
                Case 20
                    AumentoHP = RandomNumber(7, 11)
                Case 19
                    AumentoHP = RandomNumber(7, 10)
                Case 18
                    AumentoHP = RandomNumber(6, 10)
                Case 17
                    AumentoHP = RandomNumber(6, 9)
                Case 16
                    AumentoHP = RandomNumber(5, 9)
                Case 15
                    AumentoHP = RandomNumber(5, 8)
                Case 14
                    AumentoHP = RandomNumber(4, 8)
                Case 13
                    AumentoHP = RandomNumber(4, 7)
                Case 12
                    AumentoHP = RandomNumber(3, 7)
            End Select
            AumentoHIT = IIf(UserList(userIndex).Stats.ELV > 35, 2, 3)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Pirat
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(8, 12)
                Case 20
                    AumentoHP = RandomNumber(8, 11)
                Case 19
                    AumentoHP = RandomNumber(7, 11)
                Case 18
                    AumentoHP = RandomNumber(7, 10)
                Case 17
                    AumentoHP = RandomNumber(6, 10)
                Case 16
                    AumentoHP = RandomNumber(6, 9)
                Case 15
                    AumentoHP = RandomNumber(5, 9)
                Case 14
                    AumentoHP = RandomNumber(5, 8)
                Case 13
                    AumentoHP = RandomNumber(4, 8)
                Case 12
                    AumentoHP = RandomNumber(4, 7)
            End Select
            AumentoHIT = 3
            AumentoSTA = AumentoSTDef
        
        Case eClass.Paladin
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(8, 11)
                Case 20
                    AumentoHP = RandomNumber(7, 11)
                Case 19
                    AumentoHP = RandomNumber(7, 10)
                Case 18
                    AumentoHP = RandomNumber(6, 10)
                Case 17
                    AumentoHP = RandomNumber(6, 9)
                Case 16
                    AumentoHP = RandomNumber(5, 9)
                Case 15
                    AumentoHP = RandomNumber(5, 8)
                Case 14
                    AumentoHP = RandomNumber(4, 8)
                Case 13
                    AumentoHP = RandomNumber(4, 7)
                Case 12
                    AumentoHP = RandomNumber(3, 7)
            End Select
            
            AumentoHIT = IIf(UserList(userIndex).Stats.ELV > 35, 1, 3)
            AumentoMANA = UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Thief
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(8, 12)
                Case 20
                    AumentoHP = RandomNumber(8, 11)
                Case 19
                    AumentoHP = RandomNumber(7, 11)
                Case 18
                    AumentoHP = RandomNumber(7, 10)
                Case 17
                    AumentoHP = RandomNumber(6, 10)
                Case 16
                    AumentoHP = RandomNumber(6, 9)
                Case 15
                    AumentoHP = RandomNumber(5, 9)
                Case 14
                    AumentoHP = RandomNumber(5, 8)
                Case 13
                    AumentoHP = RandomNumber(4, 8)
                Case 12
                    AumentoHP = RandomNumber(4, 7)
            End Select
            AumentoHIT = 1
            AumentoSTA = AumentoSTLadron
            
        Case eClass.Mage
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(6, 8)
                Case 20
                    AumentoHP = RandomNumber(5, 8)
                Case 19
                    AumentoHP = RandomNumber(5, 7)
                Case 18
                    AumentoHP = RandomNumber(4, 7)
                Case 17
                    AumentoHP = RandomNumber(4, 6)
                Case 16
                    AumentoHP = RandomNumber(3, 6)
                Case 15
                    AumentoHP = RandomNumber(3, 5)
                Case 14
                    AumentoHP = RandomNumber(2, 5)
                Case 13
                    AumentoHP = RandomNumber(2, 4)
                Case 12
                    AumentoHP = RandomNumber(1, 4)
            End Select
            If AumentoHP < 1 Then AumentoHP = 4
            
            AumentoHIT = 1 'Nueva dist de mana para mago (ToxicWaste)
            AumentoMANA = 2.8 * UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTMago
        
        Case eClass.Lumberjack
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(8, 11)
                Case 20
                    AumentoHP = RandomNumber(7, 11)
                Case 19
                    AumentoHP = RandomNumber(7, 10)
                Case 18
                    AumentoHP = RandomNumber(6, 10)
                Case 17
                    AumentoHP = RandomNumber(6, 9)
                Case 16
                    AumentoHP = RandomNumber(5, 9)
                Case 15
                    AumentoHP = RandomNumber(5, 8)
                Case 14
                    AumentoHP = RandomNumber(4, 8)
                Case 13
                    AumentoHP = RandomNumber(4, 7)
                Case 12
                    AumentoHP = RandomNumber(3, 7)
            End Select
            
            AumentoHIT = 2
            AumentoSTA = AumentoSTLeñador
        
        Case eClass.Miner
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(8, 11)
                Case 20
                    AumentoHP = RandomNumber(7, 11)
                Case 19
                    AumentoHP = RandomNumber(7, 10)
                Case 18
                    AumentoHP = RandomNumber(6, 10)
                Case 17
                    AumentoHP = RandomNumber(6, 9)
                Case 16
                    AumentoHP = RandomNumber(5, 9)
                Case 15
                    AumentoHP = RandomNumber(5, 8)
                Case 14
                    AumentoHP = RandomNumber(4, 8)
                Case 13
                    AumentoHP = RandomNumber(4, 7)
                Case 12
                    AumentoHP = RandomNumber(3, 7)
            End Select
            
            AumentoHIT = 2
            AumentoSTA = AumentoSTMinero
        
        Case eClass.Fisher
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(8, 11)
                Case 20
                    AumentoHP = RandomNumber(7, 11)
                Case 19
                    AumentoHP = RandomNumber(7, 10)
                Case 18
                    AumentoHP = RandomNumber(6, 10)
                Case 17
                    AumentoHP = RandomNumber(6, 9)
                Case 16
                    AumentoHP = RandomNumber(5, 9)
                Case 15
                    AumentoHP = RandomNumber(5, 8)
                Case 14
                    AumentoHP = RandomNumber(4, 8)
                Case 13
                    AumentoHP = RandomNumber(4, 7)
                Case 12
                    AumentoHP = RandomNumber(3, 7)
            End Select
            
            AumentoHIT = 1
            AumentoSTA = AumentoSTPescador
        
        Case eClass.Cleric
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19
                    AumentoHP = RandomNumber(6, 9)
                Case 18
                    AumentoHP = RandomNumber(5, 9)
                Case 17
                    AumentoHP = RandomNumber(5, 8)
                Case 16
                    AumentoHP = RandomNumber(4, 8)
                Case 15
                    AumentoHP = RandomNumber(4, 7)
                Case 14
                    AumentoHP = RandomNumber(3, 7)
                Case 13
                    AumentoHP = RandomNumber(3, 6)
                Case 12
                    AumentoHP = RandomNumber(2, 6)
            End Select
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Druid
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19
                    AumentoHP = RandomNumber(6, 9)
                Case 18
                    AumentoHP = RandomNumber(5, 9)
                Case 17
                    AumentoHP = RandomNumber(5, 8)
                Case 16
                    AumentoHP = RandomNumber(4, 8)
                Case 15
                    AumentoHP = RandomNumber(4, 7)
                Case 14
                    AumentoHP = RandomNumber(3, 7)
                Case 13
                    AumentoHP = RandomNumber(3, 6)
                Case 12
                    AumentoHP = RandomNumber(2, 6)
            End Select
            
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Assasin
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19
                    AumentoHP = RandomNumber(6, 9)
                Case 18
                    AumentoHP = RandomNumber(5, 9)
                Case 17
                    AumentoHP = RandomNumber(5, 8)
                Case 16
                    AumentoHP = RandomNumber(4, 8)
                Case 15
                    AumentoHP = RandomNumber(4, 7)
                Case 14
                    AumentoHP = RandomNumber(3, 7)
                Case 13
                    AumentoHP = RandomNumber(3, 6)
                Case 12
                    AumentoHP = RandomNumber(2, 6)
            End Select
            
            AumentoHIT = IIf(UserList(userIndex).Stats.ELV > 35, 1, 3)
            AumentoMANA = UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Bard
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19
                    AumentoHP = RandomNumber(6, 9)
                Case 18
                    AumentoHP = RandomNumber(5, 9)
                Case 17
                    AumentoHP = RandomNumber(5, 8)
                Case 16
                    AumentoHP = RandomNumber(4, 8)
                Case 15
                    AumentoHP = RandomNumber(4, 7)
                Case 14
                    AumentoHP = RandomNumber(3, 7)
                Case 13
                    AumentoHP = RandomNumber(3, 6)
                Case 12
                    AumentoHP = RandomNumber(2, 6)
            End Select
            
            AumentoHIT = 2
            AumentoMANA = 2 * UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia)
            AumentoSTA = AumentoSTDef
        
        Case eClass.Blacksmith, eClass.Carpenter
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19
                    AumentoHP = RandomNumber(6, 9)
                Case 18
                    AumentoHP = RandomNumber(5, 9)
                Case 17
                    AumentoHP = RandomNumber(5, 8)
                Case 16
                    AumentoHP = RandomNumber(4, 8)
                Case 15
                    AumentoHP = RandomNumber(4, 7)
                Case 14
                    AumentoHP = RandomNumber(3, 7)
                Case 13
                    AumentoHP = RandomNumber(3, 6)
                Case 12
                    AumentoHP = RandomNumber(2, 6)
            End Select
            AumentoHIT = 2
            AumentoSTA = AumentoSTDef
            
        Case eClass.Bandit
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(7, 10)
                Case 20
                    AumentoHP = RandomNumber(6, 10)
                Case 19
                    AumentoHP = RandomNumber(6, 9)
                Case 18
                    AumentoHP = RandomNumber(5, 9)
                Case 17
                    AumentoHP = RandomNumber(5, 8)
                Case 16
                    AumentoHP = RandomNumber(4, 8)
                Case 15
                    AumentoHP = RandomNumber(4, 7)
                Case 14
                    AumentoHP = RandomNumber(3, 7)
                Case 13
                    AumentoHP = RandomNumber(3, 6)
                Case 12
                    AumentoHP = RandomNumber(2, 6)
            End Select
            
            AumentoHIT = IIf(UserList(userIndex).Stats.ELV > 35, 1, 3)
            AumentoMANA = IIf(UserList(userIndex).Stats.MaxMAN = 300, 0, UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia) - 10)
            If AumentoMANA < 4 Then AumentoMANA = 4
            AumentoSTA = AumentoSTLeñador
        Case Else
            Select Case Constitucion
                Case 21
                    AumentoHP = RandomNumber(6, 9)
                Case 20
                    AumentoHP = RandomNumber(5, 9)
                Case 19, 18
                    AumentoHP = RandomNumber(4, 8)
                Case Else
                    AumentoHP = RandomNumber(5, Constitucion \ 2) - AdicionalHPCazador
            End Select
            
            AumentoHIT = 2
            AumentoSTA = AumentoSTDef
    End Select
    
    'Actualizamos HitPoints
    UserList(userIndex).Stats.MaxHP = UserList(userIndex).Stats.MaxHP + AumentoHP
    If UserList(userIndex).Stats.MaxHP > STAT_MAXHP Then _
        UserList(userIndex).Stats.MaxHP = STAT_MAXHP
    'Actualizamos Stamina
    UserList(userIndex).Stats.MaxSta = UserList(userIndex).Stats.MaxSta + AumentoSTA
    If UserList(userIndex).Stats.MaxSta > STAT_MAXSTA Then _
        UserList(userIndex).Stats.MaxSta = STAT_MAXSTA
    'Actualizamos Mana
    UserList(userIndex).Stats.MaxMAN = UserList(userIndex).Stats.MaxMAN + AumentoMANA
    If UserList(userIndex).Stats.ELV < 36 Then
        If UserList(userIndex).Stats.MaxMAN > STAT_MAXMAN Then _
            UserList(userIndex).Stats.MaxMAN = STAT_MAXMAN
    Else
        If UserList(userIndex).Stats.MaxMAN > 9999 Then _
            UserList(userIndex).Stats.MaxMAN = 9999
    End If
    If UserList(userIndex).Clase = eClass.Bandit Then 'mana del bandido restringido hasta 300
        If UserList(userIndex).Stats.MaxMAN > 300 Then
            UserList(userIndex).Stats.MaxMAN = 300
        End If
    End If
    
    'Actualizamos Golpe Máximo
    UserList(userIndex).Stats.MaxHIT = UserList(userIndex).Stats.MaxHIT + AumentoHIT
    If UserList(userIndex).Stats.ELV < 36 Then
        If UserList(userIndex).Stats.MaxHIT > STAT_MAXHIT_UNDER36 Then _
            UserList(userIndex).Stats.MaxHIT = STAT_MAXHIT_UNDER36
    Else
        If UserList(userIndex).Stats.MaxHIT > STAT_MAXHIT_OVER36 Then _
            UserList(userIndex).Stats.MaxHIT = STAT_MAXHIT_OVER36
    End If
    
    'Actualizamos Golpe Mínimo
    UserList(userIndex).Stats.MinHIT = UserList(userIndex).Stats.MinHIT + AumentoHIT
    If UserList(userIndex).Stats.ELV < 36 Then
        If UserList(userIndex).Stats.MinHIT > STAT_MAXHIT_UNDER36 Then _
            UserList(userIndex).Stats.MinHIT = STAT_MAXHIT_UNDER36
    Else
        If UserList(userIndex).Stats.MinHIT > STAT_MAXHIT_OVER36 Then _
            UserList(userIndex).Stats.MinHIT = STAT_MAXHIT_OVER36
    End If
    
    'Notificamos al user
    If AumentoHP > 0 Then
        Call WriteConsoleMsg(userIndex, "Has ganado " & AumentoHP & " puntos de vida.", FontTypeNames.FONTTYPE_INFO)
    End If
    If AumentoSTA > 0 Then
        Call WriteConsoleMsg(userIndex, "Has ganado " & AumentoSTA & " puntos de vitalidad.", FontTypeNames.FONTTYPE_INFO)
    End If
    If AumentoMANA > 0 Then
        Call WriteConsoleMsg(userIndex, "Has ganado " & AumentoMANA & " puntos de magia.", FontTypeNames.FONTTYPE_INFO)
    End If
    If AumentoHIT > 0 Then
        Call WriteConsoleMsg(userIndex, "Tu golpe maximo aumento en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(userIndex, "Tu golpe minimo aumento en " & AumentoHIT & " puntos.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    Call LogDesarrollo(UserList(userIndex).name & " paso a nivel " & UserList(userIndex).Stats.ELV & " gano HP: " & AumentoHP)
    
    UserList(userIndex).Stats.MinHP = UserList(userIndex).Stats.MaxHP
Loop

'If it ceased to be a newbie, remove newbie items and get char away from newbie dungeon
If Not EsNewbie(userIndex) And WasNewbie Then
    Call QuitarNewbieObj(userIndex)
    If UCase$(MapInfo(UserList(userIndex).Pos.Map).Restringir) = "NEWBIE" Then
        Call WarpUserChar(userIndex, 26, 50, 50, True)
        Call WriteConsoleMsg(userIndex, "Debes abandonar el Dungeon Newbie.", FontTypeNames.FONTTYPE_INFO)
    End If
End If

'Send all gained skill points at once (if any)
If Pts > 0 Then
    Call WriteLevelUp(userIndex, Pts)
    
    UserList(userIndex).Stats.SkillPts = UserList(userIndex).Stats.SkillPts + Pts
    
    Call WriteConsoleMsg(userIndex, "Has ganado un total de " & Pts & " skillpoints.", FontTypeNames.FONTTYPE_INFO)
End If

Call WriteUpdateUserStats(userIndex)

Exit Sub

errhandler:
    Call LogError("Error en la subrutina CheckUserLevel - Error : " & Err.Number & " - Description : " & Err.description)
End Sub

Function PuedeAtravesarAgua(ByVal userIndex As Integer) As Boolean

PuedeAtravesarAgua = _
  UserList(userIndex).flags.Navegando = 1 Or _
  UserList(userIndex).flags.Vuela = 1

End Function

Sub MoveUserChar(ByVal userIndex As Integer, ByVal nHeading As eHeading)

Dim nPos As WorldPos
    
    nPos = UserList(userIndex).Pos
    Call HeadtoPos(nHeading, nPos)
    
    If LegalPos(UserList(userIndex).Pos.Map, nPos.X, nPos.Y, PuedeAtravesarAgua(userIndex)) Then
        If MapInfo(UserList(userIndex).Pos.Map).NumUsers > 1 Then
            'si no estoy solo en el mapa...

            Call SendData(SendTarget.ToPCAreaButIndex, userIndex, PrepareMessageCharacterMove(UserList(userIndex).Char.CharIndex, nPos.X, nPos.Y))

        End If
        
        'Update map and user pos
        MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).userIndex = 0
        UserList(userIndex).Pos = nPos
        UserList(userIndex).Char.Heading = nHeading
        MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).userIndex = userIndex
        
        'Actualizamos las áreas de ser necesario
        Call ModAreas.CheckUpdateNeededUser(userIndex, nHeading)
    Else
        Call WritePosUpdate(userIndex)
    End If
    
    If UserList(userIndex).Counters.Trabajando Then _
        UserList(userIndex).Counters.Trabajando = UserList(userIndex).Counters.Trabajando - 1

    If UserList(userIndex).Counters.Ocultando Then _
        UserList(userIndex).Counters.Ocultando = UserList(userIndex).Counters.Ocultando - 1
End Sub

Sub ChangeUserInv(ByVal userIndex As Integer, ByVal Slot As Byte, ByRef Object As UserOBJ)
    UserList(userIndex).Invent.Object(Slot) = Object
    Call WriteChangeInventorySlot(userIndex, Slot)
End Sub

Function NextOpenCharIndex() As Integer
    Dim LoopC As Long
    
    For LoopC = 1 To MAXCHARS
        If CharList(LoopC) = 0 Then
            NextOpenCharIndex = LoopC
            NumChars = NumChars + 1
            
            If LoopC > LastChar Then _
                LastChar = LoopC
            
            Exit Function
        End If
    Next LoopC
End Function

Function NextOpenUser() As Integer
    Dim LoopC As Long
    
    For LoopC = 1 To MaxUsers + 1
        If LoopC > MaxUsers Then Exit For
        If (UserList(LoopC).ConnID = -1 And UserList(LoopC).UserAccount.Logged = False And UserList(LoopC).flags.UserLogged = False) Then Exit For
    Next LoopC
    
    NextOpenUser = LoopC
End Function

Sub SendUserStatsTxt(ByVal sendIndex As Integer, ByVal userIndex As Integer)
Dim GuildI As Integer


    Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & UserList(userIndex).name, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Nivel: " & UserList(userIndex).Stats.ELV & "  EXP: " & UserList(userIndex).Stats.Exp & "/" & UserList(userIndex).Stats.ELU, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Salud: " & UserList(userIndex).Stats.MinHP & "/" & UserList(userIndex).Stats.MaxHP & "  Mana: " & UserList(userIndex).Stats.MinMAN & "/" & UserList(userIndex).Stats.MaxMAN & "  Vitalidad: " & UserList(userIndex).Stats.MinSta & "/" & UserList(userIndex).Stats.MaxSta, FontTypeNames.FONTTYPE_INFO)
    
    If UserList(userIndex).Invent.WeaponEqpObjIndex > 0 Then
        Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & UserList(userIndex).Stats.MinHIT & "/" & UserList(userIndex).Stats.MaxHIT & " (" & ObjData(UserList(userIndex).Invent.WeaponEqpObjIndex).MinHIT & "/" & ObjData(UserList(userIndex).Invent.WeaponEqpObjIndex).MaxHIT & ")", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & UserList(userIndex).Stats.MinHIT & "/" & UserList(userIndex).Stats.MaxHIT, FontTypeNames.FONTTYPE_INFO)
    End If
    
    If UserList(userIndex).Invent.ArmourEqpObjIndex > 0 Then
        If UserList(userIndex).Invent.EscudoEqpObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(UserList(userIndex).Invent.ArmourEqpObjIndex).MinDef + ObjData(UserList(userIndex).Invent.EscudoEqpObjIndex).MinDef & "/" & ObjData(UserList(userIndex).Invent.ArmourEqpObjIndex).MaxDef + ObjData(UserList(userIndex).Invent.EscudoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: " & ObjData(UserList(userIndex).Invent.ArmourEqpObjIndex).MinDef & "/" & ObjData(UserList(userIndex).Invent.ArmourEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(sendIndex, "(CUERPO) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)
    End If
    
    If UserList(userIndex).Invent.CascoEqpObjIndex > 0 Then
        Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: " & ObjData(UserList(userIndex).Invent.CascoEqpObjIndex).MinDef & "/" & ObjData(UserList(userIndex).Invent.CascoEqpObjIndex).MaxDef, FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "(CABEZA) Min Def/Max Def: 0", FontTypeNames.FONTTYPE_INFO)
    End If
    
    GuildI = UserList(userIndex).GuildIndex
    If GuildI > 0 Then
        Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(GuildI), FontTypeNames.FONTTYPE_INFO)
        If UCase$(modGuilds.GuildLeader(GuildI)) = UCase$(UserList(sendIndex).name) Then
            Call WriteConsoleMsg(sendIndex, "Status: Lider", FontTypeNames.FONTTYPE_INFO)
        End If
        'guildpts no tienen objeto
    End If
    
    #If ConUpTime Then
        Dim TempDate As Date
        Dim TempSecs As Long
        Dim TempStr As String
        TempDate = Now - UserList(userIndex).LogOnTime
        TempSecs = (UserList(userIndex).UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + (Hour(TempDate) * 3600) + (Minute(TempDate) * 60) + Second(TempDate))
        TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
        Call WriteConsoleMsg(sendIndex, "Logeado hace: " & Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Total: " & TempStr, FontTypeNames.FONTTYPE_INFO)
    #End If
    
    Call WriteConsoleMsg(sendIndex, "Oro: " & UserList(userIndex).Stats.GLD & "  Posicion: " & UserList(userIndex).Pos.X & "," & UserList(userIndex).Pos.Y & " en mapa " & UserList(userIndex).Pos.Map, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Dados: " & UserList(userIndex).Stats.UserAtributos(eAtributos.Fuerza) & ", " & UserList(userIndex).Stats.UserAtributos(eAtributos.Agilidad) & ", " & UserList(userIndex).Stats.UserAtributos(eAtributos.Inteligencia) & ", " & UserList(userIndex).Stats.UserAtributos(eAtributos.Carisma) & ", " & UserList(userIndex).Stats.UserAtributos(eAtributos.Constitucion), FontTypeNames.FONTTYPE_INFO)
  
End Sub

Sub SendUserMiniStatsTxt(ByVal sendIndex As Integer, ByVal userIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Shows the users Stats when the user is online.
'23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
'*************************************************
With UserList(userIndex)
    Call WriteConsoleMsg(sendIndex, "Pj: " & .name, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "CiudadanosMatados: " & .Faccion.CiudadanosMatados & " CriminalesMatados: " & .Faccion.CriminalesMatados & " NeutralesMatados: " & .Faccion.NeutralesMatados, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "NPCsMuertos: " & .Stats.NPCsMuertos, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(.Clase), FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Pena: " & .Counters.Pena, FontTypeNames.FONTTYPE_INFO)
    
    If .GuildIndex > 0 Then
        Call WriteConsoleMsg(sendIndex, "Clan: " & GuildName(.GuildIndex), FontTypeNames.FONTTYPE_INFO)
    End If
    
End With

End Sub

Sub SendUserMiniStatsTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Shows the users Stats when the user is offline.
'23/01/2007 Pablo (ToxicWaste) - Agrego de funciones y mejora de distribución de parámetros.
'*************************************************
Dim CharFile As String
Dim Ban As String
Dim BanDetailPath As String

    BanDetailPath = App.Path & "\logs\" & "BanDetail.dat"
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile) Then
        Call WriteConsoleMsg(sendIndex, "Pj: " & charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "CiudadanosMatados: " & GetVar(CharFile, "FACCIONES", "CiudMatados") & " CriminalesMatados: " & GetVar(CharFile, "FACCIONES", "CrimMatados") & " UsuariosMatados: " & GetVar(CharFile, "MUERTES", "UserMuertes"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCsMuertos: " & GetVar(CharFile, "MUERTES", "NpcsMuertes"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(GetVar(CharFile, "INIT", "Clase")), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Pena: " & GetVar(CharFile, "COUNTERS", "PENA"), FontTypeNames.FONTTYPE_INFO)

        If IsNumeric(GetVar(CharFile, "Guild", "GUILDINDEX")) Then
            Call WriteConsoleMsg(sendIndex, "Clan: " & modGuilds.GuildName(CInt(GetVar(CharFile, "Guild", "GUILDINDEX"))), FontTypeNames.FONTTYPE_INFO)
        End If
        
        Ban = GetVar(CharFile, "FLAGS", "Ban")
        Call WriteConsoleMsg(sendIndex, "Ban: " & Ban, FontTypeNames.FONTTYPE_INFO)
        If Ban = "1" Then
            Call WriteConsoleMsg(sendIndex, "Ban por: " & GetVar(CharFile, charName, "BannedBy") & " Motivo: " & GetVar(BanDetailPath, charName, "Reason"), FontTypeNames.FONTTYPE_INFO)
        End If
    Else
        Call WriteConsoleMsg(sendIndex, "El pj no existe: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Sub SendUserInvTxt(ByVal sendIndex As Integer, ByVal userIndex As Integer)
On Error Resume Next

    Dim j As Long
    
    
    Call WriteConsoleMsg(sendIndex, UserList(userIndex).name, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Tiene " & UserList(userIndex).Invent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)
    
    For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(userIndex).Invent.Object(j).ObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(UserList(userIndex).Invent.Object(j).ObjIndex).name & " Cantidad:" & UserList(userIndex).Invent.Object(j).amount, FontTypeNames.FONTTYPE_INFO)
        End If
    Next j
End Sub

Sub SendUserInvTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
On Error Resume Next

    Dim j As Long
    Dim CharFile As String, Tmp As String
    Dim ObjInd As Long, ObjCant As Long
    
    CharFile = CharPath & charName & ".chr"
    
    If FileExist(CharFile, vbNormal) Then
        Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, " Tiene " & GetVar(CharFile, "Inventory", "CantidadItems") & " objetos.", FontTypeNames.FONTTYPE_INFO)
        
        For j = 1 To MAX_INVENTORY_SLOTS
            Tmp = GetVar(CharFile, "Inventory", "Obj" & j)
            ObjInd = ReadField(1, Tmp, Asc("-"))
            ObjCant = ReadField(2, Tmp, Asc("-"))
            If ObjInd > 0 Then
                Call WriteConsoleMsg(sendIndex, " Objeto " & j & " " & ObjData(ObjInd).name & " Cantidad:" & ObjCant, FontTypeNames.FONTTYPE_INFO)
            End If
        Next j
    Else
        Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
    End If
    
End Sub

Sub SendUserSkillsTxt(ByVal sendIndex As Integer, ByVal userIndex As Integer)
On Error Resume Next
Dim j As Integer
Call WriteConsoleMsg(sendIndex, UserList(userIndex).name, FontTypeNames.FONTTYPE_INFO)
For j = 1 To NUMSKILLS
    Call WriteConsoleMsg(sendIndex, SkillsNames(j) & " = " & UserList(userIndex).Stats.UserSkills(j), FontTypeNames.FONTTYPE_INFO)
Next
Call WriteConsoleMsg(sendIndex, " SkillLibres:" & UserList(userIndex).Stats.SkillPts, FontTypeNames.FONTTYPE_INFO)
End Sub

Function DameUserIndex(SocketId As Integer) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
Do Until UserList(LoopC).ConnID = SocketId

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        DameUserIndex = 0
        Exit Function
    End If
    
Loop
  
DameUserIndex = LoopC

End Function

Function DameUserIndexConNombre(ByVal Nombre As String) As Integer

Dim LoopC As Integer
  
LoopC = 1
  
Nombre = UCase$(Nombre)

Do Until UCase$(UserList(LoopC).name) = Nombre

    LoopC = LoopC + 1
    
    If LoopC > MaxUsers Then
        DameUserIndexConNombre = 0
        Exit Function
    End If
    
Loop
  
DameUserIndexConNombre = LoopC

End Function


Function EsMascota(ByVal NpcIndex As Integer, ByVal userIndex As Integer) As Boolean
    If Npclist(NpcIndex).MaestroUser > 0 Then
        EsMascota = UserList(Npclist(NpcIndex).MaestroUser).Faccion.Alineacion = e_Alineacion.Neutro Or Not (UserList(Npclist(NpcIndex).MaestroUser).Faccion.Alineacion = UserList(userIndex).Faccion.Alineacion)
        If EsMascota Then
            Call WriteConsoleMsg(Npclist(NpcIndex).MaestroUser, "¡¡" & UserList(userIndex).name & " esta atacando tu mascota!!", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
End Function

Sub npcAttacked(ByVal NpcIndex As Integer, ByVal userIndex As Integer)
'**********************************************
'Author: Unknown
'Last Modification: 24/07/2007
'24/01/2007 -> Pablo (ToxicWaste): Agrego para que se actualize el tag si corresponde.
'24/07/2007 -> Pablo (ToxicWaste): Guardar primero que ataca NPC y el que atacas ahora.
'**********************************************

'Guardamos el usuario que ataco el npc.
Npclist(NpcIndex).flags.AttackedBy = UserList(userIndex).name

'Npc que estabas atacando.
Dim LastNpcHit As Integer
LastNpcHit = UserList(userIndex).flags.npcAttacked
'Guarda el NPC que estas atacando ahora.
UserList(userIndex).flags.npcAttacked = NpcIndex

'Revisamos robo de npc.
'Guarda el primer nick que lo ataca.
If Npclist(NpcIndex).flags.AttackedFirstBy = vbNullString Then
    'El que le pegabas antes ya no es tuyo
    If LastNpcHit <> 0 Then
        If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(userIndex).name Then
            Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
        End If
    End If
    Npclist(NpcIndex).flags.AttackedFirstBy = UserList(userIndex).name
ElseIf Npclist(NpcIndex).flags.AttackedFirstBy <> UserList(userIndex).name Then
    'Estas robando NPC
    'El que le pegabas antes ya no es tuyo
    If LastNpcHit <> 0 Then
        If Npclist(LastNpcHit).flags.AttackedFirstBy = UserList(userIndex).name Then
            Npclist(LastNpcHit).flags.AttackedFirstBy = vbNullString
        End If
    End If
End If

If Npclist(NpcIndex).MaestroUser > 0 Then Call AllMascotasAtacanUser(userIndex, Npclist(NpcIndex).MaestroUser)

If EsMascota(NpcIndex, userIndex) Then
    Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
    Npclist(NpcIndex).Hostile = 1
Else
    'hacemos que el npc se defienda
    Npclist(NpcIndex).Movement = TipoAI.NPCDEFENSA
    Npclist(NpcIndex).Hostile = 1
End If

End Sub

Function PuedeApuñalar(ByVal userIndex As Integer) As Boolean

If UserList(userIndex).Invent.WeaponEqpObjIndex > 0 Then
 PuedeApuñalar = _
 ((UserList(userIndex).Stats.UserSkills(eSkill.Apuñalar) >= MIN_APUÑALAR) _
 And (ObjData(UserList(userIndex).Invent.WeaponEqpObjIndex).Apuñala = 1)) _
 Or _
  ((UserList(userIndex).Clase = eClass.Assasin) And _
  (ObjData(UserList(userIndex).Invent.WeaponEqpObjIndex).Apuñala = 1))
Else
 PuedeApuñalar = False
End If
End Function

Sub SubirSkill(ByVal userIndex As Integer, ByVal Skill As Integer)

    If UserList(userIndex).flags.Hambre = 0 And UserList(userIndex).flags.Sed = 0 Then
        
        If UserList(userIndex).Stats.UserSkills(Skill) = MAXSKILLPOINTS Then Exit Sub
        
        Dim Lvl As Integer
        Lvl = UserList(userIndex).Stats.ELV
        
        If Lvl > UBound(LevelSkill) Then Lvl = UBound(LevelSkill)
        
        If UserList(userIndex).Stats.UserSkills(Skill) >= LevelSkill(Lvl).LevelValue Then Exit Sub
    
        Dim Aumenta As Integer
        Dim Prob As Integer
        
        If Lvl <= 3 Then
            Prob = 6
        ElseIf Lvl > 3 And Lvl < 6 Then
            Prob = 7
        ElseIf Lvl >= 6 And Lvl < 10 Then
            Prob = 8
        ElseIf Lvl >= 10 And Lvl < 20 Then
            Prob = 9
        Else
            Prob = 10
        End If
        
        Aumenta = RandomNumber(5, Prob)
        
        If Aumenta = 7 Then
            UserList(userIndex).Stats.UserSkills(Skill) = UserList(userIndex).Stats.UserSkills(Skill) + 1
            Call WriteConsoleMsg(userIndex, "¡Has mejorado tu skill " & SkillsNames(Skill) & " en un punto!. Ahora tienes " & UserList(userIndex).Stats.UserSkills(Skill) & " pts.", FontTypeNames.FONTTYPE_INFO)
            
            UserList(userIndex).Stats.Exp = UserList(userIndex).Stats.Exp + 50
            If UserList(userIndex).Stats.Exp > MAXEXP Then _
                UserList(userIndex).Stats.Exp = MAXEXP
            
            Call WriteConsoleMsg(userIndex, "¡Has ganado 50 puntos de experiencia!", FontTypeNames.FONTTYPE_FIGHT)
            
            Call WriteUpdateExp(userIndex)
            Call CheckUserLevel(userIndex)
        End If
    End If

End Sub

Sub UserDie(ByVal userIndex As Integer)
On Error GoTo ErrorHandler

    'Sonido
    If UserList(userIndex).Genero = eGenero.Mujer Then
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userIndex, e_SoundIndex.MUERTE_MUJER)
    Else
        Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, userIndex, e_SoundIndex.MUERTE_HOMBRE)
    End If
    
    'Quitar el dialogo del user muerto
    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageRemoveCharDialog(UserList(userIndex).Char.CharIndex))
    
    UserList(userIndex).Stats.MinHP = 0
    UserList(userIndex).Stats.MinSta = 0
    UserList(userIndex).flags.AtacadoPorUser = 0
    UserList(userIndex).flags.Envenenado = 0
    UserList(userIndex).flags.Muerto = 1
    
    
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
    
    '<<<< Paralisis >>>>
    If UserList(userIndex).flags.Paralizado = 1 Then
        UserList(userIndex).flags.Paralizado = 0
        Call WriteParalizeOK(userIndex)
    End If
    
    '<<< Estupidez >>>
    If UserList(userIndex).flags.Estupidez = 1 Then
        UserList(userIndex).flags.Estupidez = 0
        Call WriteDumbNoMore(userIndex)
    End If
    
    '<<<< Descansando >>>>
    If UserList(userIndex).flags.Descansar Then
        UserList(userIndex).flags.Descansar = False
        Call WriteRestOK(userIndex)
    End If
    
    '<<<< Meditando >>>>
    If UserList(userIndex).flags.Meditando Then
        UserList(userIndex).flags.Meditando = False
        Call WriteMeditateToggle(userIndex)
    End If
    
    '<<<< Invisible >>>>
    If UserList(userIndex).flags.invisible = 1 Or UserList(userIndex).flags.Oculto = 1 Then
        UserList(userIndex).flags.Oculto = 0
        UserList(userIndex).Counters.TiempoOculto = 0
        UserList(userIndex).flags.invisible = 0
        'no hace falta encriptar este NOVER
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageSetInvisible(UserList(userIndex).Char.CharIndex, False))
    End If
    
    If TriggerZonaPelea(userIndex, userIndex) <> eTrigger6.TRIGGER6_PERMITE Then
        ' << Si es newbie no pierde el inventario >>
        If Not EsNewbie(userIndex) Then
            Call TirarTodo(userIndex)
        Else
             Call TirarTodosLosItemsNoNewbies(userIndex)
        End If
    End If
    
    ' DESEQUIPA TODOS LOS OBJETOS
    'desequipar armadura
    If UserList(userIndex).Invent.ArmourEqpObjIndex > 0 Then
        Call Desequipar(userIndex, UserList(userIndex).Invent.ArmourEqpSlot)
    End If
    'desequipar arma
    If UserList(userIndex).Invent.WeaponEqpObjIndex > 0 Then
        Call Desequipar(userIndex, UserList(userIndex).Invent.WeaponEqpSlot)
    End If
    'desequipar casco
    If UserList(userIndex).Invent.CascoEqpObjIndex > 0 Then
        Call Desequipar(userIndex, UserList(userIndex).Invent.CascoEqpSlot)
    End If
    'desequipar herramienta
    If UserList(userIndex).Invent.AnilloEqpSlot > 0 Then
        Call Desequipar(userIndex, UserList(userIndex).Invent.AnilloEqpSlot)
    End If
    'desequipar municiones
    If UserList(userIndex).Invent.MunicionEqpObjIndex > 0 Then
        Call Desequipar(userIndex, UserList(userIndex).Invent.MunicionEqpSlot)
    End If
    'desequipar escudo
    If UserList(userIndex).Invent.EscudoEqpObjIndex > 0 Then
        Call Desequipar(userIndex, UserList(userIndex).Invent.EscudoEqpSlot)
    End If
    
    If UserList(userIndex).flags.Montado = 1 Then
        Call DoMontar(userIndex, ObjData(UserList(userIndex).Invent.MonturaObjIndex), UserList(userIndex).Invent.MonturaSlot)
    End If
    
    ' << Reseteamos los posibles FX sobre el personaje >>
    If UserList(userIndex).Char.loops = LoopAdEternum Then
        UserList(userIndex).Char.FX = 0
        UserList(userIndex).Char.loops = 0
    End If
    
    ' << Restauramos el mimetismo
    If UserList(userIndex).flags.Mimetizado = 1 Then
        UserList(userIndex).Char.body = UserList(userIndex).CharMimetizado.body
        UserList(userIndex).Char.head = UserList(userIndex).CharMimetizado.head
        UserList(userIndex).Char.CascoAnim = UserList(userIndex).CharMimetizado.CascoAnim
        UserList(userIndex).Char.ShieldAnim = UserList(userIndex).CharMimetizado.ShieldAnim
        UserList(userIndex).Char.WeaponAnim = UserList(userIndex).CharMimetizado.WeaponAnim
        UserList(userIndex).Counters.Mimetismo = 0
        UserList(userIndex).flags.Mimetizado = 0
    End If
    
    '<< Cambiamos la apariencia del char >>
    If UserList(userIndex).flags.Navegando = 0 Then
            UserList(userIndex).Char.body = iCuerpoMuerto
            UserList(userIndex).Char.head = iCabezaMuerto
            UserList(userIndex).Char.ShieldAnim = NingunEscudo
            UserList(userIndex).Char.WeaponAnim = NingunArma
            UserList(userIndex).Char.CascoAnim = NingunCasco
    Else
        UserList(userIndex).Char.body = iFragataFantasmal ';)
    End If
    
    
    
    Dim i As Integer
    For i = 1 To MAXMASCOTAS
        
        If UserList(userIndex).MascotasIndex(i) > 0 Then
               If Npclist(UserList(userIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                    Call MuereNpc(UserList(userIndex).MascotasIndex(i), 0)
               Else
                    Npclist(UserList(userIndex).MascotasIndex(i)).MaestroUser = 0
                    Npclist(UserList(userIndex).MascotasIndex(i)).Movement = Npclist(UserList(userIndex).MascotasIndex(i)).flags.OldMovement
                    Npclist(UserList(userIndex).MascotasIndex(i)).Hostile = Npclist(UserList(userIndex).MascotasIndex(i)).flags.OldHostil
                    UserList(userIndex).MascotasIndex(i) = 0
                    UserList(userIndex).MascotasType(i) = 0
               End If
        End If
        
    Next i
    
    UserList(userIndex).NroMacotas = 0
    
    'Nos fijamos si esta en duelo,etc...
    If UserList(userIndex).flags.EnDuelo = 1 Then
        Call WarpUserChar(userIndex, 26, 50, 50, True)
        Call WriteConsoleMsg(userIndex, "Has perdido el duelo.", FontTypeNames.FONTTYPE_INFO)
        UserList(userIndex).flags.EnDuelo = 0
    End If
    
    If UserList(userIndex).Pos.Map = MAPATORNEO Then
        'Call WarpUserChar(UserIndex, 1, 50, 50, True)
        Call WriteConsoleMsg(userIndex, "Has sido eliminado del torneo. :(", FontTypeNames.FONTTYPE_GUILD)
        Call ColaTorneo.Quitar(UserList(userIndex).name)
    End If

    'If MapInfo(UserList(UserIndex).Pos.Map).Pk Then
    '        Dim MiObj As Obj
    '        Dim nPos As WorldPos
    '        MiObj.ObjIndex = RandomNumber(554, 555)
    '        MiObj.Amount = 1
    '        nPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    '        Dim ManchaSangre As New cGarbage
    '        ManchaSangre.Map = nPos.Map
    '        ManchaSangre.X = nPos.X
    '        ManchaSangre.Y = nPos.Y
    '        Call TrashCollector.Add(ManchaSangre)
    'End If
    
    '<< Actualizamos clientes >>
    Call ChangeUserChar(userIndex, UserList(userIndex).Char.body, UserList(userIndex).Char.head, UserList(userIndex).Char.Heading, NingunArma, NingunEscudo, NingunCasco)
    Call WriteUpdateUserStats(userIndex)
    
Exit Sub

ErrorHandler:
    Call LogError("Error en SUB USERDIE. Error: " & Err.Number & " Descripción: " & Err.description)
End Sub


Sub ContarMuerte(ByVal Muerto As Integer, ByVal Atacante As Integer)

    If EsNewbie(Muerto) Then Exit Sub
    
    If UserList(Muerto).Pos.Map = MAPATORNEO Then Exit Sub
    If TriggerZonaPelea(Muerto, Atacante) = TRIGGER6_PERMITE Then Exit Sub
    
    'CONTAR MUERTE BLIZZARD
End Sub

Sub Tilelibre(ByRef Pos As WorldPos, ByRef nPos As WorldPos, ByRef Obj As Obj, ByRef Agua As Boolean, ByRef Tierra As Boolean)
'**************************************************************
'Author: Unknown
'Last Modify Date: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): El agua es ahora un TileLibre agregando las condiciones necesarias.
'**************************************************************
Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer
Dim hayobj As Boolean
    hayobj = False
    nPos.Map = Pos.Map
    
    Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y, Agua, Tierra) Or hayobj
        
        If LoopC > 15 Then
            Notfound = True
            Exit Do
        End If
        
        For tY = Pos.Y - LoopC To Pos.Y + LoopC
            For tX = Pos.X - LoopC To Pos.X + LoopC
            
                If LegalPos(nPos.Map, tX, tY, Agua, Tierra) Then
                    'We continue if: a - the item is different from 0 and the dropped item or b - the amount dropped + amount in map exceeds MAX_INVENTORY_OBJS
                    hayobj = (MapData(nPos.Map, tX, tY).ObjInfo.ObjIndex > 0 And MapData(nPos.Map, tX, tY).ObjInfo.ObjIndex <> Obj.ObjIndex)
                    If Not hayobj Then _
                        hayobj = (MapData(nPos.Map, tX, tY).ObjInfo.amount + Obj.amount > MAX_INVENTORY_OBJS)
                    If Not hayobj And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                        nPos.X = tX
                        nPos.Y = tY
                        tX = Pos.X + LoopC
                        tY = Pos.Y + LoopC
                    End If
                End If
            
            Next tX
        Next tY
        
        LoopC = LoopC + 1
        
    Loop
    
    If Notfound = True Then
        nPos.X = 0
        nPos.Y = 0
    End If

End Sub

Sub WarpUserChar(ByVal userIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal FX As Boolean = False)
    Dim OldMap As Integer
    Dim OldX As Integer
    Dim OldY As Integer
    
    'Quitar el dialogo
    Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageRemoveCharDialog(UserList(userIndex).Char.CharIndex))
    
    Call WriteRemoveAllDialogs(userIndex)
    
    OldMap = UserList(userIndex).Pos.Map
    OldX = UserList(userIndex).Pos.X
    OldY = UserList(userIndex).Pos.Y
    
    Call EraseUserChar(userIndex)
    
    If OldMap <> Map Then
        Call WriteChangeMap(userIndex, Map, MapInfo(UserList(userIndex).Pos.Map).MapVersion)
        Call WritePlayMidi(userIndex, val(ReadField(1, MapInfo(Map).Music, 45)))
        
        'Update new Map Users
        MapInfo(Map).NumUsers = MapInfo(Map).NumUsers + 1
        
        'Update old Map Users
        MapInfo(OldMap).NumUsers = MapInfo(OldMap).NumUsers - 1
        If MapInfo(OldMap).NumUsers < 0 Then
            MapInfo(OldMap).NumUsers = 0
        End If
    End If
    
    UserList(userIndex).Pos.X = X
    UserList(userIndex).Pos.Y = Y
    UserList(userIndex).Pos.Map = Map
    
    Call MakeUserChar(True, Map, userIndex, Map, X, Y)
    Call WriteUserCharIndexInServer(userIndex)
    
    'Force a flush, so user index is in there before it's destroyed for teleporting
    Call FlushBuffer(userIndex)
    
    'Seguis invisible al pasar de mapa
    If (UserList(userIndex).flags.invisible = 1 Or UserList(userIndex).flags.Oculto = 1) And (Not UserList(userIndex).flags.AdminInvisible = 1) Then
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageSetInvisible(UserList(userIndex).Char.CharIndex, True))
    End If
    
    If FX And UserList(userIndex).flags.AdminInvisible = 0 Then 'FX
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_WARP))
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCreateFX(UserList(userIndex).Char.CharIndex, FXIDs.FXWARP, 0))
    End If
    
    Call WarpMascotas(userIndex)
End Sub

Sub WarpMascotas(ByVal userIndex As Integer)
Dim i As Integer

Dim UMascRespawn  As Boolean
Dim miflag As Byte, MascotasReales As Integer
Dim prevMacotaType As Integer

Dim PetTypes(1 To MAXMASCOTAS) As Integer
Dim PetRespawn(1 To MAXMASCOTAS) As Boolean
Dim PetTiempoDeVida(1 To MAXMASCOTAS) As Integer

Dim NroPets As Integer, InvocadosMatados As Integer

NroPets = UserList(userIndex).NroMacotas
InvocadosMatados = 0

    'Matamos los invocados
    '[Alejo 18-03-2004]
    For i = 1 To MAXMASCOTAS
        If UserList(userIndex).MascotasIndex(i) > 0 Then
            ' si la mascota tiene tiempo de vida > 0 significa q fue invocada.
            If Npclist(UserList(userIndex).MascotasIndex(i)).Contadores.TiempoExistencia > 0 Then
                Call QuitarNPC(UserList(userIndex).MascotasIndex(i))
                UserList(userIndex).MascotasIndex(i) = 0
                InvocadosMatados = InvocadosMatados + 1
                NroPets = NroPets - 1
            End If
        End If
    Next i
    
    If InvocadosMatados > 0 Then
        Call WriteConsoleMsg(userIndex, "Pierdes el control de tus mascotas.", FontTypeNames.FONTTYPE_INFO)
    End If
    
    For i = 1 To MAXMASCOTAS
        If UserList(userIndex).MascotasIndex(i) > 0 Then
            PetRespawn(i) = Npclist(UserList(userIndex).MascotasIndex(i)).flags.Respawn = 0
            PetTypes(i) = UserList(userIndex).MascotasType(i)
            PetTiempoDeVida(i) = Npclist(UserList(userIndex).MascotasIndex(i)).Contadores.TiempoExistencia
            Call QuitarNPC(UserList(userIndex).MascotasIndex(i))
        End If
    Next i
    
    For i = 1 To MAXMASCOTAS
        If PetTypes(i) > 0 Then
            UserList(userIndex).MascotasIndex(i) = SpawnNpc(PetTypes(i), UserList(userIndex).Pos, False, PetRespawn(i))
            UserList(userIndex).MascotasType(i) = PetTypes(i)
            'Controlamos que se sumoneo OK
            If UserList(userIndex).MascotasIndex(i) = 0 Then
                UserList(userIndex).MascotasIndex(i) = 0
                UserList(userIndex).MascotasType(i) = 0
                If UserList(userIndex).NroMacotas > 0 Then UserList(userIndex).NroMacotas = UserList(userIndex).NroMacotas - 1
                Exit Sub
            End If
            Npclist(UserList(userIndex).MascotasIndex(i)).MaestroUser = userIndex
            Npclist(UserList(userIndex).MascotasIndex(i)).Movement = TipoAI.SigueAmo
            Npclist(UserList(userIndex).MascotasIndex(i)).Target = 0
            Npclist(UserList(userIndex).MascotasIndex(i)).TargetNPC = 0
            Npclist(UserList(userIndex).MascotasIndex(i)).Contadores.TiempoExistencia = PetTiempoDeVida(i)
            Call FollowAmo(UserList(userIndex).MascotasIndex(i))
        End If
    Next i
    
    UserList(userIndex).NroMacotas = NroPets

End Sub


Sub RepararMascotas(ByVal userIndex As Integer)
Dim i As Integer
Dim MascotasReales As Integer

    For i = 1 To MAXMASCOTAS
      If UserList(userIndex).MascotasType(i) > 0 Then MascotasReales = MascotasReales + 1
    Next i
    
    If MascotasReales <> UserList(userIndex).NroMacotas Then UserList(userIndex).NroMacotas = 0

End Sub

Sub Cerrar_Usuario(ByVal userIndex As Integer, Optional ByVal Tiempo As Integer = -1)
    If Tiempo = -1 Then Tiempo = IntervaloCerrarConexion
    
    If UserList(userIndex).flags.UserLogged And Not UserList(userIndex).Counters.Saliendo Then
        UserList(userIndex).Counters.Saliendo = True
        UserList(userIndex).Counters.Salir = IIf((UserList(userIndex).flags.Privilegios And PlayerType.User) And MapInfo(UserList(userIndex).Pos.Map).Pk, Tiempo, 0)
        
        Call WriteConsoleMsg(userIndex, "Cerrando...Se cerrará el juego en " & UserList(userIndex).Counters.Salir & " segundos...", FontTypeNames.FONTTYPE_INFO)
        
    End If
    
End Sub

'CambiarNick: Cambia el Nick de un slot.
'
'UserIndex: Quien ejecutó la orden
'UserIndexDestino: SLot del usuario destino, a quien cambiarle el nick
'NuevoNick: Nuevo nick de UserIndexDestino
Public Sub CambiarNick(ByVal userIndex As Integer, ByVal UserIndexDestino As Integer, ByVal NuevoNick As String)
Dim ViejoNick As String
Dim ViejoCharBackup As String

If UserList(UserIndexDestino).flags.UserLogged = False Then Exit Sub
ViejoNick = UserList(UserIndexDestino).name

If FileExist(CharPath & ViejoNick & ".chr", vbNormal) Then
    'hace un backup del char
    ViejoCharBackup = CharPath & ViejoNick & ".chr.old-"
    Name CharPath & ViejoNick & ".chr" As ViejoCharBackup
End If

End Sub

Public Sub Empollando(ByVal userIndex As Integer)
If MapData(UserList(userIndex).Pos.Map, UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y).ObjInfo.ObjIndex > 0 Then
    UserList(userIndex).flags.EstaEmpo = 1
Else
    UserList(userIndex).flags.EstaEmpo = 0
    UserList(userIndex).EmpoCont = 0
End If

End Sub

Sub SendUserStatsTxtOFF(ByVal sendIndex As Integer, ByVal Nombre As String)

If FileExist(CharPath & Nombre & ".chr", vbArchive) = False Then
    Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
Else
    Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & Nombre, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Nivel: " & GetVar(CharPath & Nombre & ".chr", "stats", "elv") & "  EXP: " & GetVar(CharPath & Nombre & ".chr", "stats", "Exp") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "elu"), FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Vitalidad: " & GetVar(CharPath & Nombre & ".chr", "stats", "minsta") & "/" & GetVar(CharPath & Nombre & ".chr", "stats", "maxSta"), FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Salud: " & GetVar(CharPath & Nombre & ".chr", "stats", "MinHP") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxHP") & "  Mana: " & GetVar(CharPath & Nombre & ".chr", "Stats", "MinMAN") & "/" & GetVar(CharPath & Nombre & ".chr", "Stats", "MaxMAN"), FontTypeNames.FONTTYPE_INFO)
    
    Call WriteConsoleMsg(sendIndex, "Menor Golpe/Mayor Golpe: " & GetVar(CharPath & Nombre & ".chr", "stats", "MaxHIT"), FontTypeNames.FONTTYPE_INFO)
    
    Call WriteConsoleMsg(sendIndex, "Oro: " & GetVar(CharPath & Nombre & ".chr", "stats", "GLD"), FontTypeNames.FONTTYPE_INFO)
    
#If ConUpTime Then
    Dim TempSecs As Long
    Dim TempStr As String
    TempSecs = GetVar(CharPath & Nombre & ".chr", "INIT", "UpTime")
    TempStr = (TempSecs \ 86400) & " Dias, " & ((TempSecs Mod 86400) \ 3600) & " Horas, " & ((TempSecs Mod 86400) Mod 3600) \ 60 & " Minutos, " & (((TempSecs Mod 86400) Mod 3600) Mod 60) & " Segundos."
    Call WriteConsoleMsg(sendIndex, "Tiempo Logeado: " & TempStr, FontTypeNames.FONTTYPE_INFO)
#End If

End If

End Sub

Sub SendUserOROTxtFromChar(ByVal sendIndex As Integer, ByVal charName As String)
On Error Resume Next
Dim j As Integer
Dim CharFile As String, Tmp As String
Dim ObjInd As Long, ObjCant As Long

CharFile = CharPath & charName & ".chr"

If FileExist(CharFile, vbNormal) Then
    Call WriteConsoleMsg(sendIndex, charName, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, " Tiene " & GetVar(CharFile, "STATS", "BANCO") & " en el banco.", FontTypeNames.FONTTYPE_INFO)
Else
    Call WriteConsoleMsg(sendIndex, "Usuario inexistente: " & charName, FontTypeNames.FONTTYPE_INFO)
End If

End Sub
