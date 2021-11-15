Attribute VB_Name = "Extra"
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

Public Function EsNewbie(ByVal userIndex As Integer) As Boolean
    EsNewbie = UserList(userIndex).Stats.ELV <= LimiteNewbie
End Function

Public Function esArmada(ByVal userIndex As Integer) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 23/01/2007
'***************************************************
    esArmada = (UserList(userIndex).Faccion.Alineacion = e_Alineacion.Real)
End Function

Public Function esCaos(ByVal userIndex As Integer) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 23/01/2007
'***************************************************
    esCaos = (UserList(userIndex).Faccion.Alineacion = e_Alineacion.Caos)
End Function

Public Function esNeutro(ByVal userIndex As Integer) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 23/01/2007
'***************************************************
    esNeutro = (UserList(userIndex).Faccion.Alineacion = e_Alineacion.Neutro)
End Function

Public Function Faccion(ByVal userIndex As Integer) As Byte
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 23/01/2007
'***************************************************
    Faccion = UserList(userIndex).Faccion.Alineacion
End Function

Public Function EsGM(ByVal userIndex As Integer) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 23/01/2007
'***************************************************
    EsGM = (UserList(userIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero))
End Function

Public Sub DoTileEvents(ByVal userIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Handles the Map passage of Users. Allows the existance
'of exclusive maps for Newbies, Royal Army and Caos Legion members
'and enables GMs to enter every map without restriction.
'Uses: Mapinfo(map).Restringir = "NEWBIE" (newbies), "ARMADA", "CAOS" and "NO".
'***************************************************
On Error GoTo errhandler

Dim nPos As WorldPos
Dim FxFlag As Boolean
Dim i As Byte
Dim j As Byte
'Controla las salidas
If InMapBounds(Map, X, Y) Then
    
    If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
        FxFlag = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport
    End If
    
    
    If (MapData(Map, X, Y).TileExit.Map > 0) And (MapData(Map, X, Y).TileExit.Map <= NumMaps) Then
    '¿Es mapa de newbies?
        If UCase$(MapInfo(MapData(Map, X, Y).TileExit.Map).Restringir) = "NEWBIE" Then
            '¿El usuario es un newbie?
            If EsNewbie(userIndex) Or EsGM(userIndex) Then
                If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(userIndex)) Then
                    If FxFlag Then '¿FX?
                        Call WarpUserChar(userIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, True)
                    Else
                        Call WarpUserChar(userIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, False)
                    End If
                Else
                    Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)
                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        If FxFlag Then
                            Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y, True)
                        Else
                            Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y, False)
                        End If
                    End If
                End If
            Else 'No es newbie
                Call WriteConsoleMsg(userIndex, "Mapa exclusivo para newbies.", FontTypeNames.FONTTYPE_INFO)
                Call ClosestStablePos(UserList(userIndex).Pos, nPos)
                Debug.Print UserList(userIndex).Pos.Map
                If nPos.X <> 0 And nPos.Y <> 0 Then
                    Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y, False)
                End If
            End If
        ElseIf UCase$(MapInfo(MapData(Map, X, Y).TileExit.Map).Restringir) = "ARMADA" Then '¿Es mapa de Armadas?
            '¿El usuario es Armada?
            If esArmada(userIndex) Or EsGM(userIndex) Then
                If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(userIndex)) Then
                    If FxFlag Then '¿FX?
                        Call WarpUserChar(userIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, True)
                    Else
                        Call WarpUserChar(userIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y)
                    End If
                Else
                    Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)
                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        If FxFlag Then
                            Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y, True)
                        Else
                            Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y)
                        End If
                    End If
                End If
            Else 'No es armada
                Call WriteConsoleMsg(userIndex, "Mapa exclusivo para miembros del ejercito Real", FontTypeNames.FONTTYPE_INFO)
                Call ClosestStablePos(UserList(userIndex).Pos, nPos)

                If nPos.X <> 0 And nPos.Y <> 0 Then
                        Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y)
                End If
            End If
        ElseIf UCase$(MapInfo(MapData(Map, X, Y).TileExit.Map).Restringir) = "CAOS" Then '¿Es mapa de Caos?
            '¿El usuario es Caos?
            If esCaos(userIndex) Or EsGM(userIndex) Then
                If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(userIndex)) Then
                    If FxFlag Then '¿FX?
                        Call WarpUserChar(userIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, True)
                    Else
                        Call WarpUserChar(userIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y)
                    End If
                Else
                    Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)
                    If nPos.X <> 0 And nPos.Y <> 0 Then
                        If FxFlag Then
                            Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y, True)
                        Else
                            Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y)
                        End If
                    End If
                End If
            Else 'No es caos
                Call WriteConsoleMsg(userIndex, "Mapa exclusivo para miembros del ejercito Oscuro.", FontTypeNames.FONTTYPE_INFO)
                Call ClosestStablePos(UserList(userIndex).Pos, nPos)

                If nPos.X <> 0 And nPos.Y <> 0 Then
                        Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y)
                End If
            End If
        Else 'No es un mapa de newbies, ni Armadas, ni Caos
            If LegalPos(MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, PuedeAtravesarAgua(userIndex)) Then
                If FxFlag Then
                    Call WarpUserChar(userIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y, True)
                Else
                    Call WarpUserChar(userIndex, MapData(Map, X, Y).TileExit.Map, MapData(Map, X, Y).TileExit.X, MapData(Map, X, Y).TileExit.Y)
                End If
            Else
                Call ClosestLegalPos(MapData(Map, X, Y).TileExit, nPos)
                If nPos.X <> 0 And nPos.Y <> 0 Then
                    If FxFlag Then
                        Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y, True)
                    Else
                        Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y)
                    End If
                End If
            End If
        End If
            
        'Te fusite del mapa. La criatura ya no es más tuya ni te reconoce como que vos la atacaste.
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
    End If
End If



Exit Sub

errhandler:
    Call LogError("Error en DotileEvents. Error: " & Err.Number & " - Desc: " & Err.description)
End Sub

Function InRangoVision(ByVal userIndex As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

If X > UserList(userIndex).Pos.X - MinXBorder And X < UserList(userIndex).Pos.X + MinXBorder Then
    If Y > UserList(userIndex).Pos.Y - MinYBorder And Y < UserList(userIndex).Pos.Y + MinYBorder Then
        InRangoVision = True
        Exit Function
    End If
End If
InRangoVision = False

End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, X As Integer, Y As Integer) As Boolean

If X > Npclist(NpcIndex).Pos.X - MinXBorder And X < Npclist(NpcIndex).Pos.X + MinXBorder Then
    If Y > Npclist(NpcIndex).Pos.Y - MinYBorder And Y < Npclist(NpcIndex).Pos.Y + MinYBorder Then
        InRangoVisionNPC = True
        Exit Function
    End If
End If
InRangoVisionNPC = False

End Function


Function InMapBounds(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
            
If (Map <= 0 Or Map > NumMaps) Or X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapBounds = False
Else
    InMapBounds = True
End If

End Function

Sub ClosestLegalPos(Pos As WorldPos, ByRef nPos As WorldPos, Optional PuedeAgua As Boolean = False, Optional PuedeTierra As Boolean = True)
'*****************************************************************
'Author: Unknown (original version)
'Last Modification: 24/01/2007 (ToxicWaste)
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = Pos.Map

Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y, PuedeAgua, PuedeTierra)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.Y - LoopC To Pos.Y + LoopC
        For tX = Pos.X - LoopC To Pos.X + LoopC
            
            If LegalPos(nPos.Map, tX, tY, PuedeAgua, PuedeTierra) Then
                nPos.X = tX
                nPos.Y = tY
                '¿Hay objeto?
                
                tX = Pos.X + LoopC
                tY = Pos.Y + LoopC
  
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

Sub ClosestStablePos(Pos As WorldPos, ByRef nPos As WorldPos)
'*****************************************************************
'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = Pos.Map

Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.Y - LoopC To Pos.Y + LoopC
        For tX = Pos.X - LoopC To Pos.X + LoopC
            
            If LegalPos(nPos.Map, tX, tY) And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                nPos.X = tX
                nPos.Y = tY
                '¿Hay objeto?
                
                tX = Pos.X + LoopC
                tY = Pos.Y + LoopC
  
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

Function NameIndex(ByVal name As String) As Integer

Dim userIndex As Integer
'¿Nombre valido?
If LenB(name) = 0 Then
    NameIndex = 0
    Exit Function
End If

If InStrB(name, "+") <> 0 Then
    name = UCase$(Replace(name, "+", " "))
End If

userIndex = 1
Do Until UCase$(UserList(userIndex).name) = UCase$(name)
    
    userIndex = userIndex + 1
    
    If userIndex > MaxUsers Then
        NameIndex = 0
        Exit Function
    End If
    
Loop
 
NameIndex = userIndex
 
End Function



Function IP_Index(ByVal inIP As String) As Integer
 
Dim userIndex As Integer
'¿Nombre valido?
If LenB(inIP) = 0 Then
    IP_Index = 0
    Exit Function
End If
  
userIndex = 1
Do Until UserList(userIndex).ip = inIP
    
    userIndex = userIndex + 1
    
    If userIndex > MaxUsers Then
        IP_Index = 0
        Exit Function
    End If
    
Loop
 
IP_Index = userIndex

Exit Function

End Function


Function CheckForSameIP(ByVal userIndex As Integer, ByVal UserIP As String) As Boolean
Dim LoopC As Integer
For LoopC = 1 To MaxUsers
    If UserList(LoopC).flags.UserLogged = True Then
        If UserList(LoopC).ip = UserIP And userIndex <> LoopC Then
            CheckForSameIP = True
            Exit Function
        End If
    End If
Next LoopC
CheckForSameIP = False
End Function

Function CheckForSameName(ByVal name As String) As Boolean
'Controlo que no existan usuarios con el mismo nombre
Dim LoopC As Long
For LoopC = 1 To LastUser
    If UserList(LoopC).flags.UserLogged Then
        
        'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
        'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
        'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
        'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ITEMS
        'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
        
        If UCase$(UserList(LoopC).name) = UCase$(name) Then
            CheckForSameName = True
            Exit Function
        End If
    End If
Next LoopC
CheckForSameName = False
End Function

Sub HeadtoPos(ByVal head As eHeading, ByRef Pos As WorldPos)
'*****************************************************************
'Toma una posicion y se mueve hacia donde esta perfilado
'*****************************************************************
Dim X As Integer
Dim Y As Integer
Dim tempVar As Single
Dim nX As Integer
Dim nY As Integer

X = Pos.X
Y = Pos.Y

If head = eHeading.NORTH Then
    nX = X
    nY = Y - 1
End If

If head = eHeading.SOUTH Then
    nX = X
    nY = Y + 1
End If

If head = eHeading.EAST Then
    nX = X + 1
    nY = Y
End If

If head = eHeading.WEST Then
    nX = X - 1
    nY = Y
End If

'Devuelve valores
Pos.X = nX
Pos.Y = nY

End Sub

Function LegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Checks if the position is Legal.
'***************************************************
'¿Es un mapa valido?
If (Map <= 0 Or Map > NumMaps) Or _
   (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
            LegalPos = False
Else
    If PuedeAgua And PuedeTierra Then
        LegalPos = (MapData(Map, X, Y).Blocked <> 1) And _
                   (MapData(Map, X, Y).userIndex = 0) And _
                   (MapData(Map, X, Y).NpcIndex = 0)
    ElseIf PuedeTierra And Not PuedeAgua Then
        LegalPos = (MapData(Map, X, Y).Blocked <> 1) And _
                   (MapData(Map, X, Y).userIndex = 0) And _
                   (MapData(Map, X, Y).NpcIndex = 0) And _
                   (Not HayAgua(Map, X, Y))
    ElseIf PuedeAgua And Not PuedeTierra Then
        LegalPos = (MapData(Map, X, Y).Blocked <> 1) And _
                   (MapData(Map, X, Y).userIndex = 0) And _
                   (MapData(Map, X, Y).NpcIndex = 0) And _
                   (HayAgua(Map, X, Y))
    Else
        LegalPos = False
    End If
   
End If

End Function
Function LegalPosNPC(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte) As Boolean

If (Map <= 0 Or Map > NumMaps) Or _
   (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
    LegalPosNPC = False
Else

 If AguaValida = 0 Then
   LegalPosNPC = (MapData(Map, X, Y).Blocked <> 1) And _
     (MapData(Map, X, Y).userIndex = 0) And _
     (MapData(Map, X, Y).NpcIndex = 0) And _
     (MapData(Map, X, Y).trigger <> eTrigger.POSINVALIDA) _
     And Not HayAgua(Map, X, Y)
 Else
   LegalPosNPC = (MapData(Map, X, Y).Blocked <> 1) And _
     (MapData(Map, X, Y).userIndex = 0) And _
     (MapData(Map, X, Y).NpcIndex = 0) And _
     (MapData(Map, X, Y).trigger <> eTrigger.POSINVALIDA)
 End If
 
End If


End Function

Sub SendHelp(ByVal Index As Integer)
Dim NumHelpLines As Integer
Dim LoopC As Integer

NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

For LoopC = 1 To NumHelpLines
    Call WriteConsoleMsg(Index, GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC), FontTypeNames.FONTTYPE_INFO)
Next LoopC

End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal userIndex As Integer)
    If Npclist(NpcIndex).NroExpresiones > 0 Then
        Dim randomi
        randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageChatOverHead(Npclist(NpcIndex).Expresiones(randomi), Npclist(NpcIndex).Char.CharIndex, vbWhite))
    End If
End Sub

Sub LookatTile(ByVal userIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

'Responde al click del usuario sobre el mapa
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Stat As String
Dim OBJType As Integer
Dim ft As FontTypeNames

'¿Rango Visión? (ToxicWaste)
If (Abs(UserList(userIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(userIndex).Pos.X - X) > RANGO_VISION_X) Then
    Exit Sub
End If

'¿Posicion valida?
If InMapBounds(Map, X, Y) Then
    UserList(userIndex).flags.TargetMap = Map
    UserList(userIndex).flags.targetX = X
    UserList(userIndex).flags.targetY = Y
    '¿Es un obj?
    If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
        'Informa el nombre
        UserList(userIndex).flags.TargetObjMap = Map
        UserList(userIndex).flags.TargetObjX = X
        UserList(userIndex).flags.TargetObjY = Y
        FoundSomething = 1
    ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
        'Informa el nombre
        If ObjData(MapData(Map, X + 1, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            UserList(userIndex).flags.TargetObjMap = Map
            UserList(userIndex).flags.TargetObjX = X + 1
            UserList(userIndex).flags.TargetObjY = Y
            FoundSomething = 1
        End If
    ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
        If ObjData(MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            'Informa el nombre
            UserList(userIndex).flags.TargetObjMap = Map
            UserList(userIndex).flags.TargetObjX = X + 1
            UserList(userIndex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
        If ObjData(MapData(Map, X, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            'Informa el nombre
            UserList(userIndex).flags.TargetObjMap = Map
            UserList(userIndex).flags.TargetObjX = X
            UserList(userIndex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    End If
    
    If FoundSomething = 1 Then
        UserList(userIndex).flags.TargetObj = MapData(Map, UserList(userIndex).flags.TargetObjX, UserList(userIndex).flags.TargetObjY).ObjInfo.ObjIndex
        If MostrarCantidad(UserList(userIndex).flags.TargetObj) Then
            Call WriteConsoleMsg(userIndex, ObjData(UserList(userIndex).flags.TargetObj).name & " - " & MapData(UserList(userIndex).flags.TargetObjMap, UserList(userIndex).flags.TargetObjX, UserList(userIndex).flags.TargetObjY).ObjInfo.amount & "", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(userIndex, ObjData(UserList(userIndex).flags.TargetObj).name, FontTypeNames.FONTTYPE_INFO)
        End If
    
    End If
    '¿Es un personaje?
    If Y + 1 <= YMaxMapSize Then
        If MapData(Map, X, Y + 1).userIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y + 1).userIndex
            If UserList(TempCharIndex).showName Then    ' Es GM y pidió que se oculte su nombre??
                FoundChar = 1
            End If
        End If
        If MapData(Map, X, Y + 1).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y + 1).NpcIndex
            FoundChar = 2
        End If
    End If
    '¿Es un personaje?
    If FoundChar = 0 Then
        If MapData(Map, X, Y).userIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y).userIndex
            If UserList(TempCharIndex).showName Then    ' Es GM y pidió que se oculte su nombre??
                FoundChar = 1
            End If
        End If
        If MapData(Map, X, Y).NpcIndex > 0 Then
            TempCharIndex = MapData(Map, X, Y).NpcIndex
            FoundChar = 2
        End If
    End If
    
    
    'Reaccion al personaje
    If FoundChar = 1 Then '  ¿Encontro un Usuario?
            
       If UserList(TempCharIndex).flags.AdminInvisible = 0 Or UserList(userIndex).flags.Privilegios And PlayerType.Dios Then
            
            If LenB(UserList(TempCharIndex).DescRM) = 0 Then
                If EsNewbie(TempCharIndex) Then
                    Stat = " <NEWBIE>"
                End If
                
                If UserList(TempCharIndex).Faccion.Alineacion = e_Alineacion.Real Then
                    Stat = Stat & " <Ejercito Real: " & TituloReal(TempCharIndex) & ">"
                ElseIf UserList(TempCharIndex).Faccion.Alineacion = e_Alineacion.Caos Then
                    Stat = Stat & " <Legion Oscura: " & TituloCaos(TempCharIndex) & ">"
                Else
                    Stat = Stat & " <Neutral> "
                End If
                
                If UserList(TempCharIndex).GuildIndex > 0 Then
                    Stat = Stat & " <" & modGuilds.GuildName(UserList(TempCharIndex).GuildIndex) & ">"
                End If
                
                If Len(UserList(TempCharIndex).Desc) > 0 Then
                    Stat = "Ves a " & UserList(TempCharIndex).name & Stat & " - " & UserList(TempCharIndex).Desc
                Else
                    Stat = "Ves a " & UserList(TempCharIndex).name & Stat
                End If
                
                                
                If UserList(TempCharIndex).flags.Privilegios And PlayerType.RoyalCouncil Then
                    Stat = Stat & " [CONSEJO DE BANDERBILL]"
                    ft = FontTypeNames.FONTTYPE_CONSEJOVesA
                ElseIf UserList(TempCharIndex).flags.Privilegios And PlayerType.ChaosCouncil Then
                    Stat = Stat & " [CONSEJO DE LAS SOMBRAS]"
                    ft = FontTypeNames.FONTTYPE_CONSEJOCAOSVesA
                Else
                    If Not UserList(TempCharIndex).flags.Privilegios And PlayerType.User Then
                        Stat = Stat & " <GAME MASTER>"
                        ft = FontTypeNames.FONTTYPE_GM
                    End If
                End If
            Else
                Stat = UserList(TempCharIndex).DescRM
                ft = FontTypeNames.FONTTYPE_INFOBOLD
            End If
            
            If LenB(Stat) > 0 Then
                Call WriteConsoleMsg(userIndex, Stat, ft)
            End If
            
            FoundSomething = 1
            UserList(userIndex).flags.TargetUser = TempCharIndex
            UserList(userIndex).flags.TargetNPC = 0
            UserList(userIndex).flags.TargetNpcTipo = eNPCType.Comun
       End If

    End If
    If FoundChar = 2 Then '¿Encontro un NPC?
            Dim estatus As String
            
            estatus = "(" & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & ") "
            
            If Npclist(TempCharIndex).NumQuest = 0 Then
                If Len(Npclist(TempCharIndex).Desc) > 1 Then
                    Call WriteChatOverHead(userIndex, Npclist(TempCharIndex).Desc, Npclist(TempCharIndex).Char.CharIndex, vbWhite)
                ElseIf TempCharIndex = CentinelaNPCIndex Then
                    'Enviamos nuevamente el texto del centinela según quien pregunta
                    Call modCentinela.CentinelaSendClave(userIndex)
                Else
                    If Npclist(TempCharIndex).MaestroUser > 0 Then
                        Call WriteConsoleMsg(userIndex, estatus & Npclist(TempCharIndex).name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).name, FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(userIndex, estatus & Npclist(TempCharIndex).name & ".", FontTypeNames.FONTTYPE_INFO)
                        If UserList(userIndex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                            Call WriteConsoleMsg(userIndex, "Le pegó primero: " & Npclist(TempCharIndex).flags.AttackedFirstBy & ".", FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                End If
            Else
            'Escribimos el texto sobre la cabeza de los npcs de quest
                If Npclist(TempCharIndex).NumQuest <> 0 Then
                    If UserList(userIndex).Quest.UserQuest(Npclist(TempCharIndex).NumQuest).estado = 0 Then
                        Call WriteChatOverHead(userIndex, Npclist(TempCharIndex).Desc, Npclist(TempCharIndex).Char.CharIndex, vbYellow)
                    Else
                        If UserList(userIndex).Quest.UserQuest(Npclist(TempCharIndex).NumQuest).estado = 1 Then
                            Call WriteChatOverHead(userIndex, Quest(Npclist(TempCharIndex).NumQuest).Desc, Npclist(TempCharIndex).Char.CharIndex, vbYellow)
                        Else
                            If UserList(userIndex).Quest.UserQuest(Npclist(TempCharIndex).NumQuest).estado = 2 Then
                                Call WriteChatOverHead(userIndex, Quest(Npclist(TempCharIndex).NumQuest).Desc2, Npclist(TempCharIndex).Char.CharIndex, vbYellow)
                            End If
                        End If
                    End If
                End If
            End If
            FoundSomething = 1
            UserList(userIndex).flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
            UserList(userIndex).flags.TargetNPC = TempCharIndex
            UserList(userIndex).flags.TargetUser = 0
            UserList(userIndex).flags.TargetObj = 0
        
    End If
    
    If FoundChar = 0 Then
        UserList(userIndex).flags.TargetNPC = 0
        UserList(userIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(userIndex).flags.TargetUser = 0
    End If
    
    '*** NO ENCOTRO NADA ***
    If FoundSomething = 0 Then
        UserList(userIndex).flags.TargetNPC = 0
        UserList(userIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(userIndex).flags.TargetUser = 0
        UserList(userIndex).flags.TargetObj = 0
        UserList(userIndex).flags.TargetObjMap = 0
        UserList(userIndex).flags.TargetObjX = 0
        UserList(userIndex).flags.TargetObjY = 0
        Call WriteConsoleMsg(userIndex, "No ves nada interesante.", FontTypeNames.FONTTYPE_INFO)
    End If

Else
    If FoundSomething = 0 Then
        UserList(userIndex).flags.TargetNPC = 0
        UserList(userIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(userIndex).flags.TargetUser = 0
        UserList(userIndex).flags.TargetObj = 0
        UserList(userIndex).flags.TargetObjMap = 0
        UserList(userIndex).flags.TargetObjX = 0
        UserList(userIndex).flags.TargetObjY = 0
        Call WriteConsoleMsg(userIndex, "No ves nada interesante.", FontTypeNames.FONTTYPE_INFO)
    End If
End If


End Sub

Function FindDirection(Pos As WorldPos, Target As WorldPos) As eHeading
'*****************************************************************
'Devuelve la direccion en la cual el target se encuentra
'desde pos, 0 si la direc es igual
'*****************************************************************
Dim X As Integer
Dim Y As Integer

X = Pos.X - Target.X
Y = Pos.Y - Target.Y

'NE
If Sgn(X) = -1 And Sgn(Y) = 1 Then
    FindDirection = IIf(RandomNumber(0, 1), eHeading.NORTH, eHeading.EAST)
    Exit Function
End If

'NW
If Sgn(X) = 1 And Sgn(Y) = 1 Then
    FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.NORTH)
    Exit Function
End If

'SW
If Sgn(X) = 1 And Sgn(Y) = -1 Then
    FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.SOUTH)
    Exit Function
End If

'SE
If Sgn(X) = -1 And Sgn(Y) = -1 Then
    FindDirection = IIf(RandomNumber(0, 1), eHeading.SOUTH, eHeading.EAST)
    Exit Function
End If

'Sur
If Sgn(X) = 0 And Sgn(Y) = -1 Then
    FindDirection = eHeading.SOUTH
    Exit Function
End If

'norte
If Sgn(X) = 0 And Sgn(Y) = 1 Then
    FindDirection = eHeading.NORTH
    Exit Function
End If

'oeste
If Sgn(X) = 1 And Sgn(Y) = 0 Then
    FindDirection = eHeading.WEST
    Exit Function
End If

'este
If Sgn(X) = -1 And Sgn(Y) = 0 Then
    FindDirection = eHeading.EAST
    Exit Function
End If

'misma
If Sgn(X) = 0 And Sgn(Y) = 0 Then
    FindDirection = 0
    Exit Function
End If

End Function

'[Barrin 30-11-03]
Public Function ItemNoEsDeMapa(ByVal Index As Integer) As Boolean

ItemNoEsDeMapa = ObjData(Index).OBJType <> eOBJType.otPuertas And _
            ObjData(Index).OBJType <> eOBJType.otForos And _
            ObjData(Index).OBJType <> eOBJType.otCarteles And _
            ObjData(Index).OBJType <> eOBJType.otArboles And _
            ObjData(Index).OBJType <> eOBJType.otYacimiento And _
            ObjData(Index).OBJType <> eOBJType.otTeleport
End Function
'[/Barrin 30-11-03]

Public Function MostrarCantidad(ByVal Index As Integer) As Boolean
MostrarCantidad = ObjData(Index).OBJType <> eOBJType.otPuertas And _
            ObjData(Index).OBJType <> eOBJType.otForos And _
            ObjData(Index).OBJType <> eOBJType.otCarteles And _
            ObjData(Index).OBJType <> eOBJType.otArboles And _
            ObjData(Index).OBJType <> eOBJType.otYacimiento And _
            ObjData(Index).OBJType <> eOBJType.otTeleport
End Function

Public Function EsObjetoFijo(ByVal OBJType As eOBJType) As Boolean

EsObjetoFijo = OBJType = eOBJType.otForos Or _
               OBJType = eOBJType.otCarteles Or _
               OBJType = eOBJType.otArboles Or _
               OBJType = eOBJType.otYacimiento

End Function
