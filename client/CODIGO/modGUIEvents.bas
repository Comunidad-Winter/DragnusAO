Attribute VB_Name = "modGUIEvents"
'DRAG DROP EVENTS
Option Explicit

Public Sub eventPictureDrop(ByVal targetElementID As Integer, ByVal targetControlName As String, ByVal targetControlType As eControlType, ByVal sourceForm As Integer, ByVal sourceName As String)

End Sub

Public Sub eventItemDrop(ByVal targetElementID As Integer, ByVal targetControlName As String, ByVal targetControlType As eControlType, ByVal sourceForm As Integer, ByVal sourceName As String, ByVal sourceIndex As Integer)
    Select Case targetElementID
        Case 0 'Screen
            If sourceForm = guiElements.frmMain Then
                TirarItem
            End If
        Case guiElements.frmComerciar
            If targetControlName = "userInv" Then
                If sourceName = "npcInv" Then
                    If sourceIndex > 0 Then
                        Call WriteCommerceBuy(sourceIndex, 1)
                    End If
                End If
            ElseIf targetControlName = "npcInv" Then
                If sourceName = "userInv" Then
                    If sourceIndex > 0 Then
                        Call WriteCommerceSell(sourceIndex, 1)
                    End If
                End If
            End If
    End Select
End Sub

Public Sub eventFormClose(ByVal elementID As Integer)
    Select Case elementID
        Case guiElements.frmComerciar
            WriteCommerceEnd
        Case guiElements.frmBancoObj
            WriteBankEnd
    End Select
End Sub

Public Sub eventFormShow(ByVal elementID As Integer)
    Select Case elementID
        Case guiElements.frmComerciar
            'guiFormPosSet guiElements.inventory
        Case guiElements.frmBancoObj
            'guiFormPosSet guiElements.inventory
    End Select
End Sub

'CLICK EVENTS
Public Sub eventButtonClick(ByVal elementID As Integer, ByVal controlName As String)
    Dim i As Integer

    Select Case elementID
        'BOVEDA
        Case guiElements.frmBancoObj
            If controlName = "btnRetirar" Then
                If Not guiInventoryItemGet(elementID, "npcInv") = 0 Then
                    cantidad = eCantidad.retirar
                    GUIShowForm guiElements.frmCantidad
                    lockFocus guiElements.frmCantidad
                End If
            ElseIf controlName = "btnDepositar" Then
                If Not guiInventoryItemGet(elementID, "userInv") = 0 Then
                    If Not guiInventoryEquipGet(elementID, "userInv", guiInventoryItemGet(elementID, "userInv")) Then
                        cantidad = eCantidad.depositar
                        GUIShowForm guiElements.frmCantidad
                        lockFocus guiElements.frmCantidad
                    Else
                        consoleAdd "No podes depositar el item porque lo estas usando.", D3DColorXRGB(255, 0, 0)
                        'guiListAddItem guiElements.frmMain, "rectxtConsole", "No podes depositar el item porque lo estas usando."
                    End If
                End If
            End If
        
        'MISSING - CAMBIARMTD
            
        'TIRAR
        Case guiElements.frmCantidad
            If controlName = "btnTirar" Then
                If Len(guiTextboxTextGet(elementID, "txtCantidad")) > 0 Then
                    guiTextboxTextSet elementID, "txtCantidad", "0"
                End If
                If Not IsNumeric(guiTextboxTextGet(elementID, "txtCantidad")) Then
                    guiTextboxTextSet elementID, "txtCantidad", "0"
                End If
                If guiTextboxTextGet(elementID, "txtCantidad") > 10000 Then
                    guiTextboxTextSet elementID, "txtCantidad", "10000"
                End If
            End If
            
            Select Case cantidad
                Case eCantidad.tirar
                    If controlName = "btnTodo" Then
                        If guiInventoryItemGet(guiElements.inventory, "invUser") <> FLAGORO Then
                            Call WriteDrop(guiInventoryItemGet(guiElements.inventory, "invUser"), guiInventoryAmountGet(guiElements.inventory, "invUser", guiInventoryItemGet(guiElements.inventory, "invUser")))
                        Else
                            If UserGLD > 10000 Then
                                Call WriteDrop(guiInventoryItemGet(guiElements.inventory, "invUser"), 10000)
                            Else
                                Call WriteDrop(guiInventoryItemGet(guiElements.inventory, "invUser"), UserGLD)
                            End If
                        End If
                    Else
                        Call WriteDrop(guiInventoryItemGet(guiElements.inventory, "invUser"), guiTextboxTextGet(elementID, "txtCantidad"))
                    End If
                    
                Case eCantidad.vender
                    If controlName = "btnTodo" Then
                        Call WriteCommerceSell(guiInventoryItemGet(guiElements.frmComerciar, "userInv"), guiInventoryAmountGet(guiElements.frmComerciar, "userInv", guiInventoryItemGet(guiElements.frmComerciar, "userInv")))
                    Else
                        Call WriteCommerceSell(guiInventoryItemGet(guiElements.frmComerciar, "userInv"), Int(guiTextboxTextGet(elementID, "txtCantidad")))
                    End If
                    
                Case eCantidad.comprar
                    If controlName = "btnTodo" Then
                        If UserGLD >= guiInventoryValueGet(guiElements.frmComerciar, "npcInv", guiInventoryItemGet(guiElements.frmComerciar, "npcInv")) * guiInventoryAmountGet(guiElements.frmComerciar, "npcInv", guiInventoryItemGet(guiElements.frmComerciar, "npcInv")) Then
                            Call WriteCommerceBuy(guiInventoryItemGet(guiElements.frmComerciar, "npcInv"), guiInventoryAmountGet(guiElements.frmComerciar, "npcInv", guiInventoryItemGet(guiElements.frmComerciar, "npcInv")))
                        Else
                            consoleAdd "No tenés suficiente oro.", D3DColorXRGB(255, 0, 0)
                        End If
                    Else
                        If UserGLD >= guiInventoryValueGet(guiElements.frmComerciar, "npcInv", guiInventoryItemGet(guiElements.frmComerciar, "npcInv")) * Val(guiTextboxTextGet(elementID, "txtCantidad")) Then
                            Call WriteCommerceBuy(guiInventoryItemGet(guiElements.frmComerciar, "npcInv"), Val(guiTextboxTextGet(elementID, "txtCantidad")))
                        Else
                            consoleAdd "No tenés suficiente oro.", D3DColorXRGB(255, 0, 0)
                        End If
                    End If
                    
                Case eCantidad.depositar
                    If controlName = "btnTodo" Then
                        Call WriteBankDeposit(guiInventoryItemGet(guiElements.frmBancoObj, "userInv"), guiInventoryAmountGet(guiElements.frmBancoObj, "userInv", guiInventoryItemGet(guiElements.frmBancoObj, "userInv")))
                    Else
                        Call WriteBankDeposit(guiInventoryItemGet(guiElements.frmBancoObj, "userInv"), guiTextboxTextGet(elementID, "txtCantidad"))
                    End If
                    
                Case eCantidad.retirar
                    If controlName = "btnTodo" Then
                        Call WriteBankExtractItem(guiInventoryItemGet(guiElements.frmBancoObj, "npcInv"), guiInventoryAmountGet(guiElements.frmBancoObj, "npcInv", guiInventoryItemGet(guiElements.frmBancoObj, "npcInv")))
                    Else
                        Call WriteBankExtractItem(guiInventoryItemGet(guiElements.frmBancoObj, "npcInv"), guiTextboxTextGet(elementID, "txtCantidad"))
                    End If
            End Select
                
            cantidad = 0
            guiTextboxTextSet elementID, "txtCantidad", ""
            GUIHideForm elementID
        'MISSING - CARGANDO
        
        'CARPINTERO
        Case guiElements.frmCarp
            If controlName = "btnCraft" Then
                Call WriteCraftCarpenter(ObjCarpintero(guiListIndexGet(elementID, "listCraft")))
            End If
        
        'MISSING - CHARINFO
        
        'COMERCIO
        Case guiElements.frmComerciar
            If controlName = "btnComprar" Then
                If Not guiInventoryItemGet(elementID, "npcInv") = 0 Then
                    cantidad = eCantidad.comprar
                    GUIShowForm guiElements.frmCantidad
                    lockFocus guiElements.frmCantidad
                End If
            ElseIf controlName = "btnVender" Then
                If Not guiInventoryItemGet(elementID, "userInv") = 0 Then
                    If Not guiInventoryEquipGet(elementID, "userInv", guiInventoryItemGet(elementID, "userInv")) Then
                        cantidad = eCantidad.vender
                        GUIShowForm guiElements.frmCantidad
                        lockFocus guiElements.frmCantidad
                    Else
                        consoleAdd "No podes vender el item porque lo estas usando.", D3DColorXRGB(255, 0, 0)
                        'guiListAddItem guiElements.frmMain, "rectxtConsole", "No podes vender el item porque lo estas usando."
                    End If
                End If
            End If
        
            
        'COMERCIARUSU
        Case guiElements.frmComerciarUsu
            If controlName = "btnAceptar" Then
                WriteUserCommerceOk
            ElseIf controlName = "btnRechazar" Then
                WriteUserCommerceEnd
            ElseIf controlName = "btnOfrecer" Then
                WriteUserCommerceOffer guiInventoryItemGet(elementID, "invUser"), guiTextboxTextGet(elementID, "txtCantidad")
            End If
        
        'MISSING - COMMET
        
        'CONNECT
        Case guiElements.frmConectar
            If controlName = "btnConnect" Then
                
            ElseIf controlName = "btnNewAcc" Then
                
            ElseIf controlName = "btnDeleteAcc" Then
            
            ElseIf controlName = "btnOptions" Then
            
            End If
        'MISSING - CREAR PJ
        
        'MISSING - CUENTA

        'ENTRENADOR
        Case guiElements.frmEntrenador
            If controlName = "btnAceptar" Then
                Call WriteTrain(guiListIndexGet(elementID, "lstCriaturas"))
                GUIHideForm elementID
            End If
        
        'MISSING - ESTADISTICAS
        Case guiElements.frmEstadisticas
            If controlName = "btnCerrar" Then
                GUIHideForm elementID
            End If
        'MISSING - FORO
        
        'GUILDADM
        Case guiElements.frmGuildAdm
            If controlName = "btnSolicitar" Then
                If guiListIndexGet(elementID, "lstClanes") > 0 Then _
                    GUIShowForm guiElements.frmGuildSol
            ElseIf controlName = "btnDetalles" Then
                If guiListIndexGet(elementID, "lstClanes") > 0 Then _
                    Call WriteGuildRequest(guilds(guiListIndexGet(elementID, "lstClanes")))
            End If
        
        'GUILD SOLICITUD
        Case guiElements.frmGuildSol
            If controlName = "btnAceptar" Then
                Call WriteGuildRequestMembership(guilds(guiListIndexGet(elementID, "lstClanes")), Replace(Replace(guiTextboxTextGet(elementID, "txtAplicacion"), ",", ";"), vbCrLf, "º"))
            End If
        
        'GUILD LEADER
        Case guiElements.frmGuildLeader
            If controlName = "btnDetalles" Then
                If guiListIndexGet(elementID, "lstMembers") > 0 Then _
                    Call WriteGuildMemberInfo(members(guiListIndexGet(elementID, "lstMembers")))
            ElseIf controlName = "btnEchar" Then
                If guiListIndexGet(elementID, "lstMembers") > 0 Then _
                    Call WriteGuildKickMember(members(guiListIndexGet(elementID, "lstMembers")))
            ElseIf controlName = "btnRechazar" Then
                If guiListIndexGet(elementID, "lstSolicitudes") > 0 Then _
                    Call WriteGuildRejectNewMember(solicitudes(guiListIndexGet(elementID, "lstSolicitudes")), "")
            ElseIf controlName = "btnAceptar" Then
                If guiListIndexGet(elementID, "lstSolicitudes") > 0 Then _
                    Call WriteGuildAcceptNewMember(solicitudes(guiListIndexGet(elementID, "lstSolicitudes")))
            ElseIf controlName = "btnAceptarPaz" Then
                If guiListIndexGet(elementID, "lstPaz") > 0 Then
                    Call WriteGuildAcceptPeace(paz(guiListIndexGet(elementID, "lstPaz")))
                End If
            ElseIf controlName = "btnRechazarPaz" Then
                If guiListIndexGet(elementID, "lstPaz") > 0 Then
                    Call WriteGuildRejectPeace(paz(guiListIndexGet(elementID, "lstPaz")))
                End If
            ElseIf controlName = "btnGuerra" Then
                If guiListIndexGet(elementID, "lstClanes") > 0 Then _
                    Call WriteGuildDeclareWar(guilds(guiListIndexGet(elementID, "lstClanes")))
            ElseIf controlName = "btnPaz" Then
                If guiListIndexGet(elementID, "lstClanes") > 0 Then _
                    Call WriteGuildOfferPeace(guilds(guiListIndexGet(elementID, "lstClanes")), "")
            End If
        
        'HERRERIA
        Case guiElements.frmHerrero
            If controlName = "btnConstruir" Then
                Select Case craftingBlackSmith
                    Case craftingObj.armor
                        Call WriteCraftBlacksmith(ArmadurasHerrero(guiListIndexGet(elementID, "lstCraft")).index)
                    Case craftingObj.shield
                        Call WriteCraftBlacksmith(EscudosHerrero(guiListIndexGet(elementID, "lstCraft")).index)
                    Case craftingObj.weapon
                        Call WriteCraftBlacksmith(ArmasHerrero(guiListIndexGet(elementID, "lstCraft")).index)
                    Case craftingObj.helmet
                        Call WriteCraftBlacksmith(CascosHerrero(guiListIndexGet(elementID, "lstCraft")).index)
                End Select
            ElseIf controlName = "btnArmas" Then
                guiListClear elementID, "lstCraft"
                
                craftingBlackSmith = craftingObj.weapon
                
                For i = 1 To UBound(ArmasHerrero)
                    guiListAddItem elementID, "lstCraft", ArmasHerrero(i).name
                Next i
            ElseIf controlName = "btnArmaduras" Then
                guiListClear elementID, "lstCraft"
                
                craftingBlackSmith = craftingObj.armor
                
                For i = 1 To UBound(ArmadurasHerrero)
                    guiListAddItem elementID, "lstCraft", ArmadurasHerrero(i).name
                Next i
            ElseIf controlName = "btnCascos" Then
                guiListClear elementID, "lstCraft"
                
                craftingBlackSmith = craftingObj.helmet
                
                For i = 1 To UBound(CascosHerrero)
                    guiListAddItem elementID, "lstCraft", CascosHerrero(i).name
                Next i
            ElseIf controlName = "btnEscudos" Then
                guiListClear elementID, "lstCraft"
                
                craftingBlackSmith = craftingObj.shield
                
                For i = 1 To UBound(EscudosHerrero)
                    guiListAddItem elementID, "lstCraft", EscudosHerrero(i).name
                Next i
            End If
            
        'ASIGNAR SKILLS
        Case guiElements.frmSkills3
            If controlName = "btnAgregar" Then
                If Alocados > 0 Then
                    i = guiListIndexGet(elementID, "lstSkills")
                    If i > 0 Then
                        userSkillAsign(i) = userSkillAsign(i) + 1
                        Alocados = Alocados - 1
                        guiLabelCaptionSet elementID, "lblLibres", Alocados
                        guiLabelCaptionSet elementID, "lblSkills", userSkillAsign(i)
                    End If
                End If
            ElseIf controlName = "btnQuitar" Then
                If Alocados < 10 Then
                    i = guiListIndexGet(elementID, "lstSkills")
                    If i > 0 Then
                        If userSkillAsign(i) > 0 Then
                            userSkillAsign(i) = userSkillAsign(i) - 1
                            Alocados = Alocados + 1
                            guiLabelCaptionSet elementID, "lblLibres", Alocados
                            guiLabelCaptionSet elementID, "lblSkills", userSkillAsign(i)
                        End If
                    End If
                End If
            ElseIf controlName = "btnAsignar" Then
                WriteModifySkills userSkillAsign()
                For i = 1 To NUMSKILLS
                    userSkillAsign(i) = 0
                Next i
                GUIHideForm elementID
                SkillPoints = Alocados
            End If
        
        'SPAWNLIST
        Case guiElements.frmSpawnList
            If controlName = "btnSpawn" Then
                If guiListIndexGet(elementID, "lstCriaturas") > 0 Then _
                    Call WriteSpawnCreature(guiListIndexGet(elementID, "lstCriaturas"))
            End If
            
        'HECHIZOS
        Case guiElements.spells
            'Hechizos
            If controlName = "btnLanzar" Then
                actionCast guiListIndexGet(guiElements.spells, "spellList")
            End If
        
        'INVENTARIO
        Case guiElements.inventory

    End Select
End Sub
