Attribute VB_Name = "modGame"
'EXTERNAL FUNCTIONS
'KeyInput
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long
'time between frames
Dim timer_elapsed_time As Single


'***********************
'CONSTATNS
'***********************
'Objetos


'***********************
'Type
'***********************


'***********************
'Enums
'***********************



Public Function input_key_get(ByVal key_code As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    input_key_get = GetKeyState(key_code) And &H8000
End Function

Public Sub Game_Render()
    
    '*******************
    'RenderScreen
    '*******************
    dxBeginRender
        
    '*******************************
    'Draw Map
    Engine.Map_Render
    '*******************************
    
    '*******************************
    'RenderSignal
    DibujarCartel
    '*******************************
    
    '*******************************
    'Render Dialogs
    Dialogos.MostrarTexto
    '*******************************
    
    '*******************************
    'Console Render
    consoleRender
    '*******************************
    
    '*******************************
    'Render GUI
    guiRender
    '*******************************
    
    '*******************************
    'Draw engine stats
    dxStatsRender
    '*******************************
    
    dxEndRender
    
    
    '*******************
    'RenderInventory
    '*******************
    
    
    timer_elapsed_time = General_Get_Elapsed_Time()
    SpeedCalculate (timer_elapsed_time)
End Sub

Public Sub gameMouseInput()
    inputPoll
    
    Dim mouseX As Integer
    Dim mouseY As Integer
    
    inputMouseGet mouseX, mouseY
    
    Dim tX As Long
    Dim tY As Long
    Call Engine.Input_Mouse_Tile_Get(mouseX, mouseY, tX, tY)
    
    If inputDoubleClick Then
        Debug.Print "inputDoubleClick: X:" & mouseX & " Y: " & mouseY
        If guiDragGet Then
            GUIEndDrag mouseX, mouseY
        Else
            If Not GUIInputClick(mouseX, mouseY) Then
                Call General_Screen_Double_Click(tX, tY, input_key_get(vbKeyShift))
            End If
        End If
    ElseIf inputClick Then
        Debug.Print "inputClick: X:" & mouseX & " Y: " & mouseY
        If guiDragGet Then
            GUIEndDrag mouseX, mouseY
        Else
            If Not GUIInputClick(mouseX, mouseY) Then
                Call General_Screen_Left_Click(tX, tY, input_key_get(vbKeyShift))
            End If
        End If
    End If
    
    If inputMouseDown Then
        If Not guiDragGet Then
            GUIStartDrag mouseX, mouseY
        End If
    End If
    
    If inputMouseMove Then
        If guiDragGet Then
            GUIDrag mouseX, mouseY
        End If
    End If
    
    inputReset
End Sub




Public Sub Game_CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
On Error Resume Next
    'No input allowed while Argentum is not the active window
    If Not modApi.IsAppActive() Then Exit Sub
    If Not frmMain.visible Then Exit Sub
    
    'Dont allow pressing this keys if we are moving
    If Not Engine.Player_Moving Then
        If Not UserEstupido Then
            If input_key_get(vbKeyUp) Then
                Call MoveTo(E_Heading.NORTH)
            ElseIf input_key_get(vbKeyRight) Then
                Call MoveTo(E_Heading.EAST)
            ElseIf input_key_get(vbKeyDown) Then
                Call MoveTo(E_Heading.SOUTH)
            ElseIf input_key_get(vbKeyLeft) Then
                Call MoveTo(E_Heading.WEST)
            End If
        Else
            If input_key_get(vbKeyRight) Or input_key_get(vbKeyLeft) Or input_key_get(vbKeyUp) Or input_key_get(vbKeyDown) Then
                Call RandomMove 'Si presiona cualquier tecla y es estupido se mueve para cualquier lado.
            End If
        End If
        Call ActualizarCoordenadas
    End If
    
End Sub

Public Sub Game_KeyEvents(ByVal keyCode As Integer, ByVal Shift As Integer)
    
    If Not guiTextboxFocusGet Then
        If ((keyCode >= 65 And keyCode <= 90) Or _
               (keyCode >= 48 And keyCode <= 57)) Then
            Select Case keyCode
                Case vbKeyM
                    'Audio.MusicActivated = Not Audio.MusicActivated
                    
                Case vbKeyA
                    Call AgarrarItem
                
                Case vbKeyE
                    Call EquiparItem(guiInventoryItemGet(guiElements.inventory, "invUser"))
                
                Case vbKeyN
                    Nombres = Not Nombres
                
                Case vbKeyD
                    Call WriteWork(eSkill.Domar)
                
                Case vbKeyR
                    Call WriteWork(eSkill.Robar)
                
                Case vbKeyS
                    AddtoRichTextBox frmMain.RecTxt, "Para activar o desactivar el seguro utiliza la tecla '*' (asterisco)", 255, 255, 255, False, False, False
                
                Case vbKeyO
                    Call WriteWork(eSkill.Ocultarse)
                
                Case vbKeyT
                    Call TirarItem
                
                Case vbKeyU
                    'If macrotrabajo.Enabled Then DesactivarMacroTrabajo
                        
                    'If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call UsarItem
                    'End If
                    
                
                Case vbKeyL
                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If
            End Select
        End If
        
        Select Case keyCode
            Case vbKeyDelete
                'If SendTxt.visible Then Exit Sub
                'If Not frmCantidad.visible Then
                    'SendCMSTXT.visible = True
                    'SendCMSTXT.SetFocus
                'End If
            
            Case vbKeyF2
                Call ScreenCapture
            
            Case vbKeyF4
                FPSFLAG = Not FPSFLAG
            
            Case vbKeyF5
                Call frmOpciones.Show(vbModeless, frmMain)
            
            Case vbKeyF6
                If UserMinMAN = UserMaxMAN Then Exit Sub
                
                'If Not PuedeMacrear Then
                    'AddtoRichTextBox frmMain.RecTxt, "No tan rápido..!", 255, 255, 255, False, False, False
                'Else
                    Call WriteMeditate
                    'PuedeMacrear = False
                'End If
            
            Case vbKeyF7
                'If TrainingMacro.Enabled Then
                    'DesactivarMacroHechizos
                'Else
                    'ActivarMacroHechizos
                'End If
            
            Case vbKeyF8
                'If macrotrabajo.Enabled Then
                    'DesactivarMacroTrabajo
                'Else
                    'ActivarMacroTrabajo
                'End If
            
            Case vbKeyMultiply
                If frmMain.PicSeg.visible Then
                    AddtoRichTextBox frmMain.RecTxt, "Escribe /SEG para quitar el seguro", 255, 255, 255, False, False, False
                Else
                    Call WriteSafeToggle
                End If
            
            Case vbKeyControl
                If Shift <> 0 Then Exit Sub
                
                If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
                
                'No mas intervalo golpe-hechi.
                If Not MainTimer.Check(TimersIndex.AttackSpell) Then Exit Sub 'Check if spells interval has finished.
                
                If MainTimer.Check(TimersIndex.Attack) And _
                   (Not UserDescansar) And _
                   (Not UserMeditar) Then
                        'If TrainingMacro.Enabled Then DesactivarMacroHechizos
                        'If macrotrabajo.Enabled Then DesactivarMacroTrabajo
                        Call WriteAttack
                        '[ANIM ATAK]
                        'If charlist(UserCharIndex).Arma.WeaponWalk(charlist(UserCharIndex).heading).GrhIndex <> 0 Then
                        '    charlist(UserCharIndex).Arma.WeaponWalk(charlist(UserCharIndex).heading).started = 1
                        '    charlist(UserCharIndex).Arma.WeaponAttack = 1
                        'End If
                End If
            End Select
        End If
        
        Select Case keyCode
            Case vbKeyReturn
                If guiTextboxVisibleGet(guiElements.console, "txtChat") Then
                    If Len(guiTextboxTextGet(guiElements.console, "txtChat")) > 0 Then _
                        ParseUserCommand guiTextboxTextGet(guiElements.console, "txtChat")
                    guiTextboxClear guiElements.console, "txtChat"
                    guiTextboxHide guiElements.console, "txtChat"
                Else
                    guiTextboxShow guiElements.console, "txtChat"
                End If
        End Select
End Sub
