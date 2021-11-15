Attribute VB_Name = "modGUITextbox"
Option Explicit

Public Type guiTextbox
    controlName As String
    controlRect As tRectangle
    
    Text As String
    selPos As Integer
    visible As Boolean
End Type

Public Function guiCreateTextbox(ByVal controlName As String, rectangle As tRectangle, Optional ByVal visible As Boolean = True) As guiTextbox
    guiCreateTextbox.controlName = controlName
    guiCreateTextbox.controlRect = rectangle
    guiCreateTextbox.visible = visible
End Function

Private Function drawTextBox(ByVal destX As Integer, ByVal destY As Integer, ByVal width As Integer, ByVal height As Integer)
    guiTextureRender 15991, destX, destY, 32, height - 3, D3DColorXRGB(255, 255, 255)
    guiTextureRender 15992, destX + width - 3, destY, 32, height - 3, D3DColorXRGB(255, 255, 255)
    
    'ARRIBA
    guiTextureRender 15994, destX + 30, destY, width - 32, 32, D3DColorXRGB(255, 255, 255)
    guiTextureRender 15993, destX, destY, 32, 32, D3DColorXRGB(255, 255, 255)
    guiTextureRender 15995, destX + width - 50, destY, 32, 50, 32, D3DColorXRGB(255, 255, 255)
    
    'ABAJO
    guiTextureRender 15996, destX + 30, destY + height - 32, width - 32, 32, D3DColorXRGB(255, 255, 255)
    guiTextureRender 15994, destX, destY + height - 32, 32, 32, D3DColorXRGB(255, 255, 255)
    guiTextureRender 15995, destX + width - 50, destY + height - 32, 32, 50, 32, D3DColorXRGB(255, 255, 255)
    
End Function

Public Function textboxRender(textbox As guiTextbox, ByVal destX As Integer, ByVal destY As Integer)
    With textbox
        If .visible Then
            
            
            guiTextureRender 15258, destX + .controlRect.x, destY + .controlRect.y, 64, 64, D3DColorXRGB(255, 255, 255)
            guiTextureRender 15259, destX + .controlRect.x, destY + .controlRect.y, .controlRect.width, 64, D3DColorXRGB(255, 255, 255)
            guiTextureRender 15260, destX + .controlRect.x + .controlRect.width - 64, destY + .controlRect.y, 32, 32, D3DColorXRGB(255, 255, 255)
            
            dxTextRender 1, .Text, destX + .controlRect.x + 8, destY + .controlRect.y + 8, D3DColorARGB(255, 255, 255, 255)
        End If
    End With
End Function

Public Function textboxClick(textbox As guiTextbox, ByVal mouseX As Integer, ByVal mouseY As Integer) As Boolean
    With textbox
        textboxClick = rectMouseOver(.controlRect, mouseX, mouseY)
    End With
End Function


Private Function textboxWrite(textbox As guiTextbox, ByVal KeyAscii As Byte)
    With textbox
        If KeyAscii > 31 And KeyAscii < 136 Then
            .Text = left(.Text, .selPos) + Chr(KeyAscii) + Right(.Text, Len(.Text) - .selPos)
            .selPos = .selPos + 1
        End If
    End With
End Function

Private Function textboxErase(textbox As guiTextbox)
    With textbox
        If Len(.Text) > 0 And .selPos > 0 Then
            .Text = left(.Text, .selPos - 1) + Right(.Text, Len(.Text) - .selPos)
            .selPos = .selPos - 1
            If .selPos = 0 And Len(.Text) > 0 Then
                .selPos = 1
            End If
        End If
    End With
End Function

Public Function textboxKeyPress(textbox As guiTextbox, ByVal KeyAscii As Byte) As Boolean
        With textbox
            If KeyAscii = 8 Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or (KeyAscii > 31 And KeyAscii < 127) Then
                textboxKeyPress = True
                Select Case KeyAscii
                    Case 8 'BACKSPACE
                        textboxErase textbox
                    Case vbKeyLeft
                        If .selPos > 0 Then _
                            .selPos = .selPos - 1
                    Case vbKeyRight
                        If .selPos < Len(.Text) Then _
                            .selPos = .selPos + 1
                    Case Else
                        textboxWrite textbox, KeyAscii
                End Select
            End If
    End With
End Function


'*********************
'PROPERTIES
'*********************
Public Sub textboxShow(textbox As guiTextbox)
    textbox.visible = True
End Sub

Public Sub textboxSet(textbox As guiTextbox, ByVal Text As String)
    textbox.Text = Text
    textbox.selPos = Len(textbox.Text)
End Sub

Public Sub textboxHide(textbox As guiTextbox)
    textbox.visible = False
End Sub

Public Sub textboxClear(textbox As guiTextbox)
    textbox.Text = ""
    textbox.selPos = 0
End Sub

Public Function textboxTextGet(textbox As guiTextbox) As String
    textboxTextGet = textbox.Text
End Function

Public Function textboxVisibleGet(textbox As guiTextbox) As Boolean
    textboxVisibleGet = textbox.visible
End Function

