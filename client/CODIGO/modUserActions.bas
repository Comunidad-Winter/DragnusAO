Attribute VB_Name = "modUser"
Public cantidad As eCantidad

Public Enum eCantidad
    tirar = 1
    comprar
    vender
    retirar
    depositar
End Enum

Public Sub AgarrarItem()
    Call WritePickUp
End Sub

Public Sub EquiparItem(ByVal slot As Integer)
    If slot > 0 Then _
        Call WriteEquipItem(slot)
End Sub

Public Sub UsarItem()
    'If TrainingMacro.Enabled Then DesactivarMacroHechizos
    
    If (guiInventoryItemGet(guiElements.inventory, "invUser") > 0) And (guiInventoryItemGet(guiElements.inventory, "invUser") < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteUseItem(guiInventoryItemGet(guiElements.inventory, "invUser"))
End Sub

Public Sub TirarItem()
    Dim itemIndex As Integer
    
    itemIndex = guiInventoryItemGet(guiElements.inventory, "invUser")

    If (itemIndex > 0 And itemIndex < MAX_INVENTORY_SLOTS + 1) Or (itemIndex = FLAGORO) Then
        If guiInventoryAmountGet(guiElements.inventory, "invUser", itemIndex) = 1 Then
            Call WriteDrop(itemIndex, 1)
        Else
           If guiInventoryAmountGet(guiElements.inventory, "invUser", itemIndex) > 1 Then
                cantidad = eCantidad.tirar
                GUIShowForm guiElements.frmCantidad
           End If
        End If
    End If
End Sub


Public Sub actionCast(slot As Integer)
    If slot > 0 Then
        If UserHechizos(slot) > 0 Then
            If MainTimer.Check(TimersIndex.Work, False) Then
                Call WriteCastSpell(slot)
                Call WriteWork(eSkill.Magia)
                'UsaMacro = True
            End If
        End If
    End If
End Sub

Public Sub initStatLabels()
'Iniciamos los labels con los valores de los atributos y los skills
Dim i As Integer


For i = 1 To NUMATRIBUTOS
    guiLabelCaptionSet guiElements.frmEstadisticas, "lblAtrib" & i, AtributosNames(i) & ": " & UserAtributos(i)
Next

For i = 1 To NUMSKILLS
    guiLabelCaptionSet guiElements.frmEstadisticas, "lblSkill" & i, skillsNames(i) & ": " & userSkills(i)
Next


'Label4(1).Caption = "Asesino: " & UserReputacion.AsesinoRep
'Label4(2).Caption = "Bandido: " & UserReputacion.BandidoRep
'Label4(3).Caption = "Burgues: " & UserReputacion.BurguesRep
'Label4(4).Caption = "Ladrón: " & UserReputacion.LadronesRep
'Label4(5).Caption = "Noble: " & UserReputacion.NobleRep
'Label4(6).Caption = "Plebe: " & UserReputacion.PlebeRep

'If UserReputacion.Promedio < 0 Then
    'Label4(7).ForeColor = vbRed
    'Label4(7).Caption = "Status: CRIMINAL"
'Else
    'Label4(7).ForeColor = vbBlue
    'Label4(7).Caption = "Status: Ciudadano"
'End If

With UserEstadisticas
    guiLabelCaptionSet guiElements.frmEstadisticas, "lblCriminales", "Criminales matados: " & .CriminalesMatados
    guiLabelCaptionSet guiElements.frmEstadisticas, "lblCiudadanos", "Ciudadanos matados: " & .CiudadanosMatados
    guiLabelCaptionSet guiElements.frmEstadisticas, "lblUsuarios", "Usuarios matados: " & .UsuariosMatados
    guiLabelCaptionSet guiElements.frmEstadisticas, "lblNPCs", "NPCs matados: " & .NpcsMatados
    guiLabelCaptionSet guiElements.frmEstadisticas, "lblClase", "Clase: " & .Clase
    guiLabelCaptionSet guiElements.frmEstadisticas, "lblCarcel", "Tiempo restante en carcel: " & .PenaCarcel
End With

End Sub
