Attribute VB_Name = "modEffects"
Private Const MAXEFFECTS As Byte = 15

Private Type tEffectData
    grhIndex As Integer
    particleID As Integer
    offSetX As Integer
    offSetY As Integer
    lifeTime As Integer
End Type

Private Type tEffect
    mapX As Byte
    mapY As Byte
    CharIndex As Integer
    effectIndex As Integer
    effectID As Integer
    grh As grh
    particleIndex As Integer
    lifeTime As Long
End Type

Private lastEffect As Integer
Private effectList(1 To MAXEFFECTS) As tEffect
Private effectDataList() As tEffectData
Private effectCount As Integer


Public Sub loadEffects()
    Dim i As Integer
    
    effectCount = Val(General_Var_Get(resource_path & PATH_INIT & "\Effects.dat", "INIT", "effectCount"))
    
    If effectCount > 0 Then
        ReDim effectDataList(1 To effectCount) As tEffectData
        For i = 1 To effectCount
            effectDataList(i).lifeTime = Val(General_Var_Get(resource_path & PATH_INIT & "\Effects.dat", "EFFECT" & i, "lifeTime"))
            effectDataList(i).grhIndex = Val(General_Var_Get(resource_path & PATH_INIT & "\Effects.dat", "EFFECT" & i, "grhIndex"))
            effectDataList(i).particleID = Val(General_Var_Get(resource_path & PATH_INIT & "\Effects.dat", "EFFECT" & i, "particleID"))
        Next i
    End If
End Sub

Public Sub renderEffect(ByVal effectIndex As Integer, ByVal x As Integer, ByVal y As Integer, ByRef rgbList() As Long)
    If checkEffect(effectIndex) Then
        If effectList(effectIndex).grh.grh_index > 0 Then
            Grh_Render effectList(effectIndex).grh, x + effectDataList(effectList(effectIndex).effectID).offSetX, y + effectDataList(effectList(effectIndex).effectID).offSetY, rgbList(), True, True
        End If
        If effectList(effectIndex).particleIndex > 0 Then
            particleGroupRender effectList(effectIndex).particleIndex, x, y, effectDataList(effectList(effectIndex).effectID).offSetX, effectDataList(effectList(effectIndex).effectID).offSetY
        End If
        
        'Destruimos el efecto si se termina.
        If Not effectList(effectIndex).particleIndex > 0 And Not effectList(effectIndex).grh.Started = 2 Then
            destroyEffect effectIndex
        End If
    End If
End Sub

Private Function checkEffect(ByVal effectIndex As Integer) As Boolean
    checkEffect = False
    If effectIndex > 0 And effectIndex <= lastEffect Then
        checkEffect = True
    End If
End Function

Public Function createEffect(ByVal effectID As Integer, Optional ByVal lifeTime As Integer = 0) As Integer
    Dim index As Integer
    
    index = getArrayPos()

    If index > 0 Then
        If index > lastEffect Then _
            lastEffect = index
            
        makeEffect index, effectID, lifeTime
    Else
        If lastEffect < MAXEFFECTS Then
            lastEffect = lastEffect + 1
            index = lastEffect
            
            makeEffect index, effectID, lifeTime
        End If
    End If
    
    createEffect = index
End Function

Public Sub destroyEffect(effectIndex As Integer)
    If checkEffect(effectIndex) Then
        effectList(effectIndex).effectID = 0
        
        If effectList(effectIndex).particleIndex > 0 Then _
            particleGroupDestroy effectList(effectIndex).particleIndex
        
        If effectIndex = lastEffect Then
            Do While effectList(lastEffect).effectID = 0
                lastEffect = lastEffect - 1
                If lastEffect < 1 Then Exit Do
            Loop
        End If
    End If
End Sub

Private Function getArrayPos() As Integer
    Dim i As Integer
    
    getArrayPos = 0
    
    Do While (i < lastEffect And getArrayPos = 0)
        i = i + 1
        If Not effectList(i).effectID > 0 Then
            getArrayPos = i
        End If
    Loop

End Function

Private Sub makeEffect(ByVal index As Integer, ByVal effectID As Integer, Optional ByVal lifeTime As Integer = 0)

    If effectID < 1 Or effectID > effectCount Then Exit Sub
    
    effectList(index).effectID = effectID
    
    If lifeTime = 0 Then
        lifeTime = effectDataList(effectID).lifeTime
    End If
    
    effectList(index).particleIndex = particleGroupCreate(mapX, mapY, effectDataList(effectID).particleID, lifeTime, 0, index)
    
    Grh_Initialize effectList(index).grh, effectDataList(effectID).grhIndex, False, 0, , grhLoops
    
    effectList(index).CharIndex = CharIndex
    
    effectList(index).effectIndex = effectIndex
End Sub

Public Sub destroyEffectParticle(ByVal effectIndex As Integer)
    If checkEffect(effectIndex) Then
        effectList(effectIndex).particleIndex = 0
    End If
End Sub

Public Function effectDestroyed(ByVal effectIndex As Integer) As Boolean
    If checkEffect(effectIndex) Then
        If effectList(effectIndex).effectID > 0 Then
            effectDestroyed = True
        End If
    End If
End Function
