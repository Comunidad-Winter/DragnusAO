Attribute VB_Name = "modParticles"
Private Type tParticle
    x As Single
    y As Single
    vX As Single
    vY As Single
    
    Screen_X As Single
    Screen_Y As Single
    
    Moved_X As Single
    Moved_Y As Single
    
    Created As Long
    Alive As Byte
    CurrentColor(1 To 4) As D3DCOLORVALUE
    rgb_list(3) As Long
    
    angle As Single
    
    texture_index As Integer
    
    Particle_LifeTime As Long
    Used As Boolean
    
    Delay As Integer
    DelayCounter As Integer
End Type

Private Type tParticleEmisor
    x1 As Integer
    x2 As Integer
    Y1 As Integer
    Y2 As Integer
    
    vX1 As Integer
    vY1 As Integer
    vX2 As Integer
    vY2 As Integer
    
    MoveX As Boolean
    MoveY As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    
    Particle_count As Integer
    
    Particle_Speed As Single
    Particle_Frame_Counter As Single
    
    Particle() As tParticle
    
    Gravity As Single
    
    Friction As Single
    
    RatioFriction As Single
    ColorVariation As Byte
    
    StartColor(1 To 4) As D3DCOLORVALUE
    EndColor(1 To 4) As D3DCOLORVALUE
    
    PLt1 As Integer
    PLt2 As Integer
    
    alpha_blend As Boolean
    
    WindDirection As Byte
    Wind As Single
    
    Spin As Byte
    SpinH As Integer
    SpinL As Integer
    
    Bounce_Strength As Integer
    Bounce_Y As Integer
    
    texture_count As Integer
    texture_index() As Integer
    texture_size() As Integer
    
    Ratio As Integer
    RatioVariation As Integer
    CurrentRatio As Single
    
    ParticleGroup_Type As Integer
    
    ParticlesLeft As Integer
    
    StartParticlesDestroy As Boolean 'Time
    
    KillWhenAtTarget As Boolean
    Target_X As Integer
    Target_Y As Integer
    
    StopAtTargetRatio As Boolean 'If False then it loops
    TargetRatio As Integer
    
    KillType As Byte
    
    Delay1 As Integer
    Delay2 As Integer
End Type

Private Type tParticleGroup
    Active As Byte
    'Pointer to the effect handling the particle.
    effectIndex As Integer
    
    ParticleGroup_Lifetime As Long
    defaultLifeTime As Long
    Created As Long
    ParticleEmisor As tParticleEmisor
    ParticleEmisor_Count As Integer
End Type

Public Enum e_ParticleType
    Rain = 7
End Enum

Dim particle_timer As Single

'Particles
Private particle_group_list() As tParticleGroup
Private particle_group_count As Integer

Private particle_group_data_count As Integer
Private particle_group_data() As tParticleGroup

Public Function particleSpeedCalculate(ByVal timer_elapsed_time As Single)
    particle_timer = timer_elapsed_time * 0.03
End Function

Public Function particleGroupCreate(ByVal map_x As Integer, ByVal map_y As Integer, ByVal particle_type As e_ParticleType, ByVal lifeTime As Integer, Optional ByVal char_index As Integer, Optional ByVal effectIndex As Integer) As Integer
On Error GoTo errhandler

    If particle_type = 0 Then Exit Function
        particleGroupCreate = particleGroupMake(map_x, map_y, lifeTime, particle_type, char_index, effectIndex)
    
Exit Function
errhandler:
    particleGroupCreate = 0
End Function

Private Function particleGroupMake(ByVal map_x As Byte, ByVal map_y As Byte, ByVal lifeTime As Integer, ByVal particle_type As e_ParticleType, Optional ByVal char_index As Integer, Optional ByVal effectIndex As Integer)

    
    Dim i As Integer
    Dim Particle_Index As Integer
    
    particleGroupMake = 0
    
    Call particleGetArrayPos(Particle_Index)
    
    If Particle_Index = 0 Then
        particle_group_count = particle_group_count + 1
        Particle_Index = particle_group_count
        ReDim Preserve particle_group_list(1 To particle_group_count)
    End If
    
    If lifeTime = 0 Then lifeTime = particle_group_data(particle_type).defaultLifeTime
    
    particle_group_list(Particle_Index) = particle_group_data(particle_type)
    
    With particle_group_list(Particle_Index)
        .Active = 1
        .ParticleGroup_Lifetime = lifeTime
        .Created = GetTickCount
        '.map_x = map_x
        '.map_y = map_y
        '.char_index = char_index
        
        .effectIndex = effectIndex
        
        If .ParticleEmisor.Delay1 Or .ParticleEmisor.Delay2 Then
            For i = 1 To .ParticleEmisor.Particle_count
                .ParticleEmisor.Particle(i).Delay = Val(General_Random_Number(.ParticleEmisor.Delay1, .ParticleEmisor.Delay2))
            Next i
        End If
    End With
    
    
    
    particleGroupMake = Particle_Index
End Function
Private Function particlesGroupUpdate(ByVal ParticleGroup_Index As Integer, ByVal x As Integer, ByVal y As Integer, Optional ByVal offset_x As Single, Optional ByVal offset_y As Single)
    Dim i As Integer
    Dim Time As Long
    Dim Particle_Emisor As tParticleEmisor
    
    Time = GetTickCount
    With particle_group_list(ParticleGroup_Index)
        If particleGroupCheckPermanency(ParticleGroup_Index, Time) Then
            With .ParticleEmisor
                If .Particle_Frame_Counter > .Particle_Speed Then
                    .Particle_Frame_Counter = 0
                    For i = 0 To .Particle_count
                        If .Particle(i).Created >= .Particle(i).Particle_LifeTime Then
                            .Particle(i).Alive = 0
                        ElseIf .KillWhenAtTarget Then
                            If .Particle(i).x + .Particle(i).Moved_X >= .Target_X Or .Particle(i).y + .Particle(i).Moved_Y >= .Target_Y Then
                                .Particle(i).Alive = 0
                            End If
                        End If
                    
                        If .Particle(i).Alive = 0 And Not .Particle(i).Used Then
                            If .Particle(i).DelayCounter >= .Particle(i).Delay Then
                                If .StartParticlesDestroy And Not .Particle(i).Used Then
                                    .Particle(i).Used = True
                                    .ParticlesLeft = .ParticlesLeft - 1
                                Else
                                    Select Case .ParticleGroup_Type
                                        Case 0
                                            particle1Reset ParticleGroup_Index, i
                                        Case 1
                                            particle2Reset ParticleGroup_Index, i
                                        Case 2
                                            particle3Reset ParticleGroup_Index, i
                                        Case 3
                                            particle4Reset ParticleGroup_Index, i
                                        Case 4
                                            particle5Reset ParticleGroup_Index, i
                                    End Select
                                    
                                    'Standard particle settings.
                                    .Particle(i).Alive = 1
                                    .Particle(i).Created = 0
                                    .Particle(i).vX = General_Random_Number(.vX1, .vX2)
                                    .Particle(i).vY = General_Random_Number(.vY1, .vY2)
                                    .Particle(i).Particle_LifeTime = General_Random_Number(.PLt1, .PLt2)
                                    
                                    'Reset moving status.
                                    .Particle(i).Moved_X = 0
                                    .Particle(i).Moved_Y = 0
    
                                    .Particle(i).texture_index = General_Random_Number(1, .texture_count)
                                
                                    .Particle(i).CurrentColor(1) = .StartColor(1)
                                    .Particle(i).CurrentColor(2) = .StartColor(2)
                                    .Particle(i).CurrentColor(3) = .StartColor(3)
                                    .Particle(i).CurrentColor(4) = .StartColor(4)
                                    
                                    .Particle(i).rgb_list(0) = D3DColorARGB(.Particle(i).CurrentColor(1).A, .Particle(i).CurrentColor(1).R, .Particle(i).CurrentColor(1).G, .Particle(i).CurrentColor(1).B)
                                    .Particle(i).rgb_list(1) = D3DColorARGB(.Particle(i).CurrentColor(2).A, .Particle(i).CurrentColor(2).R, .Particle(i).CurrentColor(2).G, .Particle(i).CurrentColor(2).B)
                                    .Particle(i).rgb_list(2) = D3DColorARGB(.Particle(i).CurrentColor(3).A, .Particle(i).CurrentColor(3).R, .Particle(i).CurrentColor(3).G, .Particle(i).CurrentColor(3).B)
                                    .Particle(i).rgb_list(3) = D3DColorARGB(.Particle(i).CurrentColor(4).A, .Particle(i).CurrentColor(4).R, .Particle(i).CurrentColor(4).G, .Particle(i).CurrentColor(4).B)
                                    
                                    
                                    .Particle(i).Screen_X = x + .Particle(i).x
                                    .Particle(i).Screen_Y = y + .Particle(i).y
                                End If
                            Else
                                .Particle(i).DelayCounter = .Particle(i).DelayCounter + 1
                            End If
                        Else
                            If Not .Particle(i).Used Then
                                Call particleUpdate(ParticleGroup_Index, i, x, y, Time, offset_x, offset_y)
                            End If
                        End If
                    Next i
                Else
                    .Particle_Frame_Counter = .Particle_Frame_Counter + particle_timer
                End If
            End With
        Else
            particleGroupDestroy (ParticleGroup_Index)
        End If
    End With
End Function
Public Sub particleGroupRender(ByVal Particle_Group_Index As Integer, ByVal x As Integer, ByVal y As Integer, Optional ByVal offset_x As Single, Optional ByVal offset_y As Single)
    Dim i As Integer
    Dim rgb_list(3) As Long
    Dim Size As Integer
    
    Call particlesGroupUpdate(Particle_Group_Index, x, y, offset_x, offset_y)
    
    With particle_group_list(Particle_Group_Index).ParticleEmisor
        For i = 1 To .Particle_count
            'If is destroyed...
            If .Particle(i).Alive Then
                If Not .Particle(i).texture_index = 0 Then
                    Size = .texture_size(.Particle(i).texture_index)
                    dxTextureRender .texture_index(.Particle(i).texture_index), .Particle(i).Screen_X, .Particle(i).Screen_Y, Size, Size, .Particle(i).rgb_list(), 0, 0, Size, Size, .alpha_blend, .Particle(i).angle
                End If
            End If
        Next i
    End With
End Sub

Private Sub particleGetArrayPos(ByRef Particle_Index As Integer)
    Dim i As Byte
    
    If particle_group_count = 0 Then
        Particle_Index = 0
        Exit Sub
    End If
    i = 1
    Do While particle_group_list(i).Active
        If i >= particle_group_count Then
            Particle_Index = 0
            Exit Sub
        Else
            i = i + 1
        End If
    Loop
    
    Particle_Index = i
End Sub
Public Sub particleGroupDestroy(ByVal ParticleGroup_Index As Integer)
    With particle_group_list(ParticleGroup_Index)
        destroyEffectParticle .effectIndex
        
        .Active = 0
        .effectIndex = 0
    End With
End Sub
Public Sub particleGroupDestroyAll()
    Dim i As Byte
    If particle_group_count = 0 Then Exit Sub 'Particles are already destroyed
    For i = 1 To particle_group_count
        With particle_group_list(i)
            particleGroupDestroy (i)
        End With
    Next i
    particle_group_count = 0
End Sub

Public Sub loadParticles()
    Dim particle_path As String
    Dim i As Integer
    Dim aux As String
    Dim j As Byte
    
    particle_path = resource_path & PATH_INIT & "\particles.ini"
    
    If Not General_File_Exists(particle_path, vbArchive) Then Exit Sub
    
    particle_group_data_count = General_Var_Get(particle_path, "INIT", "Total")
    
    ReDim particle_group_data(1 To particle_group_data_count)
    
    'On Error Resume Next
    
    For i = 1 To particle_group_data_count
        With particle_group_data(i)
            With .ParticleEmisor
                .ParticleGroup_Type = Val(General_Var_Get(particle_path, i, "Tipo"))
                
                .alpha_blend = CBool(General_Var_Get(particle_path, i, "AlphaBlend"))
                .Bounce_Strength = Val(General_Var_Get(particle_path, i, "Bounce_Strength"))
                .Bounce_Y = Val(General_Var_Get(particle_path, i, "BounceY"))
                If .Bounce_Y = 0 Then .Bounce_Y = 16
                
                .ColorVariation = Val(General_Var_Get(particle_path, i, "ColorVariation"))
                
                If .ColorVariation Then
                    For j = 1 To 4
                        aux = General_Var_Get(particle_path, i, "ColorSet" & j)
                        .StartColor(j) = D3DColorValueGet(Val(general_field_read(1, aux, Asc(","))), Val(general_field_read(2, aux, Asc(","))), general_field_read(3, aux, Asc(",")), Val(general_field_read(4, aux, Asc(","))))
                        aux = General_Var_Get(particle_path, i, "ColorEnd" & j)
                        .EndColor(j) = D3DColorValueGet(Val(general_field_read(1, aux, Asc(","))), Val(general_field_read(2, aux, Asc(","))), Val(general_field_read(3, aux, Asc(","))), Val(general_field_read(4, aux, Asc(","))))
                    Next j
                Else
                    For j = 1 To 4
                        aux = General_Var_Get(particle_path, i, "ColorSet" & j)
                        .StartColor(j) = D3DColorValueGet(Val(general_field_read(1, aux, Asc(","))), Val(general_field_read(2, aux, Asc(","))), Val(general_field_read(3, aux, Asc(","))), Val(general_field_read(4, aux, Asc(","))))
                    Next j
                End If
                
                .Particle_count = Val(General_Var_Get(particle_path, i, "NumOfParticles"))
                .Friction = Val(General_Var_Get(particle_path, i, "Friction"))
                .PLt1 = Val(General_Var_Get(particle_path, i, "Life1"))
                .PLt2 = Val(General_Var_Get(particle_path, i, "Life2"))
                
                'Accelerations
                .Wind = Val((General_Var_Get(particle_path, i, "Wind")))
                .Gravity = Val(General_Var_Get(particle_path, i, "Gravity"))
                
                'Rotation
                .Spin = CByte(General_Var_Get(particle_path, i, "Spin"))
                .SpinH = Val(General_Var_Get(particle_path, i, "Spin_SpeedH"))
                .SpinL = Val(General_Var_Get(particle_path, i, "Spin_SpeedL"))
                
                .texture_count = Val(General_Var_Get(particle_path, i, "NumGrhs"))
                If .texture_count > 0 Then
                    ReDim .texture_index(1 To .texture_count) As Integer
                    ReDim .texture_size(1 To .texture_count) As Integer
                    For j = 1 To .texture_count
                        .texture_index(j) = Val(general_field_read(j, General_Var_Get(particle_path, i, "Grh_List"), Asc(",")))
                        .texture_size(j) = Val(general_field_read(j, General_Var_Get(particle_path, i, "Size_List"), Asc(",")))
                    Next j
                End If
                
                .vX1 = Val(General_Var_Get(particle_path, i, "VecX1"))
                .vX2 = Val(General_Var_Get(particle_path, i, "VecX2"))
                .vY1 = Val(General_Var_Get(particle_path, i, "VecY1"))
                .vY2 = Val(General_Var_Get(particle_path, i, "VecY2"))
                
                'Particle Startup position
                .x1 = Val(General_Var_Get(particle_path, i, "X1"))
                .x2 = Val(General_Var_Get(particle_path, i, "X2"))
                .Y1 = Val(General_Var_Get(particle_path, i, "Y1"))
                .Y2 = Val(General_Var_Get(particle_path, i, "Y2"))
                
                'Speed
                .Particle_Speed = Val(General_Var_Get(particle_path, i, "Speed"))
                If .Particle_Speed = 0 Then .Particle_Speed = 0.5
                
                'For circle Effects
                .Ratio = Val(General_Var_Get(particle_path, i, "Radio"))
                .RatioFriction = Val(General_Var_Get(particle_path, i, "RatioFriction"))
                .RatioVariation = Val(General_Var_Get(particle_path, i, "RatioVariation"))
                
                'Shaking
                .MoveX = Val(General_Var_Get(particle_path, i, "XMove"))
                .MoveY = Val(General_Var_Get(particle_path, i, "YMove"))
                
                .move_x1 = Val(General_Var_Get(particle_path, i, "move_x1"))
                .move_y1 = Val(General_Var_Get(particle_path, i, "move_y1"))
                .move_x2 = Val(General_Var_Get(particle_path, i, "move_x2"))
                .move_y2 = Val(General_Var_Get(particle_path, i, "move_y2"))
                
                'Variables de efectos especiales
                .TargetRatio = Val(General_Var_Get(particle_path, i, "TargetRatio"))
                .StopAtTargetRatio = Val((General_Var_Get(particle_path, i, "StopAtTargetRatio")))
                .KillWhenAtTarget = Val((General_Var_Get(particle_path, i, "KillWhenAtTarget")))
                .Target_X = Val(General_Var_Get(particle_path, i, "TargetX"))
                .Target_Y = Val(General_Var_Get(particle_path, i, "TargetY"))
                
                'Delay
                .Delay1 = Val(General_Var_Get(particle_path, i, "Delay1"))
                .Delay2 = Val(General_Var_Get(particle_path, i, "Delay2"))
                
                '.defaultLifeTime = Val(General_Var_Get(particle_path, i, "defaultLifeTime"))
                
                .CurrentRatio = .Ratio
                .ParticlesLeft = .Particle_count
                ReDim .Particle(0 To .Particle_count)
            End With
        End With
    Next i
End Sub

Private Sub particle1Reset(GroupIndex As Integer, particleIndex As Integer)
    With particle_group_list(GroupIndex).ParticleEmisor
        .Particle(particleIndex).x = General_Random_Number(.x1, .x2)
        .Particle(particleIndex).y = General_Random_Number(.Y1, .Y2)
    End With
End Sub
Private Sub particle2Reset(GroupIndex As Integer, particleIndex As Integer)
    Dim angle As Single

    With particle_group_list(GroupIndex).ParticleEmisor
        angle = particleIndex * (360 / .Particle_count) * DegreeToRadian

        .Particle(particleIndex).x = .x1 - (Sin(angle) * .CurrentRatio)
        .Particle(particleIndex).y = .Y1 + (Cos(angle) * .CurrentRatio)
    End With
End Sub
Private Sub particle3Reset(GroupIndex As Integer, particleIndex As Integer)
    Dim R As Single
    Dim angle As Single
    
    With particle_group_list(GroupIndex).ParticleEmisor
        angle = 360 / .Particle_count * particleIndex * DegreeToRadian
        R = Rnd

        .Particle(particleIndex).x = .x1 - (Sin(angle) * (R * .CurrentRatio))
        .Particle(particleIndex).y = .Y1 + (Cos(angle) * (R * .CurrentRatio))
    End With
End Sub
Private Sub particle4Reset(GroupIndex As Integer, particleIndex As Integer)
    Dim R As Single

    With particle_group_list(GroupIndex).ParticleEmisor
        R = Sin(20 / (particleIndex + 1)) * .CurrentRatio
    
        .Particle(particleIndex).x = R * Cos(particleIndex)
        .Particle(particleIndex).y = R * Sin(particleIndex)
    End With
End Sub

Private Sub particle5Reset(GroupIndex As Integer, particleIndex As Integer)
    Dim R As Single

    With particle_group_list(GroupIndex).ParticleEmisor
        R = .CurrentRatio + Rnd * 15 * Cos(2 * particleIndex)
        
        .Particle(particleIndex).x = R * Cos(particleIndex)
        .Particle(particleIndex).y = R * Sin(particleIndex)
    End With
End Sub

Private Sub particleUpdate(ByVal GroupIndex As Integer, ByVal i As Integer, ByVal x As Integer, ByVal y As Integer, ByVal Time As Long, Optional ByVal offset_x As Single, Optional ByVal offset_y As Single)
    Dim angle As Single 'Used to calculate ratio variation.
    
    'Change color
    With particle_group_list(GroupIndex)
        With .ParticleEmisor
            .Particle(i).Created = .Particle(i).Created + 1
            
            If .ColorVariation Then
                Call D3DXColorLerp(.Particle(i).CurrentColor(1), .StartColor(1), .EndColor(1), (.Particle(i).Created) / .Particle(i).Particle_LifeTime)
                Call D3DXColorLerp(.Particle(i).CurrentColor(2), .StartColor(2), .EndColor(2), (.Particle(i).Created) / .Particle(i).Particle_LifeTime)
                Call D3DXColorLerp(.Particle(i).CurrentColor(3), .StartColor(3), .EndColor(3), (.Particle(i).Created) / .Particle(i).Particle_LifeTime)
                Call D3DXColorLerp(.Particle(i).CurrentColor(4), .StartColor(4), .EndColor(4), (.Particle(i).Created) / .Particle(i).Particle_LifeTime)
                
                .Particle(i).rgb_list(0) = D3DColorARGB(.Particle(i).CurrentColor(1).A, .Particle(i).CurrentColor(1).R, .Particle(i).CurrentColor(1).G, .Particle(i).CurrentColor(1).B)
                .Particle(i).rgb_list(1) = D3DColorARGB(.Particle(i).CurrentColor(2).A, .Particle(i).CurrentColor(2).R, .Particle(i).CurrentColor(2).G, .Particle(i).CurrentColor(2).B)
                .Particle(i).rgb_list(2) = D3DColorARGB(.Particle(i).CurrentColor(3).A, .Particle(i).CurrentColor(3).R, .Particle(i).CurrentColor(3).G, .Particle(i).CurrentColor(3).B)
                .Particle(i).rgb_list(3) = D3DColorARGB(.Particle(i).CurrentColor(4).A, .Particle(i).CurrentColor(4).R, .Particle(i).CurrentColor(4).G, .Particle(i).CurrentColor(4).B)
            End If

            'Do Shaking
            If .MoveX Then .Particle(i).vX = General_Random_Number(.move_x1, .move_x2)
            If .MoveY Then .Particle(i).vX = General_Random_Number(.move_y1, .move_y2)
                                                                                                      
            'Do Gravity
            If .Gravity Then
                .Particle(i).vY = .Particle(i).vY + .Gravity
            End If
                                
            If .Bounce_Strength <> 0 Then
                If .Particle(i).y + .Particle(i).Moved_Y > .Bounce_Y Then
                    .Particle(i).vY = .Bounce_Strength
                End If
            End If
            
            If .Spin Then .Particle(i).angle = .Particle(i).angle + General_Random_Number(.SpinL, .SpinH) / 100
                                   
            If .Wind Then
                .Particle(i).vX = .Particle(i).vX + (.Wind / .RatioFriction)
            End If
                        
            If .RatioVariation <> 0 Then
                .CurrentRatio = .CurrentRatio + .RatioVariation / .Friction
                angle = i * (360 / .Particle_count) * DegreeToRadian
                .Particle(i).x = .x1 - (Sin(angle) * .CurrentRatio)
                .Particle(i).y = .Y1 + (Cos(angle) * .CurrentRatio)
                
                If .StopAtTargetRatio Then
                     If .CurrentRatio >= .TargetRatio Then
                        .RatioVariation = 0 'Stop variation
                        .CurrentRatio = .TargetRatio
                     End If
                Else
                    If .CurrentRatio >= .TargetRatio Then
                        .CurrentRatio = .Ratio
                    End If
                End If
            End If
            
            'Move our particle
            .Particle(i).Moved_X = .Particle(i).Moved_X + .Particle(i).vX / .Friction + offset_x
            .Particle(i).Moved_Y = .Particle(i).Moved_Y + .Particle(i).vY / .Friction + offset_y
            
            .Particle(i).Screen_X = x + .Particle(i).x + .Particle(i).Moved_X
            .Particle(i).Screen_Y = y + .Particle(i).y + .Particle(i).Moved_Y
        End With
    End With
End Sub

Private Function particleGroupCheckPermanency(ByVal ParticleGroup_Index As Integer, ByVal Time As Long) As Boolean
    
    particleGroupCheckPermanency = False
    
    With particle_group_list(ParticleGroup_Index)
        If .ParticleGroup_Lifetime > Time - .Created Or .ParticleGroup_Lifetime = -1 And Not .ParticleEmisor.StartParticlesDestroy Then
            If .ParticleEmisor.ParticlesLeft > 0 Then
                particleGroupCheckPermanency = True
            End If
        Else
            If .ParticleEmisor.ParticlesLeft > 0 Then
                particleGroupCheckPermanency = True
                .ParticleEmisor.StartParticlesDestroy = True
            End If
        End If
    End With
End Function

Public Function particleGroupEnd(ByVal ParticleGroup_Index As Integer)
    With particle_group_list(ParticleGroup_Index)
        .ParticleEmisor.StartParticlesDestroy = True
    End With
End Function
