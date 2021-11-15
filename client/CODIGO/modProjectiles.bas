Attribute VB_Name = "modProjectiles"
Option Explicit

Private Const MAXPROJECTILES As Byte = 15
Private Const PROJECTILESCROLL As Byte = 12

Private Const DegreeToRadian As Single = 0.0174532925

Private Type tProjectile
    projectileID As Integer
    TargetX As Integer
    TargetY As Integer
    x As Integer
    y As Integer
    offSetX As Single
    offSetY As Single
    effectIndex As Integer
    targetCharIndex As Integer
    projectileIndex As Integer
    projectileGrh As grh
    onHitEffect As Integer
    onHitTarget As Byte
End Type

Private Type tProjectileData
    effectID As Integer
    grhIndex As Integer
End Type

Dim projectileDataList() As tProjectileData
Dim projectileList(1 To MAXPROJECTILES) As tProjectile
Dim lastProjectile As Integer

Dim projectileCount As Integer


Private Function checkIndex(ByVal projectileIndex As Integer) As Boolean
    If projectileIndex > 0 And projectileIndex <= lastProjectile Then _
        checkIndex = True
End Function

Public Function initProjectiles() As Integer
    projectileDataLoad
End Function

Public Function updateProjectileTarget(ByVal projectileIndex As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Integer
    If checkIndex(projectileIndex) Then
        projectileList(projectileIndex).TargetX = TargetX
        projectileList(projectileIndex).TargetY = TargetY
    End If
End Function

Public Sub renderProjectiles(ByVal viewX As Integer, ByVal viewY As Integer, ByVal viewWidth As Integer, ByVal viewHeight As Integer, ByVal offSetX As Single, ByVal offSetY As Single, ByRef rgbList() As Long)
    Dim projectileIndex As Integer
    Dim screenX As Single
    Dim screenY As Single
    
    For projectileIndex = 1 To lastProjectile
        With projectileList(projectileIndex)
            If .x > viewX - viewWidth / 2 And .y > viewY - viewHeight / 2 And .x < viewX + viewWidth / 2 And .y < viewY + viewHeight / 2 Then
                screenX = (.x - viewX + viewWidth / 2) * 32 + .offSetX - offSetX - 16
                screenY = (.y - viewY + viewHeight / 2) * 32 + .offSetY - offSetY - 16
                
                If .effectIndex > 0 Then
                    renderEffect .effectIndex, screenX, screenY, rgbList()
                    
                    If Not effectDestroyed(.effectIndex) Then
                        .effectIndex = 0
                    End If
                End If
                
                If .projectileGrh.grh_index > 0 Then
                    'Update the position
                    .projectileGrh.angle = DegreeToRadian * projectileGetAngle(.x, .y, .TargetX, .TargetY)
                    Call Grh_Render(.projectileGrh, screenX, screenY, rgbList(), True, True)
                End If
            End If
        End With
    Next projectileIndex
End Sub

Public Sub updateProjectiles(ByVal timerTicks)
    Dim i As Byte
    
    For i = 1 To lastProjectile
        updateProjectile i, timerTicks
    Next i
End Sub

Private Sub updateProjectile(ByVal projectileIndex As Integer, ByVal timerTicks As Single)
    If checkIndex(projectileIndex) Then
        With projectileList(projectileIndex)
            If Abs(.TargetX - .x) > 0 Or Abs(.TargetY - .y) > 0 Then
                If Abs(.TargetX - .x) > 0 Then
                    .offSetX = .offSetX + PROJECTILESCROLL * timerTicks * Sgn(.TargetX - .x)
                    If Abs(.offSetX) > 32 Then
                        .x = .x + Sgn(.TargetX - .x)
                        .offSetX = 0
                    End If
                End If
                
                If Abs(.TargetY - .y) > 0 Then
                    .offSetY = .offSetY + PROJECTILESCROLL * timerTicks * Sgn(.TargetY - .y)
                    If Abs(.offSetY) > 32 Then
                        .y = .y + Sgn(.TargetY - .y)
                        .offSetY = 0
                    End If
                End If
                
            Else
                Call destroyProjectile(projectileIndex)
            End If
        End With
    End If
End Sub
Public Sub destroyProjectile(ByVal projectileIndex As Integer)
    Dim i As Integer
    If checkIndex(projectileIndex) Then
        projectileList(projectileIndex).projectileID = 0
        
        If projectileList(projectileIndex).effectIndex Then
            destroyEffect projectileList(projectileIndex).effectIndex
            projectileList(projectileIndex).effectIndex = 0
        End If
        
        If projectileList(projectileIndex).targetCharIndex > 0 Then
            Engine.charDestroyProjectile projectileList(projectileIndex).targetCharIndex, projectileList(projectileIndex).projectileIndex
        Else
            Engine.mapDestroyProjectile projectileList(projectileIndex).TargetX, projectileList(projectileIndex).TargetY, projectileList(projectileIndex).projectileIndex
        End If
        
        If projectileList(projectileIndex).onHitEffect > 0 Then
            If projectileList(projectileIndex).onHitTarget = 0 Then
                Engine.mapCreateEffect projectileList(projectileIndex).TargetX, projectileList(projectileIndex).TargetY, projectileList(projectileIndex).onHitEffect
            Else
                If projectileList(projectileIndex).targetCharIndex > 0 Then
                    Engine.charCreateEffect projectileList(projectileIndex).targetCharIndex, projectileList(projectileIndex).onHitEffect
                End If
            End If
        End If
        
        projectileList(projectileIndex).projectileGrh.grh_index = 0
        
        If projectileIndex = lastProjectile Then
            Do While projectileList(lastProjectile).projectileID < 1
                lastProjectile = lastProjectile - 1
                If lastProjectile < 1 Then Exit Do
            Loop
        End If
    End If
End Sub

Public Function projectileDestroyed(ByVal projectileIndex As Integer) As Boolean
    If checkIndex(projectileIndex) Then
        If projectileList(projectileIndex).projectileID = 0 Then
            projectileDestroyed = True
        End If
    End If
End Function

Private Function getArrayPos() As Integer
    Dim i As Integer
    
    Do While i < lastProjectile And getArrayPos < 1
        i = i + 1
        If projectileList(i).projectileID < 1 Then _
            getArrayPos = i
    Loop
End Function

Public Function createProjectile(ByVal projectilID As Integer, ByVal x As Integer, ByVal y As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer, ByVal projectileIndex As Integer, Optional ByVal onHitEffect As Integer = 0, Optional ByVal onHitTarget As Byte = 0, Optional ByVal targetCharIndex As Integer = 0) As Integer
    Dim index As Integer
    
    index = getArrayPos()
    
    If index > 0 Then
        projectileList(index).TargetX = TargetX
        projectileList(index).TargetY = TargetY
        projectileList(index).x = x
        projectileList(index).y = y
        
        projectileList(index).onHitEffect = onHitEffect
        projectileList(index).onHitTarget = onHitTarget
        
        projectileList(index).projectileIndex = projectileIndex
        
        projectileList(index).targetCharIndex = targetCharIndex
        
        projectileList(index).projectileID = projectilID
        
        If projectileDataList(projectilID).effectID > 0 Then _
            projectileList(index).effectIndex = createEffect(projectileDataList(projectilID).effectID, -1)
            
        Grh_Initialize projectileList(index).projectileGrh, projectileDataList(projectilID).grhIndex
    Else
        If lastProjectile < MAXPROJECTILES Then
            lastProjectile = lastProjectile + 1
            
            projectileList(lastProjectile).TargetX = TargetX
            projectileList(lastProjectile).TargetY = TargetY
            projectileList(lastProjectile).x = x
            projectileList(lastProjectile).y = y
            
            projectileList(lastProjectile).onHitEffect = onHitEffect
            projectileList(lastProjectile).onHitTarget = onHitTarget
            
            projectileList(lastProjectile).projectileIndex = projectileIndex
            
            projectileList(lastProjectile).targetCharIndex = targetCharIndex

            projectileList(lastProjectile).projectileID = projectilID
            
            If projectileDataList(projectilID).effectID > 0 Then _
                projectileList(index).effectIndex = createEffect(projectileDataList(projectilID).effectID, -1)
            
            Grh_Initialize projectileList(lastProjectile).projectileGrh, projectileDataList(projectilID).grhIndex
            
            index = lastProjectile
        End If
    End If
    
    createProjectile = index
End Function

Public Sub projectileDataLoad()
    Dim i As Integer
    
    projectileCount = General_Var_Get(resource_path & PATH_INIT & "\projectiles.dat", "INIT", "projectileCount")
    ReDim projectileDataList(1 To projectileCount) As tProjectileData
    If projectileCount < 0 Then Exit Sub
    
    For i = 1 To projectileCount
        projectileDataList(i).grhIndex = Val(General_Var_Get(resource_path & PATH_INIT & "\projectiles.dat", "PROJECTILE" & i, "grhIndex"))
        projectileDataList(i).effectID = Val(General_Var_Get(resource_path & PATH_INIT & "\projectiles.dat", "PROJECTILE" & i, "effectID"))
    Next i
End Sub

Public Function projectileGetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'More info: http://www.vbgore.com/GameClient.TileEngine.projectileGetAngle
'************************************************************
Dim SideA As Single
Dim SideC As Single

    On Error GoTo ErrOut

    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then

        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            projectileGetAngle = 90

            'Check for going left (270 degrees)
        Else
            projectileGetAngle = 270
        End If

        'Exit the function
        Exit Function

    End If

    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then

        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            projectileGetAngle = 360

            'Check for going down (180 degrees)
        Else
            projectileGetAngle = 180
        End If

        'Exit the function
        Exit Function

    End If

    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)

    'Side B = CenterY

    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)

    'Calculate the angle
    projectileGetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    projectileGetAngle = (Atn(-projectileGetAngle / Sqr(-projectileGetAngle * projectileGetAngle + 1)) + 1.5708) * 57.29583

    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then projectileGetAngle = 360 - projectileGetAngle

    'Exit function

Exit Function

    'Check for error
ErrOut:

    'Return a 0 saying there was an error
    projectileGetAngle = 0

Exit Function

End Function
