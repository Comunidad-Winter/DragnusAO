Attribute VB_Name = "modGrh"
'Holds data about where a bmp can be found,
'How big it is and animation info
Public Type Grh_Data
    Active As Boolean
    texture_index As Long
    Src_X As Integer
    Src_Y As Integer
    src_width As Integer
    src_height As Integer
    
    frame_count As Integer
    frame_list(1 To 25) As Long
    frame_speed As Single
End Type

'Points to a Grh_Data and keeps animation info
Public Type grh
    grh_index As Integer
    alpha_blend As Boolean
    angle As Single
    frame_speed As Single
    frame_counter As Single
    Started As Boolean
    LoopTimes As Integer
End Type

'Grh Data Array
Dim grh_list() As Grh_Data
Dim grh_count As Long

Dim AnimBaseSpeed As Single
Dim timer_ticks_per_frame As Single

Dim base_tile_size As Integer

Public Sub Grh_Initialize(ByRef grh As grh, ByVal grh_index As Long, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single, Optional ByVal Started As Byte = 2, Optional ByVal LoopTimes As Integer = LoopAdEternum)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    If grh_index <= 0 Then Exit Sub

    'Copy of parameters
    grh.grh_index = grh_index
    grh.alpha_blend = alpha_blend
    grh.angle = angle
    grh.LoopTimes = LoopTimes
    
    'Start it if it's a animated grh
    If Started = 2 Then
        If grh_list(grh.grh_index).frame_count > 1 Then
            grh.Started = True
        Else
            grh.Started = False
        End If
    Else
        grh.Started = Started
    End If
    
    'Set frame counters
    grh.frame_counter = 1
    
    grh.frame_speed = grh_list(grh.grh_index).frame_speed
End Sub

Private Sub Grh_Load_All()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Loads Grh.dat
'**************************************************************
On Error GoTo ErrorHandler
    Dim grh As Long
    Dim Frame As Long
    Dim FileVersion As Long
    
    Dim inipath As String
    inipath = resource_path & PATH_INIT
    
    
    
    'Open files
    Open inipath & "\graficos.ind" For Binary As #1
    Seek #1, 1
    
    Get #1, , FileVersion
    
    'Get number of grhs
    Get #1, , grh_count

    'Resize arrays
    ReDim grh_list(1 To grh_count) As Grh_Data
    'Fill Grh List
    
    'Get first Grh Number
    Get #1, , grh
    
    Do Until grh <= 0
        
        grh_list(grh).Active = True
        
        'Get number of frames
        Get #1, , grh_list(grh).frame_count
        If grh_list(grh).frame_count <= 0 Then GoTo ErrorHandler
        
        If grh_list(grh).frame_count > 1 Then
        
            'Read a animation GRH set
            For Frame = 1 To grh_list(grh).frame_count
            
                Get #1, , grh_list(grh).frame_list(Frame)
                If grh_list(grh).frame_list(Frame) <= 0 Or grh_list(grh).frame_list(Frame) > grh_count Then GoTo ErrorHandler
            
            Next Frame
        
            Get #1, , grh_list(grh).frame_speed

            If grh_list(grh).frame_speed = 0 Then GoTo ErrorHandler
            
            'Compute width and height
            grh_list(grh).src_height = grh_list(grh_list(grh).frame_list(1)).src_height
            If grh_list(grh).src_height <= 0 Then GoTo ErrorHandler
            
            grh_list(grh).src_width = grh_list(grh_list(grh).frame_list(1)).src_width
            If grh_list(grh).src_width <= 0 Then GoTo ErrorHandler
        
        Else
        
            'Read in normal GRH data
            Get #1, , grh_list(grh).texture_index
            If grh_list(grh).texture_index <= 0 Then GoTo ErrorHandler
            
            Get #1, , grh_list(grh).Src_X
            If grh_list(grh).Src_X < 0 Then GoTo ErrorHandler
            
            Get #1, , grh_list(grh).Src_Y
            If grh_list(grh).Src_Y < 0 Then GoTo ErrorHandler
                
            Get #1, , grh_list(grh).src_width
            If grh_list(grh).src_width <= 0 Then GoTo ErrorHandler
            
            Get #1, , grh_list(grh).src_height
            If grh_list(grh).src_height <= 0 Then GoTo ErrorHandler
            
            grh_list(grh).frame_list(1) = grh
                
        End If
    
        'Get Next Grh Number
        Get #1, , grh
    
    Loop
    '************************************************
    
    Close #1
Exit Sub
ErrorHandler:
    Close #1
    MsgBox "Error while loading the grh.dat! Stopped at GRH number: " & grh
End Sub


Public Sub Grh_Render(ByRef grh As grh, ByVal Screen_X As Long, ByVal Screen_Y As Long, ByRef rgb_list() As Long, Optional ByVal center As Boolean, Optional ByVal vCenter As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'
'**************************************************************
    Dim tile_width As Single
    Dim tile_height As Single
    Dim grh_index As Long
    
    If grh.grh_index = 0 Then Exit Sub
    
    'Animation
    If grh.Started Then
        grh.frame_counter = grh.frame_counter + (timer_ticks_per_frame * grh.frame_speed / 1000)
        If grh.frame_counter > grh_list(grh.grh_index).frame_count Then
            If grh.LoopTimes < 2 Then
                grh.frame_counter = 1
                grh.Started = False
            Else
                grh.frame_counter = 1
                If grh.LoopTimes <> LoopAdEternum Then
                    grh.LoopTimes = grh.LoopTimes - 1
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    If grh.frame_counter <= 0 Then grh.frame_counter = 1
    grh_index = grh_list(grh.grh_index).frame_list(grh.frame_counter)
    
    If grh_index = 0 Then Exit Sub 'This is an error condition
    
    'Center Grh over X,Y pos
    If center Then
        tile_width = grh_list(grh_index).src_width / base_tile_size
        tile_height = grh_list(grh_index).src_height / base_tile_size
        If tile_width <> 1 Then
            Screen_X = Screen_X - Int(tile_width * base_tile_size / 2) + base_tile_size / 2
        End If
        If tile_height <> 1 Then
            If vCenter Then
                Screen_Y = Screen_Y - Int(tile_height * base_tile_size / 2) + base_tile_size / 2
            Else
                Screen_Y = Screen_Y - Int(tile_height * base_tile_size) + base_tile_size
            End If
        End If
    End If
    
    'Draw it to device
    dxTextureRender grh_list(grh_index).texture_index, _
        Screen_X, Screen_Y, _
        grh_list(grh_index).src_width, grh_list(grh_index).src_height, _
        rgb_list, _
        grh_list(grh_index).Src_X, grh_list(grh_index).Src_Y, _
        grh_list(grh_index).src_width, grh_list(grh_index).src_height, _
        grh.alpha_blend, _
        grh.angle
End Sub

Public Sub Grh_Render_To_Hdc(ByVal grh_index As Long, desthdc As Long, ByVal Screen_X As Long, ByVal Screen_Y As Long, Optional transparent As Boolean = False, Optional ByVal center As Boolean = False, Optional ByVal vCenter As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/02/2003
'This method is SLOW... Don't use in a loop if you care about
'speed!
'*************************************************************
    If Grh_Check(grh_index) = False Then
        Exit Sub
    End If


    Dim Src_X As Long
    Dim Src_Y As Long
    Dim src_width As Long
    Dim src_height As Long
    Dim file_index As Long

    'If it's animated switch grh_index to first frame
    If grh_list(grh_index).frame_count <> 1 Then
        grh_index = grh_list(grh_index).frame_list(1)
    End If

    file_index = grh_list(grh_index).texture_index
    Src_X = grh_list(grh_index).Src_X
    Src_Y = grh_list(grh_index).Src_Y
    src_width = grh_list(grh_index).src_width
    src_height = grh_list(grh_index).src_height
        
    Dim tile_width As Integer
    Dim tile_height As Integer
        
    'Center Grh over X,Y pos
    If center Then
        tile_width = grh_list(grh_index).src_width / base_tile_size
        tile_height = grh_list(grh_index).src_height / base_tile_size
        If tile_width <> 1 Then
            Screen_X = Screen_X - Int(tile_width * base_tile_size / 2) + base_tile_size / 2
        End If
        If tile_height <> 1 Then
            If vCenter Then
                Screen_Y = Screen_Y - Int(tile_height * base_tile_size / 2) + base_tile_size / 2
            Else
                Screen_Y = Screen_Y - Int(tile_height * base_tile_size) + base_tile_size
            End If
        End If
    End If
    
    
    Call dxTextureToHdcRender(file_index, desthdc, Screen_X, Screen_Y, Src_X, Src_Y, src_width, src_height, transparent)
End Sub

Public Function GUI_Grh_Render(ByVal grh_index As Long, ByVal x As Integer, ByVal y As Integer, Optional ByVal angle As Single, Optional ByVal alpha_blend As Boolean, Optional ByVal color As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/15/2003
'
'**************************************************************
    Dim temp_grh As grh
    Dim rpg_list(3) As Long

    If Grh_Check(grh_index) = False Then
        Exit Function
    End If

    rpg_list(0) = color
    rpg_list(1) = color
    rpg_list(2) = color
    rpg_list(3) = color

    Grh_Initialize temp_grh, grh_index, alpha_blend, angle
    
    Grh_Render temp_grh, x, y, rpg_list
    
    GUI_Grh_Render = True
End Function

Public Sub Animations_Initialize(ByVal AnimationSpeed As Single, ByVal tile_size As Integer)
    Grh_Load_All
    base_tile_size = tile_size
    AnimBaseSpeed = AnimationSpeed
End Sub

Public Sub AnimSpeedCalculate(ByVal timer_elapsed_time As Single)
    timer_ticks_per_frame = AnimBaseSpeed * timer_elapsed_time
End Sub

Public Function Grh_Check(ByVal grh_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check grh_index
    If grh_index > 0 And grh_index <= grh_count Then
        If grh_list(grh_index).Active Then
            Grh_Check = True
        End If
    End If
End Function
