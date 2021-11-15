Attribute VB_Name = "modDXEngine"
Option Explicit

Private Const DegreeToRadian As Single = 0.0174532925

'***************************
'Estructures
'***************************
'This structure describes a transformed and lit vertex.
Private Type TLVERTEX
    x As Single
    y As Single
    z As Single
    rhw As Single
    color As Long
    specular As Long
    tu As Single
    tv As Single
End Type

Private Type tGraphicChar
    Src_X As Integer
    Src_Y As Integer
End Type

Private Type tGraphicFont
    texture_index As Long
    Caracteres(0 To 255) As tGraphicChar 'Ascii Chars
    Char_Size As Byte 'In pixels
End Type

Private Type DXFont
    dFont As D3DXFont
    Size As Integer
End Type

Public Enum FontAlignment
    fa_center = DT_CENTER
    fa_top = DT_TOP
    fa_left = DT_LEFT
    fa_topleft = DT_TOP Or DT_LEFT
    fa_bottomleft = DT_BOTTOM Or DT_LEFT
    fa_bottom = DT_BOTTOM
    fa_right = DT_RIGHT
    fa_bottomright = DT_BOTTOM Or DT_RIGHT
    fa_topright = DT_TOP Or DT_RIGHT
End Enum

'***************************
'Variables
'***************************
'Major DX Objects
Public dx As DirectX8
Public d3d As Direct3D8
Public ddevice As Direct3DDevice8
Public d3dx As D3DX8

Dim d3dpp As D3DPRESENT_PARAMETERS

'Texture Manager for Dinamic Textures
Dim DXPool As New clsTextureManager

'Main form handle
Dim form_hwnd As Long

'Display variables
Dim screen_hwnd As Long
Dim screen_width As Long
Dim screen_height As Long

'FPS Counters
Dim fps_last_time As Long 'When did we last check the frame rate?
Dim fps_frame_counter As Long 'How many frames have been drawn
Dim fps As Long 'What the current frame rate is.....

Dim engine_render_started As Boolean

'Graphic Font List
Dim gfont_list() As tGraphicFont
Dim gfont_count As Long
Dim gfont_last As Long

'Font List
Private font_list() As DXFont
Private font_count As Integer


'***************************
'Constants
'***************************
'Engine
Private Const COLOR_KEY As Long = &HFF000000
Private Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
'PI
Private Const PI As Single = 3.14159265358979

'Old fashion BitBlt functions
Private Const SRCCOPY = &HCC0020
Private Const SRCPAINT = &HEE0086
Private Const SRCAND = &H8800C6
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcsrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'Initialization
Public Function initDirectX(ByVal f_hwnd As Long, ByVal s_hwnd As Long, ByVal windowed As Boolean)
'On Error GoTo errhandler
    Dim d3dcaps As D3DCAPS8
    Dim d3ddm As D3DDISPLAYMODE
    
    initDirectX = True
    
    'Main display
    screen_hwnd = s_hwnd
    form_hwnd = f_hwnd
    
    '*******************************
    'Initialize root DirectX8 objects
    '*******************************
    Set dx = New DirectX8
    'Create the Direct3D object
    Set d3d = dx.Direct3DCreate
    'Create helper class
    Set d3dx = New D3DX8
    
    '*******************************
    'Initialize video device
    '*******************************
    Dim DevType As CONST_D3DDEVTYPE
    DevType = D3DDEVTYPE_HAL
    'Get the capabilities of the Direct3D device that we specify. In this case,
    'we'll be using the adapter default (the primiary card on the system).
    Call d3d.GetDeviceCaps(D3DADAPTER_DEFAULT, DevType, d3dcaps)
    'Grab some information about the current display mode.
    Call d3d.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, d3ddm)
    
    'Now we'll go ahead and fill the D3DPRESENT_PARAMETERS type.
    With d3dpp
        .windowed = 1
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = d3ddm.Format 'current display depth
    End With
    'create device
    Set ddevice = d3d.CreateDevice(D3DADAPTER_DEFAULT, DevType, screen_hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)

    deviceRenderStates
    
    '****************************************************
    'Inicializamos el manager de texturas
    '****************************************************
    Call DXPool.Texture_Initialize(500)
    
    '****************************************************
    'Clears the buffer to start rendering
    '****************************************************
    deviceClear
    '****************************************************
    'Load Misc
    '****************************************************
    loadGraphicFonts
    loadFonts
    
    Exit Function
errhandler:
    initDirectX = False
End Function

Public Function dxBeginRender() As Boolean
On Error GoTo ErrorHandler:
    dxBeginRender = True
    
    'Check if we have the device
    If ddevice.TestCooperativeLevel <> D3D_OK Then
        Do
            DoEvents
        Loop While ddevice.TestCooperativeLevel = D3DERR_DEVICELOST
        
        DXPool.Texture_Remove_All
        fontsDestroy
        deviceReset
        
        deviceRenderStates
        loadFonts
        loadGraphicFonts
    End If
    
    '****************************************************
    'Render
    '****************************************************
    '*******************************
    'Erase the backbuffer so that it can be drawn on again
    deviceClear
    '*******************************
    '*******************************
    'Start the scene
    ddevice.BeginScene
    '*******************************
    
    engine_render_started = True
Exit Function
ErrorHandler:
    dxBeginRender = False
    MsgBox "Error in Engine_Render_Start: " & Err.Number & ": " & Err.Description
End Function

Public Function dxEndRender() As Boolean
On Error GoTo ErrorHandler:
    dxEndRender = True

    If engine_render_started = False Then
        Exit Function
    End If
    
    '*******************************
    'End scene
    ddevice.EndScene
    '*******************************
    
    '*******************************
    'Flip the backbuffer to the screen
    deviceFlip
    '*******************************
    
    '*******************************
    'Calculate current frames per second
    If GetTickCount >= (fps_last_time + 1000) Then
        fps = fps_frame_counter
        fps_frame_counter = 0
        fps_last_time = GetTickCount
    Else
        fps_frame_counter = fps_frame_counter + 1
    End If
    '*******************************
    

    
    
    engine_render_started = False
Exit Function
ErrorHandler:
    dxEndRender = False
    MsgBox "Error in Engine_Render_End: " & Err.Number & ": " & Err.Description
End Function

Private Sub deviceClear()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'Clear the back buffer
    ddevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET, 0, 1#, 0
End Sub

Private Function deviceReset() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Resets the device
'**************************************************************
On Error GoTo errhandler:
'On Error Resume Next

    'Be sure the scene is finished
    ddevice.EndScene
    'Reset device
    ddevice.Reset d3dpp
    
    deviceRenderStates
       
Exit Function
errhandler:
    deviceReset = Err.Number
End Function
Public Sub dxTextureRenderAdvance(ByVal texture_index As Long, ByVal dest_x As Long, ByVal dest_y As Long, ByVal Src_X As Long, ByVal Src_Y As Long, _
                                             ByVal dest_width As Long, ByVal dest_height As Long, ByVal src_width As Long, ByVal src_height As Long, ByRef rgb_list() As Long, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single, Optional ByVal ext As Byte = 0)
'**************************************************************
'This sub allow texture resizing
'
'**************************************************************

    
    Dim src_rect As RECT
    Dim dest_rect As RECT
    Dim temp_verts(3) As TLVERTEX
    Dim Texture As Direct3DTexture8
    Dim texture_width As Integer
    Dim texture_height As Integer

    'rgb_list(0) = RGB(255, 255, 255)
    'rgb_list(1) = RGB(255, 255, 255)
    'rgb_list(2) = RGB(255, 255, 255)
    'rgb_list(3) = RGB(255, 255, 255)
    
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + dest_height
        .left = dest_x
        .Right = dest_x + dest_width
        .top = dest_y
    End With
    
    With src_rect
        .bottom = Src_Y + src_height
        .Right = Src_X + src_width
        .top = Src_Y
        .left = Src_X
    End With
    
    Set Texture = DXPool.GetTexture(texture_index, ext)
    Call DXPool.Texture_Dimension_Get(texture_index, texture_width, texture_height)
    
    'Set up the TempVerts(3) vertices
    geometryBoxCreate temp_verts(), dest_rect, src_rect, rgb_list(), texture_width, texture_height, angle
    
    'Set Texture
    ddevice.SetTexture 0, Texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square texture
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
End Sub
Public Sub dxTextureRender(ByVal texture_index As Long, ByVal dest_x As Long, ByVal dest_y As Long, ByVal src_width As Long, _
                                            ByVal src_height As Long, ByRef rgb_list() As Long, ByVal Src_X As Long, _
                                            ByVal Src_Y As Long, ByVal dest_width As Long, ByVal dest_height As Long, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single, Optional ByVal ext As Byte = 0)
'**************************************************************
'This sub doesnt allow texture resizing
'
'**************************************************************
    Dim src_rect As RECT
    Dim dest_rect As RECT
    Dim temp_verts(3) As TLVERTEX
    Dim texture_height As Integer
    Dim texture_width As Integer
    Dim Texture As Direct3DTexture8
    
    'Set up the source rectangle
    With src_rect
        .bottom = Src_Y + src_height - 1
        .left = Src_X
        .Right = Src_X + src_width - 1
        .top = Src_Y
    End With
        
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + dest_height
        .left = dest_x
        .Right = dest_x + dest_width
        .top = dest_y
    End With
    
    'ESTO NO ME GUSTA
    Set Texture = DXPool.GetTexture(texture_index, ext)
    Call DXPool.Texture_Dimension_Get(texture_index, texture_width, texture_height)
    
    'Set up the TempVerts(3) vertices
    geometryBoxCreate temp_verts(), dest_rect, src_rect, rgb_list(), texture_height, texture_width, angle
    'Set Texture
    ddevice.SetTexture 0, Texture
    
    'Enable alpha-blending
    'ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square texture
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
    
    'Turn off alphablending after we're done
    'ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
End Sub
Private Function geometryTLVertexCreate(ByVal x As Single, ByVal y As Single, ByVal z As Single, _
                                            ByVal rhw As Single, ByVal color As Long, ByVal specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    geometryTLVertexCreate.x = x
    geometryTLVertexCreate.y = y
    geometryTLVertexCreate.z = z
    geometryTLVertexCreate.rhw = rhw
    geometryTLVertexCreate.color = color
    geometryTLVertexCreate.specular = specular
    geometryTLVertexCreate.tu = tu
    geometryTLVertexCreate.tv = tv
End Function

Private Sub geometryBoxCreate(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef texture_width As Integer, Optional ByRef texture_height As Integer, Optional ByVal angle As Single)
'**************************************************************
'Authors: Aaron Perkins;
'Last Modify Date: 5/07/2002
'
' * v1 *    v3
' |     \   |
' |     \   |
' |     \   |
' |     \   |
' |     \   |
' * v0 *    v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
    
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.left + (dest.Right - dest.left - 1) / 2
        y_center = dest.top + (dest.bottom - dest.top - 1) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
    
    
    '0 - Bottom left vertex
    If texture_width And texture_height Then
        verts(0) = geometryTLVertexCreate(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.left / texture_width, (src.bottom) / texture_height)
    Else
        verts(0) = geometryTLVertexCreate(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.left
        y_Cor = dest.top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
    
    
    '1 - Top left vertex
    If texture_width And texture_height Then
        verts(1) = geometryTLVertexCreate(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.left / texture_width, src.top / texture_height)
    Else
        verts(1) = geometryTLVertexCreate(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
    
    
    '2 - Bottom right vertex
    If texture_width And texture_height Then
        verts(2) = geometryTLVertexCreate(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right) / texture_width, (src.bottom) / texture_height)
    Else
        verts(2) = geometryTLVertexCreate(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
    
    
    '3 - Top right vertex
    If texture_width And texture_height Then
        verts(3) = geometryTLVertexCreate(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right) / texture_width, src.top / texture_height)
    Else
        verts(3) = geometryTLVertexCreate(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 0)
    End If
End Sub

Public Sub dxGraphicTextRender(Font_Index As Integer, ByVal Text As String, ByVal top As Long, ByVal left As Long, _
                                  ByVal color As Long)

    If Len(Text) > 255 Then Exit Sub
    
    Dim i As Byte
    Dim x As Integer
    Dim y As Integer
    Dim rgb_list(3) As Long
    
    For i = 0 To 3
        rgb_list(i) = color
    Next i
    
    x = -1
    Dim Char As Integer
    For i = 1 To Len(Text)
        Char = AscB(Mid$(Text, i, 1)) - 32
        
        If Char = 0 Then
            x = x + 1
        Else
            x = x + 1
            Call dxTextureRenderAdvance(gfont_list(Font_Index).texture_index, left + x * gfont_list(Font_Index).Char_Size, _
                                                        top, gfont_list(Font_Index).Caracteres(Char).Src_X, gfont_list(Font_Index).Caracteres(Char).Src_Y, _
                                                            gfont_list(Font_Index).Char_Size, gfont_list(Font_Index).Char_Size, gfont_list(Font_Index).Char_Size, gfont_list(Font_Index).Char_Size, _
                                                                rgb_list(), False)
        End If
    Next i
    
    
    
End Sub

Public Sub dxDeInit()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
On Error Resume Next

    'El manager de texturas es ahora independiente del engine.
    Call DXPool.Texture_Remove_All
    
    Set d3dx = Nothing
    Set ddevice = Nothing
    Set d3d = Nothing
    Set dx = Nothing
    Set DXPool = Nothing
End Sub

Private Sub loadChars(ByVal Font_Index As Integer)
    Dim i As Integer
    Dim x As Integer
    Dim y As Integer
    
    For i = 0 To 255
        With gfont_list(Font_Index).Caracteres(i)
            x = (i Mod 16) * gfont_list(Font_Index).Char_Size
            If x = 0 Then '16 chars per line
                y = y + 1
            End If
            .Src_X = x
            .Src_Y = (y * gfont_list(Font_Index).Char_Size) - gfont_list(Font_Index).Char_Size
        End With
    Next i
End Sub
Public Sub loadGraphicFonts()
    Dim i As Byte
    Dim file_path As String

    file_path = resource_path & PATH_INIT & "\GUIFonts.ini"

    If General_File_Exists(file_path, vbArchive) Then
        gfont_count = General_Var_Get(file_path, "INIT", "FontCount")
        If gfont_count > 0 Then
            ReDim gfont_list(1 To gfont_count) As tGraphicFont
            For i = 1 To gfont_count
                With gfont_list(i)
                    .Char_Size = General_Var_Get(file_path, "FONT" & i, "Size")
                    .texture_index = General_Var_Get(file_path, "FONT" & i, "Graphic")
                    If .texture_index > 0 Then Call DXPool.Texture_Load(.texture_index, 0)
                    loadChars (i)
                End With
            Next i
        End If
    End If
End Sub

Public Sub dxStatsRender()
    'fps
    Call dxTextRender(1, fps & " FPS", 0, 0, D3DColorXRGB(255, 255, 255))
End Sub

Private Sub deviceFlip()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'Draw the graphics to the front buffer.
    ddevice.Present ByVal 0&, ByVal 0&, screen_hwnd, ByVal 0&
End Sub

Private Sub deviceRenderStates()
    With ddevice
        'Set the vertex shader to an FVF that contains texture coords,
        'and transformed and lit vertex coords.
        .SetVertexShader FVF
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        
        'No se para q mierda sera esto.
        '.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        '.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        '.SetRenderState D3DRS_ZENABLE, True
        '.SetRenderState D3DRS_ZWRITEENABLE, False
        
        'Particle engine settings
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        '.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        '.SetRenderState D3DRS_POINTSCALE_ENABLE, 0


    End With
End Sub

Private Sub fontMake(ByVal Style As String, ByVal Size As Long, ByVal italic As Boolean, ByVal bold As Boolean)
    font_count = font_count + 1
    ReDim Preserve font_list(1 To font_count)
    
    Dim font_desc As IFont
    Dim fnt As New StdFont
    fnt.name = Style
    fnt.Size = Size
    fnt.bold = bold
    fnt.italic = italic
    
    Set font_desc = fnt
    font_list(font_count).Size = Size
    Set font_list(font_count).dFont = d3dx.CreateFont(ddevice, font_desc.hFont)
End Sub

Private Sub loadFonts()
    Dim num_fonts As Integer
    Dim i As Integer
    Dim file_path As String
    
    file_path = resource_path & PATH_INIT & "\fonts.ini"
    
    If Not General_File_Exists(file_path, vbArchive) Then Exit Sub
    
    num_fonts = General_Var_Get(file_path, "INIT", "FontCount")
    
    For i = 1 To num_fonts
        Call fontMake(General_Var_Get(file_path, "FONT" & i, "Name"), General_Var_Get(file_path, "FONT" & i, "Size"), General_Var_Get(file_path, "FONT" & i, "Cursiva"), General_Var_Get(file_path, "FONT" & i, "Negrita"))
    Next i
End Sub
Public Sub dxTextRender(ByVal Font_Index As Integer, ByVal Text As String, ByVal left As Integer, ByVal top As Integer, ByVal color As Long, Optional ByVal Alingment As Byte = DT_LEFT, Optional ByVal width As Integer = 0, Optional ByVal height As Integer = 0)
    If Not fontCheck(Font_Index) Then Exit Sub
    
    Dim TextRect As RECT 'This defines where it will be
    'Dim BorderColor As Long
    
    'Set width and height if no specified
    If width = 0 Then width = Len(Text) * (font_list(Font_Index).Size + 1)
    If height = 0 Then height = font_list(Font_Index).Size * 2
    
    'DrawBorder
    
    'BorderColor = D3DColorXRGB(0, 0, 0)
    
    'TextRect.top = top - 1
    'TextRect.left = left
    'TextRect.bottom = top + height
    'TextRect.Right = left + width
    'd3dx.DrawText font_list(Font_Index).dFont, BorderColor, Text, TextRect, Alingment
    'TextRect.top = top
    'TextRect.left = left - 1
    'TextRect.bottom = top + height
    'TextRect.Right = left + width
    'd3dx.DrawText font_list(Font_Index).dFont, BorderColor, Text, TextRect, Alingment
    'TextRect.top = top + 1
    'TextRect.left = left
    'TextRect.bottom = top + height
    'TextRect.Right = left + width
    'd3dx.DrawText font_list(Font_Index).dFont, BorderColor, Text, TextRect, Alingment
    'TextRect.top = top
    'TextRect.left = left + 1
    'TextRect.bottom = top + height
    'TextRect.Right = left + width
    'd3dx.DrawText font_list(Font_Index).dFont, BorderColor, Text, TextRect, Alingment
    
    TextRect.top = top
    TextRect.left = left
    TextRect.bottom = top + height
    TextRect.Right = left + width
    d3dx.DrawText font_list(Font_Index).dFont, color, Text, TextRect, Alingment

End Sub
Private Function fontCheck(ByVal Font_Index As Long) As Boolean
    If Font_Index > 0 And Font_Index <= font_count Then
        fontCheck = True
    End If
End Function

Private Sub fontsDestroy()
    Dim i As Integer
    
    For i = 1 To font_count
        Set font_list(i).dFont = Nothing
        font_list(i).Size = 0
    Next i
    font_count = 0
End Sub



Public Function D3DColorValueGet(ByVal A As Byte, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte) As D3DCOLORVALUE
    D3DColorValueGet.A = A
    D3DColorValueGet.R = R
    D3DColorValueGet.G = G
    D3DColorValueGet.B = B
End Function

Public Sub dxTextureToHdcRender(ByVal texture_index As Long, desthdc As Long, ByVal Screen_X As Long, ByVal Screen_Y As Long, ByVal SX As Integer, ByVal SY As Integer, ByVal sW As Integer, ByVal sH As Integer, Optional transparent As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/02/2003
'This method is SLOW... Don't use in a loop if you care about
'speed!
'*************************************************************

    Dim file_path As String
    Dim Src_X As Long
    Dim Src_Y As Long
    Dim src_width As Long
    Dim src_height As Long
    Dim hdcsrc As Long

    file_path = resource_path & "\graphics\" & texture_index & ".bmp"
    
    Src_X = SX
    Src_Y = SY
    src_width = sW
    src_height = sH

    hdcsrc = CreateCompatibleDC(desthdc)
    
    SelectObject hdcsrc, LoadPicture(file_path)
    
    If transparent = False Then
        BitBlt desthdc, Screen_X, Screen_Y, src_width, src_height, hdcsrc, Src_X, Src_Y, SRCCOPY
    Else
        TransparentBlt desthdc, Screen_X, Screen_Y, src_width, src_height, hdcsrc, Src_X, Src_Y, src_width, src_height, COLOR_KEY
    End If
        
    DeleteDC hdcsrc
End Sub

Public Sub dxBeginSecondaryRender()
    deviceClear
    ddevice.BeginScene
End Sub
Public Sub dxEndSecondaryRender(ByVal hwnd As Long, ByVal width As Integer, ByVal height As Integer)
    Dim DR As RECT
    DR.left = 0
    DR.top = 0
    DR.bottom = height
    DR.Right = width
    
    ddevice.EndScene
    ddevice.Present DR, ByVal 0&, hwnd, ByVal 0&
End Sub


Public Sub dxDrawBox(ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal color As Long, Optional ByVal borderColor As Long, Optional ByVal border_width = 1)
    Dim DR As RECT
    Dim VertexB(3) As TLVERTEX
    Dim box_rect As RECT
    Dim border_rect As RECT
    
    With box_rect
        .bottom = y + height
        .left = x
        .Right = x + width
        .top = y
    End With
    
    With border_rect
        .bottom = y + height + border_width * 2
        .left = x - border_width
        .Right = x + width + border_width * 2
        .top = y - border_width
    End With
        
    ddevice.SetTexture 0, Nothing
    
    'Border
    VertexB(0) = geometryTLVertexCreate(border_rect.left, border_rect.bottom, 0, 1, borderColor, 0, 0, 0)
    VertexB(1) = geometryTLVertexCreate(border_rect.left, border_rect.top, 0, 1, borderColor, 0, 0, 0)
    VertexB(2) = geometryTLVertexCreate(border_rect.Right, border_rect.bottom, 0, 2, borderColor, 0, 0, 0)
    VertexB(3) = geometryTLVertexCreate(border_rect.Right, border_rect.top, 0, 2, borderColor, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    'Box
    VertexB(0) = geometryTLVertexCreate(box_rect.left, box_rect.bottom, 0, 1, color, 0, 0, 0)
    VertexB(1) = geometryTLVertexCreate(box_rect.left, box_rect.top, 0, 1, color, 0, 0, 0)
    VertexB(2) = geometryTLVertexCreate(box_rect.Right, box_rect.bottom, 0, 2, color, 0, 0, 0)
    VertexB(3) = geometryTLVertexCreate(box_rect.Right, box_rect.top, 0, 2, color, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    
End Sub
Public Sub D3DColorToRgbList(rgb_list() As Long, color As D3DCOLORVALUE)
    rgb_list(0) = D3DColorARGB(color.A, color.R, color.G, color.B)
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

