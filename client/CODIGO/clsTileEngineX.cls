VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTileEngineX"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsTileEngineX - ORE Direct3D 8 Tile Engine - v0.5.0
'
'Graphics Engine.
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

'*****************************************************************
'Contributors History
'   When releasing modifications to this source file please add your
'   date of release, name, email, and any info to the top of this list.
'   Follow this template:
'    XX/XX/200X - Your Name Here (Your Email Here)
'       - Your Description Here
'       Sub Release Contributors:
'           XX/XX/2003 - Sub Contributor Name Here (SC Email Here)
'               - SC Description Here
'*****************************************************************
'
'05/27/2003 - Aaron Perkins (aaron@baronsoft.com)
'   -Change: Light_Move now called Light_Map_Pos_Set and Light_Move is a new method.
'   -Add: Particle_Group_Map_Pos_Set and Move
'   Sub Release Contributors:
'       05/25/2003 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'           - Add: Map_Bounds_Get_From_File
'           - Add: Light_Move
'
'05/15/2003 - Aaron Perkins (aaron@baronsoft.com)
'   -Add: GUI_Grh_Render
'   -Change: Text_Render is now called GUI_Text_Render
'       -Added word_wrap option
'   Sub Release Contributors:
'       05/15/2003 - Fredrik Alexandersson (fredrik@oraklet.zzn.com)
'           - Add: GUI_Box_Filled_Render
'           - Add: GUI_Box_Outline_Render
'
'05/12/2003 - Aaron Perkins (aaron@baronsoft.com)
'   -Add: Char_Data_Grh_Index_Get
'
'05/08/2003 - Aaron Perkins (aaron@baronsoft.com)
'   -Change: Moved all the General_X methods to a external modGeneral.bas
'
'05/07/2003 - Aaron Perkins (aaron@baronsoft.com)
'   -Fix: Map_Save and Map_Load format changed. Added angle and alpha values.
'   Sub Release Contributors:
'       05/07/2003 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'           - Fix: Geometry_Create_Sqaure - fixed rotation
'
'05/02/2003 - Aaron Perkins (aaron@baronsoft.com)
'   - Fix: bug with Light_Move and base light values
'   - Fix: Grh_Get_Info and Grh_Render_To_Hdc when using animated grhs
'   - Add: Base light values can now be specfied for each corner
'      - Changed .map format. Old maps will not work anymore
'      - Modified Map_Base_Light_Set
'   - Remove: Light_Remove_XY (Use Light_Remove Map_Light_Get(X,Y) instead)
'
'04/29/2003 - Aaron Perkins (aaron@baronsoft.com)
'   - Add: Grh_Render_To_Hdc
'   - Add: Map_Load_From_File
'   - Add: Text_Render method
'   - Change: file name to clsTileEngineX.cls
'   - Change: Split Engine_Render into Engine_Render_Start and Engine_Render_End
'   - Change: Method Names
'       -Engine_Stats_Render
'       -Engine_Blocked_Tiles_Show_Toggle
'   Sub Release Contributors:
'       04/28/2003 - Fredrik Alexandersson (fredrik@oraklet.zzn.com)
'           - Add: Grh_Add_GrhList_To_ListBox method
'           - Change: Particle_Render (added 3 new stream types)
'
'04/07/2003 - Aaron Perkins (aaron@baronsoft.com)
'   - First release
'   Sub Release Contributors:
'       03/31/2003 - Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'           - Add: Light_Erase method
'*****************************************************************

'***************************
'Required Externals
'***************************
'Reference to dx8vb.dll
'   - URL: http://www.microsoft.com/directx
'***************************
Option Explicit

Private Const MAXEFFECTS As Byte = 5
Private Const MAXPROJECTILES As Byte = 5

Private Type offset
    x As Integer
    y As Integer
End Type

'Char Body
Private Type Char_Data_Body
    Body(1 To 4) As grh
    HeadOffset As offset
End Type

'Char Head
Private Type Char_Data_Head
    Head(1 To 4) As grh
End Type

'Char Weapons
Private Type Char_Data_Weapon
    WeaponWalk(1 To 4) As grh
    '[ANIM ATAK]
    'WeaponAttack As Byte
End Type

Private Type tAuraData
    grhIndex As Integer
    rgb_list(3) As Long
End Type

Private Type Char_Data_Aura
    Aura As Integer
    AuraGrh As grh
End Type

'Char Shields
Private Type Char_Data_Shield
    ShieldWalk(1 To 4) As grh
End Type

Private Type tBonfire
    mapX As Byte
    mapY As Byte
    lightIndex As Integer
    effectIndex As Integer
End Type

Private Type Char_Data_Fx
    fx_grh_index As Integer
    fx_offset As offset
    ParticleType As Integer
    ParticleLifeTime As Integer
End Type

Private Type Char_Data_Fx_Grh
    fx As Integer
    FxLoopTimes As Integer
    FxGrh As grh
    Particle_Index As Integer
End Type
'Char Data
Private Type Char_Data
    BodyData As Char_Data_Body
    HeadData As Char_Data_Head
    WeaponData As Char_Data_Weapon
    ShieldData As Char_Data_Shield
    CascoData As Char_Data_Head
    FxData As Char_Data_Fx_Grh
    AuraData As Char_Data_Aura
    
    'lastEffect
    
    lastEffect As Integer
    effectList(1 To MAXEFFECTS) As Integer
    
    'Pointer to projectiles used to update the projectile target when char moves.
    projectileIndex(1 To MAXPROJECTILES) As Integer
    lastProjectile As Integer
End Type

Private Type Char_Data_List
    'Para que mierda es esto?
    NumWeaponAnims As Integer
    NumShieldAnims As Integer

    BodyData() As Char_Data_Body
    HeadData() As Char_Data_Head
    WeaponData() As Char_Data_Weapon
    ShieldData() As Char_Data_Shield
    CascoData() As Char_Data_Head
    FxData() As Char_Data_Fx
    AuraData() As tAuraData
    
    
End Type
'***************************
'Lista de cabezas (Utilizado para cargar la lista)
Private Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

Private Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Private Type tIndiceFx
    Animacion As Integer
    offSetX As Integer
    offSetY As Integer
    ParticleType As Integer
    ParticleLifeTime As Integer
End Type
'***************************
'Hold info about a character
Private Type Char
    Active As Boolean
    heading As Long
    id As Long
    map_x As Long
    map_y As Long
    
    pie As Boolean
    
    chr_data As Char_Data
    chr_data_body_index As Long
    
    label As String
    label_font_index As Long
    label_offset_x As Long
    label_offset_y As Long
    
    scroll_on As Boolean
    scroll_offset_counter_x As Single
    scroll_offset_counter_y As Single
    scroll_direction_x As Long
    scroll_direction_y As Long
    
    'Flags
    Invisible As Boolean
    priv As Byte
    faccion As Byte
    
    light_index As Integer
End Type

Private Type Map_Exit
    exit_map_name As String
    exit_map_x As Long
    exit_map_y As Long
    
    c_map_x As Long
    c_map_y As Long
End Type

Private Type Map_NPC
    npc_data_index As Long
    
    c_char_data_index As Long
    c_map_x As Long
    c_map_y As Long
End Type

Private Type Map_Item
    item_data_index As Long
    item_amount As Long
    
    c_grh_index As Long
    c_map_x As Long
    c_map_y As Long
End Type

'Map Tile structure
Private Type Map_Tile
    grh(1 To 5) As grh
    Blocked As Byte
    
    bonfireIndex As Integer
    
    effectList(1 To MAXEFFECTS) As Integer
    lastEffect As Integer
    
    char_index As Long
    npc_index As Long
    
    base_light(0 To 3) As Boolean 'Indica si el tile tiene luz propia.
    light_base_value(0 To 3) As Long 'Luz propia del tile.
    light_value(0 To 3) As Long 'Color de luz con el que esta siendo renderizado.
    light_index As Integer
    
    exit_index As Long
    item_index As Long
    Trigger As Integer
    
    'Pointer to projectiles used to update the projectile target when char moves.
    projectileList(1 To MAXPROJECTILES) As Integer
    lastProjectile As Integer
End Type

'Posicion en el Mundo
Private Type World_Pos
    Map As Integer
    x As Integer
    y As Integer
End Type

'Info de cada mapa
Private Type Map_Info
    Music As String
    name As String
    StartPos As World_Pos
    MapVersion As Integer
    
    'ME Only
    Changed As Byte
    Ambient As Byte
    MapAmbientColor As D3DCOLORVALUE
    Llueve As Byte
End Type


'Map structure
Private Type Map
    map_grid() As Map_Tile
    map_x_max As Long
    map_x_min As Long
    map_y_max As Long
    map_y_min As Long
    Map_Info As Map_Info
End Type


' TRIGGERS
'
' @param NADA nada
' @param BAJOTECHO bajo techo
' @param trigger_2 ???
' @param POSINVALIDA los npcs no pueden pisar tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
' @param Clan hall

Public Enum eTrigger
    NADA = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
End Enum

'The light structure
Private Type Light
    x As Integer
    y As Integer
    Active As Boolean 'Do we ignore this light?
    id As Long
    map_x As Long 'Coordinates
    map_y As Long
    color As Long 'Start colour
    range As Long
    RGBColor As D3DCOLORVALUE
End Type

'Map view area (where the game is played)
Dim view_screen_top As Long 'In pixels
Dim view_screen_left As Long 'In pixels
Dim view_screen_bottom As Long
Dim view_screen_right As Long
Dim view_screen_tile_width As Long 'In tiles
Dim view_screen_tile_height As Long 'In tiles
Dim view_screen_width As Long
Dim view_screen_height As Long

'Buffer area (used to draw object outside the map area but may still show up on the screen)
Dim view_tile_buffer As Long 'In tiles

Dim light_buffer As Long

'Base tile size (smallest possible tile size: must be square)
Dim base_tile_size As Long 'In pixels

'View position: In tiles
Dim view_pos_x As Long
Dim view_pos_y As Long

'Scrolling stuff
Dim scroll_on As Boolean
Dim scroll_direction_x As Long
Dim scroll_direction_y As Long
Dim scroll_offset_counter_x As Single
Dim scroll_offset_counter_y As Single
Dim scroll_pixels_per_frame As Long

'ticks per frame
Dim timer_ticks_per_frame As Single

'base speed for the engine
Dim engine_base_speed As Single

'windowed or not windowed
Dim engine_windowed As Boolean

'show engine stats
Dim engine_show_stats As Boolean

'show blocked tiles on map
Dim engine_show_blocked_tiles As Boolean
Dim engine_show_special_tiles As Boolean

'***************************
'Arrays
'***************************

'Char list
Dim char_list(1 To 32000) As Char
Dim char_count As Long
Dim char_last As Long

Dim user_char_index As Long

'Char data list
Dim Char_Data_List As Char_Data_List

'Current Map
Dim map_current As Map

'Light list
Dim light_list() As Light
Dim light_count As Long
Dim light_last As Long

Dim exit_list() As Map_Exit
Dim npc_list() As Map_NPC
Dim item_list() As Map_Item

'Fogatas
Private bonfireList() As tBonfire
Private bonfireCount As Integer

'Screen parameters.
Dim view_min_y As Integer
Dim view_max_y As Integer
Dim view_min_x As Integer
Dim view_max_x As Integer
    
Dim view_buffer_min_y As Integer
Dim view_buffer_max_y As Integer
Dim view_buffer_min_x As Integer
Dim view_buffer_max_x As Integer

Dim AlphaRoof As Byte 'For transparent roofs
Private Const AlphaRoofInterval As Byte = 1


'***************************
'External Functions
'***************************

'MouseInput
Private Type PointAPI
    x As Long
    y As Long
End Type


Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

'Gets number of ticks since windows started
Private Declare Function GetTickCount Lib "kernel32" () As Long


Private Sub Class_Initialize()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
End Sub

Private Sub Class_Terminate()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
End Sub

Private Function Convert_Tile_To_View_Y(ByVal y As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Convert tile position into position in view area
'**************************************************************
    If engine_windowed Then
        Convert_Tile_To_View_Y = ((y - 1) * base_tile_size)
    Else
        Convert_Tile_To_View_Y = view_screen_top + ((y - 1) * base_tile_size)
    End If
End Function

Private Sub Convert_Screen_To_View(ByVal Screen_X As Long, ByVal Screen_Y As Long, ByRef view_x As Long, ByRef view_y As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    view_x = Screen_X - view_screen_left
    view_y = Screen_Y - view_screen_top
End Sub

Private Sub Convert_View_To_Map(ByVal view_x As Long, ByVal view_y As Long, ByRef map_x As Long, ByRef map_y As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim half_view_tile_width As Long
    Dim half_view_tile_height As Long
    
    half_view_tile_width = (view_screen_tile_width \ 2)
    half_view_tile_height = (view_screen_tile_height \ 2)
    
    map_x = (view_x \ base_tile_size)
    map_y = (view_y \ base_tile_size)
    
    If map_x > half_view_tile_width Then
        map_x = (map_x - half_view_tile_width)
    
    Else
        If map_x < half_view_tile_width Then
            map_x = (0 - (half_view_tile_width - map_x))
        Else
            map_x = 0
        End If
    End If
    
    If map_y > half_view_tile_height Then
        map_y = (0 - (half_view_tile_height - map_y))
    Else
        If map_y < half_view_tile_height Then
            map_y = (map_y - half_view_tile_height)
        Else
            map_y = 0
        End If
    End If
    
    map_x = view_pos_x + map_x
    map_y = view_pos_y + map_y
End Sub

Private Sub Convert_Map_To_Direction(ByVal map_x As Long, ByVal map_y As Long, ByRef direction_x As Long, ByRef direction_y As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp_x As Long
    Dim temp_y As Long
    
    temp_x = map_x - view_pos_x
    temp_y = map_y - view_pos_y
    
    If temp_x <> 0 Then
        direction_x = temp_x \ Abs(temp_x)
    Else
        direction_x = 0
    End If
    If temp_y <> 0 Then
        direction_y = temp_y \ Abs(temp_y)
    Else
        direction_y = 0
    End If
End Sub

Private Function Convert_Direction_To_Heading(ByVal direction_x As Long, ByVal direction_y As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'North
    If direction_y = -1 Then
        Convert_Direction_To_Heading = 1
    End If
    'East
    If direction_x = 1 Then
        Convert_Direction_To_Heading = 2
    End If
    'South
    If direction_y = 1 Then
        Convert_Direction_To_Heading = 3
    End If
    'West
    If direction_x = -1 Then
        Convert_Direction_To_Heading = 4
    End If
    
End Function

Private Sub Convert_Heading_to_Direction(ByVal heading As Long, ByRef direction_x As Long, ByRef direction_y As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim addY As Long
    Dim addX As Long
    
    'Figure out which way to move
    Select Case heading
    
        Case 1
            addY = -1
    
        Case 2
            addX = 1
    
        Case 3
            addY = 1
            
        Case 4
            addX = -1
            
    End Select
    
    direction_x = direction_x + addX
    direction_y = direction_y + addY
End Sub

Private Function Convert_Tile_To_View_X(ByVal x As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Convert tile position into position in view area
'**************************************************************
    If engine_windowed Then
        Convert_Tile_To_View_X = ((x - 1) * base_tile_size)
    Else
        Convert_Tile_To_View_X = view_screen_left + ((x - 1) * base_tile_size)
    End If
End Function
Public Sub Engine_Blocked_Tiles_Show_Toggle()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/17/2003

'*****************************************************************
    If engine_show_blocked_tiles Then
        engine_show_blocked_tiles = False
    Else
        engine_show_blocked_tiles = True
    End If
End Sub

Public Sub Engine_Special_Tiles_Show_Toggle()
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003

'*****************************************************************
    If engine_show_special_tiles Then
        engine_show_special_tiles = False
    Else
        engine_show_special_tiles = True
    End If
End Sub

Public Function Engine_View_Pos_Set(ByVal map_x As Long, ByVal map_y As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Sets the user postion
'If valid, returns True, else False
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        view_pos_x = map_x
        view_pos_y = map_y
    End If
End Function

Public Sub Engine_View_Pos_Get(ByRef map_x As Long, ByRef map_y As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the user postion
'*****************************************************************
    map_x = view_pos_x
    map_y = view_pos_y
End Sub

Public Sub Engine_Base_Speed_Set(ByVal Speed As Single)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'*****************************************************************
    engine_base_speed = Speed
End Sub

Public Function Engine_Base_Speed_Get() As Single
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the target fps that the engine
'*****************************************************************
    Engine_Base_Speed_Get = engine_base_speed
End Function

Public Function TileEngine_Initialize(ByVal windowed As Boolean, Optional ByVal v_left As Long = 0, Optional ByVal v_top As Long = 0, Optional ByVal v_width_in_tiles As Long = 0, _
                                    Optional ByVal v_height_in_tiles As Long = 0, Optional ByVal tile_size As Long = 32) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'**************************************************************
On Error GoTo ErrorHandler:
    TileEngine_Initialize = True
    
    '****************************************************
    'Fill in global variables
    '****************************************************
    
    'Fill in base tile size (must always be square)
    base_tile_size = tile_size 'In pixels
    
    'Fill in viewport sizes. How large the game area is.
    If v_width_in_tiles = 0 Or v_height_in_tiles = 0 Then
        view_screen_tile_width = 19 'In tiles
        view_screen_tile_height = 15 'In tiles
    Else
        view_screen_tile_width = v_width_in_tiles 'In tiles
        view_screen_tile_height = v_height_in_tiles 'In tiles
    End If
    

    Dim target As RECT
    GetWindowRect frmMain.MainViewPic, target
    view_screen_top = target.top
    view_screen_left = target.left
    view_screen_right = target.Right
    view_screen_bottom = target.bottom
    view_screen_width = target.Right - target.left
    view_screen_height = target.bottom - target.top
    engine_windowed = True
    
    '****************************************************
    'Get external data
    '****************************************************
    
    'Load body data for characters
    Char_Load_Char_Data
    
    '****************************************************
    'Setup Map
    '****************************************************
    'Buffer area
    view_tile_buffer = 9
    
    'Light area
    light_buffer = 3
    
    'How many pixels to move per frame when scrolling
    scroll_pixels_per_frame = 8
    
    'User start position
    view_pos_x = 1
    view_pos_y = 1
    
    'Create default map
    Map_Create 100, 100
    
    AlphaRoof = 255
Exit Function
ErrorHandler:
    MsgBox "Error in Engine_Initialization: " & Err.Number & ": " & Err.Description
    TileEngine_Initialize = False
End Function

Public Function Engine_View_Move(ByVal heading As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim x As Long
    Dim y As Long
    Dim nX As Long
    Dim nY As Long
    
    'Don't move if we are already moving....
    If scroll_on Then
        Engine_View_Move = False
        Exit Function
    End If
    
    'Invalid heading
    If heading < 1 Or heading > 8 Then
        Engine_View_Move = False
        Exit Function
    End If
    
    x = view_pos_x
    y = view_pos_y
    nX = x
    nY = y
    Convert_Heading_to_Direction heading, nX, nY
    
    'See if out new position is legal
    If Map_In_Bounds(nX, nY) Then
        'start the scrolling process
        view_pos_x = nX
        view_pos_y = nY
        
        scroll_offset_counter_x = (base_tile_size * (x - nX))
        scroll_offset_counter_y = (base_tile_size * (y - nY))
        scroll_direction_x = nX - x
        scroll_direction_y = nY - y
        scroll_on = True

        Engine_View_Move = True
    Else
        'not legal don't move
        scroll_direction_x = 0
        scroll_direction_y = 0
    End If
End Function

Private Function NPC_Ini_Char_Data_Index_Get(ByVal s_npc_data_index As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'
'**************************************************************
    NPC_Ini_Char_Data_Index_Get = CLng(General_Var_Get(resource_path & PATH_SCRIPTS & "\npc.ini", "NPC" & s_npc_data_index, "npc_char_data_index"))
End Function

Private Function Item_Ini_Grh_Index_Get(ByVal s_item_data_index As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'
'**************************************************************
    Item_Ini_Grh_Index_Get = CLng(General_Var_Get(resource_path & PATH_SCRIPTS & "\item.ini", "ITEM" & s_item_data_index, "item_grh_index"))
End Function

Public Function Map_NPC_Add(ByVal s_map_x As Long, ByVal s_map_y As Long, ByVal s_npc_data_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_Legal_Char_Pos(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).npc_index = 0 Then
            ReDim Preserve npc_list(0 To UBound(npc_list) + 1)
            npc_list(UBound(npc_list)).npc_data_index = s_npc_data_index
            npc_list(UBound(npc_list)).c_char_data_index = NPC_Ini_Char_Data_Index_Get(s_npc_data_index)
            npc_list(UBound(npc_list)).c_map_x = s_map_x
            npc_list(UBound(npc_list)).c_map_y = s_map_y
            map_current.map_grid(s_map_x, s_map_y).npc_index = UBound(npc_list)
            Map_NPC_Add = True
        End If
    End If
End Function

Public Function Map_NPC_Remove(ByVal s_map_x As Long, ByVal s_map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).npc_index <> 0 Then
            npc_list(map_current.map_grid(s_map_x, s_map_y).npc_index).npc_data_index = 0
            npc_list(map_current.map_grid(s_map_x, s_map_y).npc_index).c_char_data_index = 0
            npc_list(map_current.map_grid(s_map_x, s_map_y).npc_index).c_map_x = 0
            npc_list(map_current.map_grid(s_map_x, s_map_y).npc_index).c_map_y = 0
            map_current.map_grid(s_map_x, s_map_y).npc_index = 0
            Map_NPC_Remove = True
        End If
    End If
End Function

Public Function Map_NPC_Get(ByVal s_map_x As Long, ByVal s_map_y As Long, ByRef r_npc_data_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'Returns NPC's data index
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).npc_index <> 0 Then
            r_npc_data_index = npc_list(map_current.map_grid(s_map_x, s_map_y).npc_index).npc_data_index
            Map_NPC_Get = True
        End If
    End If
End Function

Public Function Map_Exit_Add(ByVal s_map_x As Long, ByVal s_map_y As Long, ByVal s_exit_map_name As String, ByVal s_exit_map_x As Long, ByVal s_exit_map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).exit_index = 0 Then
            ReDim Preserve exit_list(0 To UBound(exit_list) + 1)
            exit_list(UBound(exit_list)).exit_map_name = s_exit_map_name
            exit_list(UBound(exit_list)).exit_map_x = s_exit_map_x
            exit_list(UBound(exit_list)).exit_map_y = s_exit_map_y
            exit_list(UBound(exit_list)).c_map_x = s_map_x
            exit_list(UBound(exit_list)).c_map_y = s_map_y
            map_current.map_grid(s_map_x, s_map_y).exit_index = UBound(exit_list)
            Map_Exit_Add = True
        End If
    End If
End Function

Public Function Map_Exit_Remove(ByVal s_map_x As Long, ByVal s_map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).exit_index <> 0 Then
            exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).exit_map_name = ""
            exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).exit_map_x = 0
            exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).exit_map_y = 0
            exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).c_map_x = 0
            exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).c_map_y = 0
            map_current.map_grid(s_map_x, s_map_y).exit_index = 0
            Map_Exit_Remove = True
        End If
    End If
End Function

Public Function Map_Exit_Get(ByVal s_map_x As Long, ByVal s_map_y As Long, ByRef r_exit_map_name As String, ByRef r_exit_map_x As Long, ByRef r_exit_map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'Returns exit information
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).exit_index <> 0 Then
            r_exit_map_name = exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).exit_map_name
            r_exit_map_x = exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).exit_map_x
            r_exit_map_y = exit_list(map_current.map_grid(s_map_x, s_map_y).exit_index).exit_map_y
            Map_Exit_Get = True
        End If
    End If
End Function

Public Function Map_Item_Add(ByVal s_map_x As Long, ByVal s_map_y As Long, grh_index) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If grh_index = FOgata Then
            mapBonfireCreate s_map_x, s_map_y
        End If
        map_current.map_grid(s_map_x, s_map_y).grh(5).grh_index = grh_index
        Call Grh_Initialize(map_current.map_grid(s_map_x, s_map_y).grh(5), map_current.map_grid(s_map_x, s_map_y).grh(5).grh_index)
        Map_Item_Add = True
    End If
End Function

Public Function Map_Item_Remove(ByVal s_map_x As Long, ByVal s_map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).grh(5).grh_index = FOgata Then
            mapBonfireDestroy map_current.map_grid(s_map_x, s_map_y).bonfireIndex
        End If
        map_current.map_grid(s_map_x, s_map_y).grh(5).grh_index = 0
        Map_Item_Remove = True
    End If
End Function

Public Function Map_Item_Get(ByVal s_map_x As Long, ByVal s_map_y As Long, ByRef r_item_data_index As Long, ByRef r_item_amount As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2002
'Returns item information
'**************************************************************
    If Map_In_Bounds(s_map_x, s_map_y) Then
        If map_current.map_grid(s_map_x, s_map_y).item_index <> 0 Then
            r_item_data_index = item_list(map_current.map_grid(s_map_x, s_map_y).item_index).item_data_index
            r_item_amount = item_list(map_current.map_grid(s_map_x, s_map_y).item_index).item_amount
            Map_Item_Get = True
        End If
    End If
End Function

Public Function Map_Load_Map(ByVal map_name As String, Optional ByVal load_ini As Boolean = False) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/25/2003
'Load map data from file using an map id number
'*****************************************************************
    Dim map_path As String

    'Get map file path
    map_path = resource_path & PATH_MAPS & "\Mapa" & map_name & ".map"
    
    Map_Load_Map = Map_Load_Map_From_File(map_path, load_ini)

End Function

Public Function Map_Load_Map_From_File(ByVal file_path As String, Optional ByVal load_ini As Boolean = False) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/25/2003
'Load map data from file using an filepath
'*****************************************************************
    Dim loopc As Long
    Dim LoopC2 As Long
    Dim TempByte As Byte
    Dim TempInt As Integer
    Dim TempLng As Long
    Dim TempByte1 As Byte
    Dim TempByte2 As Byte
    Dim TempByte3 As Byte
    Dim TempSgl As Single
    Dim TempBln As Boolean
    Dim TempLngList() As Long
    Dim y As Long
    Dim x As Long
    Dim map_path As String
    
    'Get map file path
    map_path = file_path
    
    'If file doesn't exists, exit
    If Not (General_File_Exists(map_path, vbNormal)) Then
        MsgBox "Error en los mapas, algún archivo ha sido modificado o esta dañado."
        prgRun = False
        'I think this is not necessary at all
        Exit Function
    End If
    
    'Erase Chars, Lights, and Particle Groups
    Char_Remove_All
    Light_Remove_All
    mapDestroyAllEffects
    
    'Open .map file
    Open map_path For Binary As #1
    Seek #1, 1
    
    'map Header
    Get #1, , map_current.Map_Info.MapVersion
    Get #1, , MiCabecera
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    Get #1, , TempInt
    
    map_current.map_y_min = 1
    map_current.map_y_max = 100
    
    map_current.map_x_min = 1
    map_current.map_x_max = 100
    
    'Clear out and resize map
    ReDim map_current.map_grid(1 To 100, _
                                1 To 100) As Map_Tile
    
    'Read .map file
    For y = map_current.map_y_min To map_current.map_y_max
        For x = map_current.map_x_min To map_current.map_x_max
            
            '.map file
            
            'Blocked
            Get #1, , TempByte
            map_current.map_grid(x, y).Blocked = (TempByte And 1)
        
            'Layer 1
            Get #1, , TempInt
            If TempInt > 0 Then Grh_Initialize map_current.map_grid(x, y).grh(1), TempInt
            
            'Layer 2 Used?
            If (TempByte And 2) Then
                Get #1, , TempInt
                If TempInt > 0 Then Grh_Initialize map_current.map_grid(x, y).grh(2), TempInt
            Else
                map_current.map_grid(x, y).grh(2).grh_index = 0
            End If
            
            'Layer 3 Used?
            If (TempByte And 4) Then
                Get #1, , TempInt
                If TempInt > 0 Then Grh_Initialize map_current.map_grid(x, y).grh(3), TempInt
            Else
                map_current.map_grid(x, y).grh(3).grh_index = 0
            End If

            'Layer 4 Used?
            If (TempByte And 8) Then
                Get #1, , TempInt
                If TempInt > 0 Then Grh_Initialize map_current.map_grid(x, y).grh(4), TempInt
            Else
                map_current.map_grid(x, y).grh(4).grh_index = 0
            End If
            
            'Trigger used?
            If (TempByte And 16) Then
                Get #1, , map_current.map_grid(x, y).Trigger
            Else
                map_current.map_grid(x, y).Trigger = 0
            End If
            
            If TempByte And 32 Then
                Get 1#, , TempLng
                Get 1#, , TempByte1
                Get 1#, , TempByte2
                Get 1#, , TempByte3
                Call Map_LightCreate(x, y, TempByte1, TempByte2, TempByte3, TempLng)
            End If
            
            If TempByte And 64 Then
                Get 1#, , map_current.map_grid(x, y).base_light(0)
                Get 1#, , map_current.map_grid(x, y).base_light(1)
                Get 1#, , map_current.map_grid(x, y).base_light(2)
                Get 1#, , map_current.map_grid(x, y).base_light(3)
                
                If map_current.map_grid(x, y).base_light(0) Then _
                    Get 1#, , map_current.map_grid(x, y).light_base_value(0)
                
                If map_current.map_grid(x, y).base_light(1) Then _
                    Get 1#, , map_current.map_grid(x, y).light_base_value(1)
                
                If map_current.map_grid(x, y).base_light(2) Then _
                    Get 1#, , map_current.map_grid(x, y).light_base_value(2)
                
                If map_current.map_grid(x, y).base_light(3) Then _
                    Get 1#, , map_current.map_grid(x, y).light_base_value(3)
            End If
            
            If TempByte And 128 Then
                Get 1#, , TempInt
                'map_current.map_grid(x, y).Particle_Group_Index = particleGroupCreate(x, y, TempInt, -1)
            End If
        Next x
    Next y
    
    Call Map_Info_Load(left(map_path, Len(map_path) - 3) & "dat")
    
    'Close .map file
    Close #1
    
    Map_Load_Map_From_File = True
End Function

Public Function Map_Fill(ByVal grh_index As Long, ByVal layer As Long, Optional ByVal light_base_color As Long = -1, _
                        Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    Dim x As Long
    Dim y As Long
    Dim loop_counter As Long
    
    Dim temp_list() As Long
    
    For y = map_current.map_y_min To map_current.map_y_max
        For x = map_current.map_x_min To map_current.map_x_max
        
            'Grh
            If Map_Grh_Set(x, y, grh_index, layer, alpha_blend, angle) = False Then
                Exit Function
            End If
        
            'Base light color
            If light_base_color <> -1 Then
                If Map_Base_Light_Set(x, y, light_base_color) = False Then
                    Exit Function
                End If
            End If
        
        Next x
    Next y
    
    Map_Fill = True
End Function

Public Function Map_Edges_Blocked_Set(ByVal edge_distance_x As Long, ByVal edge_distance_y As Long, ByVal Blocked As Byte) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/27/2003
'
'**************************************************************
    Dim x As Long
    Dim y As Long
    
    For y = map_current.map_y_min To map_current.map_y_max
        For x = map_current.map_x_min To map_current.map_x_max
            If x <= edge_distance_x Or y <= edge_distance_y Then
                map_current.map_grid(x, y).Blocked = Blocked
            End If

            If x > map_current.map_x_max - edge_distance_x Or y > map_current.map_y_max - edge_distance_y Then
                map_current.map_grid(x, y).Blocked = Blocked
            End If
        Next x
    Next y
    
    Map_Edges_Blocked_Set = True
End Function

Public Function Map_Create(ByVal map_x_max As Long, ByVal map_y_max As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'Set size
    map_current.map_x_min = 1
    map_current.map_x_max = map_x_max
    map_current.map_y_min = 1
    map_current.map_y_max = map_y_max
    
    'Erase Chars, Lights, and Particle Groups
    Char_Remove_All
    Light_Remove_All
    
    ReDim npc_list(0)
    ReDim item_list(0)
    ReDim exit_list(0)
    
    'Erase map
    ReDim map_current.map_grid(map_current.map_x_min To map_current.map_x_max, map_current.map_y_min To map_current.map_y_max) As Map_Tile
    
    'Fill in the map with grh 1 so ther is something to render
    Map_Fill 1, 1, &HAAAAAA
    
    Map_Create = True
End Function

Public Function Map_Bounds_Get(ByRef map_x_max As Long, ByRef map_y_max As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/17/2003
'
'**************************************************************
    'Get size
    map_x_max = map_current.map_x_max
    map_y_max = map_current.map_y_max
    Map_Bounds_Get = True
End Function

Public Function Map_Bounds_Get_From_File(ByVal map_name As String, ByRef max_x As Long, ByRef max_y As Long) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/27/2003
'
'**************************************************************
    Dim TempLng As Long
    
    map_name = resource_path & PATH_MAPS & "map" & map_name & ".map"
    
    'If file doesn't exists, exit
    If Not (General_File_Exists(map_name, vbNormal)) Then
        Exit Function
    End If
    
    'Open .map file
    Open map_name For Binary As #1
    Seek #1, 1
    
    'map Header
    Get #1, , TempLng
    Get #1, , max_x
    Get #1, , TempLng
    Get #1, , max_y
    
    Close #1
    
    Map_Bounds_Get_From_File = True
End Function


Public Function Map_Base_Light_Fill(ByVal light_base_color As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim x As Long
    Dim y As Long
    
    For y = map_current.map_y_min To map_current.map_y_max
        For x = map_current.map_x_min To map_current.map_x_max
            
            'Base light color
            If Map_Base_Light_Set(x, y, light_base_color) = False Then
                Exit Function
            End If
    
        Next x
    Next y
    
    Map_Base_Light_Fill = True
End Function

Public Function Map_Grh_Set(ByVal map_x As Long, ByVal map_y As Long, ByVal grh_index As Long, _
                            ByVal layer As Long, Optional ByVal alpha_blend As Boolean, _
                            Optional ByVal angle As Single) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'**************************************************************
    'Check
    If layer < 1 Or layer > 5 Then
        Exit Function
    End If
    If Map_In_Bounds(map_x, map_y) = False Then
        Exit Function
    End If
    If Grh_Check(grh_index) = False Then
        Exit Function
    End If
    
    'Do it
    Grh_Initialize map_current.map_grid(map_x, map_y).grh(layer), grh_index, alpha_blend, angle
    
    Map_Grh_Set = True
End Function

Public Function Map_Grh_UnSet(ByVal map_x As Long, ByVal map_y As Long, ByVal layer As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'**************************************************************
    Dim grh_index As Long
    
    'Check
    If layer < 1 Or layer > 5 Then
        Exit Function
    End If
    If Map_In_Bounds(map_x, map_y) = False Then
        Exit Function
    End If
    grh_index = map_current.map_grid(map_x, map_y).grh(layer).grh_index
    If Grh_Check(grh_index) = False Then
        Exit Function
    End If
    
    'Do it
    'Grh_Uninitialize map_current.map_grid(map_x, map_y).grh(layer)
    
    Map_Grh_UnSet = True
End Function

Public Function Map_Base_Light_Set(ByVal map_x As Long, ByVal map_y As Long, ByVal light_base_value As Long, Optional corner As Long = -1) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim loop_counter As Long
    
    'Check
    If Map_In_Bounds(map_x, map_y) = False Then
        Exit Function
    End If
    If corner < -1 Or corner > 3 Then
        Exit Function
    End If
    
    'Do it
    If corner = -1 Then
        'Set all corners
        For loop_counter = 0 To 3
            map_current.map_grid(map_x, map_y).light_base_value(loop_counter) = light_base_value
            map_current.map_grid(map_x, map_y).light_value(loop_counter) = light_base_value
        Next loop_counter
    Else
        'Set just one
        map_current.map_grid(map_x, map_y).light_base_value(corner) = light_base_value
        map_current.map_grid(map_x, map_y).light_value(corner) = light_base_value
    End If
    
    Map_Base_Light_Set = True
End Function


Public Function Map_Base_Light_Get(ByVal map_x As Long, ByVal map_y As Long, Optional corner As Long = 0) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Base_Light_Get = map_current.map_grid(map_x, map_y).light_base_value(corner)
    End If
End Function

Public Function Map_In_Bounds(ByVal map_x As Long, ByVal map_y As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If map_x < map_current.map_x_min Or map_x > map_current.map_x_max Or map_y < map_current.map_y_min Or map_y > map_current.map_y_max Then
        Map_In_Bounds = False
        Exit Function
    End If
    
    Map_In_Bounds = True
End Function

Public Function Map_Legal_Char_Pos(ByVal map_x As Long, ByVal map_y As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/17/2003
'Checks to see if a map position is a legal pos for a char
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) = False Then
        Exit Function
    End If
    If map_current.map_grid(map_x, map_y).Blocked Then
        Exit Function
    End If
    If map_current.map_grid(map_x, map_y).char_index Then
        Exit Function
    End If
    If UserNavegando Then
        If Not WaterFound(map_x, map_y) Then
            Exit Function
        End If
    Else
        If WaterFound(map_x, map_y) Then
            Exit Function
        End If
    End If
    Map_Legal_Char_Pos = True
End Function

Public Function Map_Legal_Char_Pos_By_Heading(ByVal char_index As Long, ByVal heading As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/17/2003
'Checks to see if a map position is a legal pos for a char
'*****************************************************************
    'Invalid heading
    If heading < 1 Or heading > 4 Then
        Exit Function
    End If
    
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        Dim nX As Long
        Dim nY As Long
        nX = char_list(char_index).map_x
        nY = char_list(char_index).map_y
        Convert_Heading_to_Direction heading, nX, nY
            
        Map_Legal_Char_Pos_By_Heading = Map_Legal_Char_Pos(nX, nY)
    End If
End Function
Public Function Map_Char_Get(ByVal map_x As Long, ByVal map_y As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a char_index and return it
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Char_Get = map_current.map_grid(map_x, map_y).char_index
    Else
        Map_Char_Get = 0
    End If
End Function

Public Function Map_Description_Get() As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'*****************************************************************
    Map_Description_Get = map_current.Map_Info.name
End Function

Public Function Map_Description_Set(ByVal s_map_description As String) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'*****************************************************************
    map_current.Map_Info.name = s_map_description
    Map_Description_Set = True
End Function

Public Function Map_Blocked_Get(ByVal map_x As Long, ByVal map_y As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Checks to see if a tile position is blocked
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Blocked_Get = map_current.map_grid(map_x, map_y).Blocked
    Else
        Map_Blocked_Get = True
    End If
End Function

Public Function Map_Blocked_Set(ByVal map_x As Long, ByVal map_y As Long, ByVal Blocked As Byte) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Sets a tile position to blocked
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        map_current.map_grid(map_x, map_y).Blocked = Blocked
        Map_Blocked_Set = True
    End If
End Function

Public Function Map_Grh_Get(ByVal map_x As Long, ByVal map_y As Long, ByVal layer As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a grh and return it
'*****************************************************************
    'Check
    If layer < 1 Or layer > 4 Then
        Map_Grh_Get = 0
        Exit Function
    End If
    
    If Map_In_Bounds(map_x, map_y) Then
        Map_Grh_Get = map_current.map_grid(map_x, map_y).grh(layer).grh_index
    Else
        Map_Grh_Get = 0
    End If
End Function

Public Function Map_Light_Get(ByVal map_x As Long, ByVal map_y As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a light_index and return it
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).map_x = map_x And light_list(loopc).map_y = map_y
        If loopc = light_last Then
            Map_Light_Get = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Map_Light_Get = loopc
Exit Function
ErrorHandler:
    Map_Light_Get = 0
End Function

Public Sub Map_Render()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/14/2003
'
'**************************************************************
    Dim map_x As Integer
    Dim map_y As Integer
    Dim view_x As Integer
    Dim view_y As Integer
    Dim Screen_X As Integer
    Dim Screen_Y As Integer
    
    Dim i As Integer
    
    Dim roofrgb_list(3) As Long
    
    '*********************
    'Handle scrolling
    'counters
    '*********************
    If scroll_on Then
        '****** Move screen Left and Right if needed ******
        If scroll_direction_x <> 0 Then
            scroll_offset_counter_x = scroll_offset_counter_x + (scroll_pixels_per_frame * timer_ticks_per_frame * scroll_direction_x)
            If Sgn(scroll_offset_counter_x) = scroll_direction_x Then
                scroll_offset_counter_x = 0
                scroll_direction_x = 0
            End If
        End If
        '****** Move screen Up and Down if needed ******
        If scroll_direction_y <> 0 Then
            scroll_offset_counter_y = scroll_offset_counter_y + (scroll_pixels_per_frame * timer_ticks_per_frame * scroll_direction_y)
            If Sgn(scroll_offset_counter_y) = scroll_direction_y Then
                scroll_offset_counter_y = 0
                scroll_direction_y = 0
            End If
        End If
        'End scrolling if needed
        If scroll_direction_x = 0 And scroll_direction_y = 0 Then
            scroll_on = False
        End If
    End If
    
    'Figure out ends and starts of view area
    view_min_y = ((view_pos_y) - (view_screen_tile_height \ 2)) - 2
    view_max_y = ((view_pos_y) + (view_screen_tile_height \ 2)) + 4
    view_min_x = ((view_pos_x) - (view_screen_tile_width \ 2)) - 2
    view_max_x = ((view_pos_x) + (view_screen_tile_width \ 2)) + 3
    
    'Add the buffer
    view_buffer_min_y = view_min_y - 2
    view_buffer_max_y = view_max_y + view_tile_buffer
    view_buffer_min_x = view_min_x - 2
    view_buffer_max_x = view_max_x + view_tile_buffer
    
    'Roof Light
    Call RoofAlphaCalculate(roofrgb_list)
    
    '*********************
    'Layer 1
    '*********************
    view_y = -1
    For map_y = view_min_y To view_max_y
        view_x = -1
        For map_x = view_min_x To view_max_x
    
            If Map_In_Bounds(map_x, map_y) Then
                '*** Start Layer 1 ***
                If map_current.map_grid(map_x, map_y).grh(1).grh_index Then
                    Screen_X = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                    Screen_Y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                    Grh_Render map_current.map_grid(map_x, map_y).grh(1), Screen_X, Screen_Y, map_current.map_grid(map_x, map_y).light_value(), False
                End If
                '*** End Layer 1 ***
            End If
            
            view_x = view_x + 1
        Next map_x
        view_y = view_y + 1
    Next map_y
    
    '*********************
    'Layer 2 and 5
    '*********************
    view_y = -3
    For map_y = view_buffer_min_y To view_buffer_max_y
        view_x = -3
        For map_x = view_buffer_min_x To view_buffer_max_x
            If Map_In_Bounds(map_x, map_y) Then
                '*** Start Layer 2 ***
                If map_current.map_grid(map_x, map_y).grh(2).grh_index Then
                    Screen_X = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                    Screen_Y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                    Grh_Render map_current.map_grid(map_x, map_y).grh(2), Screen_X, Screen_Y, map_current.map_grid(map_x, map_y).light_value(), True
                End If
                '*** End Layer 2 ***
            End If
            view_x = view_x + 1
        Next map_x
        view_y = view_y + 1
    Next map_y
    
    '*********************
    'Middle layer
    '*********************
    view_y = -3
    For map_y = view_buffer_min_y To view_buffer_max_y
        view_x = -3
        For map_x = view_buffer_min_x To view_buffer_max_x
            If Map_In_Bounds(map_x, map_y) Then
                'figure out tile screen position
                Screen_X = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                Screen_Y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
            
                '*** Start Layer 5 *** 'Special layer that is not saved and used for items
               If map_current.map_grid(map_x, map_y).grh(5).grh_index Then
                    Grh_Render map_current.map_grid(map_x, map_y).grh(5), Screen_X, Screen_Y, map_current.map_grid(map_x, map_y).light_value(), True
                End If
                '*** End Layer 5 ***
                '*** Start Layer 3 ***
                If map_current.map_grid(map_x, map_y).grh(3).grh_index Then
                    Grh_Render map_current.map_grid(map_x, map_y).grh(3), Screen_X, Screen_Y, map_current.map_grid(map_x, map_y).light_value(), True
                End If
                '*** End Layer 3 ***
                '*** Start Characters ***
                If map_current.map_grid(map_x, map_y).char_index Then
                    Char_Render char_list(map_current.map_grid(map_x, map_y).char_index), Screen_X, Screen_Y, map_current.map_grid(map_x, map_y).light_value, map_current.map_grid(map_x, map_y).char_index
                End If
                '*** End Characters ***
                
                'Fogata, despues se puede implementar algo para mayor versatilidad no? xD
                If map_current.map_grid(map_x, map_y).bonfireIndex > 0 Then
                    Call mapRenderBonfire(map_current.map_grid(map_x, map_y).bonfireIndex, Screen_X, Screen_Y)
                End If
                
                'If map_x = 43 And map_y = 43 Then Stop
                
                '*** Start Effects ***
                If map_current.map_grid(map_x, map_y).lastEffect > 0 Then
                    For i = 1 To map_current.map_grid(map_x, map_y).lastEffect
                        Call renderEffect(map_current.map_grid(map_x, map_y).effectList(i), Screen_X, Screen_Y, map_current.map_grid(map_x, map_y).light_value)
                        If Not effectDestroyed(map_current.map_grid(map_x, map_y).effectList(i)) Then
                            Call mapDestroyEffect(map_x, map_y, i)
                        End If
                    Next i
                End If
                '*** End Particles ***
            End If
            view_x = view_x + 1
        Next map_x
        view_y = view_y + 1
    Next map_y
    
    '*********************
    'Layer 4
    '*********************
    view_y = -3
    For map_y = view_buffer_min_y To view_buffer_max_y
        view_x = -3
        For map_x = view_buffer_min_x To view_buffer_max_x
            If Map_In_Bounds(map_x, map_y) Then
                '*** Start Layer 4 ***
                If map_current.map_grid(map_x, map_y).grh(4).grh_index Then
                    Screen_X = Convert_Tile_To_View_X(view_x) - scroll_offset_counter_x
                    Screen_Y = Convert_Tile_To_View_Y(view_y) - scroll_offset_counter_y
                    Grh_Render map_current.map_grid(map_x, map_y).grh(4), Screen_X, Screen_Y, roofrgb_list(), True
                End If
                '*** End Layer 4 ***
            End If
            view_x = view_x + 1
        Next map_x
        view_y = view_y + 1
    Next map_y
    
    updateProjectiles timer_ticks_per_frame
    renderProjectiles view_pos_x, view_pos_y, view_screen_tile_width, view_screen_tile_height, scroll_offset_counter_x, scroll_offset_counter_y, roofrgb_list()
    
    'Render Lights
    Map_LightsAreaRender view_min_x, view_max_x, view_min_y, view_max_y
    
    '*******************************
    'Draw Rain
    Meteo.MeteoParticle_Render scroll_direction_y
    '*******************************
End Sub
Public Function Char_Create(ByVal char_index As Integer, ByVal map_x As Long, ByVal map_y As Long, ByVal heading As Long, _
                            ByVal body_index As Long, ByVal head_index As Integer, ByVal casco_index As Integer, _
                            ByVal weapon_index As Integer, ByVal shield_index As Integer, ByVal privs As Byte, _
                            ByVal fx As Integer, ByVal FxLoopTimes As Integer, ByVal Nombre As String, ByVal criminal As Byte, ByVal Aura As Integer) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/13/2003
'Returns the char_index if successful, else 0
'**************************************************************
    'Invalid heading
    If heading < 1 Or heading > 4 Then
        Exit Function
    End If

    If Map_Char_Get(map_x, map_y) = 0 Then
        Char_Make char_index, map_x, map_y, heading, body_index, head_index, casco_index, _
            weapon_index, shield_index, privs, fx, FxLoopTimes, Nombre, criminal, Aura
    End If
End Function

Private Function Char_Check(ByVal char_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check char_index
    If char_index > 0 And char_index <= char_last Then
        If char_list(char_index).Active Then
            Char_Check = True
        End If
    End If
End Function

Private Function Char_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until char_list(loopc).Active = False
        If loopc = char_last Then
            Char_Next_Open = char_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Char_Next_Open = loopc
Exit Function
ErrorHandler:
    Char_Next_Open = 1
End Function

Public Function Char_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until char_list(loopc).id = id
        If loopc = char_last Then
            Char_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Char_Find = loopc
Exit Function
ErrorHandler:
    Char_Find = 0
End Function

Public Function Char_Move(ByVal char_index As Long, ByVal heading As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns true if successful, else false
'**************************************************************
    Dim temp_x As Long
    Dim temp_y As Long
    
    'Invalid heading
    If heading < 1 Or heading > 4 Then
        Char_Move = False
        Exit Function
    End If
    
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        'Make sure it's a legal move
        temp_x = char_list(char_index).map_x
        temp_y = char_list(char_index).map_y
        Convert_Heading_to_Direction heading, temp_x, temp_y
        If Map_In_Bounds(temp_x, temp_y) Then
            'check for another char_index
            If map_current.map_grid(temp_x, temp_y).char_index = 0 Then
                'Move it
                Char_Move_By_Heading char_index, heading
                Char_Move = True
                DoPasosFx char_index
            End If
        End If
    End If
End Function

Public Function Char_Remove(ByVal char_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Char_Check(char_index) Then
        Char_Destroy char_index
        Char_Remove = True
    End If
End Function

Public Function Char_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim index As Long
    
    For index = 1 To char_last
        'Make sure it's a legal index
        If Char_Check(index) Then
            Char_Destroy index
        End If
    Next index
    
    Char_Remove_All = True
End Function

Private Sub Char_Make(ByVal char_index As Integer, ByVal map_x As Long, ByVal map_y As Long, ByVal heading As Long, _
                            ByVal body_index As Long, ByVal head_index As Integer, ByVal casco_index As Integer, _
                            ByVal weapon_index As Integer, ByVal shield_index As Integer, ByVal privs As Byte, _
                            ByVal fx As Integer, ByVal FxLoopTimes As Integer, ByVal Nombre As String, ByVal criminal As Byte, ByVal Aura As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Makes a new character and puts it on the map
'*****************************************************************
    'Update array size
    If char_index > char_last Then
        char_last = char_index
    '    ReDim Preserve char_list(1 To char_last)
    End If
    char_count = char_count + 1
    
    'Make active
    char_list(char_index).Active = True
    
    'Heading
    char_list(char_index).heading = heading
    
    'Update char data
    If body_index > 0 Then char_list(char_index).chr_data.BodyData = Char_Data_List.BodyData(body_index)
    If head_index > 0 Then char_list(char_index).chr_data.HeadData = Char_Data_List.HeadData(head_index)
    If shield_index > 0 Then char_list(char_index).chr_data.ShieldData = Char_Data_List.ShieldData(shield_index)
    If weapon_index > 0 Then char_list(char_index).chr_data.WeaponData = Char_Data_List.WeaponData(weapon_index)
    If casco_index > 0 Then char_list(char_index).chr_data.CascoData = Char_Data_List.CascoData(casco_index)
    
    'Fx data
    char_list(char_index).chr_data.FxData.fx = fx
    char_list(char_index).chr_data.FxData.FxLoopTimes = FxLoopTimes
    Call Grh_Initialize(char_list(char_index).chr_data.FxData.FxGrh, Char_Data_List.FxData(fx).fx_grh_index, , , 1, True)
    
    char_list(char_index).label = Nombre
    
    char_list(char_index).priv = privs
    'Label Offset 'esto es una constante
    'char_list(char_index).label_offset_x = Char_Data_List(char_data_index).label_offset_x
    'char_list(char_index).label_offset_y = Char_Data_List(char_data_index).label_offset_y
    
    'Reset moving stats
    char_list(char_index).scroll_on = False
    char_list(char_index).scroll_direction_x = 0
    char_list(char_index).scroll_direction_y = 0
    char_list(char_index).scroll_offset_counter_y = 0
    char_list(char_index).scroll_offset_counter_x = 0
    
    'Update position
    char_list(char_index).map_x = map_x
    char_list(char_index).map_y = map_y
    
    Char_Set_Char_Aura char_index, Aura
    
    Debug.Print char_list(char_index).Active
    'Plot on map
    map_current.map_grid(map_x, map_y).char_index = char_index
    
    'ACORDARSE
    'char_list(char_index).light_index = Map_LightCreate(map_x, map_y, , 220, 200, 9)
End Sub

Public Function Char_Label_Set(ByVal char_index As Long, ByVal label As String, Optional ByVal label_font_index As Long = 1) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/27/2003
'Changes the character label
'*****************************************************************
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        'Label
        char_list(char_index).label = label
        char_list(char_index).label_font_index = label_font_index
        Char_Label_Set = True
    End If
End Function

Public Function Char_Heading_Set(ByVal char_index As Long, ByVal heading As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Changes the character heading
'*****************************************************************
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        char_list(char_index).heading = heading
        Char_Heading_Set = True
        Exit Function
    End If
End Function

Public Function Char_Map_Pos_Get(ByVal char_index As Long, map_x As Long, map_y As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Changes the character label
'*****************************************************************
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        'Get map pos
        map_x = char_list(char_index).map_x
        map_y = char_list(char_index).map_y
        
        Char_Map_Pos_Get = True
    End If
End Function

Public Function Char_Map_Pos_Set(ByVal char_index As Long, ByVal map_x As Long, ByVal map_y As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Changes the character label
'*****************************************************************
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        'Check map pos
        If Map_In_Bounds(map_x, map_y) Then
            'Move char
            map_current.map_grid(char_list(char_index).map_x, char_list(char_index).map_y).char_index = 0
            char_list(char_index).map_x = map_x
            char_list(char_index).map_y = map_y
            map_current.map_grid(char_list(char_index).map_x, char_list(char_index).map_y).char_index = char_index
            Char_Map_Pos_Set = True
        End If
    End If
End Function

Public Sub Char_Move_By_Pos(ByVal char_index As Integer, ByVal nX As Integer, ByVal nY As Integer)
    
    'Make sure it's a legal char_index
    If Not Char_Check(char_index) Then Exit Sub
    
    'Movemos al char a partir de una posicion. (esto no me gusta el server deberia dar la direccion...)
    Dim x As Long, y As Long
    
    Dim add_to_x As Integer
    Dim add_to_y As Integer
    
    Call Char_Map_Pos_Get(char_index, x, y)
    
    add_to_x = nX - x
    add_to_y = nY - y
    
    Call Char_Move_By_Heading(char_index, Convert_Direction_To_Heading(Sgn(add_to_x), Sgn(add_to_y)))
    
    'Aca se deberia borrar el pj si sale del area?
    If Not isPcArea(nX, nY) Then
        Char_Remove char_index
    End If
End Sub

Private Sub Char_Move_By_Heading(ByVal char_index As Long, ByVal heading As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim x As Long
    Dim y As Long
    Dim nX As Long
    Dim nY As Long
    
    x = char_list(char_index).map_x
    y = char_list(char_index).map_y
    
    nX = x
    nY = y
    Convert_Heading_to_Direction heading, nX, nY
    
    map_current.map_grid(nX, nY).char_index = char_index
    char_list(char_index).map_x = nX
    char_list(char_index).map_y = nY
    map_current.map_grid(x, y).char_index = 0
    
    char_list(char_index).scroll_offset_counter_x = (base_tile_size * (x - nX))
    char_list(char_index).scroll_offset_counter_y = (base_tile_size * (y - nY))
    char_list(char_index).scroll_direction_x = nX - x
    char_list(char_index).scroll_direction_y = nY - y
    
    char_list(char_index).scroll_on = True
    char_list(char_index).heading = heading
    'Set char to walk
    char_list(char_index).chr_data_body_index = 2
End Sub

Private Sub Char_Destroy(ByVal char_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As Char
    
    char_list(char_index).Active = False
    map_current.map_grid(char_list(char_index).map_x, char_list(char_index).map_y).char_index = 0
    
    'Update array size
    If char_index = char_last Then
        Do Until char_list(char_last).Active
            char_last = char_last - 1
            If char_last = 0 Then
                char_count = 0
                Exit Sub
            End If
        Loop
    End If
    char_count = char_count - 1
End Sub

Private Sub Char_Load_Char_Data()
'*****************************************************************
'Carga cabezas, cuerpos, cascos, armas, fx 'Los FXs Son del usuario?
'*****************************************************************
    Call Char_Data_Body_Load
    Call Char_Data_Head_Load
    Call Char_Data_Shield_Load
    Call Char_Data_Weapon_Load
    Call Char_Data_Casco_Load
    Call Char_Data_Fx_Load
    Call Char_Data_Aura_Load
    
End Sub
'TODO: Esta funcion tiene que cambiar para renderizar los cuerpos del ao ya que la cabeza es independiente al cuerpo.
Private Sub Char_Render(ByRef temp_char As Char, ByVal Screen_X As Long, ByVal Screen_Y As Long, ByRef light_value() As Long, ByVal char_index As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 12/28/2002
'Renders a character at a paticular screen point
'*****************************************************************
    Dim Pos As Long
    Dim color As Long
    
    'Scrolling
    If temp_char.scroll_on Then
        '****** Move Left and Right if needed ******
        If temp_char.scroll_direction_x <> 0 Then
            temp_char.scroll_offset_counter_x = temp_char.scroll_offset_counter_x + (scroll_pixels_per_frame * timer_ticks_per_frame * temp_char.scroll_direction_x)
            If Sgn(temp_char.scroll_offset_counter_x) = temp_char.scroll_direction_x Then
                temp_char.scroll_offset_counter_x = 0
                temp_char.scroll_direction_x = 0
            End If
        End If
        '****** Move Up and Down if needed ******
        If temp_char.scroll_direction_y <> 0 Then
            temp_char.scroll_offset_counter_y = temp_char.scroll_offset_counter_y + (scroll_pixels_per_frame * timer_ticks_per_frame * temp_char.scroll_direction_y)
            If Sgn(temp_char.scroll_offset_counter_y) = temp_char.scroll_direction_y Then
                temp_char.scroll_offset_counter_y = 0
                temp_char.scroll_direction_y = 0
            End If
        End If
        'End scrolling if needed
        If temp_char.scroll_direction_x = 0 And temp_char.scroll_direction_y = 0 Then
            'Turn off scrolling
            temp_char.scroll_on = False
        Else
            temp_char.chr_data.BodyData.Body(temp_char.heading).Started = 1
            temp_char.chr_data.WeaponData.WeaponWalk(temp_char.heading).Started = 1
            temp_char.chr_data.ShieldData.ShieldWalk(temp_char.heading).Started = 1
        End If
    Else
        'Set char to stand
        temp_char.chr_data.BodyData.Body(temp_char.heading).Started = 0
        temp_char.chr_data.WeaponData.WeaponWalk(temp_char.heading).Started = 0
        temp_char.chr_data.ShieldData.ShieldWalk(temp_char.heading).Started = 0
        'Set the anim to stand
        temp_char.chr_data.BodyData.Body(temp_char.heading).frame_counter = 1
        temp_char.chr_data.WeaponData.WeaponWalk(temp_char.heading).frame_counter = 1
        temp_char.chr_data.ShieldData.ShieldWalk(temp_char.heading).frame_counter = 1
    End If
    
    'Find screen position
    Screen_X = Screen_X + temp_char.scroll_offset_counter_x
    Screen_Y = Screen_Y + temp_char.scroll_offset_counter_y
    
    'Actualizamos la posicion del objetivo del projectil.
    'updateProjectileTarget temp_char.chr_data.projectileIndex, temp_char.map_x * 32, temp_char.map_y * 32
    
    'Por ahora no lo dibujamos
    
    'Render Body Grh
    If temp_char.chr_data.AuraData.Aura Then _
        Grh_Render temp_char.chr_data.AuraData.AuraGrh, Screen_X, Screen_Y - 16, Char_Data_List.AuraData(temp_char.chr_data.AuraData.Aura).rgb_list, True, True
    
    'Render Body Grh
    If temp_char.chr_data.BodyData.Body(temp_char.heading).grh_index Then _
        Grh_Render temp_char.chr_data.BodyData.Body(temp_char.heading), Screen_X, Screen_Y, light_value(), True
        
    'Render Head Grh
    If temp_char.chr_data.HeadData.Head(temp_char.heading).grh_index Then _
        Grh_Render temp_char.chr_data.HeadData.Head(temp_char.heading), Screen_X + temp_char.chr_data.BodyData.HeadOffset.x, Screen_Y + temp_char.chr_data.BodyData.HeadOffset.y, light_value(), True
        
    'Render Weapon Grh
    If temp_char.chr_data.WeaponData.WeaponWalk(temp_char.heading).grh_index Then _
        Grh_Render temp_char.chr_data.WeaponData.WeaponWalk(temp_char.heading), Screen_X, Screen_Y, light_value(), True
    
    'Render Shield Grh
    If temp_char.chr_data.ShieldData.ShieldWalk(temp_char.heading).grh_index Then _
        Grh_Render temp_char.chr_data.ShieldData.ShieldWalk(temp_char.heading), Screen_X, Screen_Y, light_value(), True
    
    'Render Helmet Grh
    If temp_char.chr_data.CascoData.Head(temp_char.heading).grh_index Then _
        Grh_Render temp_char.chr_data.CascoData.Head(temp_char.heading), Screen_X + temp_char.chr_data.BodyData.HeadOffset.x, Screen_Y + temp_char.chr_data.BodyData.HeadOffset.y, light_value(), True
        
    'Name
    If temp_char.label <> "" Then
        If temp_char.priv = 0 Then
            If temp_char.faccion = 0 Then
                color = D3DColorXRGB(ColoresPJ(50).R, ColoresPJ(50).G, ColoresPJ(50).B)
            ElseIf temp_char.faccion = 1 Then
                color = D3DColorXRGB(ColoresPJ(49).R, ColoresPJ(49).G, ColoresPJ(49).B)
            Else
                color = D3DColorXRGB(ColoresPJ(51).R, ColoresPJ(51).G, ColoresPJ(51).B)
            End If
        Else
            color = D3DColorXRGB(ColoresPJ(temp_char.priv).R, ColoresPJ(temp_char.priv).G, ColoresPJ(temp_char.priv).B)
        End If

        Pos = InStr(temp_char.label, "<")
        
        If Pos Then
            Call dxTextRender(1, left$(temp_char.label, Pos - 2), Screen_X + temp_char.label_offset_x - base_tile_size, Screen_Y + temp_char.label_offset_y + 30, color, DT_TOP Or DT_CENTER, 100, 20)
            Call dxTextRender(1, Mid$(temp_char.label, Pos), Screen_X + temp_char.label_offset_x - base_tile_size, Screen_Y + temp_char.label_offset_y + 45, color, DT_TOP Or DT_CENTER, 100, 20)
        Else
            Call dxTextRender(1, temp_char.label, Screen_X - base_tile_size + temp_char.label_offset_x, Screen_Y + temp_char.label_offset_y + 30, color, DT_TOP Or DT_CENTER, 100, 20)
        End If
    End If
    
    'Update Light Pos
    If temp_char.light_index Then
        Call Map_LightMove(temp_char.light_index, temp_char.map_x, temp_char.map_y, temp_char.map_x * 32 + temp_char.scroll_offset_counter_x, temp_char.map_y * 32 + temp_char.scroll_offset_counter_y)
    End If
    
    'RENDER EFFECTS!
    Dim i As Integer

    For i = 1 To temp_char.chr_data.lastEffect
        renderEffect temp_char.chr_data.effectList(i), Screen_X, Screen_Y, light_value()
        If Not effectDestroyed(temp_char.chr_data.effectList(i)) Then _
            charDestroyEffect char_index, i
    Next i
    
    'If the char moved then update projectile target.
    If temp_char.chr_data.lastProjectile > 0 Then
        For i = 1 To temp_char.chr_data.lastProjectile
            updateProjectileTarget temp_char.chr_data.projectileIndex(i), temp_char.map_x, temp_char.map_y
        Next i
    End If
    
    'Update dialogs
    If Dialogos.CantidadDialogos > 0 Then
        Call Dialogos.Update_Dialog_Pos(Screen_X + temp_char.chr_data.BodyData.HeadOffset.x, Screen_Y + temp_char.chr_data.BodyData.HeadOffset.y, char_index)
    End If
    
End Sub

Public Function Light_Remove(ByVal light_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        Map_LightDestroy light_index
        Light_Remove = True
    End If
End Function

Public Function Light_Color_Value_Get(ByVal light_index As Long, ByRef color_value As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        color_value = light_list(light_index).color
        Light_Color_Value_Get = True
    End If
End Function

Public Function Map_LightCreate(ByVal map_x As Long, ByVal map_y As Long, Optional ByVal R As Byte = 255, _
                            Optional ByVal G As Byte = 255, Optional ByVal B As Byte = 255, _
                            Optional ByVal range As Long = 1, _
                            Optional ByVal id As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns the light_index if successful, else 0
'**************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_LightCreate = Light_Next_Open
        Light_Make Map_LightCreate, map_x, map_y, R, G, B, range, id
    End If
End Function
Private Sub Light_Make(ByVal light_index As Long, ByVal map_x As Long, _
                        ByVal map_y As Long, ByVal R As Long, _
                        ByVal G As Long, ByVal B As Long, _
                        ByVal range As Long, Optional ByVal id As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    'Update array size
    If light_index > light_last Then
        light_last = light_index
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count + 1
    
    'Make active
    light_list(light_index).Active = True
    
    light_list(light_index).map_x = map_x
    light_list(light_index).map_y = map_y
    light_list(light_index).x = map_x * base_tile_size
    light_list(light_index).y = map_y * base_tile_size
    light_list(light_index).color = D3DColorXRGB(R, G, B)
    light_list(light_index).RGBColor.R = R
    light_list(light_index).RGBColor.G = G
    light_list(light_index).RGBColor.B = B
    light_list(light_index).range = range
    light_list(light_index).id = id
    
    'set light on tile
    'map_current.map_grid(map_x, map_y).light_index = light_index
End Sub

Private Function Light_Check(ByVal light_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check light_index
    If light_index > 0 And light_index <= light_last Then
        If light_list(light_index).Active Then
            Light_Check = True
        End If
    End If
End Function

Public Sub Map_LightsRenderAll()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim loop_counter As Long
            
    For loop_counter = 1 To light_count
        
        If light_list(loop_counter).Active Then
            Map_LightRender loop_counter
        End If
    
    Next loop_counter
End Sub

Private Sub Map_LightRenderSquare(ByVal light_index, Optional ByVal map_x As Long = -1, Optional ByVal map_y As Long = -1, _
                        Optional ByVal rgb_value As Long = -1, Optional ByVal range As Long = -1)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim x As Integer
    Dim y As Integer
    Dim color As Long
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    'Set color
    color = light_list(light_index).color
    
    'Arrange corners
    'NE
    If Map_In_Bounds(min_x, min_y) Then
        map_current.map_grid(min_x, min_y).light_value(2) = color
    End If
    'NW
    If Map_In_Bounds(max_x, min_y) Then
        map_current.map_grid(max_x, min_y).light_value(0) = color
    End If
    'SW
    If Map_In_Bounds(max_x, max_y) Then
        map_current.map_grid(max_x, max_y).light_value(1) = color
    End If
    'SE
    If Map_In_Bounds(min_x, max_y) Then
        map_current.map_grid(min_x, max_y).light_value(3) = color
    End If
    
    'Arrange borders
    'Upper border
    For x = min_x + 1 To max_x - 1
        If Map_In_Bounds(x, min_y) Then
            map_current.map_grid(x, min_y).light_value(0) = color
            map_current.map_grid(x, min_y).light_value(2) = color
        End If
    Next x
    
    'Lower border
    For x = min_x + 1 To max_x - 1
        If Map_In_Bounds(x, max_y) Then
            map_current.map_grid(x, max_y).light_value(1) = color
            map_current.map_grid(x, max_y).light_value(3) = color
        End If
    Next x
    
    'Left border
    For y = min_y + 1 To max_y - 1
        If Map_In_Bounds(min_x, y) Then
            map_current.map_grid(min_x, y).light_value(2) = color
            map_current.map_grid(min_x, y).light_value(3) = color
        End If
    Next y
    
    'Right border
    For y = min_y + 1 To max_y - 1
        If Map_In_Bounds(max_x, y) Then
            map_current.map_grid(max_x, y).light_value(0) = color
            map_current.map_grid(max_x, y).light_value(1) = color
        End If
    Next y
    
    'Set the inner part of the light
    For x = min_x + 1 To max_x - 1
        For y = min_y + 1 To max_y - 1
            If Map_In_Bounds(x, y) Then
                map_current.map_grid(x, y).light_value(0) = color
                map_current.map_grid(x, y).light_value(1) = color
                map_current.map_grid(x, y).light_value(2) = color
                map_current.map_grid(x, y).light_value(3) = color
            End If
        Next y
    Next x
End Sub

Private Function Light_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).Active = False
        If loopc = light_last Then
            Light_Next_Open = light_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Next_Open = loopc
Exit Function
ErrorHandler:
    Light_Next_Open = 1
End Function

Public Function Light_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).id = id
        If loopc = light_last Then
            Light_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Find = loopc
Exit Function
ErrorHandler:
    Light_Find = 0
End Function

Public Function Light_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim index As Long
    
    For index = 1 To light_last
        'Make sure it's a legal index
        If Light_Check(index) Then
            Map_LightDestroy index
        End If
    Next index
    
    Light_Remove_All = True
End Function

Private Sub Map_LightDestroy(ByVal light_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As Light
    
    light_list(light_index) = temp
    
    'Update array size
    If light_index = light_last Then
        Do Until light_list(light_last).Active
            light_last = light_last - 1
            If light_last = 0 Then
                light_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count - 1
End Sub

Public Function User_Char_Index_Set(ByVal index As Integer)
    user_char_index = index
    'Nueva posicion para el usuario.
    Call Engine.Engine_View_Pos_Set(char_list(user_char_index).map_x, char_list(user_char_index).map_y)
End Function

Public Function User_Char_Index_Get() As Integer
    User_Char_Index_Get = user_char_index
End Function

Public Function Char_Heading_Get(ByVal char_index As Integer) As Long
    Char_Heading_Get = char_list(char_index).heading
End Function

Public Function Char_Label_Get(ByVal char_index As Integer) As String
    Char_Label_Get = char_list(char_index).label
End Function

Public Function Char_Faccion_Set(ByVal char_index As Integer, ByVal value As Integer)
    char_list(char_index).faccion = value
End Function

Public Function Char_Invisible_Get(ByVal char_index As Integer) As Boolean
    Char_Invisible_Get = char_list(char_index).Invisible
End Function

Public Function Char_Invisible_Set(ByVal char_index As Integer, ByVal value As Boolean) As Boolean
    char_list(char_index).Invisible = value
End Function

Public Function IsIndoor() As Boolean
    With map_current.map_grid(char_list(user_char_index).map_x, char_list(user_char_index).map_y)
        If .Trigger = eTrigger.BAJOTECHO Then IsIndoor = True
    End With
End Function
Private Sub RoofAlphaCalculate(roofrgb_list() As Long)
    Dim color As D3DCOLORVALUE
    Static last_tick As Long

    With map_current.map_grid(char_list(user_char_index).map_x, char_list(user_char_index).map_y)
        If .Trigger = eTrigger.BAJOTECHO Then
            If AlphaRoof > 0 Then
                If GetTickCount - last_tick >= AlphaRoofInterval Then
                    AlphaRoof = AlphaRoof - 5
                    last_tick = GetTickCount
                End If
            End If
        Else
            If AlphaRoof < 255 Then
                If GetTickCount - last_tick >= AlphaRoofInterval Then
                    AlphaRoof = AlphaRoof + 5
                    last_tick = GetTickCount
                End If
            End If
        End If
    End With
    
    'Get roof light.
    Meteo.Get_AmbientLight color
    color.A = AlphaRoof 'Por ahora
    D3DColorToRgbList roofrgb_list(), color

    
End Sub
Public Function Char_Set_Char_Fx(ByVal char_index As Integer, ByVal fx As Integer, ByVal FxLoopTimes As Integer)
    If fx < 1 Then Exit Function
    If FxLoopTimes = LoopAdEternum Then
        Call charCreateEffect(char_index, fx, -1)
    Else
        Call charCreateEffect(char_index, fx)
    End If
End Function
Public Sub Char_Set_Char_Aura(ByVal char_index As Integer, ByVal Aura As Integer)
    char_list(char_index).chr_data.AuraData.Aura = Aura
    If Aura > 0 Then _
        Grh_Initialize char_list(char_index).chr_data.AuraData.AuraGrh, Char_Data_List.AuraData(Aura).grhIndex, True
End Sub
Public Function Char_Set_Char_Body(ByVal char_index As Integer, ByVal body_index As Integer)
    If body_index > UBound(Char_Data_List.BodyData) Or body_index < LBound(Char_Data_List.BodyData) Then Exit Function
    char_list(char_index).chr_data.BodyData = Char_Data_List.BodyData(body_index)
End Function

Public Function Char_Set_Char_Head(ByVal char_index As Integer, ByVal head_index As Integer)
    If head_index > UBound(Char_Data_List.HeadData) Or head_index < LBound(Char_Data_List.HeadData) Then Exit Function
    char_list(char_index).chr_data.HeadData = Char_Data_List.HeadData(head_index)
End Function

Public Function Char_Set_Char_Casco(ByVal char_index As Integer, ByVal casco_index As Integer)
    If casco_index > UBound(Char_Data_List.CascoData) Or casco_index < LBound(Char_Data_List.CascoData) Then Exit Function
    char_list(char_index).chr_data.CascoData = Char_Data_List.CascoData(casco_index)
End Function

Public Function Char_Set_Char_Weapon(ByVal char_index As Integer, ByVal weapon_index As Integer)
    If weapon_index > UBound(Char_Data_List.WeaponData) Or weapon_index < LBound(Char_Data_List.WeaponData) Then Exit Function
    char_list(char_index).chr_data.WeaponData = Char_Data_List.WeaponData(weapon_index)
End Function

Public Function Char_Set_Char_Shield(ByVal char_index As Integer, ByVal shield_index As Integer)
    If shield_index > UBound(Char_Data_List.ShieldData) Or shield_index < LBound(Char_Data_List.ShieldData) Then Exit Function
    char_list(char_index).chr_data.ShieldData = Char_Data_List.ShieldData(shield_index)
End Function
'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
'TODO : Me parece que esta al pedo...
Public Sub Char_Refresh_All()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopc As Long
    Dim x As Integer
    Dim y As Integer
    
    For loopc = 1 To char_last
        If char_list(loopc).Active Then
            map_current.map_grid(char_list(loopc).map_x, char_list(loopc).map_y).char_index = loopc
        End If
    Next loopc
End Sub
Public Sub Char_Data_Body_Load()
    On Error Resume Next
    
    Dim N As Integer, i As Integer
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile
    Open resource_path & PATH_INIT & "\Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim Char_Data_List.BodyData(0 To NumCuerpos + 1) As Char_Data_Body
    ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        With Char_Data_List
            Get #N, , MisCuerpos(i)
            Grh_Initialize .BodyData(i).Body(1), MisCuerpos(i).Body(1), , , 0
            Grh_Initialize .BodyData(i).Body(2), MisCuerpos(i).Body(2), , , 0
            Grh_Initialize .BodyData(i).Body(3), MisCuerpos(i).Body(3), , , 0
            Grh_Initialize .BodyData(i).Body(4), MisCuerpos(i).Body(4), , , 0
            .BodyData(i).HeadOffset.x = MisCuerpos(i).HeadOffsetX
            .BodyData(i).HeadOffset.y = MisCuerpos(i).HeadOffsetY
        End With
    Next i
    
    Close #N
End Sub

Public Sub Char_Data_Head_Load()
    On Error Resume Next
    Dim N As Integer, i As Integer, Numheads As Integer, index As Integer
    
    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile
    Open resource_path & PATH_INIT & "\Cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim Char_Data_List.HeadData(0 To Numheads + 1) As Char_Data_Head
    ReDim Miscabezas(0 To Numheads + 1) As tIndiceCabeza
    
    For i = 1 To Numheads
        With Char_Data_List
        Get #N, , Miscabezas(i)
            Grh_Initialize .HeadData(i).Head(1), Miscabezas(i).Head(1), , , 0
            Grh_Initialize .HeadData(i).Head(2), Miscabezas(i).Head(2), , , 0
            Grh_Initialize .HeadData(i).Head(3), Miscabezas(i).Head(3), , , 0
            Grh_Initialize .HeadData(i).Head(4), Miscabezas(i).Head(4), , , 0
        End With
    Next i
    
    Close #N
End Sub

Sub Char_Data_Aura_Load()
    Dim AurasCount As Integer
    Dim i As Integer
    
    AurasCount = Val(General_Var_Get(resource_path & PATH_INIT & "\Auras.dat", "INIT", "AurasCount"))
    
    ReDim Char_Data_List.AuraData(1 To AurasCount) As tAuraData
    
    For i = 1 To AurasCount
        Char_Data_List.AuraData(i).grhIndex = Val(General_Var_Get(resource_path & PATH_INIT & "\Auras.dat", "AURA" & i, "GrhIndex"))
        Char_Data_List.AuraData(i).rgb_list(0) = D3DColorARGB(Val(General_Var_Get(resource_path & PATH_INIT & "\Auras.dat", "AURA" & i, "A")), Val(General_Var_Get(resource_path & PATH_INIT & "\Auras.dat", "AURA" & i, "R")), Val(General_Var_Get(resource_path & PATH_INIT & "\Auras.dat", "AURA" & i, "G")), Val(General_Var_Get(resource_path & PATH_INIT & "\Auras.dat", "AURA" & i, "B")))
        Char_Data_List.AuraData(i).rgb_list(1) = D3DColorARGB(Val(General_Var_Get(resource_path & PATH_INIT & "\Auras.dat", "AURA" & i, "A")), Val(General_Var_Get(resource_path & PATH_INIT & "\Auras.dat", "AURA" & i, "R")), Val(General_Var_Get(resource_path & PATH_INIT & "\Auras.dat", "AURA" & i, "G")), Val(General_Var_Get(resource_path & PATH_INIT & "\Auras.dat", "AURA" & i, "B")))
        Char_Data_List.AuraData(i).rgb_list(2) = D3DColorARGB(Val(General_Var_Get(resource_path & PATH_INIT & "\Auras.dat", "AURA" & i, "A")), Val(General_Var_Get(resource_path & PATH_INIT & "\Auras.dat", "AURA" & i, "R")), Val(General_Var_Get(resource_path & PATH_INIT & "\Auras.dat", "AURA" & i, "G")), Val(General_Var_Get(resource_path & PATH_INIT & "\Auras.dat", "AURA" & i, "B")))
        Char_Data_List.AuraData(i).rgb_list(3) = D3DColorARGB(Val(General_Var_Get(resource_path & PATH_INIT & "\Auras.dat", "AURA" & i, "A")), Val(General_Var_Get(resource_path & PATH_INIT & "\Auras.dat", "AURA" & i, "R")), Val(General_Var_Get(resource_path & PATH_INIT & "\Auras.dat", "AURA" & i, "G")), Val(General_Var_Get(resource_path & PATH_INIT & "\Auras.dat", "AURA" & i, "B")))
    Next i
    
End Sub

Sub Char_Data_Fx_Load()
On Error Resume Next
    Dim N As Integer, i As Integer
    Dim NumFxs As Integer
    Dim MisFxs() As tIndiceFx
    
    N = FreeFile
    
    Open resource_path & PATH_INIT & "\Fxs.ind" For Binary Access Read As #N
    
        'cabecera
        Get #N, , MiCabecera
        
        'num de cabezas
        Get #N, , NumFxs
    
        'Resize array
        ReDim Char_Data_List.FxData(0 To NumFxs + 1) As Char_Data_Fx
        ReDim MisFxs(0 To NumFxs + 1) As tIndiceFx
        
        For i = 1 To NumFxs
            Get #N, , MisFxs(i)
            With Char_Data_List
                .FxData(i).fx_grh_index = MisFxs(i).Animacion
                .FxData(i).fx_offset.x = MisFxs(i).offSetX
                .FxData(i).fx_offset.y = MisFxs(i).offSetY
            End With
        Next i
    
    Close #N
    
End Sub

Sub Char_Data_Casco_Load()
On Error Resume Next
Dim N As Integer, i As Integer, NumCascos As Integer, index As Integer

Dim Miscabezas() As tIndiceCabeza

N = FreeFile
Open resource_path & PATH_INIT & "\Cascos.ind" For Binary Access Read As #N

'cabecera
Get #N, , MiCabecera

'num de cabezas
Get #N, , NumCascos

'Resize array
ReDim Char_Data_List.CascoData(0 To NumCascos + 1) As Char_Data_Head
ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza

For i = 1 To NumCascos
    Get #N, , Miscabezas(i)
    With Char_Data_List
        Grh_Initialize .CascoData(i).Head(1), Miscabezas(i).Head(1), , , 0
        Grh_Initialize .CascoData(i).Head(2), Miscabezas(i).Head(2), , , 0
        Grh_Initialize .CascoData(i).Head(3), Miscabezas(i).Head(3), , , 0
        Grh_Initialize .CascoData(i).Head(4), Miscabezas(i).Head(4), , , 0
    End With
Next i

Close #N

End Sub

Sub Char_Data_Weapon_Load()
On Error Resume Next
    
    Dim loopc As Long
    Dim arch As String
    
    arch = resource_path & PATH_INIT & "\armas.dat"
    
    Char_Data_List.NumWeaponAnims = Val(General_Var_Get(arch, "INIT", "NumArmas"))
    
    ReDim Char_Data_List.WeaponData(0 To Char_Data_List.NumWeaponAnims) As Char_Data_Weapon
    
    For loopc = 1 To Char_Data_List.NumWeaponAnims
        With Char_Data_List
        Grh_Initialize .WeaponData(loopc).WeaponWalk(1), Val(General_Var_Get(arch, "ARMA" & loopc, "Dir1")), , , 0
        Grh_Initialize .WeaponData(loopc).WeaponWalk(2), Val(General_Var_Get(arch, "ARMA" & loopc, "Dir2")), , , 0
        Grh_Initialize .WeaponData(loopc).WeaponWalk(3), Val(General_Var_Get(arch, "ARMA" & loopc, "Dir3")), , , 0
        Grh_Initialize .WeaponData(loopc).WeaponWalk(4), Val(General_Var_Get(arch, "ARMA" & loopc, "Dir4")), , , 0
        End With
    Next loopc
End Sub

Sub Char_Data_Shield_Load()
On Error Resume Next

    Dim loopc As Long
    
    Dim arch As String
    
    arch = resource_path & PATH_INIT & "\escudos.dat"
    
    Char_Data_List.NumShieldAnims = Val(General_Var_Get(arch, "INIT", "NumEscudos"))
    
    ReDim Char_Data_List.ShieldData(0 To Char_Data_List.NumShieldAnims) As Char_Data_Shield
    
    For loopc = 1 To Char_Data_List.NumShieldAnims
        With Char_Data_List
            Grh_Initialize .ShieldData(loopc).ShieldWalk(1), Val(General_Var_Get(arch, "ESC" & loopc, "Dir1")), , , 0
            Grh_Initialize .ShieldData(loopc).ShieldWalk(2), Val(General_Var_Get(arch, "ESC" & loopc, "Dir2")), , , 0
            Grh_Initialize .ShieldData(loopc).ShieldWalk(3), Val(General_Var_Get(arch, "ESC" & loopc, "Dir3")), , , 0
            Grh_Initialize .ShieldData(loopc).ShieldWalk(4), Val(General_Var_Get(arch, "ESC" & loopc, "Dir4")), , , 0
        End With
    Next loopc
End Sub

Public Function Look_For_Name_In_Char_List(ByVal Nombre As String) As Integer

    If Nombre = "" Then Exit Function
    Dim i As Integer
    
    i = 1
    
    Do While i <= char_last
        If char_list(i).label = Nombre Then
            Look_For_Name_In_Char_List = i
            Exit Function
        Else
            i = i + 1
        End If
    Loop
    
    'Si llegamos hasta ak no lo encontro
    Look_For_Name_In_Char_List = 0
End Function

Public Function Player_Moving() As Boolean
    Player_Moving = scroll_on
End Function

Public Function User_Char_Map_Pos_Set(ByVal map_x As Integer, ByVal map_y As Integer)
    'Move char
    map_current.map_grid(char_list(user_char_index).map_x, char_list(user_char_index).map_y).char_index = 0
    char_list(user_char_index).map_x = map_x
    char_list(user_char_index).map_y = map_y
    map_current.map_grid(char_list(user_char_index).map_x, char_list(user_char_index).map_y).char_index = user_char_index
    'Set the screen`s new position
    Call Engine_View_Pos_Set(map_x, map_y)
End Function
Public Function Input_Mouse_Tile_Get(ByVal input_mouse_view_x As Integer, ByVal input_mouse_view_y As Integer, ByRef tX As Long, ByRef tY As Long) As Boolean
    Call Convert_View_To_Map(input_mouse_view_x, input_mouse_view_y, tX, tY)
End Function

Public Function TileEngineSpeedCalculate(ByVal timer_elapsed_time As Single)
    'Get timing info
    timer_ticks_per_frame = (timer_elapsed_time * engine_base_speed)
End Function
Public Sub Map_LightsAreaRender(ByVal min_x As Integer, ByVal max_x As Integer, ByVal min_y As Integer, ByVal max_y As Integer)
    Dim i As Integer
    
    Call Map_LightAreaClear(min_x - light_buffer, max_x + light_buffer, min_y - light_buffer, max_y + light_buffer)
    
    For i = 1 To light_count
        If light_list(i).map_x > min_x - light_buffer And light_list(i).map_x < max_x + light_buffer Then
            If light_list(i).map_y > min_y - light_buffer And light_list(i).map_y < max_y + light_buffer Then
                Map_LightRender (i)
            End If
        End If
    Next i
End Sub
Public Sub Map_LightsRender()
    Dim i As Integer
    
    Call Map_LightsClear
    
    For i = 1 To light_count
        Map_LightRender (i)
    Next i
End Sub
Private Sub Map_LightRender(ByVal light_index As Integer)
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim color As Long
    Dim Ya As Integer
    Dim Xa As Integer
    
    Dim TileLight As D3DCOLORVALUE
    Dim AmbientColor As D3DCOLORVALUE
    Dim LightColor As D3DCOLORVALUE
    
    Dim XCoord As Integer
    Dim YCoord As Integer
        
        LightColor = light_list(light_index).RGBColor
        Meteo.Get_AmbientLight AmbientColor
        
        min_x = light_list(light_index).map_x - light_list(light_index).range
        max_x = light_list(light_index).map_x + light_list(light_index).range
        min_y = light_list(light_index).map_y - light_list(light_index).range
        max_y = light_list(light_index).map_y + light_list(light_index).range
        
        For Ya = min_y To max_y
            For Xa = min_x To max_x
                If Map_In_Bounds(Xa, Ya) Then
                    XCoord = Xa * base_tile_size
                    YCoord = Ya * base_tile_size
                    'Color = LightCalculate(light_list(light_index).range, light_list(light_index).x, light_list(light_index).y, XCoord, YCoord, map_current.map_grid(Xa, Ya).light_value(1), LightColor, AmbientColor)
                    map_current.map_grid(Xa, Ya).light_value(1) = LightCalculate(light_list(light_index).range, light_list(light_index).x, light_list(light_index).y, XCoord, YCoord, map_current.map_grid(Xa, Ya).light_value(1), LightColor, AmbientColor)

                    XCoord = Xa * base_tile_size + base_tile_size
                    YCoord = Ya * base_tile_size
                    map_current.map_grid(Xa, Ya).light_value(3) = LightCalculate(light_list(light_index).range, light_list(light_index).x, light_list(light_index).y, XCoord, YCoord, map_current.map_grid(Xa, Ya).light_value(3), LightColor, AmbientColor)
                       
                    XCoord = Xa * base_tile_size
                    YCoord = Ya * base_tile_size + base_tile_size
                    map_current.map_grid(Xa, Ya).light_value(0) = LightCalculate(light_list(light_index).range, light_list(light_index).x, light_list(light_index).y, XCoord, YCoord, map_current.map_grid(Xa, Ya).light_value(0), LightColor, AmbientColor)
    
                    XCoord = Xa * base_tile_size + base_tile_size
                    YCoord = Ya * base_tile_size + base_tile_size
                    map_current.map_grid(Xa, Ya).light_value(2) = LightCalculate(light_list(light_index).range, light_list(light_index).x, light_list(light_index).y, XCoord, YCoord, map_current.map_grid(Xa, Ya).light_value(2), LightColor, AmbientColor)
                End If
            Next Xa
        Next Ya
End Sub
Private Sub Map_LightMove(ByVal light_index As Integer, ByVal map_x As Byte, ByVal map_y As Byte, ByVal x As Integer, ByVal y As Integer)
    If light_index = 0 Then Exit Sub
    If Map_In_Bounds(map_x, map_y) Then
        light_list(light_index).map_x = map_x
        light_list(light_index).map_y = map_y
        light_list(light_index).x = x
        light_list(light_index).y = y
    End If
End Sub

Private Function Map_LightAreaClear(ByVal min_x As Integer, ByVal max_x As Integer, ByVal min_y As Integer, ByVal max_y As Integer)
    Dim x As Integer
    Dim y As Integer
    
    Dim AmbientColor As D3DCOLORVALUE
    Dim color As Long
    
    Meteo.Get_AmbientLight AmbientColor
    color = D3DColorXRGB(AmbientColor.R, AmbientColor.G, AmbientColor.B)
    
    For x = min_x To max_x
        For y = min_y To max_y
            If Map_In_Bounds(x, y) Then
                With map_current.map_grid(x, y)
                    If .base_light(0) Then 'Si tiene luz propia, la seteamos.
                        .light_value(0) = .light_base_value(0)
                    Else
                        .light_value(0) = color
                    End If
                    If .base_light(1) Then
                        .light_value(1) = .light_base_value(1)
                    Else
                        .light_value(1) = color
                    End If
                    If .base_light(2) Then
                        .light_value(2) = .light_base_value(2)
                    Else
                        .light_value(2) = color
                    End If
                    If .base_light(3) Then
                        .light_value(3) = .light_base_value(3)
                    Else
                        .light_value(3) = color
                    End If
                End With
            End If
        Next y
    Next x
    
End Function
Private Function Map_LightsClear()
    Dim x As Integer
    Dim y As Integer
    
    Dim AmbientColor As D3DCOLORVALUE
    Dim color As Long
    
    Meteo.Get_AmbientLight AmbientColor
    color = D3DColorXRGB(AmbientColor.R, AmbientColor.G, AmbientColor.B)
    
    For x = map_current.map_x_min To map_current.map_x_max
        For y = map_current.map_y_min To map_current.map_y_max
            If Map_In_Bounds(x, y) Then
                With map_current.map_grid(x, y)
                    If .base_light(0) Then 'Si tiene luz propia, la seteamos.
                        .light_value(0) = .light_base_value(0)
                    Else
                        .light_value(0) = color
                    End If
                    If .base_light(1) Then
                        .light_value(1) = .light_base_value(1)
                    Else
                        .light_value(1) = color
                    End If
                    If .base_light(2) Then
                        .light_value(2) = .light_base_value(2)
                    Else
                        .light_value(2) = color
                    End If
                    If .base_light(3) Then
                        .light_value(3) = .light_base_value(3)
                    Else
                        .light_value(3) = color
                    End If
                End With
            End If
        Next y
    Next x
    
End Function
Private Function LightCalculate(ByVal cRadio As Integer, ByVal LightX As Integer, ByVal LightY As Integer, ByVal XCoord As Integer, ByVal YCoord As Integer, TileLight As Long, LightColor As D3DCOLORVALUE, AmbientColor As D3DCOLORVALUE) As Long
    Dim XDist As Single
    Dim YDist As Single
    Dim VertexDist As Single
    Dim pRadio As Integer
    
    Dim CurrentColor As D3DCOLORVALUE
    
    pRadio = cRadio * 32
    
    XDist = LightX + 16 - XCoord
    YDist = LightY + 16 - YCoord
    
    VertexDist = Sqr(XDist * XDist + YDist * YDist)
    
    If VertexDist <= pRadio Then
        Call D3DXColorLerp(CurrentColor, LightColor, AmbientColor, VertexDist / pRadio)
        LightCalculate = D3DColorXRGB(CurrentColor.R, CurrentColor.G, CurrentColor.B)
        If TileLight > LightCalculate Then LightCalculate = TileLight
    Else
        LightCalculate = TileLight
    End If
End Function

Private Function Map_Info_Load(ByVal map_path As String)
    With map_current.Map_Info
        .Ambient = Val(General_Var_Get(map_path, "INIT", "Ambient"))
        .MapAmbientColor = D3DColorValueGet(Val(General_Var_Get(map_path, "INIT", "AmbientA")), Val(General_Var_Get(map_path, "INIT", "AmbientR")), Val(General_Var_Get(map_path, "INIT", "AmbientG")), Val(General_Var_Get(map_path, "INIT", "AmbientB")))
        .Llueve = Val(General_Var_Get(map_path, "INIT", "Lluvia"))
    End With
End Function

Private Function WaterFound(ByVal map_x As Byte, ByVal map_y As Byte) As Boolean
    WaterFound = (map_current.map_grid(map_x, map_y).grh(1).grh_index >= 1505 And map_current.map_grid(map_x, map_y).grh(1).grh_index <= 1520 Or _
        map_current.map_grid(map_x, map_y).grh(1).grh_index >= 5665 And map_current.map_grid(map_x, map_y).grh(1).grh_index <= 5680 Or _
        map_current.map_grid(map_x, map_y).grh(1).grh_index >= 13547 And map_current.map_grid(map_x, map_y).grh(1).grh_index <= 13562) _
        And map_current.map_grid(map_x, map_y).grh(2).grh_index = 0
End Function

Private Sub DoPasosFx(ByVal CharIndex As Integer)
    With char_list(CharIndex)
        If Not UserNavegando Then
            If .priv = 0 Then
                .pie = Not .pie
                If .pie Then
                    Call modSound.Sound_Play3D(SND_PASOS1, .map_x, .map_y)
                Else
                    Call modSound.Sound_Play3D(SND_PASOS2, .map_x, .map_y)
                End If
            End If
        Else
            'TODO : Actually we would have to check if the CharIndex char is in the water or not....
            Call modSound.Sound_Play3D(SND_NAVEGANDO, .map_x, .map_y)
        End If
    End With
End Sub

Public Sub Char_ParticleGroup_Destroy(ByVal char_index As Integer)
    If char_index > 0 Then
        char_list(char_index).chr_data.FxData.Particle_Index = 0
    End If
End Sub

Public Sub getBodyData(ByVal iBodyIndex As Integer, ByRef iGrhIndex As Integer, ByRef iOffsetY As Integer)
    If Not (iBodyIndex > 0 And iBodyIndex <= UBound(Char_Data_List.BodyData)) Then Exit Sub
    iGrhIndex = Char_Data_List.BodyData(iBodyIndex).Body(E_Heading.SOUTH).grh_index
    iOffsetY = Char_Data_List.BodyData(iBodyIndex).HeadOffset.y
End Sub

Public Sub getShieldData(ByVal iShieldIndex As Integer, iGrhIndex As Integer)
    If Not (iShieldIndex > 0 And iShieldIndex <= UBound(Char_Data_List.ShieldData)) Then Exit Sub
    iGrhIndex = Char_Data_List.ShieldData(iShieldIndex).ShieldWalk(E_Heading.SOUTH).grh_index
End Sub

Public Sub getWeaponData(ByVal iWeaponIndex As Integer, iGrhIndex As Integer)
    If Not (iWeaponIndex > 0 And iWeaponIndex <= UBound(Char_Data_List.WeaponData)) Then Exit Sub
    iGrhIndex = Char_Data_List.WeaponData(iWeaponIndex).WeaponWalk(E_Heading.SOUTH).grh_index
End Sub

Public Sub getHeadData(ByVal iHeadIndex As Integer, iGrhIndex As Integer)
    If Not (iHeadIndex > 0 And iHeadIndex <= UBound(Char_Data_List.HeadData)) Then Exit Sub
    iGrhIndex = Char_Data_List.HeadData(iHeadIndex).Head(E_Heading.SOUTH).grh_index
End Sub

Public Sub getHelmData(ByVal iHelmIndex As Integer, iGrhIndex As Integer)
    If Not (iHelmIndex > 0 And iHelmIndex <= UBound(Char_Data_List.CascoData)) Then Exit Sub
    iGrhIndex = Char_Data_List.CascoData(iHelmIndex).Head(E_Heading.SOUTH).grh_index
End Sub

Private Function charGetEffectArrayPos(ByVal CharIndex As Integer) As Integer
    Dim i As Integer
    
    Do While i < char_list(CharIndex).chr_data.lastEffect And Not charGetEffectArrayPos = 0
        i = i + 1
        If char_list(CharIndex).chr_data.effectList(i) < 1 Then _
            charGetEffectArrayPos = i
    Loop
End Function

Public Sub charCreateEffect(ByVal CharIndex As Integer, ByVal effectID As Integer, Optional ByVal lifeTime As Integer = 0)
    Dim effectIndex As Integer
    
    effectIndex = charGetEffectArrayPos(CharIndex)
    
    If effectIndex > 0 Then
        char_list(CharIndex).chr_data.effectList(effectIndex) = createEffect(effectID, lifeTime)
    Else
        If char_list(CharIndex).chr_data.lastEffect < MAXEFFECTS Then
            char_list(CharIndex).chr_data.lastEffect = char_list(CharIndex).chr_data.lastEffect + 1
            char_list(CharIndex).chr_data.effectList(char_list(CharIndex).chr_data.lastEffect) = createEffect(effectID, lifeTime)
        End If
    End If
End Sub

Public Sub charDestroyEffect(ByVal CharIndex As Integer, ByVal effectIndex As Integer)
    char_list(CharIndex).chr_data.effectList(effectIndex) = 0
    If effectIndex = char_list(CharIndex).chr_data.lastEffect Then
        Do While char_list(CharIndex).chr_data.effectList(char_list(CharIndex).chr_data.lastEffect) < 1
            char_list(CharIndex).chr_data.lastEffect = char_list(CharIndex).chr_data.lastEffect - 1
            If char_list(CharIndex).chr_data.lastEffect < 1 Then Exit Do
        Loop
    End If
End Sub

Public Sub mapDestroyEffect(ByVal mapX As Integer, ByVal mapY As Integer, ByVal effectIndex As Integer)
    If mapCheckEffect(mapX, mapY, effectIndex) Then
        map_current.map_grid(mapX, mapY).effectList(effectIndex) = 0
        If effectIndex = map_current.map_grid(mapX, mapY).lastEffect Then
            Do While map_current.map_grid(mapX, mapY).effectList(map_current.map_grid(mapX, mapY).lastEffect) < 1
                map_current.map_grid(mapX, mapY).lastEffect = map_current.map_grid(mapX, mapY).lastEffect - 1
                If map_current.map_grid(mapX, mapY).lastEffect < 1 Then Exit Do
            Loop
        End If
    End If
End Sub

Private Function mapCheckEffect(ByVal mapX As Integer, ByVal mapY As Integer, ByVal effectIndex As Integer) As Boolean
    If effectIndex > 0 And effectIndex <= map_current.map_grid(mapX, mapY).lastEffect Then
        mapCheckEffect = True
    End If
End Function


Public Sub mapCreateEffect(ByVal mapX As Byte, ByVal mapY As Byte, ByVal effectID As Integer, Optional ByVal lifeTime As Integer = 0)
    Dim effectIndex As Integer
    effectIndex = mapGetEffectArrayPos(mapX, mapY)
    
    If effectIndex > 0 Then
        If effectIndex > map_current.map_grid(mapX, mapY).lastEffect Then _
            map_current.map_grid(mapX, mapY).lastEffect = effectIndex
        
        map_current.map_grid(mapX, mapY).effectList(effectIndex) = createEffect(effectID, lifeTime)
    Else
        If map_current.map_grid(mapX, mapY).lastEffect < MAXEFFECTS Then
            map_current.map_grid(mapX, mapY).lastEffect = map_current.map_grid(mapX, mapY).lastEffect + 1
            map_current.map_grid(mapX, mapY).effectList(map_current.map_grid(mapX, mapY).lastEffect) = createEffect(effectID, lifeTime)
        End If
    End If
    
End Sub

Private Function mapGetEffectArrayPos(ByVal mapX As Byte, ByVal mapY As Byte) As Integer
    Dim i As Byte
    
    With map_current.map_grid(mapX, mapY)
        Do While i < .lastEffect And mapGetEffectArrayPos < 1
            i = i + 1
            If Not .effectList(i) > 0 Then
                mapGetEffectArrayPos = i
            End If
        Loop
    End With
End Function

Public Function mapDestroyAllEffects()
    Dim x As Byte
    Dim y As Byte
    Dim i As Integer
    
    For x = map_current.map_x_min To map_current.map_x_max
        For y = map_current.map_y_min To map_current.map_y_max
            For i = 1 To map_current.map_grid(x, y).lastEffect
                If map_current.map_grid(x, y).effectList(i) > 0 Then _
                    destroyEffect map_current.map_grid(x, y).effectList(i)
                    mapDestroyEffect x, y, i
            Next i
        Next y
    Next x

End Function

Private Function mapBonfireCreate(ByVal mapX As Byte, ByVal mapY As Byte) As Integer
    bonfireCount = bonfireCount + 1
    ReDim Preserve bonfireList(1 To bonfireCount) As tBonfire
    bonfireList(bonfireCount).mapX = mapX
    bonfireList(bonfireCount).mapY = mapY
    
    bonfireList(bonfireCount).lightIndex = Map_LightCreate(mapX, mapY, 255, 255, 255, 4)
    bonfireList(bonfireCount).effectIndex = createEffect(2, -1)
    
    mapBonfireCreate = bonfireCount
    map_current.map_grid(mapX, mapY).bonfireIndex = bonfireCount
End Function

Private Sub mapRenderBonfire(bonfireIndex As Integer, ByVal screenX As Integer, ByVal screenY As Integer)
    If bonfireIndex <= bonfireCount Then _
        Call renderEffect(bonfireList(bonfireIndex).effectIndex, screenX, screenY, map_current.map_grid(bonfireList(bonfireIndex).mapX, bonfireList(bonfireIndex).mapY).light_value)
End Sub

Private Sub mapBonfireDestroy(bonfireIndex As Integer)
    Dim i As Integer
    
    map_current.map_grid(bonfireList(bonfireIndex).mapX, bonfireList(bonfireIndex).mapY).bonfireIndex = 0
    
    If bonfireIndex < bonfireCount Then
        For i = bonfireIndex To bonfireCount - 1
            bonfireList(i) = bonfireList(i + 1)
        Next i
    End If
    
    bonfireCount = bonfireCount - 1
    
    If bonfireCount > 0 Then _
        ReDim Preserve bonfireList(1 To bonfireCount)
End Sub

Private Function charGetProjectileArrayPos(ByVal CharIndex As Integer) As Integer
    Dim i As Byte
    Do While i < char_list(CharIndex).chr_data.lastProjectile And charGetProjectileArrayPos < 1
        i = i + 1
        If char_list(CharIndex).chr_data.projectileIndex(i) = 0 Then _
            charGetProjectileArrayPos = i
    Loop
End Function

Public Sub charCreateProjectile(CharIndex As Integer, ByVal origX As Integer, ByVal origY As Integer, projectileID As Integer, Optional ByVal onHitEffect As Integer = 0, Optional ByVal onHitTarget As Byte = 0)
    Dim index As Integer
    
    index = charGetProjectileArrayPos(CharIndex)
    
    If index > 0 Then
        char_list(CharIndex).chr_data.projectileIndex(index) = createProjectile(projectileID, origX, origY, char_list(CharIndex).map_x, char_list(CharIndex).map_y, index, onHitEffect, onHitTarget, CharIndex)
    Else
        If char_list(CharIndex).chr_data.lastProjectile < MAXPROJECTILES Then
        
            char_list(CharIndex).chr_data.lastProjectile = char_list(CharIndex).chr_data.lastProjectile + 1
            index = char_list(CharIndex).chr_data.lastProjectile
            
            char_list(CharIndex).chr_data.projectileIndex(index) = createProjectile(projectileID, origX, origY, char_list(CharIndex).map_x, char_list(CharIndex).map_y, index, onHitEffect, onHitTarget, CharIndex)
        End If
    End If
End Sub

Public Sub charDestroyProjectile(CharIndex As Integer, projectileIndex As Integer)
    Dim i As Byte
    
    If projectileIndex < 1 Or projectileIndex > char_list(CharIndex).chr_data.lastProjectile Then Exit Sub
    
    char_list(CharIndex).chr_data.projectileIndex(projectileIndex) = 0
    
    If projectileIndex = char_list(CharIndex).chr_data.lastProjectile Then
        Do While char_list(CharIndex).chr_data.projectileIndex(char_list(CharIndex).chr_data.lastProjectile) < 1
            char_list(CharIndex).chr_data.lastProjectile = char_list(CharIndex).chr_data.lastProjectile - 1
            If char_list(CharIndex).chr_data.lastProjectile < 1 Then Exit Do
        Loop
    End If

End Sub

Private Function mapProjectileGetArrayPos(ByVal mapX As Byte, ByVal mapY As Byte)
    Dim i As Byte
    
    With map_current.map_grid(mapX, mapY)
        Do While i < .lastProjectile And mapProjectileGetArrayPos < 1
            i = i + 1
            If .projectileList(i) = 0 Then _
                mapProjectileGetArrayPos = i
        Loop
    End With
    
End Function

Public Sub mapCreateProjectile(ByVal mapX As Integer, ByVal mapY As Integer, ByVal origX As Integer, ByVal origY As Integer, projectileID As Integer, Optional ByVal onHitEffect As Integer = 0, Optional ByVal onHitTarget As Byte = 0)
    With map_current.map_grid(mapX, mapY)
        Dim index As Integer
        
        index = mapProjectileGetArrayPos(mapY, mapY)
        
        If index > 0 Then
            .projectileList(index) = createProjectile(projectileID, origX, origY, mapX, mapY, index, onHitEffect, onHitTarget, 0)
        Else
            If .lastProjectile < MAXPROJECTILES Then
                .lastProjectile = .lastProjectile + 1
                index = .lastProjectile
                .projectileList(index) = createProjectile(projectileID, origX, origY, mapX, mapY, index, onHitEffect, onHitTarget, 0)
            End If
        End If
    End With
End Sub

Public Sub mapDestroyProjectile(mapX As Integer, mapY As Integer, projectileIndex As Integer)
    With map_current.map_grid(mapX, mapY)
        
        If projectileIndex < 1 Or projectileIndex > .lastProjectile Then Exit Sub
        
        .projectileList(projectileIndex) = 0
        
        If projectileIndex = .lastProjectile Then
            Do While .projectileList(.lastProjectile) < 1
                .lastProjectile = .lastProjectile - 1
                If .lastProjectile < 1 Then Exit Do
            Loop
        End If
    End With
End Sub


