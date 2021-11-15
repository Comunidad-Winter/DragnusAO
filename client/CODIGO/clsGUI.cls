VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class was made by Agustin Andreucci (Blizzard)
'for the specific use of Argentum Online comunity using Ore Engine.

Private Enum e_button_states
    Reposo = 1
    Mouse_Over = 2
    Mouse_Down = 3
End Enum

Private Type gui_button
    X As Integer
    Y As Integer
    width As Integer
    height As Integer
    filename(2) As Integer 'for the 3 states
    state As e_button_states
    caption As String
    visible As Byte
End Type

Private Type gui_label
    X As Integer
    Y As Integer
    width As Integer
    height As Integer
    caption As String
    visible As Byte
End Type

Private Type gui_textbox
    X As Integer
    Y As Integer
    width As Integer
    height As Integer
    Text As String
    locked As Byte
    Active As Byte
    sel_text_pos As Integer
    max_chars As Integer
    color As Long
    visible As Byte
End Type

Private Type gui_checkbox
    X As Integer
    Y As Integer
    width As Integer
    height As Integer
    value As Byte
    caption As String
    filenum(1) As Integer
    visible As Byte
End Type

Private Type gui_scrollbar
    X As Integer
    Y As Integer
    width As Integer
    height As Integer
    pos As Integer
    max_value As Integer
    sensibility As Integer
    visible As Byte
End Type

Private Type gui_image
    image As Integer
    X As Integer
    Y As Integer
    width As Integer
    height As Integer
    alpha_blend As Boolean
    visible As Byte
End Type

Private Type gui_inventory
    inventory_id As Byte 'User or Npc Inventory
    X As Integer
    Y As Integer
    width As Integer
    height As Integer
    visible As Byte
End Type

Private Type gui_listbox
    X As Integer
    Y As Integer
    height As Integer
    width As Integer
    max_elements As Integer
    list_index As Integer
    btn_width As Integer
    full_height As Integer
    full As Boolean
    Font_Index As Integer
    List() As String
    elem_count As Integer
    scrollbar As gui_scrollbar
    visible As Byte
End Type

Private Type Recttext_line
    Text As String
    color As Long
    Font_Index As Byte
End Type

Private Type gui_rectxt
    X As Integer
    Y As Integer
    lines() As Recttext_line
    max_lines As Integer
    visible As Byte
    scrollbar As gui_scrollbar
    line_count As Integer
    width As Integer
    height As Integer
End Type

Private Type gui_combobox
    X As Integer
    Y As Integer
    width As Integer
    height As Integer
    list_index As Integer
    element_count As Integer
    max_elements As Integer
    List() As String
    scrollbar As gui_scrollbar
    visible As Byte
End Type

'Representa un formulario con sus controles.
Private Type menu
    id As e_Menus
    'Menu properties
    X As Integer
    Y As Integer
    width As Integer
    height As Integer
    background_file As String
    caption As String
    
    num_buttons As Integer
    num_textboxs As Integer
    num_scrollbars As Integer
    num_checkboxes As Integer
    num_labels As Integer
    num_images As Integer
    num_inventorys As Integer
    num_listboxs As Integer
    num_recttxts As Integer
    num_comboboxes As Integer
    active_textbox As Integer
    
    'Menu controls
    Button() As gui_button
    textbox() As gui_textbox
    scrollbar() As gui_scrollbar
    checkbox() As gui_checkbox
    label() As gui_label
    image() As gui_image
    inventory() As gui_inventory
    listbox() As gui_listbox
    recttxt() As gui_rectxt
    combobox() As gui_combobox
End Type

'Mouse coordinates, used to render cursor.
Dim mouse_x As Integer
Dim mouse_y As Integer

Dim Z_last_menu As Integer
'Como Maximo mostramos 6 menues... desp lo hago dinamico...
Private Const max_menus As Integer = 6
Dim menu_array() As menu

'END Declarations
Public Function GUI_Window_Change(ByVal window_id As e_Menus)
    
    If Z_last_menu > 0 Then
        'Si hay una ventana en el array nos fijamos que no sea la misma.
        If menu_array(1).id = window_id Then Exit Function
    Else
        'Si no hay ventana...
        'Redimensionamos el array.
        Z_last_menu = 1
        ReDim menu_array(1 To Z_last_menu) As menu
    End If
    
    'Reemplazamos la 1ra ventana
    
    'Borrar los controles por las dudas de que haya quedado algo.
    Menu_Controls_Remove 1
    With menu_array(1)
        .id = window_id
        .X = 0
        .Y = 0
    'RECORDAR AGREGAR ALTO Y ANCHO DE LA VENTANA!
        Select Case window_id
            Case e_Menus.Main
                .width = 800
                .height = 600
                .background_file = 12072
                'Recttxt
                Call Recttxt_Create(1, 14, 9, 500, 94, 10) 'Consola
                'Comboboxes
                Call Combobox_Create(1, 50, 50, 50, 50, 34) 'Lista hechizos
                'Buttons
                Call Button_Create(1, 665, 326, 10, 10) 'Mover hechi
                Call Button_Create(1, 665, 340, 10, 10) 'Mover hechi
                Call Button_Create(1, 790, 0, 10, 10) 'Cerrar
                Call Button_Create(1, 587, 327, 54, 17) 'Lanzar
                Call Button_Create(1, 680, 327, 54, 17) 'Info
                Call Button_Create(1, 677, 142, 84, 18) 'Hechizos
                Call Button_Create(1, 502, 142, 84, 18) 'Inventario
                Call Button_Create(1, 555, 389, 68, 23) 'Estadisticas
                Call Button_Create(1, 680, 389, 80, 23) 'Opciones
                Call Button_Create(1, 680, 389, 70, 23) 'Clanes
                'Textboxes
                Call TextBox_Create(1, 10, 123, 535, 18, "", 0, 33, RGB(255, 255, 255), 0) 'Chat
                'Labels
                
                'Inventario
                Call Inventory_Create(1, 582, 170, 177, 183, 1)
            Case e_Menus.Conectar
                .width = 800
                .height = 600
                .background_file = 500
                Call Button_Create(1, 309, 354, 185, 46)
                Call Button_Create(1, 309, 393, 185, 46)
                Call Button_Create(1, 309, 432, 185, 46)
                Call TextBox_Create(1, 325, 253, 100, 30, "hola", 0, 30, RGB(255, 255, 255))
                Call TextBox_Create(1, 325, 302, 100, 30, "hola", 1, 30, RGB(255, 255, 255))
            Case e_Menus.CrearPj
            Case e_Menus.Cargando
        End Select
    End With
End Function
Public Function GUI_Show_Menu(menu_id As e_Menus, ByVal X As Integer, ByVal Y As Integer, Optional ByVal index As Integer = 0)
    If Z_last_menu >= max_menus Then Exit Function
    If Z_last_menu = 0 Then Exit Function 'No podemos mostrar ningun menu si no hay ninguna ventana abierta (Index=1)
    
    If index = 0 Then
        'Incrementamos la cantidad de menues
        Z_last_menu = Z_last_menu + 1
        'Ahora la cantidad de ventanas es dinamica.
        ReDim Preserve menu_array(1 To Z_last_menu) As menu
        index = Z_last_menu
    Else
        Menu_Controls_Remove (index)
    End If
    
    With menu_array(index)
        .X = X
        .Y = Y

        .id = menu_id
        
        'Que menu debemos crear?
        'RECORDAR AGREGAR ALTO Y ANCHO DE LA VENTANA!
        Select Case menu_id
            Case e_Menus.Estadisticas
                'ATRIBUTO ...Ahi ta bien tabulado :P (MaxTus)...
                Call Label_Create(index, 67, 14, 59, 17, "ATRIBUTOS")
                Call Label_Create(index, 20, 34, 26, 13, "Pablo")
                Call Label_Create(index, 20, 49, 26, 13, "Pablo")
                Call Label_Create(index, 20, 63, 26, 13, "Pablo")
                Call Label_Create(index, 20, 77, 26, 13, "Pablo")
                Call Label_Create(index, 20, 91, 26, 13, "Pablo")
                'REPUTACION
                Call Label_Create(index, 63, 131, 70, 14, "REPUTACION")
                Call Label_Create(index, 19, 149, 60, 13, "Reputacion1")
                Call Label_Create(index, 19, 165, 60, 13, "Reputacion1")
                Call Label_Create(index, 19, 180, 60, 13, "Reputacion1")
                Call Label_Create(index, 19, 196, 60, 13, "Reputacion1")
                Call Label_Create(index, 19, 212, 60, 13, "Reputacion1")
                Call Label_Create(index, 19, 228, 60, 13, "Reputacion1")
                Call Label_Create(index, 19, 244, 60, 13, "Reputacion1")
                'ESTADISTICAS
                Call Label_Create(index, 63, 272, 72, 14, "ESTADISTICAS")
                Call Label_Create(index, 16, 292, 165, 13, "Estadisticas1")
                Call Label_Create(index, 16, 308, 165, 13, "Estadisticas1")
                Call Label_Create(index, 16, 324, 165, 13, "Estadisticas1")
                Call Label_Create(index, 16, 340, 165, 13, "Estadisticas1")
                Call Label_Create(index, 16, 356, 165, 13, "Estadisticas1")
                Call Label_Create(index, 16, 372, 372, 13, "Estadisticas1")
                'SKILLS
                Call Label_Create(index, 310, 17, 31, 16, "SKILLS")
                Call Label_Create(index, 240, 48, 32, 13, "Pablo2")
                Call Label_Create(index, 240, 63, 32, 13, "Pablo2")
                Call Label_Create(index, 240, 78, 32, 13, "Pablo2")
                Call Label_Create(index, 240, 94, 32, 13, "Pablo2")
                Call Label_Create(index, 240, 110, 32, 13, "Pablo2")
                Call Label_Create(index, 240, 125, 32, 13, "Pablo2")
                Call Label_Create(index, 240, 141, 32, 13, "Pablo2")
                Call Label_Create(index, 240, 157, 32, 13, "Pablo2")
                Call Label_Create(index, 240, 172, 32, 13, "Pablo2")
                Call Label_Create(index, 240, 188, 32, 13, "Pablo2")
                Call Label_Create(index, 240, 204, 32, 13, "Pablo2")
                Call Label_Create(index, 240, 219, 32, 13, "Pablo2")
                Call Label_Create(index, 240, 235, 32, 13, "Pablo2")
                Call Label_Create(index, 240, 251, 32, 13, "Pablo2")
                Call Label_Create(index, 240, 266, 32, 13, "Pablo2")
                Call Label_Create(index, 240, 282, 32, 13, "Pablo2")
                Call Label_Create(index, 240, 298, 32, 13, "Pablo2")
                Call Label_Create(index, 240, 313, 32, 13, "Pablo2")
                Call Label_Create(index, 240, 329, 32, 13, "Pablo2")
                Call Label_Create(index, 240, 345, 32, 13, "Pablo2")
                Call Label_Create(index, 240, 360, 32, 13, "Pablo2")
                'Call Label_Create(index, 240, 375, 32, 13, "Pablo2") Skill Alquimia
                'Call Label_Create(index, 240, 390, 32, 13, "Pablo2") Skill Cortar Raices
                'Call Label_Create(index, 240, 405, 32, 13, "Pablo2") Skill Sastrería
                'BOTÓN
                Call Button_Create(index, 4, 392, 454, 25, , , "OK")
            Case e_Menus.BancoObj
                 '(Faltan los ListBox) MaxTus
                 '(Faltan el PictureBox) MaxTus
                Call Label_Create(index, 151, 450, 43, 13, "Cantidad:")
                Call TextBox_Create(index, 207, 446, 72, 19, 1, False, 5, &H80000005)
                Call Button_Create(index, 413, 452, 31, 16, , , "Aceptar")
            Case e_Menus.Tirar
                Call Label_Create(index, 39, 11, 161, 25, "Escriba la cantidad:")
                Call TextBox_Create(index, 22, 35, 175, 25, , False, 5, &H80000005)
                Call Button_Create(index, 22, 69, 112, 23, , , "Aceptar")
                Call Button_Create(index, 135, 69, 62, 23, , , "Todo")
            Case e_Menus.Carpintero
                '(Faltan el ListBox) MaxTus
                Call Button_Create(index, 17, 178, 114, 29, , , "Salir")
                Call Button_Create(index, 176, 178, 114, 29, , , "Construir")
            Case e_Menus.CharInfo
                'GENERAL
                Call Label_Create(index, 135, 360, 5640, 255, "Nombre:")
                Call Label_Create(index, 120, 600, 2880, 255, "Raza:")
                Call Label_Create(index, 120, 840, 3270, 255, "Clase:")
                Call Label_Create(index, 120, 1080, 2895, 255, "Genero:")
                Call Label_Create(index, 120, 1320, 3105, 255, "Nivel:")
                Call Label_Create(index, 120, 1560, 2805, 255, "Oro:")
                Call Label_Create(index, 3030, 600, 2880, 255, "Clan Actual:")
                Call Label_Create(index, 3060, 844, 2880, 255, "Faccion:")
                Call Label_Create(index, 3060, 1080, 2850, 255, "Ciudadanos Asesinados:")
                Call Label_Create(index, 3060, 1325, 2900, 255, "Criminales Asesinados:")
                Call Label_Create(index, 3060, 1560, 2445, 255, "Reputacion:")
                Call Label_Create(index, 3060, 1800, 2760, 255, "Status:")
                 'Clanes
                Call Label_Create(index, 135, 270, 2985, 255, "Ultimas membresías solicitadas:")
                Call TextBox_Create(index, 135, 450, 5790, 1110, "", False, , &H80000005)
            Case e_Menus.Comerciar
            Case e_Menus.ComerciarUsu
            Case e_Menus.Commet
            Case e_Menus.EligeAlineacion
            Case e_Menus.Entrenador
            Case e_Menus.Foro
            Case e_Menus.GuildAdm
            Case e_Menus.GuildBrief
            Case e_Menus.GuildDetails
            Case e_Menus.GuildFoundation
            Case e_Menus.GuildLeader
            Case e_Menus.GuildNews
            Case e_Menus.GuildSol
            Case e_Menus.GuildURL
            Case e_Menus.Herrero
            Case e_Menus.KeyPad
            Case e_Menus.Mensaje
            Case e_Menus.SOS
            Case e_Menus.Opciones
            Case e_Menus.PanelGM
            Case e_Menus.Passwd
            Case e_Menus.PeaceProp
            Case e_Menus.Premios
            Case e_Menus.Pres
            Case e_Menus.Retos
            Case e_Menus.ScreenShots
            Case e_Menus.Skills
            Case e_Menus.SpawnList
            Case e_Menus.Subasta
            Case e_Menus.tip
            Case e_Menus.Torneo
            Case e_Menus.UserRequest
        End Select
    End With
    
End Function
Public Sub Menu_Hide(menu_id As e_Menus)
    If Z_last_menu = 0 Then Exit Sub

    Dim menu_index As Integer
    Dim i As Integer
    menu_index = Menu_Obtain_Index(menu_id)
    
    If menu_index = 0 Then Exit Sub
    
    If menu_index = Z_last_menu Then
        Z_last_menu = Z_last_menu - 1
    Else
        'Si no es el ultimo...
        For i = menu_index To Z_last_menu - 1
            menu_array(i) = menu_array(i + 1)
        Next i
        Z_last_menu = Z_last_menu - 1
    End If
    
    'Redimensionamos el array de ventanas.
    ReDim Preserve menu_array(1 To Z_last_menu) As menu
    
End Sub
Private Function Menu_Obtain_Index(ByVal menu_id As e_Menus) As Integer
    Dim i As Integer
    i = 1
    Do While i <= Z_last_menu
        If menu_array(i).id = menu_id Then
            Menu_Obtain_Index = i
            Exit Function
        Else
            i = i + 1
        End If
    Loop
End Function

Private Sub Menu_Controls_Remove(ByVal index As Integer)
    If index = 0 Then Exit Sub
    
    'Set num controls to -1
    menu_array(index).num_buttons = -1
    menu_array(index).num_checkboxes = -1
    menu_array(index).num_labels = -1
    menu_array(index).num_scrollbars = -1
    menu_array(index).num_textboxs = -1
    menu_array(index).num_images = -1
    menu_array(index).num_inventorys = -1
    menu_array(index).num_listboxs = -1
    menu_array(index).num_recttxts = -1
    menu_array(index).num_comboboxes = -1
    'Descargamos los arrays
    ReDim menu_array(index).Button(0) As gui_button
    ReDim menu_array(index).checkbox(0) As gui_checkbox
    ReDim menu_array(index).scrollbar(0) As gui_scrollbar
    ReDim menu_array(index).label(0) As gui_label
    ReDim menu_array(index).textbox(0) As gui_textbox
    ReDim menu_array(index).image(0) As gui_image
    ReDim menu_array(index).inventory(0) As gui_inventory
    ReDim menu_array(index).listbox(0) As gui_listbox
    ReDim menu_array(index).recttxt(0) As gui_rectxt
End Sub

Private Function Image_Create(ByVal index As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal image As Integer, Optional ByVal alpha_blend As Boolean = False)
    Dim image_index As Integer
    
    With menu_array(index)
        .num_images = .num_images + 1
        ReDim Preserve .image(0 To .num_images) As gui_image
        image_index = .num_images
        
        .image(image_index).X = X
        .image(image_index).Y = Y
        .image(image_index).width = width
        .image(image_index).height = height
        .image(image_index).image = image
        .image(image_index).alpha_blend = alpha_blend
    End With
End Function
Private Function Listbox_Create(ByVal index As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal max_elements As Integer, ByVal btn_width As Integer, Optional ByVal visible As Byte = 1)
    With menu_array(index)
        .num_listboxs = .num_listboxs + 1
        ReDim Preserve .listbox(0 To .num_listboxs) As gui_listbox
        
        .listbox(.num_listboxs).X = X
        .listbox(.num_listboxs).Y = Y
        .listbox(.num_listboxs).width = width
        .listbox(.num_listboxs).height = height
        .listbox(.num_listboxs).btn_width = btn_width
        .listbox(.num_listboxs).max_elements = max_elements
        .listbox(.num_listboxs).list_index = -1
        .listbox(.num_listboxs).elem_count = -1
        .listbox(.num_listboxs).full = 0
        .listbox(.num_listboxs).visible = visible
    End With
End Function
Private Function Recttxt_Create(ByVal index As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal max_lines As Integer, Optional ByVal visible As Byte = 1)
    Dim recttxt_index As Integer
    menu_array(index).num_recttxts = menu_array(index).num_recttxts + 1
    ReDim Preserve menu_array(index).recttxt(0 To menu_array(index).num_recttxts)
    recttxt_index = menu_array(index).num_recttxts
    
    menu_array(index).recttxt(recttxt_index).X = X
    menu_array(index).recttxt(recttxt_index).Y = Y
    menu_array(index).recttxt(recttxt_index).width = width
    menu_array(index).recttxt(recttxt_index).height = height
    menu_array(index).recttxt(recttxt_index).visible = visible
    menu_array(index).recttxt(recttxt_index).max_lines = max_lines
    menu_array(index).recttxt(recttxt_index).line_count = -1
    menu_array(index).recttxt(recttxt_index).scrollbar.height = menu_array(index).recttxt(recttxt_index).height
    
End Function
Private Function Combobox_Create(ByVal index As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal max_elements As Integer, Optional ByVal visible As Byte = 1)
    Dim combobox_index As Integer
    menu_array(index).num_comboboxes = menu_array(index).num_comboboxes + 1
    ReDim Preserve menu_array(index).combobox(0 To menu_array(index).num_comboboxes)
    combobox_index = menu_array(index).num_comboboxes
    With menu_array(index).combobox(combobox_index)
        .X = X
        .Y = Y
        .width = width
        .height = height
        .max_elements = max_elements
        .element_count = -1
        .visible = visible
    End With
End Function
Private Function Inventory_Create(ByVal index As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal width As Integer, ByVal height As Integer, Optional ByVal id As Byte = 1, Optional ByVal visible As Byte = 1)
    With menu_array(index)
        .num_inventorys = .num_inventorys + 1
        ReDim Preserve .inventory(0 To .num_inventorys) As gui_inventory
        .inventory(.num_inventorys).X = X
        .inventory(.num_inventorys).Y = Y
        .inventory(.num_inventorys).width = width
        .inventory(.num_inventorys).height = height
        .inventory(.num_inventorys).visible = visible
        .inventory(.num_inventorys).inventory_id = id
    End With
End Function
Private Function Button_Create(ByVal index As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal width As Integer, ByVal height As Integer, Optional ByVal graphic_unpressed As Integer, Optional ByVal graphic_pressed As Integer, Optional ByVal caption As String, Optional ByVal visible As Byte = 1)
    Dim button_index As Integer
    
    menu_array(index).num_buttons = menu_array(index).num_buttons + 1
    ReDim Preserve menu_array(index).Button(0 To menu_array(index).num_buttons) As gui_button
    button_index = menu_array(index).num_buttons
        
    menu_array(index).Button(button_index).X = X
    menu_array(index).Button(button_index).Y = Y
    menu_array(index).Button(button_index).width = width
    menu_array(index).Button(button_index).height = height
    menu_array(index).Button(button_index).caption = caption
    menu_array(index).Button(button_index).visible = visible
    'menu_array(menu_id).button(menu_array(menu_id).num_buttons).filename() = filename()

End Function

Private Function Label_Create(ByVal index As Integer, ByVal X As Integer, _
                                ByVal Y As Integer, ByVal width As Integer, ByVal height As Integer, _
                                    ByVal caption As String, Optional ByVal visible As Byte = 1)
    Dim label_index As Integer
                                    
    menu_array(index).num_labels = menu_array(index).num_labels + 1
    ReDim Preserve menu_array(index).label(0 To menu_array(index).num_labels) As gui_label
    label_index = menu_array(index).num_labels
        
    'We store in the control the screen x and y
    menu_array(index).label(label_index).X = X
    menu_array(index).label(label_index).Y = Y
    menu_array(index).label(label_index).width = width
    menu_array(index).label(label_index).height = height
    menu_array(index).label(label_index).caption = caption
    menu_array(index).label(label_index).visible = visible

End Function

Private Function TextBox_Create(ByVal index As Integer, ByVal X As Integer, ByVal Y As Integer, _
                                    ByVal width As Integer, ByVal height As Integer, ByVal Text As String, ByVal locked As Byte, _
                                        ByVal max_chars As Integer, ByVal color As Long, Optional ByVal visible As Byte = 1)
    Dim textbox_index As Integer
    
    menu_array(index).num_textboxs = menu_array(index).num_textboxs + 1
    ReDim Preserve menu_array(index).textbox(0 To menu_array(index).num_textboxs) As gui_textbox
    textbox_index = menu_array(index).num_textboxs
    'El activo es el ultimo creado.
    menu_array(index).active_textbox = textbox_index
    
    'We store in the control the screen x and y
    menu_array(index).textbox(textbox_index).X = X
    menu_array(index).textbox(textbox_index).Y = Y
    menu_array(index).textbox(textbox_index).width = width
    menu_array(index).textbox(textbox_index).height = height
    menu_array(index).textbox(textbox_index).Text = caption
    menu_array(index).textbox(textbox_index).max_chars = max_chars
    menu_array(index).textbox(textbox_index).color = color
    menu_array(index).textbox(textbox_index).sel_text_pos = 0
    menu_array(index).textbox(textbox_index).visible = visible
    menu_array(index).textbox(textbox_index).locked = locked
    
End Function

Private Function ScrollBar_Create(ByVal index As Integer, ByVal X As Integer, ByVal Y As Integer, _
                                    ByVal width As Integer, ByVal height As Integer, ByVal max_value As Integer, _
                                        ByVal sensivity As Integer, Optional ByVal visible As Byte = 1)
    Dim scrollbar_index As Integer
        
    menu_array(index).num_scrollbars = menu_array(index).num_scrollbars + 1
    ReDim Preserve menu_array(index).scrollbar(0 To menu_array(index).num_scrollbars) As gui_scrollbar
    scrollbar_index = menu_array(index).num_scrollbars
        
    'We store in the control the screen x and y
    menu_array(index).scrollbar(scrollbar_index).X = X
    menu_array(index).scrollbar(scrollbar_index).Y = Y
    menu_array(index).scrollbar(scrollbar_index).width = width
    menu_array(index).scrollbar(scrollbar_index).height = height
    menu_array(index).scrollbar(scrollbar_index).max_value = max_value
    menu_array(index).scrollbar(scrollbar_index).pos = 0
    menu_array(index).scrollbar(scrollbar_index).sensibility = sensivity
    menu_array(index).scrollbar(scrollbar_index).visible = visible

End Function

Private Function CheckBox_Create(ByVal index As Integer, ByVal checkbox_index As Integer, ByVal X As Integer, _
                                ByVal Y As Integer, ByVal width As Integer, ByVal height As Integer, _
                                    ByVal caption As String, ByVal state As Byte, Optional ByVal visible As Byte = 1)
       
    Dim checkbox_index As Byte
       
    menu_array(index).num_checkboxes = menu_array(index).num_checkboxes + 1
    ReDim Preserve menu_array(index).checkbox(0 To menu_array(index).num_checkboxes) As gui_checkbox
    checkbox_index = menu_array(index).num_checkboxes
    
    'We store in the control the screen x and y
    menu_array(index).checkbox(checkbox_index).X = X
    menu_array(index).checkbox(checkbox_index).Y = Y
    menu_array(index).checkbox(checkbox_index).width = width
    menu_array(index).checkbox(checkbox_index).height = height
    menu_array(index).checkbox(checkbox_index).caption = caption
    menu_array(index).checkbox(checkbox_index).state = state
    menu_array(index).checkbox(checkbox_index).visible = visible
    
End Function

Private Function Menu_Render(ByVal index As Integer)

    Dim i As Integer
    
    With menu_array(index)
        Call GUI_Render_Image(.background_file, .X, .Y, .width, .height)
        'Render Menu buttons
        'For i = 1 To .num_buttons
        '    Engine.GUI_Render_Button .button(i).x, .button(i).y
        'Next i
        'Render Menu checkbox
        'For i = 1 To .num_checkboxes
        '    Engine.GUI_Render_Checkbox
        'Next i
        'Render Menu scrollbar
        If .num_scrollbars > -1 Then
            For i = 0 To .num_scrollbars
                If .scrollbar(i).visible Then
                    Call GUI_Render_Scrollbar(.X + .scrollbar(i).X, .Y + .scrollbar(i).Y, .scrollbar(i).height, .scrollbar(i).width, .scrollbar(i).max_value, .scrollbar(i).pos)
                End If
            Next i
        End If
        'Render Menu listboxes
        If .num_listboxs > -1 Then
            For i = 0 To .num_listboxs
                Call GUI_Render_Listbox(index, i)
            Next i
        End If
        'Render Menu textboxes
        If .num_textboxs > -1 Then
            For i = 0 To .num_textboxs
                If .textbox(i).visible Then
                    Call GUI_Render_Textbox(index, i)
                End If
            Next i
        End If
        'Render images
        If .num_images > -1 Then
            For i = 0 To .num_images
                If .image(i).visible Then
                    Call GUI_Render_Image(.image(i).image, .X + .image(i).X, .Y + .image(i).Y, .image(i).width, .image(i).height) ', .image(i).alpha_blend)
                End If
            Next i
        End If
        'Render Menu labels
        If .num_labels > -1 Then
            For i = 0 To .num_labels
                If .label(i).visible Then
                    Call GUI_Render_Label(index, i)
                End If
            Next i
        End If
        'Render Menu inventorys
        If .num_inventorys > -1 Then
            For i = 0 To .num_inventorys
                If .inventory(i).visible Then
                    Call Inventario(.inventory(i).inventory_id).Inventory_Render(.X + .inventory(i).X, .Y + .inventory(i).Y, .inventory(i).width, .inventory(i).height)
                End If
            Next i
        End If
        'Render Menu Recttxts
        If .num_recttxts > -1 Then
            For i = 0 To .num_recttxts
                If .recttxt(i).visible Then
                    Call GUI_Render_RecTxt(index, i)
                End If
            Next i
        End If
        
        'Render Menu Comboboxes
        If .num_comboboxes > -1 Then
            For i = 0 To .num_comboboxes
                If .combobox(i).visible Then
                    Call GUI_Render_Combobox(index, i)
                End If
            Next i
        End If
        Call DXEngine_GraphicTextRender(1, mouse_x & ";" & mouse_y, 216, 10, RGB(255, 255, 255))
        Call GUI_Render_Image(16001, mouse_x, mouse_y, 16, 16)
    End With
End Function

Public Function GUI_Render_All()
    Dim i As Integer
    
    i = Z_last_menu
    
    Do While i > 0
        Menu_Render (i)
        i = i - 1
    Loop
    
End Function

Public Function GUI_Click(ByVal X As Integer, ByVal Y As Integer, ByVal Button As Integer) As Byte
    Dim Menu_Clicked As Byte
    Dim tX As Long, tY As Long
    Menu_Clicked = Menu_Test(X, Y)
    
    If Not Menu_Clicked > 0 Then Exit Function
    
    
    'Si la ventana no es la de mas arriba ni es la principal (la 1), la ponemos arriba de todo.
    If Menu_Clicked <> Z_last_menu And Menu_Clicked <> 1 Then
        If Button = vbLeftButton Then
            Menu_Set_Focus Menu_Clicked
        End If
    Else
        If Menu_Clicked = 1 And menu_array(Menu_Clicked).id = e_Menus.Main And Engine.Input_Mouse_In_View(X, Y) Then
            Call Engine.Input_Mouse_Tile_Get(X, Y, tX, tY)
            If Button = vbLeftButton Then
                Call General_Screen_Left_Click(tX, tY, Input_Key_Get(vbKeyShift))
            ElseIf Button = vbRightButton Then
                Call General_Screen_Right_Click(tX, tY)
            End If
        Else
            If Button = vbLeftButton Then
                
                If Not GUI_Textbox_Test(X, Y, Menu_Clicked) = 0 Then
                    GUI_Click = 1
                Else
                    If GUI_Button_Test(X, Y, Menu_Clicked) Then
                        GUI_Click = 1
                    ElseIf GUI_Checkbox_Test(X, Y, Menu_Clicked) Then
                        GUI_Click = 1
                    ElseIf GUI_Listbox_Test(X, Y, Menu_Clicked) Then
                        GUI_Click = 1
                    ElseIf GUI_Scrollbar_Test(X, Y, Menu_Clicked) Then
                        GUI_Click = 1
                    ElseIf GUI_Combobox_Test(X, Y, Menu_Clicked) Then
                        GUI_Click = 1
                    ElseIf GUI_Recttxt_Test(Menu_Clicked, X, Y) Then
                        GUI_Click = 1
                    End If
                    menu_array(Menu_Clicked).active_textbox = -1
                End If
                
    
                    '    If Not GUI_Click_Checkbox_Test(x, y, Menu_Clicked) = 0 Then
                    '        If Not GUI_Click_scrollbar_Test(x, y, Menu_Clicked) Then
                    '            GUI_Click_label_Test x, y, Menu_Clicked
                    '        End If
                    '    End If
            End If
        End If
    End If

    
End Function
Private Function GUI_Listbox_Test(ByVal X As Integer, ByVal Y As Integer, ByVal menu_index As Integer) As Integer
    Dim i As Byte
    Dim add_height As Integer
    With menu_array(menu_index)
        If .num_listboxs <> -1 Then
            For i = 0 To .num_listboxs
                If .listbox(i).visible Then
                    If .listbox(i).full Then _
                        add_height = .listbox(i).full_height
                    If X > .X + .listbox(i).X And Y > .Y + .listbox(i).Y And X < .X + .listbox(i).X + .listbox(i).width And _
                        Y < .Y + .listbox(i).Y + .listbox(i).height + add_height Then
                        'Return listbox Y Coordinates
                        Call GUI_Listbox_Click(i, menu_index, Y, X)
                        GUI_Listbox_Test = 1
                        Exit Function
                    End If
                End If
            Next i
        End If
    End With
    
    GUI_Listbox_Test = 0
End Function
Private Function GUI_Combobox_Test(ByVal X As Integer, ByVal Y As Integer, ByVal menu_index As Integer) As Integer
    Dim i As Byte
    With menu_array(menu_index)
        If .num_comboboxes > -1 Then
            For i = 0 To .num_comboboxes
                If .combobox(i).visible Then
                    If X > .X + .combobox(i).X And X < .X + .combobox(i).X + .combobox(i).width And Y > .Y + .combobox(i).Y And Y < .Y + .combobox(i).Y + .combobox(i).height Then
                        Call GUI_Combobox_Click(i, menu_index, X, Y)
                        GUI_Combobox_Test = 1
                        Exit Function
                    End If
                End If
            Next i
        End If
    End With
    GUI_Combobox_Test = 0
End Function
Private Function GUI_Scrollbar_Test(ByVal X As Integer, ByVal Y As Integer, ByVal menu_id As Integer) As Integer
    Dim i As Byte
    With menu_array(menu_id)
        If .num_scrollbars > -1 Then
            For i = 0 To .num_scrollbars
                If .scrollbar(i).visible Then
                    If X > .X + .scrollbar(i).X And X < .X + .scrollbar(i).X + .scrollbar(i).width And _
                        Y > .Y + .scrollbar(i).Y And Y < .Y + .scrollbar(i).Y + .scrollbar(i).height Then
                        Call GUI_Scrollbar_Click(i, menu_id, Y)
                        GUI_Scrollbar_Test = 1
                        Exit Function
                    End If
                End If
            Next i
        End If
    End With
    GUI_Scrollbar_Test = 0
End Function
Private Function GUI_Button_Test(ByVal X As Integer, ByVal Y As Integer, ByVal menu_id As Integer) As Integer
    Dim i As Byte
    
    If menu_array(menu_id).num_buttons <> -1 Then
        For i = 0 To menu_array(menu_id).num_buttons
            If menu_array(menu_id).Button(i).visible Then
                If (X > menu_array(menu_id).Button(i).X + menu_array(menu_id).X) And _
                        (X < menu_array(menu_id).Button(i).X + menu_array(menu_id).X + menu_array(menu_id).Button(i).width) And _
                            (Y > menu_array(menu_id).Button(i).Y + menu_array(menu_id).Y) And _
                                (Y < menu_array(menu_id).Button(i).Y + menu_array(menu_id).Y + menu_array(menu_id).Button(i).height) Then
                    Call GUI_Button_Click(i, menu_id)
                    GUI_Button_Test = 1
                    Exit Function
                End If
            End If
        Next i
    End If
    
    GUI_Button_Test = 0
End Function

Private Function GUI_Checkbox_Test(ByVal X As Integer, ByVal Y As Integer, ByVal menu_index As Integer) As Integer
    Dim i As Byte
    With menu_array(menu_index)
        If .num_checkboxes = -1 Then Exit Function
        For i = 0 To .num_checkboxes
            If .checkbox(i).visible Then
                If X > .X + .checkbox(i).X And X < .X + .checkbox(i).X + .checkbox(i).width _
                    And Y > .Y + .checkbox(i).Y And Y < .checkbox(i).Y + .checkbox(i).width Then
                    GUI_Checkbox_Test = 1
                    Call GUI_Checkbox_Click(i, menu_index)
                    Exit Function
                End If
            End If
        Next i
    End With
    
    GUI_Checkbox_Test = 0
End Function
Private Sub GUI_Combobox_Click(ByVal combobox_index As Integer, ByVal menu_index As Byte, ByVal X As Integer, ByVal Y As Integer)
    If Not Combobox_Check(menu_index, combobox_index) Then Exit Sub
    With menu_array(menu_index)
        If X > .X + .combobox(combobox_index).X + .combobox(combobox_index).width - 16 Then
            Scrollbar_Move .combobox(combobox_index).scrollbar, Y
        Else
            .combobox(combobox_index).list_index = Fix((Y - .Y - .combobox(combobox_index).Y) / 16)
            If .combobox(combobox_index).list_index > .combobox(combobox_index).element_count Then
                .combobox(combobox_index).list_index = -1
            End If
        End If
    End With
End Sub
Private Sub GUI_Scrollbar_Click(ByVal scrollbar_index As Integer, ByVal menu_index As Byte, ByVal Y As Integer)
    If Not Scrollbar_Check(menu_index, scrollbar_index) Then Exit Sub
    
    Dim sY As Integer
    With menu_array(menu_index)
        sY = Y - .Y - .scrollbar(scrollbar_index).Y
        Call Scrollbar_Move(.scrollbar(scrollbar_index), sY)
    End With
End Sub
Private Sub GUI_Listbox_Click(ByVal listbox_index As Integer, ByVal menu_index As Byte, ByVal Y As Integer, ByVal X As Integer)
    If Not Listbox_Check(menu_index, listbox_index) Then Exit Sub

    Dim elem_clicked As Integer
    Dim ly As Integer
    Dim lX As Integer
    With menu_array(menu_index)
        If .listbox(listbox_index).full Then
            ly = Y - (.Y + .listbox(listbox_index).Y) - .listbox(listbox_index).height
            lX = X - .X - .listbox(listbox_index).X
            If lX > .listbox(listbox_index).width - .listbox(listbox_index).btn_width Then
                Call Scrollbar_Move(.listbox(listbox_index).scrollbar, ly)
            Else
                elem_clicked = Fix(ly / .listbox(listbox_index).height + .listbox(listbox_index).scrollbar.pos)
                If elem_clicked <= .listbox(listbox_index).elem_count And elem_clicked > -1 Then
                    .listbox(listbox_index).list_index = elem_clicked
                End If
                .listbox(listbox_index).full = Not .listbox(listbox_index).full
            End If
        Else
            'Recalculate full_height
            If .listbox(listbox_index).elem_count > 9 Then
                .listbox(listbox_index).full_height = .listbox(i).height * 10
            Else
                .listbox(listbox_index).full_height = .listbox(i).height * .listbox(listbox_index).elem_count
            End If
            .listbox(listbox_index).full = Not .listbox(listbox_index).full
        End If
    End With
End Sub

Private Function GUI_Checkbox_Click(ByVal chk_index As Integer, ByVal menu_index As Byte)
    If chk_index < 0 Then Exit Function
    If menu_array(menu_index).num_checkboxes < 0 Then Exit Function
    'Toggle checkbox value
    menu_array(menu_index).checkbox(chk_index).value = Not menu_array(menu_index).checkbox(chk_index).value
End Function
Private Function GUI_Textbox_Test(ByVal X As Integer, ByVal Y As Integer, ByVal menu_index As Byte) As Integer
    Dim i As Byte
    
    With menu_array(menu_index)
        If .num_textboxs < 0 Then Exit Function
        For i = 0 To .num_textboxs
            If X > .X + .textbox(i).X And X < .X + .textbox(i).X + .textbox(i).width _
                And Y > .Y + .textbox(i).Y And Y < .Y + .textbox(i).Y + .textbox(i).height Then
                Call GUI_Textbox_Click(i, menu_index)
                GUI_Textbox_Test = 1
                Exit Function
            End If
        Next i
    End With
    
    GUI_Textbox_Test = 0
End Function
Private Function GUI_Recttxt_Test(ByVal menu_index As Byte, ByVal X As Integer, ByVal Y As Integer) As Integer
    Dim i As Integer
    With menu_array(menu_index)
        If .num_recttxts < 0 Then Exit Function
        For i = 0 To .num_recttxts
            If X > .X + .recttxt(i).X And X < .X + .recttxt(i).X + .recttxt(i).width And Y > .Y + .recttxt(i).Y And Y < Y + .recttxt(i).Y + .recttxt(i).height Then
                Call GUI_Recttxt_Click(i, menu_index, X, Y)
                GUI_Recttxt_Test = 1
                Exit Function
            End If
        Next i
    End With
    GUI_Recttxt_Test = 0
End Function
Private Function GUI_Recttxt_Click(ByVal recttxt_index As Integer, ByVal menu_index As Byte, ByVal X As Integer, ByVal Y As Integer)
    If Not Recttxt_Check(menu_index, recttxt_index) Then Exit Function
    With menu_array(menu_index)
        If .recttxt(recttxt_index).scrollbar.visible Then
            If X > .X + .recttxt(recttxt_index).X + .recttxt(recttxt_index).width - 16 And X < .X + .recttxt(recttxt_index).X + .recttxt(recttxt_index).width Then
                Scrollbar_Move .recttxt(recttxt_index).scrollbar, Y - (.Y + .recttxt(recttxt_index).Y)
            End If
        End If
    End With
End Function
Private Function GUI_Textbox_Click(ByVal txt_index As Integer, ByVal menu_index As Byte)
    If Not Textbox_Check(menu_index, txt_index) Then Exit Function
    'Set the active textbox to the clicked textbox
    menu_array(menu_index).active_textbox = txt_index
    'Desp podria poner para correr el cursor..
End Function

Private Function GUI_Button_Click(ByVal btn_index As Integer, ByVal menu_index As Byte) As Byte
    If Not Button_Check(menu_index, btn_index) Then Exit Function
    
    Dim menu_id As Byte
    menu_id = menu_array(menu_index).id
    
    Select Case menu_id
        Case e_Menus.Cargando
            Select Case btn_index
            End Select
        Case e_Menus.Conectar
            Select Case btn_index
                Case 0 'Boton Conectar
                    Call General_Connect(menu_array(menu_index).textbox(0).Text, menu_array(menu_index).textbox(1).Text)
                Case 1 'Boton Crear Pj
                    Call General_Char_Create
                Case 2 'Boton Salir
                    prgRun = False
            End Select
        Case e_Menus.Main
            Select Case btn_index
                Case 0 Or 1 'Boton hechizos
                    If Combobox_Get_Index(e_Menus.Main, 0) = -1 Then Exit Function
                    Call Combobox_Move_Item(e_Menus.Main, 0, btn_index)
                    Call WriteMoveSpell(index, Combobox_Get_Index(e_menu.Main, 0) + 1)
                Case 2 'Boton Salir
                    prgRun = False
                Case 3 'Lanzar
                    If UCase(Combobox_Get_Element(menu_id, 0, Combobox_Get_Index(menu_id, 0))) <> UCase("(None)") And MainTimer.Check(TimersIndex.Work, False) Then
                        Call WriteCastSpell(Combobox_Get_Index(menu_id, 0) + 1)
                        Call WriteWork(eSkill.Magia)
                        UsaMacro = True
                    End If
                Case 4 'Info
                    Call WriteSpellInfo(Combobox_Get_Index(menu_id, 0) + 1)
                Case 5 'Hechizos
                Case 6 'Inventario
            End Select
        Case e_Menus.CrearPj
            Select Case btn_index
            End Select
        Case e_Menus.Estadisticas
            Select Case btn_index
                Case 0
                    Menu_Hide (menu_id)
            End Select
        Case e_Menus.BancoObj
            Select Case btn_index
            End Select
        Case e_Menus.Tirar
            Select Case btn_index
            End Select
        Case e_Menus.Carpintero
            Select Case btn_index
            End Select
        Case e_Menus.CharInfo
            Select Case btn_index
            End Select
        Case e_Menus.Comerciar
            Select Case btn_index
            End Select
        Case e_Menus.ComerciarUsu
            Select Case btn_index
            End Select
        Case e_Menus.Commet
            Select Case btn_index
            End Select
        Case e_Menus.EligeAlineacion
            Select Case btn_index
            End Select
        Case e_Menus.Entrenador
            Select Case btn_index
            End Select
        Case e_Menus.Foro
            Select Case btn_index
            End Select
        Case e_Menus.GuildAdm
            Select Case btn_index
            End Select
        Case e_Menus.GuildBrief
            Select Case btn_index
            End Select
        Case e_Menus.GuildDetails
            Select Case btn_index
            End Select
        Case e_Menus.GuildFoundation
            Select Case btn_index
            End Select
        Case e_Menus.GuildLeader
            Select Case btn_index
            End Select
        Case e_Menus.GuildNews
            Select Case btn_index
            End Select
        Case e_Menus.GuildSol
            Select Case btn_index
            End Select
        Case e_Menus.GuildURL
            Select Case btn_index
            End Select
        Case e_Menus.Herrero
            Select Case btn_index
            End Select
        Case e_Menus.KeyPad
            Select Case btn_index
            End Select
        Case e_Menus.Mensaje
            Select Case btn_index
            End Select
        Case e_Menus.SOS
            Select Case btn_index
            End Select
        Case e_Menus.Opciones
            Select Case btn_index
            End Select
        Case e_Menus.PanelGM
            Select Case btn_index
            End Select
        Case e_Menus.Passwd
            Select Case btn_index
            End Select
        Case e_Menus.PeaceProp
            Select Case btn_index
            End Select
        Case e_Menus.Premios
            Select Case btn_index
            End Select
        Case e_Menus.Pres
            Select Case btn_index
            End Select
        Case e_Menus.Retos
            Select Case btn_index
            End Select
        Case e_Menus.ScreenShots
            Select Case btn_index
            End Select
        Case e_Menus.Skills
            Select Case btn_index
            End Select
        Case e_Menus.SpawnList
            Select Case btn_index
            End Select
        Case e_Menus.Subasta
            Select Case btn_index
            End Select
        Case e_Menus.tip
            Select Case btn_index
            End Select
        Case e_Menus.Torneo
            Select Case btn_index
            End Select
        Case e_Menus.UserRequest
            Select Case btn_index
            End Select
    End Select
    
End Function

Private Function Menu_Set_Focus(ByVal menu_index As Byte)
    Dim aux_menu As menu
    Dim i As Byte
    
    aux_menu = menu_array(menu_index)
    
    For i = menu_index To Z_last_menu - 1
        menu_array(i) = menu_array(i + 1)
    Next i
    'Llevamos el menu al principio.
    menu_array(Z_last_menu) = aux_menu
    
    'Seteamos al primer textbox como activo.
    If menu_array(Z_last_menu).num_textboxs > -1 Then
        menu_array(Z_last_menu).active_textbox = 0
    End If
End Function
Public Function GUI_Is_Visible(ByVal id As e_Menus) As Boolean
    If Menu_Obtain_Index(id) Then
        GUI_Is_Visible = True
    End If
End Function
Public Sub GUI_Active_Textbox_Set(ByVal menu_id As Byte, ByVal textbox_index As Integer)
    Dim menu_index As Byte
    menu_index = Menu_Obtain_Index(menu_id)
    If Not Textbox_Check(menu_index, textbox_index) Then Exit Sub
    
    menu_array(menu_index).active_textbox = textbox_index
    
End Sub
Public Function GUI_Active_Textbox_Change() As Integer
    If menu_array(Z_last_menu).num_textboxs = -1 Then Exit Function
    
    If menu_array(Z_last_menu).active_textbox = menu_array(Z_last_menu).num_textboxs Then
        menu_array(Z_last_menu).active_textbox = 0
    Else
        menu_array(Z_last_menu).active_textbox = menu_array(Z_last_menu).active_textbox + 1
    End If
    GUI_Active_Textbox_Change = menu_array(Z_last_menu).active_textbox
End Function

Public Function GUI_Textbox_Get(ByVal menu_id As e_Menus, ByVal textbox_index As Byte) As String
    If menu_id = 0 Or textbox_index = -1 Then Exit Function
    Dim index As Integer
    index = Menu_Obtain_Index(menu_id)
    
    GUI_Textbox_Get = menu_array(index).textbox(textbox_index).Text
End Function

Public Function GUI_Checkbox_Get(ByVal menu_id As e_Menus, ByVal checkbox_index As Byte) As Byte
    If menu_id = 0 Or checkbox_index = 0 Then Exit Function
    Dim index As Integer
    index = Menu_Obtain_Index(menu_id)
    
    GUI_Checkbox_Get = menu_array(index).checkbox(checkbox_index).value
End Function

Public Function GUI_Label_Set(ByVal menu_id As e_Menus, ByVal label_index As Byte, ByVal value As String)
    If menu_id = 0 Or label_index = 0 Then Exit Function
    Dim index As Integer
    index = Menu_Obtain_Index(menu_id)
    
    menu_array(index).label(label_index).caption = value
End Function
Public Function Window_Get() As Byte
    Window_Get = menu_array(Z_last_menu).id
End Function
Public Sub GUI_Key_Press(ByVal key_code As Integer)
    If menu_array(Z_last_menu).active_textbox = -1 Then Exit Sub
    'We have an active textbox in the focused form?
    Call Textbox_Key_Press(key_code, Z_last_menu)

End Sub
Private Function Textbox_Key_Press(ByVal key_code As Integer, ByVal menu_index As Byte) As Boolean
    If menu_array(menu_index).active_textbox <> -1 And menu_array(menu_index).num_textboxs <> -1 Then
        Select Case key_code
            Case vbKeyTab
                Call GUI_Active_Textbox_Change
                Textbox_Key_Press = True
            Case vbKeyBack
                Call GUI_Textbox_Remove_Char(Z_last_menu, menu_array(menu_index).active_textbox)
                Textbox_Key_Press = True
            Case vbKeyRight
                Call GUI_Textbox_Move(Z_last_menu, menu_array(menu_index).active_textbox, 2)
                Textbox_Key_Press = True
            Case vbKeyLeft
                Call GUI_Textbox_Move(Z_last_menu, menu_array(menu_index).active_textbox, 1)
                Textbox_Key_Press = True
            Case Else
                If key_code >= vbKeySpace And key_code <= 126 Then
                    Call GUI_Textbox_Add_Char(menu_index, menu_array(menu_index).active_textbox, Chr(key_code))
                    Debug.Print menu_array(menu_index).textbox(menu_array(menu_index).active_textbox).Text
                    Textbox_Key_Press = True
                End If
        End Select
    End If
End Function
Private Sub GUI_Textbox_Add_Char(ByVal menu_index As Integer, ByVal textbox_index As Integer, ByVal Char As String)
    If textbox_index = -1 Then Exit Sub
    If Len(menu_array(menu_index).textbox(textbox_index).Text) >= menu_array(menu_index).textbox(textbox_index).max_chars Then Exit Sub
    With menu_array(menu_index).textbox(textbox_index)
        .Text = Left(.Text, .sel_text_pos) & Char & Right(.Text, Len(.Text) - .sel_text_pos)
        .sel_text_pos = menu_array(menu_index).textbox(textbox_index).sel_text_pos + 1
    End With
End Sub
Private Function GUI_Textbox_Move(ByVal menu_index As Integer, ByVal textbox_index As Integer, ByVal direction As Byte)
    With menu_array(menu_index)
        Select Case direction
            Case 1
                If .textbox(textbox_index).sel_text_pos = 0 Then Exit Function
                .textbox(textbox_index).sel_text_pos = .textbox(textbox_index).sel_text_pos - 1
            Case 2
                If .textbox(textbox_index).sel_text_pos = Len(.textbox(textbox_index).Text) Then Exit Function
                .textbox(textbox_index).sel_text_pos = .textbox(textbox_index).sel_text_pos + 1
        End Select
    End With
End Function
Private Function GUI_Textbox_Remove_Char(menu_index As Integer, textbox_index As Integer)
    With menu_array(menu_index).textbox(textbox_index)
        If .sel_text_pos = 0 Then Exit Function
        If Len(.Text) = 0 Then Exit Function
        .Text = Left$(.Text, .sel_text_pos - 1) & Right$(.Text, Len(.Text) - .sel_text_pos)
        .sel_text_pos = .sel_text_pos - 1
    End With
End Function

Public Sub GUI_Mouse_Move(ByVal X As Integer, ByVal Y As Integer)
    mouse_x = X
    mouse_y = Y
End Sub

Private Function Button_Check(ByVal menu_index As Byte, ByVal btn_index As Integer) As Boolean
    With menu_array(menu_index)
        If btn_index > .num_buttons Or btn_index < 0 Then
            Button_Check = False
        Else
            Button_Check = True
        End If
    End With
End Function

Private Function Textbox_Check(ByVal menu_index As Byte, ByVal txt_index As Integer) As Boolean
    With menu_array(menu_index)
        If txt_index > .num_textboxs Or txt_index < 0 Then
            Textbox_Check = False
        Else
            Textbox_Check = True
        End If
    End With
End Function
Private Function Listbox_Check(ByVal menu_index As Byte, ByVal lst_index As Integer) As Boolean
    With menu_array(menu_index)
        If lst_index > .num_listboxs Or lst_index < 0 Then
            Listbox_Check = False
        Else
            Listbox_Check = True
        End If
    End With
End Function
Private Function Combobox_Check(ByVal menu_index As Byte, ByVal combobox_index As Integer) As Boolean
    With menu_array(menu_index)
        If combobox_index > .num_comboboxes Or combobox_index < 0 Then
            Combobox_Check = False
        Else
            Combobox_Check = True
        End If
    End With
End Function
Private Function Checkbox_Check(ByVal menu_index As Byte, ByVal chk_index As Integer) As Boolean
    With menu_array(menu_index)
        If chk_index > .num_checkboxes Or chk_index < 0 Then
            Checkbox_Check = False
        Else
            Checkbox_Check = True
        End If
    End With
End Function
Private Function Scrollbar_Check(ByVal menu_index As Byte, ByVal scrollbar_index As Integer) As Boolean
    With menu_array(menu_index)
        If scrollbar_index > .num_scrollbars Or scrollbar_index < 0 Then
            Scrollbar_Check = False
        Else
            Scrollbar_Check = True
        End If
    End With
End Function
Private Function Recttxt_Check(ByVal menu_index As Byte, ByVal recttxt_index As Integer) As Boolean
    With menu_array(menu_index)
        If recttxt_index > .num_recttxts Or recttxt_index < 0 Then
            Recttxt_Check = False
        Else
            Recttxt_Check = True
        End If
    End With
End Function
Private Function Menu_Test(ByVal X As Integer, ByVal Y As Integer) As Byte
    If Z_last_menu < 1 Then Exit Function
    Dim i As Byte

    i = Z_last_menu
    
    Do While i > 0
        If X > menu_array(i).X And X < menu_array(i).X + menu_array(i).width And Y > menu_array(i).Y And Y < menu_array(i).Y + menu_array(i).height Then
            Menu_Test = i
            Exit Function
        Else
            i = i - 1
        End If
    Loop
    
    Menu_Test = 0
End Function

Public Sub Set_Visible_Control(ByVal menu_id As Byte, control_index As Byte, ByVal control_type As e_control_type, ByVal value As Boolean)
    Dim menu_index As Byte
    menu_index = Menu_Obtain_Index(menu_id)
    
    If menu_index = 0 Then Exit Sub
    
    With menu_array(menu_index)
        Select Case control_type
            Case e_control_type.mButton
                .Button(control_index).visible = value
            Case e_control_type.mCheckbox
                .checkbox(control_index).visible = value
            Case e_control_type.mImage
                .image(control_index).visible = value
            Case e_control_type.mInventory
                .inventory(control_index).visible = value
            Case e_control_type.mLabel
                .label(control_index).visible = value
            Case e_control_type.mListbox
                .listbox(control_index).visible = value
            Case e_control_type.mScrollbar
                .scrollbar(control_index).visible = value
            Case e_control_type.mTextbox
                .textbox(control_index).visible = value
        End Select
    End With
End Sub

Private Sub Listbox_Add_Element(ByVal menu_id As e_Menus, ByVal lst_index As Integer, ByVal Element As String)
    Dim menu_index As Integer
    menu_index = Menu_Obtain_Index(menu_id)
    
    If menu_array(menu_index).listbox(lst_index).elem_count = menu_array(menu_index).listbox(lst_index).max_elements Then Exit Sub
    
    
    With menu_array(menu_index).listbox(lst_index)
        .elem_count = .elem_count + 1
        ReDim Preserve .List(0 To .elem_count) As String
        .List(.elem_count) = Element
        .list_index = .elem_count
        If .elem_count > 9 Then
            .scrollbar.max_value = .elem_count - 9
            .scrollbar.pos = .scrollbar.max_value 'Vamos al ultimo
            .scrollbar.width = .btn_width
            .scrollbar.X = .X + menu_array(menu_index).X + .width - .btn_width
            .scrollbar.Y = .Y + menu_array(menu_index).Y + .height
            .scrollbar.height = 10 * .height
        End If
    End With
End Sub

Private Sub Listbox_Remove_Element(ByVal menu_id As e_Menus, ByVal lst_index As Integer, ByVal index As Integer)
    If elem_count = -1 Then Exit Sub
    
    elem_count = elem_count - 1
    With menu_array(Menu_Obtain_Index(menu_id)).listbox(lst_index)
        If index <> UBound(.List) Then
            For i = index To elem_count
                .List(i) = .List(i + 1)
            Next i
        End If
        ReDim Preserve .List(0 To elem_count) As String
    End With
End Sub
Private Sub Scrollbar_Move(ByRef scrollbar As gui_scrollbar, ByVal ly As Integer)
    If ly < 16 And ly > -1 Then
        If scrollbar.pos > 0 Then
            scrollbar.pos = scrollbar.pos - 1
        End If
    ElseIf ly > scrollbar.height - 16 And ly < scrollbar.height Then
        If scrollbar.pos < scrollbar.max_value Then
            scrollbar.pos = scrollbar.pos + 1
        End If
    End If
End Sub
Public Sub Combobox_Add_Element(ByVal menu_index As e_Menus, ByVal combobox_index, ByVal Element As String)
    If menu_index < 1 Then Exit Sub
    If Not Combobox_Check(menu_index, combobox_index) Then Exit Sub
    With menu_array(menu_index).combobox(combobox_index)
        If .element_count > .max_elements Then Exit Sub
        .element_count = .element_count + 1
        ReDim Preserve .List(0 To .element_count)
        .List(.element_count) = Element
    End With
End Sub
Public Sub Combobox_Remove_Element(ByVal menu_id As e_Menus, ByVal combobox_index, ByVal Element As Integer)
    Dim menu_index As Byte
    Dim i As Integer
    menu_index = Menu_Obtain_Index(menu_id)
    If menu_index < 1 Then Exit Sub
    If Not Combobox_Check(menu_index, combobox_index) Then Exit Sub
    
    With menu_array(menu_index).combobox(combobox_index)
        If Element < 0 Or Element > .element_count Then Exit Sub
        If Element = .element_count Then
            .element_count = .element_count - 1
            ReDim Preserve .List(0 To .element_count)
        Else
            For i = Element To .element_count - 1
                .List(i) = .List(i + 1)
                .element_count = .element_count - 1
                ReDim Preserve .List(0 To .element_count)
            Next i
        End If
        If .list_index = Element Then .list_index = -1
    End With
End Sub
Public Sub Combobox_Clear(ByVal menu_id As e_Menus, ByVal combobox_index)
    Dim menu_index As Byte
    menu_index = Menu_Obtain_Index(menu_id)
    If menu_index < 1 Then Exit Sub
    If Not Combobox_Check(menu_index, combobox_index) Then Exit Sub
    With menu_array(menu_index).combobox(combobox_index)
        ReDim .List(0)
        .element_count = -1
        .list_index = -1
    End With
End Sub
Public Sub Recttxt_Add_Text(ByVal menu_id As e_Menus, ByVal recttxt_index As Integer, ByVal Text As String, ByVal Font_Index As Byte, ByVal color As Long)
    Dim i As Integer
    Dim chars_per_line As Integer
    Dim menu_index As Byte
    Dim j As Integer
    Dim last_space As Integer

    menu_index = Menu_Obtain_Index(menu_id)
    If menu_index < 1 Then Exit Sub
    If Not Recttxt_Check(menu_index, recttxt_index) Then Exit Sub
    With menu_array(menu_index).recttxt(recttxt_index)
        If .line_count > 18 Then
            For i = 0 To 9
                .lines(i) = .lines(i + 10)
            Next i
            .line_count = 9
            ReDim Preserve .lines(0 To .line_count)
        End If
        'Caracteres por linea
        chars_per_line = .width / 8
    
        If Len(Text) > chars_per_line Then
            i = 1
            Do While i <= Len(Text)
                For j = 1 To chars_per_line
                    If Mid(Mid(Text, i, chars_per_line), j, 1) = " " Then
                        last_space = j
                    End If
                Next j
                If last_space > 1 And Len(Text) - i > chars_per_line Then
                    .line_count = .line_count + 1
                    ReDim Preserve .lines(0 To .line_count)
                    .lines(.line_count).Text = Mid(Text, i, last_space)
                    .lines(.line_count).color = color
                    i = i + last_space
                Else
                    .line_count = .line_count + 1
                    ReDim Preserve .lines(0 To .line_count)
                    If last_space = 1 Then
                        .lines(.line_count).Text = Mid(Text, i + 1, chars_per_line)
                    Else
                        .lines(.line_count).Text = Mid(Text, i, chars_per_line)
                    End If
                    .lines(.line_count).color = color 'color
                    i = i + chars_per_line
                End If
            Loop
        Else
            .line_count = .line_count + 1
            ReDim Preserve .lines(0 To .line_count)
            .lines(.line_count).color = color
            .lines(.line_count).Text = Text
        End If
        .scrollbar.max_value = .line_count - .height \ 10
        
        If .scrollbar.max_value < 0 Then .scrollbar.max_value = 0
        
        If .scrollbar.max_value > 0 Then
            .scrollbar.visible = True
            .scrollbar.pos = .scrollbar.max_value
        Else
            .scrollbar.visible = False
        End If
    End With
End Sub

Public Sub Recttxt_Clear(ByVal menu_id As Byte, ByVal recttxt_index As Integer)
    Dim menu_index As Byte
    
    menu_index = Menu_Obtain_Index(menu_id)
    If menu_index = 0 Then Exit Sub
    If Not Recttxt_Check(menu_index, recttxt_index) Then Exit Sub
    
    With menu_array(menu_index).recttxt(recttxt_index)
        .line_count = -1
        .scrollbar.pos = 0
        .scrollbar.max_value = 0
        .scrollbar.visible = False
    End With
End Sub
Private Function GUI_Render_RecTxt(ByVal menu_index As Integer, ByVal rectxt_index As Integer)
    Dim i As Integer
    Dim last_line As Integer
    With menu_array(menu_index)
        last_line = .recttxt(rectxt_index).height \ 10
        If last_line > .recttxt(rectxt_index).line_count Then
            last_line = .recttxt(rectxt_index).line_count
        End If
        If .recttxt(rectxt_index).scrollbar.visible Then Call GUI_Render_Scrollbar(.X + .recttxt(rectxt_index).X + .recttxt(rectxt_index).width - 16, .Y + .recttxt(rectxt_index).Y, .recttxt(rectxt_index).scrollbar.height, 16, .recttxt(rectxt_index).scrollbar.max_value, .recttxt(rectxt_index).scrollbar.pos)
        For i = 0 To last_line
            Call DXEngine_TextRender(1, .recttxt(rectxt_index).lines(i + .recttxt(rectxt_index).scrollbar.pos).Text, .X + .recttxt(rectxt_index).X, .Y + .recttxt(rectxt_index).Y + i * 10, .recttxt(rectxt_index).lines(i + .recttxt(rectxt_index).scrollbar.pos).color)
        Next i
    End With
End Function
Private Function GUI_Render_Button(ByVal X As Integer, Y As Integer, ByVal filenum As Integer, ByVal caption As String)

End Function
Private Function GUI_Render_Checkbox()

End Function
Private Function GUI_Render_Combobox(ByVal menu_index As Integer, ByVal combobox_index As Integer)
    Dim last_index As Integer
    Dim i As Integer
    With menu_array(menu_index).combobox(combobox_index)
        If .element_count > .scrollbar.pos + .height / 16 Then
            last_index = .scrollbar.pos + .height / 16
        Else
            last_index = .element_count
        End If
        If .list_index > -1 Then
            'Call Engine.GUI_Box_Filled_Render(menu_array(menu_index).x + .x, menu_array(menu_index).y + .y + .list_index * 16, .width, 16, RGB(200, 200, 200))
        End If
        For i = 0 To last_index
            Call DXEngine_GraphicTextRender(1, .List(i + .scrollbar.pos), menu_array(menu_index).Y + .Y + 16 * i, menu_array(menu_index).X + .X, RGB(255, 255, 255))
        Next i
        
    End With
End Function
Private Function GUI_Render_Scrollbar(ByVal X As Integer, ByVal Y As Integer, ByVal height As Integer, ByVal width As Integer, ByVal max_value As Integer, ByVal pos As Integer)
    'Alto de la barra
    Dim bar_height As Integer

    'Call GUI_Box_Filled_Render(x, y, 16, height, RGB(0, 0, 0))
    Call GUI_Render_Image(16001, X, Y, 16, 16)
    Call GUI_Render_Image(16002, X, Y + height - 17, 15, 16)
    bar_height = (height - 32) / (max_value + 1)
    'Si la barra es menor a 10 pixels entonces la dibujamos de 10.
    If bar_height < 10 Then
        bar_height = 10
        Call GUI_Render_Image(16003, X, Y + 15 + (height - 30) / (max_value + 1) * pos, 16, bar_height)
    Else
        Call GUI_Render_Image(16003, X, Y + 15 + bar_height * pos, 16, bar_height)
    End If
    
End Function
Private Function GUI_Render_Listbox(ByVal menu_index As Byte, ByVal listbox_index As Integer) ' List() As String, ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal btn_width As Integer, ByVal font_index As Integer, ByVal list_index As Integer, ByVal elem_offset As Integer, Optional ByVal full_height As Integer = 0)
    Dim i As Integer
    Dim max_chars As Integer
    Dim last_element As Integer
    
    With menu_array(menu_index)
    'Lo calculo una sola vez para que sea mas rapido.
    max_chars = width / 16
        If .listbox(listbox_index).full_height And .listbox(listbox_index).full Then
            If .listbox(listbox_index).element_count > 9 Then
                last_element = 9
                'Call Engine.GUI_Box_Filled_Render(x - 1, y + height - 1, width + 2, 10 * height + 2, RGB(255, 255, 255), , , , False)
                'Call Engine.GUI_Box_Filled_Render(x, y + height, width, 10 * height, RGB(0, 0, 0), , , , False)
                Call GUI_Render_Scrollbar(X + width - 17, Y + height, 10 * height, 16, .listbox(listbox_index).scrollbar.max_value, .listbox(listbox_index).scrollbar.pos)
            Else
                last_element = .listbox(listbox_index).element_count
                'Call Engine.GUI_Box_Filled_Render(x - 1, y + height - 1, width + 2, (.listbox(listbox_index).element_count) * height + 2, RGB(255, 255, 255), , , , False)
                'Call Engine.GUI_Box_Filled_Render(x, y + height, width, (.listbox(listbox_index).element_count) * height, RGB(0, 0, 0), , , , False)
            End If
            
            If list_index > -1 Then
                Call DXEngine_GraphicTextRender(1, List(list_index), Y, X, RGB(255, 255, 255))
            End If
    
            For i = 0 To last_element
                'Call GUI_Box_Filled_Render(x, y + (i + 2) * height - 1, width, height, RGB(0, 0, 0), , , , False)
                Call DXEngine_GraphicTextRender(1, .listbox(listbox_index).List(i + .listbox(listbox_index).scrollbar.pos), Y + height + i * 16, X, RGB(255, 255, 255))
            Next i
        End If
        
        'Call Engine.GUI_Box_Filled_Render(x - 1, y - 1, width + 2, height + 2, RGB(255, 255, 255), , , , False)
        'Call Engine.GUI_Box_Filled_Render(x, y, width, height, RGB(0, 0, 0), , , , False)
        Call GUI_Render_Image(16002, X + width - 17, Y, 16, 15) 'flechita
        If list_index > -1 Then
            Call DXEngine_GraphicTextRender(1, List(list_index), Y, X, RGB(255, 255, 255))
        End If
    End With
End Function

Private Sub GUI_Render_Textbox(ByVal menu_index As Byte, ByVal textbox_index As Integer)
    Dim i As Byte
    Dim aux As String
    With menu_array(menu_index)
        If .textbox(textbox_index).locked Then
            For i = 1 To Len(.textbox(textbox_index).Text)
                aux = aux & "*"
            Next i
            Call DXEngine_GraphicTextRender(1, aux, .Y + .textbox(textbox_index).Y, .X + .textbox(textbox_index).X, RGB(255, 255, 255))
        Else
            Call DXEngine_GraphicTextRender(1, .textbox(textbox_index).Text, .Y + .textbox(textbox_index).Y, .X + .textbox(textbox_index).X, RGB(255, 255, 255))
        End If
    End With
End Sub

Private Sub GUI_Render_Label(ByVal menu_index As Byte, ByVal label_index As Integer)
    With menu_array(menu_index)
        Call DXEngine_GraphicTextRender(1, .label(label_index).caption, .Y + .label(label_index).Y, .X + .label(label_index).X, RGB(255, 255, 255))
    End With
End Sub
Public Sub Gui_Menu_Remove_All()
    If Z_last_menu <= 1 Then Exit Sub
    
    Dim i As Byte
    
    For i = UBound(menu_array) To 2
        Menu_Remove (i)
    Next i
End Sub
Private Sub Menu_Remove(ByVal menu_index As Byte)
    Dim i As Byte
    If menu_index > 1 And menu_index <= Z_last_menu Then
        If Not menu_index = Z_last_menu Then
            For i = menu_index To Z_last_menu - 1
                menu_array(i) = menu_array(i + 1)
            Next i
        End If
        Z_last_menu = Z_last_menu - 1
        ReDim Preserve menu_array(1 To Z_last_menu) As menu
    End If
End Sub
Public Function Combobox_Get_Index(ByVal menu_id As e_Menus, ByVal combobox_index) As Integer
    Dim menu_index As Byte
    
    menu_index = Menu_Obtain_Index(menu_id)
    If menu_index = 0 Then Combobox_Get_Index = -1: Exit Function
    If Not Combobox_Check(menu_index, combobox_index) Then Combobox_Get_Index = -1: Exit Function
    
    Combobox_Get_Index = menu_array(menu_index).combobox(combobox_index).list_index
End Function

Public Sub Combobox_Move_Item(ByVal menu_id As e_Menus, ByVal combobox_index As Integer, ByVal direccion As Byte)
    Dim menu_index As Byte
    Dim aux As String
    
    menu_index = Menu_Obtain_Index(menu_id)
    If menu_index = 0 Then Exit Sub
    If Not Combobox_Check(menu_index, combobox_index) Then Exit Sub
    
    With menu_array(menu_index).combobox(combobox_index)
        Select Case direccion
            Case 1
                If .list_index = 0 Then Exit Sub
                aux = .List(.list_index)
                .List(.list_index) = .List(.list_index - 1)
                .List(.list_index - 1) = aux
                .list_index = .list_index - 1
            Case 0
                If .list_index = .element_count Then Exit Sub
                aux = .List(.list_index)
                .List(.list_index) = .List(.list_index + 1)
                .List(.list_index + 1) = aux
                .list_index = .list_index + 1
        End Select
    End With
End Sub
Public Function Combobox_Get_Visible(ByVal menu_id As e_Menus, ByVal combobox_index As Integer) As Byte
    Dim menu_index As Byte
    menu_index = Menu_Obtain_Index(menu_id)
    If Not Combobox_Check(menu_index, combobox_index) Then Exit Function
    
    Combobox_Get_Visible = menu_array(menu_index).combobox(combobox_index).visible
End Function
Public Sub Combobox_Set_Element(ByVal menu_id As e_Menus, ByVal combobox_index As Integer, ByVal Element As String, Optional ByVal slot As Integer = 0)
    Dim menu_index As Byte
    menu_index = Menu_Obtain_Index(menu_id)
    If menu_index = 0 Then Exit Sub
    If Not Combobox_Check(menu_index, combobox_index) Then Exit Sub
    
    With menu_array(menu_index).combobox(combobox_index)
        If slot <> 0 And slot <= .element_count Then
            .List(slot - 1) = Element
        Else
            Call Combobox_Add_Element(menu_id, combobox_index, Element)
        End If
    End With
End Sub
Public Function Combobox_Get_Element(ByVal menu_id As e_Menus, ByVal combobox_index As Integer, ByVal list_index As Integer) As String
    Dim menu_index As Byte
    menu_index = Menu_Obtain_Index(menu_id)
    If menu_index = 0 Then Exit Function
    If Not Combobox_Check(menu_index, combobox_index) Then Exit Function
    
    With menu_array(menu_index).combobox(combobox_index)
        If list_index = -1 Or list_index > .element_count Then Exit Function
        Combobox_Get_Element = .List(list_index)
    End With
End Function
Private Sub GUI_Render_Image(ByVal texture_index As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal width As Integer, ByVal height As Integer)
    Dim rgb_list(0 To 3) As Long
    rgb_list(0) = RGB(255, 255, 255)
    rgb_list(1) = RGB(255, 255, 255)
    rgb_list(2) = RGB(255, 255, 255)
    rgb_list(3) = RGB(255, 255, 255)
    Call DXEngine_TextureRenderAdvance(texture_index, X, Y, 0, 0, width, height, width, height, rgb_list())
End Sub

Public Function Textbox_Get_Visible(ByVal menu_id As Byte, ByVal textbox_index As Integer) As Byte
    Dim menu_index As Byte
    menu_index = Menu_Obtain_Index(menu_id)
    If menu_index = 0 Then Exit Function
    If Not Textbox_Check(menu_index, textbox_index) Then Exit Function
    
    Textbox_Get_Visible = menu_array(menu_index).textbox(textbox_index).visible
End Function
Public Function Textbox_Get_Text(ByVal menu_id As Byte, ByVal textbox_index As Integer) As String
    Dim menu_index As Byte
    menu_index = Menu_Obtain_Index(menu_id)
    If menu_index = 0 Then Exit Function
    If Not Textbox_Check(menu_index, textbox_index) Then Exit Function
    Textbox_Get_Text = menu_array(menu_index).textbox(textbox_index).Text
End Function
Public Function Textbox_Clear(ByVal menu_id As Byte, ByVal textbox_index As Integer)
    Dim menu_index As Byte
    menu_index = Menu_Obtain_Index(menu_id)
    If menu_index = 0 Then Exit Function
    If Not Textbox_Check(menu_index, textbox_index) Then Exit Function
    menu_array(menu_index).textbox(textbox_index).Text = ""
    menu_array(menu_index).textbox(textbox_index).sel_text_pos = 0
End Function
