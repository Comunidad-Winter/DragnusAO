Attribute VB_Name = "modTextures"
Private mFreeMemoryBytes As Long
Dim MaxEntries As Integer

'Structure to hold all the graphics for the game
Private Type DX_Texture
    file_name As Integer
    width As Integer
    height As Integer
    d3dtexture As Direct3DTexture8
    UltimoAcceso As Long
    Size As Long
    borrable As Byte
End Type

'Keys
Dim lKeys() As Long
'Array of surfaces
Dim texture_list() As DX_Texture
Dim texture_count As Long

'Private DX objets
Private Dev As Direct3DDevice8
Private Loader As D3DX8

Private resource_path As String
Private Const PATH_GRAPHICS = "\graphics"

Private Const COLOR_KEY As Long = &HFF000000

'END DECLARATIONS

Private Function BorraMenosUsado() As Integer
    Dim Valor As Long
    Dim i As Long
    
    'Inicializamos todo
    Valor = texture_list(1).UltimoAcceso
    BorraMenosUsado = 1
    
    'Buscamos cual es el que lleva más tiempo sin ser utilizado
    For i = 1 To texture_count
        If texture_list(i).UltimoAcceso < Valor And texture_list(i).borrable Then
            Valor = texture_list(i).UltimoAcceso
            BorraMenosUsado = i
        End If
    Next i
    
    'Disminuimos el contador
    texture_count = texture_count - 1
    
    'Quitamos la clave
    Texture_Remove (BorraMenosUsado)
    
End Function
Public Function GetTexture(ByVal file_name As Integer) As Direct3DTexture8
    On Error GoTo ErrorHandler
    
    If lKeys(file_name) <> 0 Then
        With texture_list(lKeys(file_name))
            'Ultimo acceso
            .UltimoAcceso = GetTickCount
            'Devuelvo una texture con el grafico cargado
            Set GetTexture = .d3dtexture
        End With
    Else    'Gráfico no cargado
        'Vemos si puedo agregar uno a la lista
        If MaxEntries = texture_count Then
            'Sacamos el que hace más que no usamos, y utilizamos el slot
            lKeys(file_name) = Load_Texture(file_name, BorraMenosUsado())
            Set GetTexture = texture_list(index).d3dtexture
        Else
            lKeys(file_name) = Load_Texture(file_name)
            Set GetTexture = texture_list(lKeys(file_name)).d3dtexture
        End If
    End If
    
ErrorHandler:
End Function
Private Function ObtenerIndice(ByVal filename As Integer) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'Busqueda binaria para hallar la texture deseada
'**************************************************************
    Dim max As Integer  'Max index
    Dim min As Integer  'Min index
    Dim mid As Integer  'Middle index
    
    min = 1
    max = texture_count
    
    Do While min <= max
        mid = (min + max) / 2
        If filename < texture_list(mid).file_name Then
            'El índice no existe
            If max = mid Then
                max = max - 1
            Else
                max = mid
            End If
        ElseIf filename > texture_list(mid).file_name Then
            'El índice no existe
            If min = mid Then
                min = min + 1
            Else
                min = mid
            End If
        Else
            ObtenerIndice = mid
            Exit Function
        End If
    Loop
End Function

Private Sub OrdenarGraficos(ByVal primero As Integer, ByVal ultimo As Integer)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'Ordenamos los gráficos por nombre usando QuickSort
'**************************************************************
    Dim min As Integer      'Primer item de la lista
    Dim max As Integer      'Ultimo item de la lista
    Dim Comp As Integer     'Item usado para comparar
    Dim temp As DX_Texture
    
    min = primero
    max = ultimo
    
    Comp = texture_list((min + max) / 2).file_name
    
    Do While min <= max
        Do While texture_list(min).file_name < Comp And min < ultimo
            min = min + 1
        Loop
        Do While texture_list(max).file_name > Comp And max > primero
            max = max - 1
        Loop
        If min <= max Then
            temp = texture_list(min)
            texture_list(min) = texture_list(max)
            texture_list(max) = temp
            min = min + 1
            max = max - 1
        End If
    Loop
    If primero < max Then OrdenarGraficos primero, max
    If min < ultimo Then OrdenarGraficos min, ultimo
End Sub

Public Function Load_Texture(ByVal file_name As Integer, Optional ByVal texture_index As Integer = -1, Optional ByVal borrable As Byte = 1) As Integer
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
'On Error GoTo ErrHandler
    Dim texture_info As D3DXIMAGE_INFO
      
    If texture_index = -1 Then
        'Agregamos al final de la lista
        texture_index = texture_count + 1
        ReDim Preserve texture_list(1 To texture_index)
    End If
    
    'Call GetTextureHeader(GrhPath & Archivo & ".bmp", BMPInfo)  'para alto y ancho de la texture
    
    With texture_list(texture_index)
        'Nombre
        .file_name = file_name
        'Ultimo acceso
        .UltimoAcceso = GetTickCount
        'Cargamos el gráfico y seteamos la Color Key
        Texture_Load_From_File .d3dtexture, file_name & ".bmp", texture_info
        .width = texture_info.width
        .height = texture_info.height
        .borrable = borrable
    End With
    
    'Aumentamos la cantidad de gráficos
    texture_count = texture_count + 1
    
    'Devolvemos el ìndice en que lo cargamos
    Load_Texture = texture_index
Exit Function

errhandler:

End Function

Public Function Texture_Remove_All()
    Dim i As Long
    
    If texture_count <= 0 Then Exit Function
    
    For i = 1 To UBound(texture_list)
        Set texture_list(i).d3dtexture = Nothing
        texture_list(i).borrable = 1
        texture_list(i).width = 0
        texture_list(i).height = 0
        texture_list(i).file_name = 0
        texture_list(i).borrable = 1
        texture_list(i).UltimoAcceso = 0
        texture_list(i).Size = 0
    Next i
    texture_count = 0
    
End Function
Public Function Texture_Initialize(ByVal max_entries As Long, ByVal t_path As String, Device As Direct3DDevice8)
    Set Dev = Device
    Set Loader = New D3DX8
    
    ReDim lKeys(1 To 16003) As Long
    resource_path = t_path
    MaxEntries = max_entries
End Function

Public Sub Texture_Load(ByVal texture_index As Long, Optional ByVal borrable As Byte = 1)
    If texture_index = 0 Then Exit Sub
    'Call Load_Texture(texture_index, , borrable)
End Sub

Public Function Texture_Dimension_Get(ByVal file_name As Integer, texture_width As Integer, texture_height As Integer) As Integer
    If file_name > 0 Then
        texture_width = texture_list(lKeys(file_name)).width
        texture_height = texture_list(lKeys(file_name)).height
    End If
End Function

'TODO: No estoy muy seguro de que esto deba quedar aca (Blizzard)
Private Sub Texture_Load_From_File(Texture As Direct3DTexture8, ByVal file As String, ByRef texture_info As D3DXIMAGE_INFO)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*************************************************************+
On Error GoTo errhandler

    file = resource_path & PATH_GRAPHICS & "\" & file
    
    Set Texture = Loader.CreateTextureFromFileEx(Dev, file, D3DX_DEFAULT, _
    D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
    D3DX_FILTER_POINT, D3DX_FILTER_POINT, COLOR_KEY, texture_info, ByVal 0)
    
errhandler:
    'Ak solamente si no esta el archivo.
    prgRun = False
End Sub
Private Sub Texture_Remove(ByVal texture_index As Integer)
    'Quitamos la clave
    lKeys(texture_list(texture_index).file_name) = 0
    'Borramos la textura
    Set texture_list(i).d3dtexture = Nothing
    texture_list(i).borrable = 1
    texture_list(i).width = 0
    texture_list(i).height = 0
    texture_list(i).file_name = 0
    texture_list(i).borrable = 0
    texture_list(i).UltimoAcceso = 0
End Sub


