VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextureManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private mFreeMemoryBytes As Long
Dim MaxEntries As Integer

Private Enum eExt
    bmp
    png
End Enum

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

Private Type D3DXIMAGE_INFO_A
    width As Long
    height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type

'Keys
Dim lKeys() As Long
'Array of surfaces
Dim texture_list() As DX_Texture
Dim texture_count As Long

Private Const COLOR_KEY As Long = &HFF000000

'END DECLARATIONS

Private Function BorraMenosUsado() As Integer
    Dim valor As Long
    Dim i As Long
    
    'Inicializamos todo
    valor = texture_list(1).UltimoAcceso
    BorraMenosUsado = 1
    
    'Buscamos cual es el que lleva m�s tiempo sin ser utilizado
    For i = 1 To texture_count
        If texture_list(i).UltimoAcceso < valor And texture_list(i).borrable Then
            valor = texture_list(i).UltimoAcceso
            BorraMenosUsado = i
        End If
    Next i
    
    'Disminuimos el contador
    texture_count = texture_count - 1
    
    'Quitamos la clave
    Texture_Remove (BorraMenosUsado)
    
End Function
Public Function GetTexture(ByVal file_name As Integer, Optional ByVal ext As Byte = 0) As Direct3DTexture8
    On Error GoTo ErrorHandler
    If file_name = 0 Then Exit Function
    If lKeys(file_name) <> 0 Then
        With texture_list(lKeys(file_name))
            'Ultimo acceso
            .UltimoAcceso = GetTickCount
            'Devuelvo una texture con el grafico cargado
            Set GetTexture = .d3dtexture
        End With
    Else    'Gr�fico no cargado
        'Vemos si puedo agregar uno a la lista
        If MaxEntries = texture_count Then
            'Sacamos el que hace m�s que no usamos, y utilizamos el slot
            lKeys(file_name) = Load_Texture(file_name, BorraMenosUsado(), , ext)
            Set GetTexture = texture_list(index).d3dtexture
        Else
            lKeys(file_name) = Load_Texture(file_name, , , ext)
            Set GetTexture = texture_list(lKeys(file_name)).d3dtexture
        End If
    End If
Exit Function
ErrorHandler:
    Debug.Print "GetTexture Error " & Err.Description
End Function

Public Function Load_Texture(ByVal file_name As Integer, Optional ByVal texture_index As Integer = -1, Optional ByVal borrable As Byte = 1, Optional ByVal ext As Byte = 0) As Integer
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
'On Error GoTo ErrHandler
    Dim texture_info As D3DXIMAGE_INFO_A
      
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
        'Cargamos el gr�fico y seteamos la Color Key
        Texture_Load_From_File .d3dtexture, file_name & "." & extToString(ext), texture_info
        .width = texture_info.width
        .height = texture_info.height
        .borrable = borrable
    End With
    
    'Aumentamos la cantidad de gr�ficos
    texture_count = texture_count + 1
    
    'Devolvemos el �ndice en que lo cargamos
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
    
    For i = 1 To UBound(lKeys)
        lKeys(i) = 0
    Next i
    
    texture_count = 0
    
End Function
Public Function Texture_Initialize(ByVal max_entries As Long)
    
    ReDim lKeys(1 To 32000) As Long
    MaxEntries = max_entries
End Function

Public Sub Texture_Load(ByVal texture_index As Long, Optional ByVal borrable As Byte = 1)
    If texture_index = 0 Then Exit Sub
    Call Load_Texture(texture_index, , borrable)
End Sub

Public Function Texture_Dimension_Get(ByVal file_name As Integer, texture_width As Integer, texture_height As Integer) As Integer
    If file_name > 0 Then
        texture_width = texture_list(lKeys(file_name)).width
        texture_height = texture_list(lKeys(file_name)).height
    End If
End Function

'TODO: No estoy muy seguro de que esto deba quedar aca (Blizzard)
Private Sub Texture_Load_From_File(Texture As Direct3DTexture8, ByVal file As String, ByRef texture_info As D3DXIMAGE_INFO_A)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*************************************************************+
On Error GoTo errhandler

    file = resource_path & PATH_GRAPHICS & "\" & file
    
    Set Texture = d3dx.CreateTextureFromFileEx(ddevice, file, D3DX_DEFAULT, _
    D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, _
    D3DX_FILTER_POINT, D3DX_FILTER_POINT, COLOR_KEY, texture_info, ByVal 0)
    Exit Sub
errhandler:
    'Ak solamente si no esta el archivo.
   
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

Private Function extToString(ByVal ext As eExt) As String
    Select Case ext
        Case eExt.bmp
            extToString = "bmp"
        Case eExt.png
            extToString = "png"
    End Select
End Function
