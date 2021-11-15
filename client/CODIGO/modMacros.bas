Attribute VB_Name = "modDXInput"
Private Enum eMouseState
    up
    down
End Enum

Dim mouseState As eMouseState

Private dInput As DirectInput8
Private ddeviceMouse As DirectInputDevice8

Dim mouseLeft As Boolean
Dim mouseRigh As Boolean

Dim mouseLeftLast As Long
Dim mouseRightLast As Long

Dim mouseClick As Boolean
Dim mouseDoubleClick As Boolean

Dim mouseMove As Boolean

Dim mouseDown As Boolean

Dim mouseClickLast As Long

Dim mouseX As Integer
Dim mouseY As Integer

Dim hwnd As Long

'MouseInput
Private Type PointAPI
    x As Long
    y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Public Sub inputInit(ByVal formHwnd As Long)
    hwnd = formHwnd
    
    'Input object
    Set dInput = dx.DirectInputCreate
    
    Set ddeviceMouse = dInput.CreateDevice("guid_SysMouse")
    
    Call ddeviceMouse.SetCommonDataFormat(DIFORMAT_MOUSE)
    Call ddeviceMouse.SetCooperativeLevel(formHwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE)
    
    Dim diProp As DIPROPLONG
    diProp.lHow = DIPH_DEVICE
    diProp.lObj = 0
    diProp.lData = 1000
    
    Call ddeviceMouse.SetProperty("DIPROP_BUFFERSIZE", diProp)
    ddeviceMouse.Acquire
End Sub

Public Sub inputReset()
    If mouseDown Then
        mouseDown = False
        mouseState = eMouseState.up
    End If
    mouseClick = False
    mouseDoubleClick = False
End Sub

Public Sub inputDeInit()
    'ddevice_mouse.Unacquire
    Set ddeviceMouse = Nothing
    Set dInput = Nothing
End Sub

Public Sub inputPoll()
    'Get the mouse event buffer
    Dim devData(1 To 1000) As DIDEVICEOBJECTDATA
    Dim nEvents As Long
    nEvents = ddeviceMouse.GetDeviceData(devData, DIGDD_DEFAULT)
    
    'Check buffer for clicks
    Dim i As Long
    For i = 1 To nEvents
        Select Case devData(i).lOfs
            Case DIMOFS_BUTTON0
                mouseLeft = (devData(i).lData And &H80)
                
                'Released? Then click.
                If Not mouseLeft Then
                    mouseState = eMouseState.up
                    
                    mouseDown = False
                    
                    mouseClick = True
                    If (GetTickCount - mouseLeftLast) < 300 Then
                        If GetTickCount - mouseClickLast < 300 Then
                            mouseDoubleClick = True
                        Else
                            mouseClick = True
                            mouseClickLast = GetTickCount
                        End If
                    End If
                Else
                    mouseState = eMouseState.down
                    mouseLeftLast = GetTickCount
                End If
        End Select
    Next i
    
    If mouseState = eMouseState.down And Not mouseDown Then
        If GetTickCount - mouseLeftLast > 95 Then
            mouseDown = True
        End If
    End If
    
    Dim target As RECT
    GetWindowRect hwnd, target
    
    'Use a API to get the mouse cordinates.
    Dim tempPoint As PointAPI
    GetCursorPos tempPoint
    
    If tempPoint.x - target.left <> mouseX Or tempPoint.y - target.top <> mouseY Then
        mouseMove = True
    End If
    
    mouseX = tempPoint.x - target.left
    mouseY = tempPoint.y - target.top
    
End Sub

Public Function inputDoubleClick() As Boolean
    inputDoubleClick = mouseDoubleClick
End Function

Public Function inputClick() As Boolean
    inputClick = mouseClick
End Function

Public Sub inputMouseGet(x As Integer, y As Integer)
    x = mouseX
    y = mouseY
End Sub

Public Function inputMouseMove() As Boolean
    inputMouseMove = mouseMove
End Function

Public Function inputMouseDown() As Boolean
    inputMouseDown = mouseDown
End Function

