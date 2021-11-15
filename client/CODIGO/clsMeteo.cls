VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMeteo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This Object provides info about the Ambient.
'Handles the rain, day and night.

'Time info
Private mHour As Byte
Private mMinutes As Byte
Private mSeconds As Byte

Private RainParticleIndex As Integer
Private Rainning As Byte
'Light Info
Private MedioDia As D3DCOLORVALUE
Private MediaNoche As D3DCOLORVALUE
Private LuzActual As D3DCOLORVALUE

Public Sub Initialize()
    MedioDia.R = 255
    MedioDia.G = 255
    MedioDia.B = 255
    
    MediaNoche.R = 100
    MediaNoche.G = 100
    MediaNoche.B = 120
End Sub

Public Sub Set_Time(ByVal Hour As Byte, ByVal Minutes As Byte)
    mHour = Hour
    mMinutes = Minutes
    mSeconds = 0
    
    AmbientCalculate
End Sub
Private Sub AmbientCalculate()
    If mHour >= 12 Then
        Call D3DXColorLerp(LuzActual, MedioDia, MediaNoche, mHour / 23)
    ElseIf mHour < 12 Then
        Call D3DXColorLerp(LuzActual, MediaNoche, MedioDia, mHour / 12)
    End If
    
    'Render Lights
    'Call Engine.Map_LightsRender(view_buffer_min_x, view_buffer_max_x, view_buffer_min_y, view_buffer_max_y)
End Sub
Public Sub Get_AmbientLight(color As D3DCOLORVALUE)
    color = LuzActual
End Sub
Public Sub Meteo_Check()
    If MainTimer.Check(TimersIndex.Minute) Then
        If mMinutes >= 59 Then
            mMinutes = 0
            If mHour >= 23 Then
                mHour = 0
            Else
                mHour = mHour + 1
            End If
            AmbientCalculate
        Else
            mMinutes = mMinutes + 1
        End If
    End If
End Sub

Public Sub MeteoParticle_Render(ByVal DirectionY As Integer)
    If Rainning Then
        particleGroupRender RainParticleIndex, 0, 0, 0, -DirectionY * 1.9
    End If
End Sub

Public Sub Rain_Toggle()
    If Rainning Then
        Rain_Stop
    Else
        Rain_Start
    End If
End Sub
Private Sub Rain_Stop()
    Rainning = 0
    If RainParticleIndex > 0 Then
        particleGroupDestroy RainParticleIndex
    End If
    RainParticleIndex = 0
End Sub
Private Sub Rain_Start()
    If Rainning Then Exit Sub
    
    Rainning = 1
    
    RainParticleIndex = particleGroupCreate(0, 0, e_ParticleType.Rain, -1)
End Sub

Public Sub Get_AmbientColorRGB(rgb_list() As Long)
    rgb_list(0) = D3DColorARGB(LuzActual.A, LuzActual.R, LuzActual.G, LuzActual.B)
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub
