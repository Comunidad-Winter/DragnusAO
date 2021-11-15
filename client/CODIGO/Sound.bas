Attribute VB_Name = "modSound"
Option Explicit

Private Type Sound
    SoundID As Integer
    Buffer As DirectSoundSecondaryBuffer8
    Looping As Byte
End Type

Public UseSfx As Byte
Public UseMusic As Byte

'***********SOUND************
Private Const NumSfx As Integer = 30

Public DS As DirectSound8
Public DSBDesc As DSBUFFERDESC

Private DSBuffer() As Sound

'********** MUSIC ***********
Public Const Music_MaxVolume As Long = 100

Dim CurrentMusicFile As String
'DirectMusic's Performance object
Dim Performance As DirectMusicPerformance8
'Currently loaded segment
Dim Segment As DirectMusicSegment8
'The one and only DirectMusic Loader
Dim Loader As DirectMusicLoader8
'State of the currently loaded segment
Dim SegState As DirectMusicSegmentState8

Public Function Sound_Init() As Byte
'************************************************************
'Initialize the 3D sound device
'More info: http://www.vbgore.com/GameClient.Sound.Sound_Init
'************************************************************
On Error GoTo ErrOut
    UseSfx = 0
    UseMusic = 0
    
    Sound_Init = 1
    
    'Create the DirectSound device (with the default device)
    Set DS = dx.DirectSoundCreate("")
    DS.SetCooperativeLevel frmMain.hwnd, DSSCL_PRIORITY
    
    'Set up the buffer description for later use
    'We are only using panning and volume - combined, we will use this to create a custom 3D effect
    DSBDesc.lFlags = DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
    
    'Resize the souSound_Setnd buffer array
    ReDim DSBuffer(1 To NumSfx)
    
    Dim mus_Params As DMUS_AUDIOPARAMS
    Set Loader = dx.DirectMusicLoaderCreate()
   
    Set Performance = dx.DirectMusicPerformanceCreate()
    Call Performance.InitAudio(frmMain.hwnd, DMUS_AUDIOF_ALL, mus_Params, Nothing, DMUS_APATH_DYNAMIC_STEREO, 128)
    Call Performance.SetMasterAutoDownload(True)        'Enable auto download of instruments
   
    Performance.SetMasterTempo 1
    Performance.SetMasterVolume 1
Exit Function
ErrOut:
    'Failure loading sounds, so we won't use them
    UseSfx = 0
    UseMusic = 0
    Sound_Init = 0
End Function
Public Function Sound_Play(ByVal SoundID As Integer, Optional ByVal flags As CONST_DSBPLAYFLAGS = DSBPLAY_DEFAULT)
'************************************************************
'Used for non area-specific sound effects, such as weather
'More info: http://www.vbgore.com/GameClient.Sound.Sound_Play
'************************************************************
    Dim buffer_index As Integer

    'Make sure we are using sound
    If UseSfx = 0 Then Exit Function
    If SoundID = 0 Then Exit Function
    
    buffer_index = BufferIndex_Get(SoundID, flags = DSBPLAY_LOOPING)
    
    Sound_Set DSBuffer(buffer_index).Buffer, SoundID
    DSBuffer(buffer_index).SoundID = SoundID
    DSBuffer(buffer_index).Looping = flags = DSBPLAY_LOOPING
    
    'Confirm the buffer exists
    If Not DSBuffer(buffer_index).Buffer Is Nothing Then
        
        'Reset the sounds values (in case they were ever changed)
        DSBuffer(buffer_index).Buffer.SetCurrentPosition 0
        Sound_Pan DSBuffer(buffer_index).Buffer, 0
        Sound_Volume DSBuffer(buffer_index).Buffer, 0
        
        'Play the sound
        DSBuffer(buffer_index).Buffer.Play flags
    End If
   
End Function

Public Sub Sound_Stop(Optional ByVal buffer_index As Integer = 0)
'************************************************************
'Erase the sound buffer
'More info: http://www.vbgore.com/GameClient.Sound.Sound_Erase
'************************************************************
    
    'Make sure the object exists
    If buffer_index > 0 And buffer_index <= NumSfx Then
        If Not DSBuffer(buffer_index).Buffer Is Nothing Then
        
            'If it is playing, we have to stop it first
            If DSBuffer(buffer_index).Buffer.GetStatus > 0 Then DSBuffer(buffer_index).Buffer.Stop
            
            'Clear the object
            Set DSBuffer(buffer_index).Buffer = Nothing
            DSBuffer(buffer_index).SoundID = 0
            DSBuffer(buffer_index).Looping = 0
        End If
    Else
        Dim i As Byte
        For i = 1 To NumSfx
            If Not DSBuffer(i).Buffer Is Nothing Then
                If DSBuffer(i).Buffer.GetStatus <> 0 And DSBuffer(i).Buffer.GetStatus <> DSBSTATUS_BUFFERLOST Then
                    Call DSBuffer(i).Buffer.Stop
                End If
            End If
        Next i
    End If
End Sub



Public Sub Sound_Set(ByRef SoundBuffer As DirectSoundSecondaryBuffer8, ByVal SoundID As Integer)
'************************************************************
'Set the SoundID to the sound buffer
'More info: http://www.vbgore.com/GameClient.Sound.Sound_Set
'************************************************************

    If UseSfx = 0 Then Exit Sub

    'Clean the buffer
    Set SoundBuffer = Nothing
    
    'Set the buffer
    If General_File_Exists(resource_path & SfxPath & "\" & SoundID & ".wav", vbNormal) Then
        Set SoundBuffer = DS.CreateSoundBufferFromFile(resource_path & SfxPath & "\" & SoundID & ".wav", DSBDesc)
    End If

End Sub

Public Sub Sound_Play3D(ByVal SoundID As Integer, ByVal TileX As Integer, ByVal TileY As Integer)
'************************************************************
'Play a pseudo-3D sound by the sound buffer ID
'More info: http://www.vbgore.com/GameClient.Sound.Sound_Play3D
'************************************************************
    Dim SX As Long
    Dim SY As Long

    Dim buffer_index As Integer
    
    If UseSfx = 0 Then Exit Sub
    'Check for a valid sound
    If SoundID <= 0 Then Exit Sub
    
    buffer_index = BufferIndex_Get(SoundID, 0)
    
    Sound_Set DSBuffer(buffer_index).Buffer, SoundID
    DSBuffer(buffer_index).SoundID = SoundID
    DSBuffer(buffer_index).Looping = 0
    
    'Clear the position (used in case the sound was already playing - we can only have one of each sound play at a time)
    DSBuffer(buffer_index).Buffer.SetCurrentPosition 0
    
    'Set the user's position to sX/sY
    Engine.Engine_View_Pos_Get SX, SY
    
    'Calculate the panning
    Sound_Pan DSBuffer(buffer_index).Buffer, Sound_CalcPan(SX, TileX)
    
    'Calculate the volume
    Sound_Volume DSBuffer(buffer_index).Buffer, Sound_CalcVolume(SX, SY, TileX, TileY)
    
    'Play the sound
    DSBuffer(buffer_index).Buffer.Play DSBPLAY_DEFAULT
    
End Sub

Public Function Sound_CalcPan(ByVal x1 As Integer, ByVal x2 As Integer) As Long
'************************************************************
'Calculate the panning for 3D sound based on the user's position and the sound's position
'More info: http://www.vbgore.com/GameClient.Sound.Sound_CalcPan
'************************************************************

    If UseSfx = 0 Then Exit Function

    Sound_CalcPan = (x1 - x2) * 75 '* ReverseSound
    
End Function

Public Function Sound_CalcVolume(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal x2 As Integer, ByVal Y2 As Integer) As Long
'************************************************************
'Calculate the volume for 3D sound based on the user's position and the sound's position
'the (Abs(sX - TileX) * 25) is put on the end to make up for the simulated
' volume loss during panning (since one speaker gets muted to create the panning)
'More info: http://www.vbgore.com/GameClient.Sound.Sound_CalcVolume
'************************************************************
Dim Dist As Single

    If UseSfx = 0 Then Exit Function

    'Store the distance
    Dist = Sqr(((Y1 - Y2) * (Y1 - Y2)) + ((x1 - x2) * (x1 - x2)))
    
    'Apply the initial value
    Sound_CalcVolume = 200 * (1 - Dist / 20) '-(Dist * 80) + (Abs(x1 - x2) * 25)
    
    'Once we get out of the screen (>= 13 tiles away) then we want to fade fast
    If Dist > 15 Then Sound_CalcVolume = Sound_CalcVolume - ((Dist - 13) * 120)
    
End Function

Private Sub Sound_Pan(ByRef SoundBuffer As DirectSoundSecondaryBuffer8, ByVal value As Long)
'************************************************************
'Pan the selected SoundID (-10,000 to 10,000)
'More info: http://www.vbgore.com/GameClient.Sound.Sound_Pan
'************************************************************

    If UseSfx = 0 Then Exit Sub

    If SoundBuffer Is Nothing Then Exit Sub
    SoundBuffer.SetPan value

End Sub

Private Sub Sound_Volume(ByRef SoundBuffer As DirectSoundSecondaryBuffer8, ByVal value As Long)
'************************************************************
'Pan the selected SoundID (-10,000 to 0)
'More info: http://www.vbgore.com/GameClient.Sound.Sound_Volume
'************************************************************

    If UseSfx = 0 Then Exit Sub

    If SoundBuffer Is Nothing Then Exit Sub
    If value > 0 Then value = 0
    If value < -10000 Then value = -10000
    SoundBuffer.SetVolume value

End Sub

Public Function Music_Load(ByVal file As String) As Boolean
'************************************************************
'Loads a mp3 by the specified path
'More info: http://www.vbgore.com/GameClient.Sound.Music_Load
'************************************************************

    If UseMusic = 0 Then Exit Function

    On Error GoTo errhandler
    If Not General_File_Exists(resource_path & MusicPath & "\" & file, vbArchive) Then Exit Function
    
    Call Music_Stop
    
    'Destroy old object
    Set Segment = Nothing
    
    Set Segment = Loader.LoadSegment(resource_path & MusicPath & "\" & file)
    
    If Segment Is Nothing Then GoTo errhandler
    
    Call Segment.SetStandardMidiFile
    
    Music_Load = True
    
Exit Function
errhandler:
    Music_Load = False
End Function

Public Sub Music_Play(Optional ByVal file As String = "", Optional ByVal Loops As Long = -1)
'************************************************************
'Plays the mp3 in the specified buffer
'More info: http://www.vbgore.com/GameClient.Sound.Music_Play
'************************************************************
On Error GoTo errhandler
    If LenB(file) > 0 Then _
        CurrentMusicFile = file
    
    If UseMusic = 0 Then Exit Sub
    
    If PlayingMusic Then Music_Stop
    
    If LenB(file) > 0 Then
        If Not Music_Load(file) Then Exit Sub
    Else
        'Make sure we have a loaded segment
        If Segment Is Nothing Then Exit Sub
    End If
    
    'Play it
    Call Segment.SetRepeats(Loops)
    
    Set SegState = Performance.PlaySegmentEx(Segment, 0, 0)
Exit Sub
errhandler:
End Sub

Public Sub Music_Stop()
'************************************************************
'Stops the mp3 in the specified buffer
'More info: http://www.vbgore.com/GameClient.Sound.Music_Stop
'************************************************************

On Error GoTo errhandler

    If PlayingMusic Then
        Call Performance.StopEx(Segment, 0, DMUS_SEGF_DEFAULT)
    End If
    
Exit Sub
errhandler:
End Sub

Private Function BufferIndex_Get(ByVal SoundID As Integer, ByVal Looping As Byte) As Integer
    Dim i As Byte
    
    i = 0
    Do While i < NumSfx
        i = i + 1
        If DSBuffer(i).Buffer Is Nothing Then  'If its nothing then we load it.
            BufferIndex_Get = i
            Exit Function
        End If
    Loop
    
    i = 0
    Do While i < NumSfx
        i = i + 1
        If Not DSBuffer(i).Buffer Is Nothing Then  'If its nothing then we load it.
            If DSBuffer(i).Buffer.GetStatus = DSBSTATUS_BUFFERLOST Or Not DSBuffer(i).Buffer.GetStatus = DSBSTATUS_PLAYING Then
                BufferIndex_Get = i
                Exit Function
            End If
        End If
    Loop
    
    i = 0
    Do While i < NumSfx
        i = i + 1
        If DSBuffer(i).SoundID = SoundID Then
            BufferIndex_Get = i
            Exit Function
        End If
    Loop
    
    If Looping Then 'Si no es importante lo saltamos.
        i = 0
        Do While i < NumSfx
            i = i + 1
            If DSBuffer(i).Looping = 0 Then
                BufferIndex_Get = i
                Exit Function
            End If
        Loop
    End If

End Function

Private Function PlayingMusic() As Boolean
    If Segment Is Nothing Then Exit Function
    PlayingMusic = Performance.IsPlaying(Segment, SegState)
End Function

Public Sub soundDeactivate()
    Dim i As Integer
    For i = 1 To NumSfx
        Sound_Stop i
    Next i
End Sub
