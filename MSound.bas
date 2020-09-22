Attribute VB_Name = "MSound"
Option Explicit

' Sound Code
'----------------------------
Public FireBuffer           As DirectSoundBuffer
Public MissileBuffer        As DirectSoundBuffer

Public SndExplodeBuffer()   As DirectSoundBuffer ' Explode more 1
Public Const ExpSndCount    As Byte = 20
Type SndExpProperties
    Active      As Boolean
    TimeKill    As Byte     ' Calculation End Sound
    CurrFreq    As Long
End Type
Public SndExplode(ExpSndCount) As SndExpProperties
Public NumberSoundCount As Byte

Public DSound As DirectSound

Dim Ds As DirectSound

Dim DsBuffer As DirectSoundBuffer

Dim DsDesc As DSBUFFERDESC
Dim DsWave As WAVEFORMATEX

'Dim CurrFreq As Long

Sub SoundInit(ByRef NameForm As Form)
    'Label1.Caption = "Initialising DirectSound"
    Set Ds = DX.DirectSoundCreate("")
    
    ' It is best to check for errors before continuing
    If Err.Number <> 0 Then
        MsgBox "Unable to Continue, Error creating Directsound object."
        'Label1.Caption = "Error"
        Exit Sub
    End If
    
    Ds.SetCooperativeLevel NameForm.hWnd, DSSCL_NORMAL
    
    DsDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    DsWave.nFormatTag = WAVE_FORMAT_PCM 'Sound Must be PCM otherwise we get errors
    DsWave.nChannels = 2                '1= Mono, 2 = Stereo
    DsWave.lSamplesPerSec = 22050
    DsWave.nBitsPerSample = 16          '16 =16bit, 8=8bit
    DsWave.nBlockAlign = DsWave.nBitsPerSample / 8 * DsWave.nChannels
    DsWave.lAvgBytesPerSec = DsWave.lSamplesPerSec * DsWave.nBlockAlign
    
    Set FireBuffer = Ds.CreateSoundBufferFromFile(App.Path & "\Laser.wav", DsDesc, DsWave)
    Set MissileBuffer = Ds.CreateSoundBufferFromFile(App.Path & "\Missile.wav", DsDesc, DsWave)
    
    Dim i As Byte
    ReDim SndExplodeBuffer(ExpSndCount)
    For i = 0 To ExpSndCount
        Set SndExplodeBuffer(i) = Ds.CreateSoundBufferFromFile(App.Path & "\Explodes.wav", DsDesc, DsWave)
        SndExplode(i).CurrFreq = 0 'SndExplodeBuffer(i).GetFrequency
    Next i
    
End Sub

Sub CreateSndExp(PanLong As Long)
    Dim i As Byte
    For i = 0 To ExpSndCount
        If SndExplode(i).Active = False Then
            SndExplode(i).Active = True
            SndExplode(i).TimeKill = 0
            SndExplodeBuffer(i).SetPan PanLong * 10  ' Left Speaker
            PlaySound SndExplodeBuffer(i), False, False
            Exit Sub
        End If
    Next i
End Sub

' Kill Explode sound
Sub KillExpSnd()
    Dim i As Byte
    NumberSoundCount = 0
    For i = 0 To ExpSndCount
        If SndExplode(i).Active = True Then
            SndExplode(i).TimeKill = SndExplode(i).TimeKill + 1
            If SndExplode(i).TimeKill > 100 Then
                SndExplodeBuffer(i).Stop
                SndExplodeBuffer(i).SetCurrentPosition 0
                SndExplode(i).Active = False
            End If
        NumberSoundCount = NumberSoundCount + 1
        End If
    Next i
End Sub

Function CreateSound(filename As String) As DirectSoundBuffer
    DsDesc.lFlags = DSBCAPS_STATIC Or DSBCAPS_CTRLVOLUME
    
    Set CreateSound = Ds.CreateSoundBufferFromFile(filename, DsDesc, DsWave)
    If Err.Number <> 0 Then
        MsgBox "Unable to find sound file"
        MsgBox Err.Description
        End
    End If
End Function

Sub StopSound()
    DsBuffer.Stop
    DsBuffer.SetCurrentPosition 0
End Sub

Sub PlaySound(Sound As DirectSoundBuffer, CloseFirst As Boolean, LoopSound As Boolean)
    If CloseFirst Then
        Sound.Stop
        Sound.SetCurrentPosition 0
    End If
    If LoopSound Then
        Sound.Play 1
    Else
        Sound.Play 0
    End If
End Sub

