Attribute VB_Name = "MGames"
Option Explicit

' Get Mouse PointAPI
'Type PointAPI
'   x As Long
'   y As Long
'End Type
'Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
'Public MousePoint As PointAPI

' Used to hide or show the mouse cursor
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public GameFinish         As Boolean
Public GameOverTime       As Byte

Public ScrWidth           As Long     ' Set Width and Height Screen
Public ScrHeight          As Long

' For Player
Public PlayerPower        As Byte    ' HP
Public PlayerPowerDevide  As Single  ' To devide with 11 (picture for power bar have 11 picture)
Public PlayerWeapon       As Byte    ' Add Player Weapon
Public PlayerWeaponDelay  As Byte    ' More Weapon More Delay
Public PlayerTimeFire     As Byte

' For Enemy
'Public EnemyPower        As Integer ' Add for enemy if level up more difficully
Public EnemyLevel         As Byte
Public EnemySpeed         As Single
Public EnemyTurn          As Single
Public EnemyFireDelay     As Byte
Public EnemyDestroyCount  As Integer
Public EnemyLimitCreate   As Byte
Public EnemyMaxCount      As Byte     ' Count Enemy Create

Public BigAsteroidLimit   As Byte    ' Limit for Create BigAsteroid
Public BigAsteroidCount   As Byte    ' Limit for Create BigAsteroid
Public SmallAsteroidCount As Byte    ' Limit for Small Asteroid in Screen
                                     ' if to more kill if out of screen

Sub Init_Games()
    ' Create Player unit (FoF=0)
    CreateObject 0, (ScrWidth / 2), -(ScrHeight / 2), 90, 3, 0, 0, 15, False, 250
    PlayerPower = Unit(0).HP
    PlayerPowerDevide = (Unit(0).HP / 11)
    PlayerWeapon = 0            ' 0
    PlayerWeaponDelay = 5       ' 5
    
    ' Set Enemy
    EnemyDestroyCount = 0       ' 0
    EnemyLevel = 0              ' 0
    EnemySpeed = 2.5            ' 2.5  -> Max:5
    EnemyTurn = 1               ' 1    -> Max:2
    EnemyFireDelay = 50         ' 50
    EnemyLimitCreate = 1        ' 1
    
    BigAsteroidLimit = 2        ' Limit Create Asteroid will up if Level up
                                ' every 10 enemy ship destroy
    
    GameOverTime = 0            ' 0
End Sub

Sub GameLoops()
    Dim RectFighter As RECT
    
    ' Hide Cursor
    ShowCursor 0
    
    Do                                      ' Loop main until GameFinish=True
        On Local Error Resume Next
        DoEvents
        D3D_ViewPort.Clear D3DRMCLEAR_TARGET Or D3DRMCLEAR_ZBUFFER 'ClS Viewport.
        D3D_Device.Update                   ' Update the Direct3D Device.
       'D3D_ViewPort.Render FrameRoot       ' Render the 3D Objects, but if put in these
                                            ' after object Delete still have shadow object
        DelayGame 25    ' 25 = Set 39 FPS   '                |
                                            '                |
        DisplayMap                          '                |
                                            '                |
        MoveUnit                            ' Use 3D Object  |
                                            '                |
        MoveWeapon                          ' Use 3D Object  |
                                            '                |
        ModuleMove                          ' Use 3D Object  |
                                            '                |
                                            '                |
        CreateEnemy                         '                |
                                            '                |
        D3D_ViewPort.Render FrameRoot       '           <----+ Then i put after Declarations MoveUnit and MoveWeapon
                                            '                  because these Declarations render 3D Object
        DisplayHit
        
        DisplayExplode
        
        KillExpSnd
        
        If Unit(0).Active = False Then
            DrawGameOver (ScrWidth / 2) - 136, (ScrHeight / 2) - 26
            GameOverTime = GameOverTime + 1
            If GameOverTime > 150 Then
                Init_Games
            End If
        End If
        
        KeyBoard
               
        ' If Destroy every 5 Enemy Fighter, Level Up (set more difficuly Like Enemy speed and Turn)
        '---------------------------------------------------------
        If EnemyLevel <> Int(EnemyDestroyCount / 5) Then
            EnemyLevel = Int(EnemyDestroyCount / 5)
            
            ' Limited EnemySpeed, effect at EnemyTurn too
            '--------------------------------------------
            If EnemySpeed < 5 Then
                EnemySpeed = EnemySpeed + 0.5
                EnemyTurn = EnemyTurn + 0.2
            End If
            
            ' Fire Delay for Enemy
            '--------------------------------------------
            EnemyLimitCreate = EnemyLimitCreate + 1
            If EnemyFireDelay > 20 Then
                EnemyFireDelay = EnemyFireDelay - 5
            End If
            
            ' If Destroy every 10 Enemy Fighter add more Big Asteroid to Create
            '--------------------------------------------
            If InStr(1, (EnemyDestroyCount / 10), ".") = 0 Then
                BigAsteroidLimit = BigAsteroidLimit + 1
            End If
        End If
        
        ' Sorry Menu not Finish .......?
        'DrawMenu 360, 300
        
        DrawTextInfo
        
        DrawNumber 40, 39, PicBuffer, Trim(Str(EnemyDestroyCount))
        
        DrawBarPower 0, 0
        
        Primary.Flip Nothing, DDFLIP_WAIT   ' Flip the BackBuffer with the FrontBuffer.
    Loop Until GameFinish = True
    
    ' Show Cursor
    ShowCursor -1
    
    End
    
End Sub

Private Sub KeyBoard()
    Dim i As Byte
    Dim j As Byte
    
    ' Get the array of keyboard keys and their current states
    DI_Device.GetDeviceStateKeyboard DI_State
    
    ' ESC = Exit Program
    If DI_State.Key(DIK_ESCAPE) <> 0 Then
        Call Delete_All
        GameFinish = True
    End If
    
    ' Fire player ship
    If DI_State.Key(DIK_LCONTROL) <> 0 Or DI_State.Key(DIK_RCONTROL) <> 0 Then
        If Unit(0).Active = True Then
            PlayerTimeFire = PlayerTimeFire + 1
            If PlayerTimeFire > PlayerWeaponDelay Then
                PlayerTimeFire = 0
                ' Fire Red Laser
                For j = 0 To PlayerWeapon
                    Weapon2Fire 0, Unit(i).Angle, 5, 5 * j, 25, 0, 10, 0
                Next j
                ' Create sound Fire
                PlaySound FireBuffer, False, False
            End If
        End If
    End If
        
    ' Control Player ship
    i = 0   ' I alwasy use player ship value = 0 (if only 1 player)
    If Unit(i).Type = 0 Then
        If DI_State.Key(DIK_LEFT) <> 0 Then
            ' Turn Player ship
            Unit(i).Angle = Unit(i).Angle + 3
            ' make player ship spin if turn
            Unit(i).AngleTurn = Unit(i).AngleTurn - 5
            If Unit(i).AngleTurn < -70 Then Unit(i).AngleTurn = -70
        End If
        If DI_State.Key(DIK_RIGHT) <> 0 Then
            ' Turn Player ship
            Unit(i).Angle = Unit(i).Angle - 3
            ' make player ship spin if turn
            Unit(i).AngleTurn = Unit(i).AngleTurn + 5
            If Unit(i).AngleTurn > 70 Then Unit(i).AngleTurn = 70
        End If
        
        ' Test in 3D View
        'Dim ZoomView As Single
        'ZoomView = 1
        'FrameCamera.SetOrientation Nothing, (Pi / 2), 0, 1, 1, 0, 0
        'FrameCamera.SetPosition Nothing, Unit(i).x - (1000 / ZoomView), Unit(i).y, zCamera + 250 '400
        
        ' Make Camera following player ship, if not have background nott good
        'FrameCamera.SetPosition Nothing, Unit(i).x, Unit(i).y, zCamera '- 0
    End If
                    
End Sub

Private Sub DisplayMap()
    MapFrame.AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, (Pi / 2)
    MapFrame.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, DegreeToRadian(0)
    MapFrame.AddRotation D3DRMCOMBINE_BEFORE, 0, 0, 1, -DegreeToRadian(90)    ' 90=Position Object upper
    MapFrame.AddScale D3DRMCOMBINE_AFTER, 21, 16, 10
    MapFrame.SetPosition Nothing, 400, -300, 20
End Sub

Private Sub DrawTextInfo()
    Dim TxtIntro    As String
    Dim TxtMid      As Integer

    '--------------------------------------------------------------
    ' Text Introduction about "..EngineAAK.."
    '--------------------------------------------------------------
    TxtIntro = "ASTEROID FIGHTER"
    TxtMid = (ScrWidth / 2) - (Len(TxtIntro) * 4)
    BackBuffer.DrawText TxtMid, 0, TxtIntro, False
            
    '--------------------------------------------------------------
    ' Text Introduction about "..EngineAAK.."
    '--------------------------------------------------------------
    TxtIntro = "EngineAAK use for Arcade like ASTEROID Clone  (Esc = Exit)"
    TxtMid = (ScrWidth / 2) - (Len(TxtIntro) * 3.5)
    BackBuffer.DrawText TxtMid, 580, TxtIntro, False
End Sub

Private Sub DrawNumber(x As Integer, y As Integer, surface As DirectDrawSurface4, Txt As String)
    Dim i As Integer
    Dim ValText As Integer
    Dim RECTvar As RECT
    
    For i = 0 To Len(Txt) - 1
        ValText = Val(Mid(Txt, i + 1, 1))
        With RECTvar
            .Top = 0
            .Left = (ValText) * 14
            .Right = (ValText + 1) * 14
            .Bottom = 23
        End With
        BackBuffer.BltFast x + (i * 14), y, surface, RECTvar, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY  'DDBLTFAST_WAIT
    Next i
End Sub

Private Sub DrawGameOver(x As Integer, y As Integer)
    Dim RECTvar As RECT
    
    With RECTvar
        .Left = 169
        .Top = 0
        .Right = 439
        .Bottom = 57
    End With
    BackBuffer.BltFast x, y, PicBuffer, RECTvar, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY    'DDBLTFAST_WAIT
End Sub

' Power Bar Player
Private Sub DrawBarPower(x As Integer, y As Integer)
    Dim RECTvar As RECT
    Dim HP       As Integer
    
    With RECTvar
        .Left = 755
        .Top = 12
        .Right = 814
        .Bottom = 111
    End With
    BackBuffer.BltFast x, y, PicBuffer, RECTvar, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY   'DDBLTFAST_WAIT
    
    If Unit(0).HP > 0 Then
        HP = Int((PlayerPower / PlayerPowerDevide) - (Unit(0).HP / PlayerPowerDevide))
    Else
        HP = 12 'PowerBar Picture No.12 Blank 'Int(PlayerPower / PlayerPowerDevide)
    End If
    
    With RECTvar
        .Left = (HP) * 60
        .Top = 117
        .Right = (HP + 1) * 60
        .Bottom = 216
    End With
    BackBuffer.BltFast x, y, PicBuffer, RECTvar, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY   'DDBLTFAST_WAIT
End Sub

Sub DrawMenu(x As Integer, y As Integer)
    Dim RECTvar As RECT
    
    ' Play
    With RECTvar
        .Left = 443
        .Top = 1
        .Right = 531
        .Bottom = 24
    End With
    BackBuffer.BltFast x, y - 50, PicBuffer, RECTvar, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY  'DDBLTFAST_WAIT
    
    ' Info
    With RECTvar
        .Left = 533
        .Top = 1
        .Right = 615
        .Bottom = 24
    End With
    BackBuffer.BltFast x, y, PicBuffer, RECTvar, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY   'DDBLTFAST_WAIT
    
    ' Exit
    With RECTvar
        .Left = 618
        .Top = 1
        .Right = 697
        .Bottom = 24
    End With
    BackBuffer.BltFast x, y + 50, PicBuffer, RECTvar, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY  'DDBLTFAST_WAIT
    
End Sub

