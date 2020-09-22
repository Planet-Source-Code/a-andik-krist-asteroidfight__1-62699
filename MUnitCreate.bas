Attribute VB_Name = "MUnitCreate"
Option Explicit

' this to add create more ship
Global Const UnitCount As Byte = 30 ' 31 ship can create

' Unit Properties for unit (ship)
Type UnitProperties
    Active          As Boolean      ' Active Ship
    FoF             As Byte         ' Friend or Foe, Use for: (Player value=0, Asteroid/Enemy value=1)
                                    ' If Player ship shot a homing missile then missile hit player
                                    ' FoF Player = Missile = 0 then not make damage, but if missile hit
                                    ' enemy ship/Asteroid, FoF Asteroid(1) <> Missile(0) then will
                                    ' make damage, condition same we create Enemy ship and shot missile (FoF=1)...
    Type            As Byte         ' Only for Direct3D Object
                                    ' Set Type, 0 = Player Fighter
                                    '           1 = Enemy Fighter
                                    '           2 = Asteroid
    HP              As Integer      ' HP for Player/Enemy Fighter and Asteroid
    Size            As Byte         ' Size Object
    x               As Single       ' Position x, y
    y               As Single       '
    Angle           As Single       ' Direction unit
    AngleTurn       As Single       ' Angle to spin body left or Right if unit turn
    Speed           As Single       ' Speed Unit
    Turn            As Single       ' For turn unit if Object close to destiny turn value will big
    '-------------------------------------------------------------
    WeaponTime      As Byte         ' Weapon Battery Charge for Enemy
                                    ' Enemy can fire only direct player ship
    '-------------------------------------------------------------
    ' Check front have object (Asteroid, Friend) or not, if have then
    ' make avoid
    ' Ide : Use check all
    '-------------------------------------------------------------
    AvoidCollision   As Boolean     '
    AvoidTime        As Byte        ' set time for avoid collision
    GetUnitCollision As Byte        ' Get number unit collision
                                    ' Check range if < 50..100 (Size), get ship number
    RangeCollision   As Single
    
End Type
Public Unit(UnitCount)      As UnitProperties

'----------------------------------------------------------------
Public CountTimeToCreate    As Byte

'----------------------------------------------------------------
' Store data for random create Fighter
Public Type RndUnitProperties
    x As Single
    y As Single
    Angle As Single
End Type
Public RndUnit As RndUnitProperties

Sub CreateObject(FoF As Byte, x As Single, y As Single, Angle As Single, Speed As Single, Turn As Single, TypeObject As Byte, Size As Byte, Optional AvoidCollision As Boolean = False, Optional HP As Integer)
    Dim i As Byte
    
    For i = 0 To UnitCount
        ' Check i=0 only use by player
        If i = 0 And TypeObject <> 0 Then GoTo ExitNotPlayer
        If Unit(i).Active = False Then
            Unit(i).Active = True
            Unit(i).FoF = FoF
            Unit(i).Type = TypeObject
            Unit(i).Size = Size
            If TypeObject = 2 Or TypeObject = 3 Then
                Unit(i).HP = Size * 5
            Else
                Unit(i).HP = HP
            End If
            Unit(i).x = x
            Unit(i).y = y
            Unit(i).Angle = Angle
            Unit(i).AngleTurn = 0
            Unit(i).Speed = Speed
            Unit(i).Turn = Turn
            Unit(i).WeaponTime = 30
            Unit(i).AvoidCollision = AvoidCollision
            Unit(i).AvoidTime = 0
            
            UnitFrame(i).AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, (Pi / 2)
            UnitFrame(i).AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, DegreeToRadian(-Unit(i).Angle)
            UnitFrame(i).AddRotation D3DRMCOMBINE_BEFORE, 0, 0, 1, -DegreeToRadian(90 + Unit(i).AngleTurn)  ' 90=Position ship up
            UnitFrame(i).AddScale D3DRMCOMBINE_AFTER, 1, 1, 1
            UnitFrame(i).SetPosition Nothing, Unit(i).x, Unit(i).y, 0
            ' Add Visual Direct 3D Object
            UnitFrame(i).AddVisual UnitObject(TypeObject)
            
            ' Count Enemy Fighter
            If TypeObject = 1 Then EnemyMaxCount = EnemyMaxCount + 1
            
            ' Count Enemy Fighter
            If TypeObject = 2 Or TypeObject = 3 Then
                If Size = 20 Then
                    BigAsteroidCount = BigAsteroidCount + 1
                Else
                    SmallAsteroidCount = SmallAsteroidCount + 1
                End If
            End If
            
            Exit Sub
        End If
ExitNotPlayer:
    Next i
End Sub

' Create Enemy /Asteroid
Sub CreateEnemy()
    Dim x        As Single
    Dim y        As Single
    Dim Angle    As Single
    
    CountTimeToCreate = CountTimeToCreate + 1
    If CountTimeToCreate > 100 Then
        ' set back CountTimeToCreate = 0
        CountTimeToCreate = 0
        
        ' Set Max enemy create
        If EnemyMaxCount < EnemyLimitCreate Then
            ' Set random position for create enemy
            PlaceRndUnit
            ' Get value from PlaceRndUnit
            x = RndUnit.x
            y = RndUnit.y
            Angle = RndUnit.Angle
            
            CreateObject 1, x, y, Angle, EnemySpeed, EnemyTurn, 1, 15, False, 100
        End If
        
        If BigAsteroidCount < BigAsteroidLimit Then
            ' Set random position for create enemy
            PlaceRndUnit
            ' Get value from PlaceRndUnit
            x = RndUnit.x
            y = RndUnit.y
            Angle = RndUnit.Angle
            
            CreateObject 2, x, y, Angle, 1, 0.5, 2, 20
        End If
    
    End If
End Sub

' Create Random for put enemy Fighter
Sub PlaceRndUnit()
    Dim SideCreateEnemy As Byte
    Dim x               As Single
    Dim y               As Single
    Dim Angle       As Single

    ' Apper enemy ship from side 1=Up, 2=Left ....(Clockwwise)
    SideCreateEnemy = Int(Rnd * 4 + 1)
    If SideCreateEnemy = 1 Then         ' from UP
        x = (Rnd * ((ScrWidth / 2) + 0))
        y = -10
        ' Make direction enemy Fighter always to player
        Angle = Trigonometri(Unit(0).x, Unit(0).y, x, y) + 180
    End If
    If SideCreateEnemy = 2 Then         ' from RIGHT
        x = ScrWidth - 10
        y = -Rnd * ((ScrHeight / 2))
        Angle = AddLessDegree(Trigonometri(Unit(0).x, Unit(0).y, x, y), 180)
    End If
    If SideCreateEnemy = 3 Then         ' from DOWN
        x = (Rnd * ((ScrWidth / 2) + 20))
        y = -ScrHeight + 10
        Angle = Trigonometri(Unit(0).x, Unit(0).y, x, y) - 180
    End If
    If SideCreateEnemy = 4 Then         ' from LEFT
        x = 10
        y = -Rnd * ((ScrHeight / 2))
        Angle = AddLessDegree(Trigonometri(Unit(0).x, Unit(0).y, x, y), 180)
    End If
        
    ' Store result
    RndUnit.x = x
    RndUnit.y = y
    RndUnit.Angle = Angle
    
End Sub
