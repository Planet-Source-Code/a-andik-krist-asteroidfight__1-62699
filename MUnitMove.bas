Attribute VB_Name = "MUnitMove"
Option Explicit

Sub MoveUnit()
    Dim i               As Byte
    Dim xTarget         As Single
    Dim yTarget         As Single
    Dim Speed           As Single
    Dim Turn            As Single
    Dim Angle           As Single
    Dim Zoom            As Single
    '--------------------------------------
    Dim xsudut          As Single
    Dim ysudut          As Single
    '--------------------------------------
    Dim DetectUnitFront As Byte

    For i = 0 To UnitCount
        If Unit(i).Active = True Then
            
            Speed = Unit(i).Speed
            Turn = Unit(i).Turn
                        
            ' Detect Object in Front (add radar angle 40), Collision Unit with Unit
            DetectUnitFront = Multi_DetObjfront_ColUnitWtUnit(i, 40 + (EnemyLevel * 5))
            
            ' Set enemy fighter chase Player / Asteroid too (because turn to small then not to see  if asteroid chase player)
            xTarget = Unit(0).x
            yTarget = Unit(0).y
            
            If i = 0 Then
                '[----------------------------------------------------
                '[ Routine for Player
                '[----------------------------------------------------
            Else
                If Unit(i).Type = 1 Then
                    '[----------------------------------------------------
                    '[ Routine for Enemy Fighter
                    '[----------------------------------------------------
                    
                    ' Make unit enemy avoid object in front (not for player and asteroid)
                    If DetectUnitFront <> 255 Or DetectUnitFront <> 0 Then
                        ' If Close avoid object in front
                        If Unit(i).RangeCollision < 200 Then
                            ' Detect turn Left or Right
                            If DirectTurn(Unit(i).Angle, Unit(DetectUnitFront).Angle, 5) >= 0 Then
                                xsudut = 200 * Sin(DegreeToRadian(Unit(i).Angle))
                                ysudut = 200 * Cos(DegreeToRadian(Unit(i).Angle)) * -1
                                xTarget = Unit(i).x + xsudut
                                yTarget = Unit(i).y + ysudut
                            Else
                                xsudut = -200 * Sin(DegreeToRadian(Unit(i).Angle))
                                ysudut = -200 * Cos(DegreeToRadian(Unit(i).Angle)) * -1
                                xTarget = Unit(i).x + xsudut
                                yTarget = Unit(i).y + ysudut
                            End If
                        End If
                    End If
                                    
                    ' Enemy Fire, only have Player in front, Angle radar Add 20
                    If DetectUnitFront = 0 Then
                        EnemyFire i
                    End If
                Else
                    '[----------------------------------------------------
                    '[ Routine for Asteroid
                    '[----------------------------------------------------
                    ' For Rotation Asteroid (Not Use Turn and Spin like Ship
                    Unit(i).AngleTurn = Unit(i).AngleTurn + 2
                    If Unit(i).AngleTurn > 360 Then Unit(i).AngleTurn = 0
                    
                    ' Zoom for Asteroid see by Size
                    Zoom = (Unit(i).Size / 15)
                End If
            End If
                            
         '[-------------------------------------------------]
         '[ ENGINEAAK : Calculation moving unit             ]
         '[-------------------------------------------------]
            Engine Unit(i).Angle, Unit(i).x, Unit(i).y, xTarget, yTarget, Speed, Turn
         '[-------------------------------------------------]
         '[ ENGINEAAK : Don't forget replace with new value ]
         '[-------------------------------------------------]
            Unit(i).x = EngineResult.x
            Unit(i).y = EngineResult.y
            Unit(i).Angle = EngineResult.Angle
         '[-------------------------------------------------]
            
            Angle = Unit(i).Angle
            If Unit(i).Type > 1 Then Angle = Unit(i).AngleTurn
            
            ' Asteroid
            If Unit(i).AvoidCollision = True Then
                Unit(i).AvoidTime = Unit(i).AvoidTime + 1
                If Unit(i).AvoidTime > 30 Then
                    Unit(i).AvoidCollision = False
                    Unit(i).AvoidTime = 0
                End If
            End If
                        
            ' Turn and Spin Fighter (Player and Enemy) not for Asteroid
            If Unit(i).Type < 2 Then
                ' If a ship turn left/right then body ship will make spin
                SpinBodyUnit i, xTarget, yTarget
                ' if ship spin after turn, then make body back to normal position
                SpinBodyUnitToNormal i
                
                ' Zoom for Fihgter
                Zoom = 1 * (ScrWidth / 800)
            End If
            
            SetPosRotObj UnitFrame, i, Angle, Unit(i).AngleTurn, Unit(i).x, Unit(i).y, 0, Zoom
            
            ' Kill small asteroid if to much, limit is 15
            If Unit(i).Type > 1 And Unit(i).Size <> 20 Then
                If SmallAsteroidCount > 15 Then
                    If Unit(i).x < 0 Or Unit(i).x > ScrWidth Or Unit(i).y > 0 Or Unit(i).y < -ScrHeight Then
                        KillUnit i, False, False
                    End If
                End If
            End If
            
            ' Player, Enemy and Asteroid if Out Screen
            If Unit(i).x < 0 Then Unit(i).x = ScrWidth
            If Unit(i).x > ScrWidth Then Unit(i).x = 0
            If Unit(i).y > 0 Then Unit(i).y = -ScrHeight
            If Unit(i).y < -ScrHeight Then Unit(i).y = 0
            
            ' Only Text, show info
            'BackBuffer.DrawText Unit(i).x - 50, -Unit(i).y + Unit(i).Size, "Angle:" & Unit(i).Angle, False
            'BackBuffer.DrawText Unit(i).x - 50, -Unit(i).y + 10, "ObjectFront:" & DetectUnitFront, False
            'BackBuffer.DrawText Unit(i).x - 50, -Unit(i).y + 25, "ObjectRange:" & Unit(i).RangeCollision, False
            
        End If
    Next i
End Sub

Private Sub EnemyFire(i As Byte)
    Unit(i).WeaponTime = Unit(i).WeaponTime + 1
    If Unit(i).WeaponTime > EnemyFireDelay Then
        Unit(i).WeaponTime = 0
        PlaySound MissileBuffer, True, False
        Weapon2Fire i, Unit(i).Angle, 5, 0, 5, 1, 10, 0
    End If
End Sub

' Mutli Function :
' Can use detect Object in front
' Can use collision unit with unit
Function Multi_DetObjfront_ColUnitWtUnit(i As Byte, Angle As Byte) As Byte
    Dim j        As Byte
    Dim ChkAngle As Single
    Dim Range    As Single
    Dim RangeMax As Single
    Dim HPUnit As Integer
    
    Unit(i).RangeCollision = 1000
    RangeMax = 600 ' make set for long range
    Multi_DetObjfront_ColUnitWtUnit = 255
    For j = 0 To UnitCount
        If Unit(j).Active = True And j <> i Then
            ' Check angle for Object in front
            ChkAngle = Trigonometri(Unit(i).x, Unit(i).y, Unit(j).x, Unit(j).y)
            
            ' Can use check for collision unit with unit
            Range = Trigonometri(Unit(i).x, Unit(i).y, Unit(j).x, Unit(j).y, RESULT_RADIUS)
            
            ' Check Object/Unit in Front
            If Unit(i).Angle + Angle > ChkAngle And Unit(i).Angle < ChkAngle + Angle Then
                If Range < RangeMax Then
                    RangeMax = Range
                    Unit(i).RangeCollision = RangeMax
                    Multi_DetObjfront_ColUnitWtUnit = j
                End If
            End If
            
            ' Check Collision
            If Range < Unit(i).Size Then
                If Unit(j).Type > 1 And Unit(i).Type > 1 Then
                    '[----------------------------------------------------
                    '[ Routine collision Asteroid with Asteroid, only
                    '[ direction change not explode
                    '[----------------------------------------------------
                        If Unit(j).Type > 1 Then
                            If Unit(j).AvoidCollision = False Then
                                Unit(j).Angle = AddLessDegree(Unit(j).Angle, 180)
                            End If
                            ' Use Unit(...).AvoidCollision=True, this mean not check collision for now (use time)
                            If Unit(j).AvoidCollision = False Then Unit(j).AvoidCollision = True
                        End If
                        
                        If Unit(i).Type > 1 Then
                            If Unit(i).AvoidCollision = False Then
                                Unit(i).Angle = AddLessDegree(Unit(i).Angle, 180)
                            End If
                            If Unit(i).AvoidCollision = False Then Unit(i).AvoidCollision = True
                        End If
                Else
                    HPUnit = Unit(j).HP
                    
                    Unit(j).HP = Unit(j).HP - Unit(i).HP
                    Unit(i).HP = Unit(i).HP - HPUnit
                    
                    If Unit(j).HP < 1 Then
                        KillUnit j
                    Else
                        If Unit(i).HP < 1 Then
                            KillUnit i
                        End If
                    End If
                    'Exit Function
                End If
            End If
        End If
    Next j
End Function

Sub KillUnit(i As Byte, Optional ByPlayer As Boolean = False, Optional SoundExplode As Boolean = True)
    Dim xsudut As Single
    Dim ysudut As Single
    
    ' Calculation make Explode mid Unit
    xsudut = -35 * Sin(DegreeToRadian(45))
    ysudut = -35 * Cos(DegreeToRadian(45)) * -1

    Unit(i).Active = False
    UnitFrame(i).DeleteVisual UnitObject(Unit(i).Type)
    ' Create Explode and sound
    If SoundExplode = True Then
        CreateExplode Unit(i).x + xsudut, Unit(i).y + ysudut
        CreateSndExp 1000
    End If
    
    ' Check for Decrease Count Enemy Fighter
    If Unit(i).Type = 1 Then
        EnemyMaxCount = EnemyMaxCount - 1
        If ByPlayer = True Then EnemyDestroyCount = EnemyDestroyCount + 1
    End If
   
    ' Check for Decrease Count Asteroid Big or Small
    If Unit(i).Type = 2 Or Unit(i).Type = 3 Then
        If Unit(i).Size = 20 Then
            BigAsteroidCount = BigAsteroidCount - 1
            CreateSmallAsteroid i
        Else
            SmallAsteroidCount = SmallAsteroidCount - 1
        End If
    End If
End Sub

Sub CreateSmallAsteroid(j As Byte)
    Dim SizeAsteroid As Byte
    Dim GetAngle     As Single
    
    ' must store Angle If use Unit(j).Angle direction not Right
    GetAngle = Unit(j).Angle
    
    SizeAsteroid = (Unit(j).Size - 10)
    
    CreateObject 2, Unit(j).x, Unit(j).y, AddLessDegree(GetAngle, 0), 2, 0, 2, SizeAsteroid, True
    CreateObject 2, Unit(j).x, Unit(j).y, AddLessDegree(GetAngle, 90), 2, 0, 2, SizeAsteroid, True
    CreateObject 2, Unit(j).x, Unit(j).y, AddLessDegree(GetAngle, 180), 2, 0, 2, SizeAsteroid, True
    CreateObject 2, Unit(j).x, Unit(j).y, AddLessDegree(GetAngle, -90), 2, 0, 2, SizeAsteroid, True
    
    Dim RndModule As Byte
    If EnemyLevel < 5 Then
        RndModule = Int(Rnd * (EnemyLevel + 3) + 1) ' 2 : 5, for more difficule set 2 : 9
    Else
        RndModule = Int(Rnd * 7 + 1)
    End If
    
    If RndModule = 1 Or RndModule = 2 Then
        ' Create Module
        ModuleCreate Unit(j).x, Unit(j).y, Unit(j).Angle, RndModule
    End If
End Sub

' Spin body if ship turn left or right
Sub SpinBodyUnit(j As Byte, xDest As Single, yDest As Single)
    Dim PosDegrees As Single
    Dim AddSpin    As Byte
    
    PosDegrees = Int(Trigonometri(Unit(j).x, Unit(j).y, xDest, yDest))
    '---------------------------------------------------------------
    ' If a ship turn left/right then body ship will make little spin
    '---------------------------------------------------------------
    AddSpin = 3     ' is 5 = Standart for Fighter
    
    If PosDegrees + AddSpin > Unit(j).Angle And PosDegrees - AddSpin < Unit(j).Angle Then
    Else
        ' to now spin ship Left/Right i am use -> DirectSpin in module * EngineAAK *
        If DirectTurn(PosDegrees, Unit(j).Angle, 1) < 0 Then     ' Value minus turn Left/Kiri
            Unit(j).AngleTurn = Unit(j).AngleTurn + Unit(j).Turn * AddSpin
            If Unit(j).AngleTurn > 70 Then Unit(j).AngleTurn = 70
        Else                                                    ' Value plus turn Right/Kanan
           Unit(j).AngleTurn = Unit(j).AngleTurn - Unit(j).Turn * AddSpin
           If Unit(j).AngleTurn < -70 Then Unit(j).AngleTurn = -70
        End If
    End If
End Sub

' After ship spin then make back to normal
Sub SpinBodyUnitToNormal(j As Byte)
    If Unit(j).AngleTurn <> 0 Then
        If Unit(j).AngleTurn < 0 Then
            Unit(j).AngleTurn = Unit(j).AngleTurn + 2
        Else
            Unit(j).AngleTurn = Unit(j).AngleTurn - 2
        End If
    Else
        Unit(j).AngleTurn = 0
    End If
End Sub

' Set Rotation/Spin, Zoom and Position Frame Object : Ship & Missile ..ect
Sub SetPosRotObj(FrameObj() As Direct3DRMFrame3, j As Byte, Angle As Single, AngleTurn As Single, x As Single, y As Single, z As Single, Zoom As Single)
    
    '---------------------------------------------------------------
    ' Set for Rotation/Spin, Zoom and Position Object
    '---------------------------------------------------------------
    FrameObj(j).AddRotation D3DRMCOMBINE_REPLACE, 0, 1, 0, (Pi / 2)
    FrameObj(j).AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, DegreeToRadian(-Angle)
    FrameObj(j).AddRotation D3DRMCOMBINE_BEFORE, 0, 0, 1, -DegreeToRadian(90 + AngleTurn)  ' 90=Position Object upper
    FrameObj(j).AddScale D3DRMCOMBINE_AFTER, Zoom, Zoom, Zoom
    FrameObj(j).SetPosition Nothing, x, y, z
    
End Sub

